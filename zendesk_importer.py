import base64
import logging
from typing import Callable, Dict, List, Optional, Tuple
from pathlib import Path

import requests
import pandas as pd


class ZendeskImporter:
    """
    Imports Help Center articles into a specific section.

    Multi-brand safe:
      • Resolves the brand host from a section’s html_url and POSTs there
      • Sends X-Zendesk-Brand-Id when the brand is known

    Supports:
      • Excel (.xlsx/.xls) and Word (.docx) inputs
      • Update-in-place by id / exact title / overlapping labels
      • Label merge or replace
      • Per-row section_id override (lets a single sheet target multiple sections)
      • Default locale for rows that lack a locale (e.g., from Word)
    """

    def __init__(self, subdomain: str, email: str, api_token: str):
        self._brands_cache: Optional[List[Dict]] = None
        self._host_to_brand: Optional[Dict[str, Dict]] = None
        self._section_cache: Dict[int, Dict] = {}
        self.subdomain = subdomain
        self.email = email
        self.api_token = api_token
        self.base_url = f"https://{subdomain}.zendesk.com"

        creds = base64.b64encode(f"{email}/token:{api_token}".encode()).decode()
        self.headers = {"Authorization": f"Basic {creds}", "Accept": "application/json"}

        self.session = requests.Session()
        self.session.headers.update(self.headers)

    # -----------------------------
    # HTTP helper
    # -----------------------------
    def safe_request(self, method: str, url: str, quiet_404: bool = False, **kwargs) -> requests.Response:
        try:
            r = self.session.request(method, url, timeout=90, **kwargs)
            if r.status_code == 404 and quiet_404:
                logging.info("API %s %s -> 404 (quiet)", method, url)
            elif r.status_code >= 400:
                logging.error("API %s %s failed (%s): %s", method, url, r.status_code, r.text)
            if r.status_code == 401:
                raise RuntimeError("Zendesk rejected the credentials (401). Check the email and API token.")
            if r.status_code == 403:
                raise RuntimeError("Zendesk refused the request (403). The user needs Guide agent or admin rights.")
            r.raise_for_status()
            return r
        except requests.RequestException as e:
            if quiet_404 and getattr(e, "response", None) is not None and e.response.status_code == 404:
                # treat as soft miss
                raise
            logging.error("API %s %s exception: %s", method, url, e)
            raise

    # -----------------------------
    # Multi-brand host resolution
    # -----------------------------
    def _load_brands_map(self) -> None:
        if self._brands_cache is not None and self._host_to_brand is not None:
            return
        url = f"{self.base_url}/api/v2/brands.json"
        r = self.safe_request('GET', url)
        brands = r.json().get('brands', [])
        self._brands_cache = brands

        hostmap: Dict[str, Dict] = {}
        for b in brands:
            sub = (b.get("subdomain") or "").strip()
            custom = (b.get("host_mapping") or "").strip().lower()
            if sub:
                hostmap[f"{sub}.zendesk.com"] = b
            if custom:
                hostmap[custom] = b
        self._host_to_brand = hostmap

    def _get_section(self, section_id: int) -> Dict:
        """
        Fetch section (cached). Try account host first; on 404, try each brand host.
        """
        sid = int(section_id)
        if sid in self._section_cache:
            return self._section_cache[sid]

        # Account host first
        url_acct = f"{self.base_url}/api/v2/help_center/sections/{sid}.json"
        try:
            r = self.safe_request('GET', url_acct, quiet_404=True)
            sec = r.json().get('section', {}) or {}
            self._section_cache[sid] = sec
            return sec
        except Exception:
            logging.info("Section %s not visible on account host; probing brand hosts…", sid)

        # Probe brand hosts
        self._load_brands_map()
        if not self._brands_cache:
            raise ValueError(f"Unable to resolve brands to locate section {sid}")

        candidates = []
        for b in self._brands_cache:
            custom = (b.get("host_mapping") or "").strip().lower()
            sub = (b.get("subdomain") or "").strip()
            if custom:
                candidates.append((custom, b))
            if sub:
                candidates.append((f"{sub}.zendesk.com", b))

        for host, b in candidates:
            brand_base = f"https://{host}"
            headers = dict(self.session.headers)
            if b.get("id"):
                headers["X-Zendesk-Brand-Id"] = str(b["id"])
            try:
                r = self.session.get(
                    f"{brand_base}/api/v2/help_center/sections/{sid}.json",
                    headers=headers, timeout=90
                )
                if r.status_code == 404:
                    continue
                r.raise_for_status()
                sec = (r.json() or {}).get("section", {}) or {}
                if sec:
                    self._section_cache[sid] = sec
                    return sec
            except requests.RequestException:
                continue

        raise ValueError(f"Section {sid} not found on any brand host")

    def _brand_base_for_section(self, section_id: int) -> Tuple[str, Dict, Optional[Dict]]:
        """
        Resolve base host + brand for a section.
        Returns (brand_base_url, section_obj, brand_obj_or_None).
        """
        from urllib.parse import urlparse as _urlparse

        self._load_brands_map()
        sec = self._get_section(section_id)
        html = (sec.get("html_url") or "").strip().lower()
        host = _urlparse(html).netloc if html else ""
        if host:
            brand_base = f"https://{host}"
            brand_obj = self._host_to_brand.get(host) if self._host_to_brand else None
            return brand_base, sec, brand_obj
        return self.base_url, sec, None

    # -----------------------------
    # Parsing helpers
    # -----------------------------
    def parse_file(self, file_path: str) -> List[Dict]:
        """Dispatch on extension."""
        ext = Path(file_path).suffix.lower()
        if ext in ('.xlsx', '.xls'):
            return self.parse_excel_file(file_path)
        if ext == '.docx':
            return self.parse_word_file(file_path)
        raise ValueError(f"Unsupported file format: {ext}")

    def parse_excel_file(self, file_path: str) -> List[Dict]:
        """Parse Excel file to list of article dicts."""
        df = pd.read_excel(file_path)
        if df.empty:
            return []

        column_mapping = {
            'id':          ['id', 'article_id', 'Article ID', 'ARTICLE_ID'],
            'title':       ['title', 'Title', 'TITLE', 'article_title', 'name'],
            'body':        ['body_text', 'body', 'content', 'Body', 'BODY', 'article_body', 'text'],
            'locale':      ['locale', 'Locale', 'LOCALE', 'language', 'lang'],
            'section_id':  ['section_id', 'Section ID', 'SECTION_ID', 'section'],
            'draft':       ['draft', 'Draft', 'DRAFT', 'is_draft'],
            'promoted':    ['promoted', 'Promoted', 'PROMOTED', 'is_promoted'],
            'labels':      ['label_names', 'labels', 'Labels', 'LABELS', 'tags'],
        }

        actual: Dict[str, str] = {}
        lower_cols = {str(c).strip().lower(): c for c in df.columns}
        for field, candidates in column_mapping.items():
            for cand in candidates:
                key = cand.strip().lower()
                if key in lower_cols:
                    actual[field] = lower_cols[key]
                    break

        if 'title' not in actual and 'id' not in actual:
            raise ValueError("Excel must contain at least 'title' or 'id'")

        articles: List[Dict] = []
        for _, row in df.iterrows():
            art: Dict = {}

            if 'id' in actual:
                v = row.get(actual['id'])
                if pd.notna(v) and str(v).strip():
                    try:
                        art['id'] = int(v)
                    except (ValueError, TypeError):
                        pass

            if 'title' in actual:
                v = row.get(actual['title'])
                art['title'] = ("" if pd.isna(v) else str(v)).strip()
            else:
                art['title'] = ""

            if 'body' in actual:
                v = row.get(actual['body'])
                art['body'] = ("" if pd.isna(v) else str(v)).strip()
            else:
                art['body'] = ""

            if 'locale' in actual:
                v = row.get(actual['locale'])
                loc = ("" if pd.isna(v) else str(v)).strip().lower()
                if loc:
                    art['locale'] = loc

            if 'draft' in actual:
                v = row.get(actual['draft'])
                if pd.notna(v):
                    art['draft'] = _to_bool(v)

            if 'promoted' in actual:
                v = row.get(actual['promoted'])
                if pd.notna(v):
                    art['promoted'] = _to_bool(v)

            if 'section_id' in actual:
                v = row.get(actual['section_id'])
                if pd.notna(v):
                    try:
                        art['section_id'] = int(v)
                    except (ValueError, TypeError):
                        pass

            if 'labels' in actual:
                v = row.get(actual['labels'])
                if pd.notna(v):
                    labels = [s.strip() for s in str(v).split(',') if s and s.strip()]
                    if labels:
                        art['label_names'] = labels

            # Skip if both missing
            if not art.get('id') and not art['title'] and not art['body']:
                continue

            # Skip create without a body
            if not art.get('id') and not art['body']:
                continue

            articles.append(art)

        return articles

    def parse_word_file(self, file_path: str) -> List[Dict]:
        """
        Parse Word (.docx) into a single article using Mammoth to preserve styling as HTML.
        Title = first non-empty paragraph or bold run; Body = full HTML from the document.
        """
        try:
            import mammoth  # preferred for HTML preservation
        except Exception:
            # Fallback: simple text extraction with python-docx
            logging.warning("mammoth not available; falling back to plain-text parsing for %s", file_path)
            from docx import Document as _Doc
            d = _Doc(file_path)
            title, body_parts = "", []
            for p in d.paragraphs:
                text = (p.text or "").strip()
                if not text:
                    continue
                if not title and (p.runs and any(r.bold for r in p.runs)):
                    title = text
                elif not title:
                    title = text
                else:
                    body_parts.append(text)
            if not title and not body_parts:
                raise ValueError("Could not extract content from the Word document")
            return [{
                "title": title or "Imported Article",
                "body": "".join(f"<p>{_esc(t)}</p>" for t in body_parts),
            }]

        # Mammoth path (preferred)
        try:
            # Word's Title style becomes a <title> tag so we can pick it out first.
            # The list lines let nested bullets our own exporter writes come back nested.
            style_map = "\n".join([
                "p[style-name='Title'] => title:fresh",
                "p[style-name='List Bullet 2'] => ul > li > ul > li:fresh",
                "p[style-name='List Bullet 3'] => ul > li > ul > li > ul > li:fresh",
                "p[style-name='List Number 2'] => ol > li > ol > li:fresh",
                "p[style-name='List Number 3'] => ol > li > ol > li > ol > li:fresh",
            ])
            with open(file_path, "rb") as f:
                result = mammoth.convert_to_html(f, style_map=style_map)  # .value (html), .messages
            html = (result.value or "").strip()

            # Heuristic title: Title style, else first heading, else first <p><strong>, else first <p>.
            # Whichever wins is removed from the body so the title is not repeated.
            title = ""
            import re
            for pat in (r"<title[^>]*>(.+?)</title>",
                        r"<h[1-3][^>]*>(.+?)</h[1-3]>",
                        r"<p[^>]*>\s*<strong>(.+?)</strong>\s*</p>",
                        r"<p[^>]*>(.+?)</p>"):
                m = re.search(pat, html, re.IGNORECASE | re.DOTALL)
                if m:
                    import html as _html
                    cand = _html.unescape(re.sub(r"<.*?>", "", m.group(1))).strip()
                    if cand:
                        title = cand
                        html = (html[:m.start()] + html[m.end():]).strip()
                        break

            # Drop the metadata line our own exporter writes under the title
            html = re.sub(r"^<p>\s*Article ID: .*?</p>", "", html, count=1, flags=re.IGNORECASE | re.DOTALL).strip()

            if not title:
                title = "Imported Article"

            return [{
                "title": title,
                "body": html,       # keep full HTML; Zendesk supports HTML body
            }]

        except Exception as e:
            logging.error("Failed to parse Word file %s with mammoth: %s", file_path, e)
            raise ValueError(f"Error parsing Word file: {e}")

    # -----------------------------
    # Reference data
    # -----------------------------
    def get_sections_and_categories(self):
        """Fetch sections and categories (fallback only)."""
        try:
            s = requests.Session()
            s.headers.update(self.headers)

            rc = s.get(f"{self.base_url}/api/v2/help_center/categories.json", timeout=30)
            rc.raise_for_status()
            categories = rc.json().get('categories', [])

            rs = s.get(f"{self.base_url}/api/v2/help_center/sections.json", timeout=30)
            rs.raise_for_status()
            sections = rs.json().get('sections', [])

            logging.info("Found %d categories and %d sections", len(categories), len(sections))
            return categories, sections
        except Exception as e:
            logging.error("Failed to fetch sections/categories: %s", e)
            return [], []

    def get_permission_groups_and_segments(self):
        """Fetch permission groups and user segments; safe fallbacks if missing."""
        try:
            s = requests.Session()
            s.headers.update(self.headers)

            rpg = s.get(f"{self.base_url}/api/v2/guide/permission_groups.json", timeout=30)
            rpg.raise_for_status()
            permission_groups = rpg.json().get('permission_groups', [])

            user_segments: List[Dict] = []
            try:
                rus = s.get(f"{self.base_url}/api/v2/help_center/user_segments.json", timeout=30)
                if rus.ok:
                    user_segments = rus.json().get('user_segments', [])
                else:
                    logging.info("User segments endpoint unavailable; using default access")
            except Exception:
                logging.info("User segments endpoint unavailable; using default access")

            logging.info("Found %d permission groups and %d user segments", len(permission_groups), len(user_segments))
            return permission_groups, user_segments

        except Exception as e:
            logging.error("Failed to fetch permission groups / user segments: %s", e)
            return [], []

    # -----------------------------
    # Lookups / updates
    # -----------------------------
    def _list_section_articles(self, section_id: int) -> List[Dict]:
        """List ALL articles in a section (paginated) from the section’s brand host."""
        brand_base, _, brand_obj = self._brand_base_for_section(section_id)
        headers = dict(self.session.headers)
        if brand_obj and brand_obj.get("id"):
            headers["X-Zendesk-Brand-Id"] = str(brand_obj["id"])

        out: List[Dict] = []
        url = f"{brand_base}/api/v2/help_center/sections/{int(section_id)}/articles.json?per_page=100"
        while url:
            r = self.session.get(url, headers=headers, timeout=90)
            if r.status_code >= 400:
                logging.error("API GET %s failed (%s): %s", url, r.status_code, r.text)
            r.raise_for_status()
            j = r.json() or {}
            out.extend(j.get('articles', []))
            url = j.get('next_page')
        return out

    def _find_existing_article(
        self,
        section_id: int,
        candidate_id: Optional[int],
        title: str,
        labels: List[str]
    ) -> Optional[Dict]:
        """Find existing article in section by ID, exact title, or overlapping labels."""
        # --- 1) Try by numeric ID, but use the section's brand host (avoid account-host 404s)
        if candidate_id:
            try:
                brand_base, _, brand_obj = self._brand_base_for_section(section_id)
                headers = dict(self.session.headers)
                if brand_obj and brand_obj.get("id"):
                    headers["X-Zendesk-Brand-Id"] = str(brand_obj["id"])
                r = self.session.get(
                    f"{brand_base}/api/v2/help_center/articles/{int(candidate_id)}.json",
                    headers=headers, timeout=90
                )
                if r.status_code == 200:
                    art = (r.json() or {}).get('article') or {}
                    if art and art.get('section_id') == section_id:
                        return art
            except requests.RequestException:
                pass  # treat as not found and continue

        # --- 2) Pull section inventory once (brand-aware)
        articles = self._list_section_articles(section_id)

        # --- 3) Exact title match (case-insensitive)
        if title:
            tnorm = title.strip().lower()
            for a in articles:
                if (a.get('title') or '').strip().lower() == tnorm:
                    return a

        # --- 4) Overlapping labels
        want = set(x.strip().lower() for x in (labels or []) if str(x).strip())
        if want:
            for a in articles:
                have = set(x.strip().lower() for x in (a.get('label_names') or []) if str(x).strip())
                if have & want:
                    return a

        return None

    # -----------------------------
    # Payload build
    # -----------------------------
    def _build_payload_from_row(
        self,
        article_data: Dict,
        permission_group_id: Optional[int],
        user_segment_id: Optional[int],
        update_labels_mode: str
    ) -> Dict:
        """Build a Zendesk article payload from a parsed row."""
        title = str(article_data.get('title', '')).strip()
        body = str(article_data.get('body', '')).strip()
        if not title and not body:
            raise ValueError("Row has neither title nor body")

        payload: Dict = {
            "article": {
                "locale": article_data.get('locale', 'en-us'),
                "draft": bool(article_data.get('draft', False)),
            }
        }
        if title:
            payload["article"]["title"] = title
        if body:
            payload["article"]["body"] = body

        if permission_group_id is not None:
            payload["article"]["permission_group_id"] = permission_group_id
        if user_segment_id is not None:
            payload["article"]["user_segment_id"] = user_segment_id

        labels = article_data.get('label_names')
        if labels:
            if isinstance(labels, str):
                labels = [s.strip() for s in labels.split(',') if s.strip()]
            labels = [str(x).strip() for x in labels if str(x).strip()]
            payload["article"]["label_names"] = labels

        if 'promoted' in article_data:
            payload["article"]["promoted"] = bool(article_data['promoted'])

        return payload

    def _prepare_permissions(self, permission_group_id, user_segment_id) -> Tuple[int, Optional[int]]:
        """Fetch default permission group / user segment if needed."""
        if permission_group_id is not None and user_segment_id is not None:
            return permission_group_id, user_segment_id

        pgs, segs = self.get_permission_groups_and_segments()
        if permission_group_id is None:
            if pgs:
                permission_group_id = pgs[0]['id']
                logging.info("Using default permission group: %s (ID: %s)",
                             pgs[0].get('name', 'Unknown'), permission_group_id)
            else:
                permission_group_id = 1
                logging.warning("Permission groups unavailable; using fallback ID=1")
        if user_segment_id is None:
            user_segment_id = segs[0]['id'] if segs else None
            if user_segment_id:
                logging.info("Using default user segment: ID %s", user_segment_id)
            else:
                logging.info("Using user_segment_id=None (visible to everyone)")
        return permission_group_id, user_segment_id

    # -----------------------------
    # Core create/update
    # -----------------------------
    def create_article(
        self,
        article_data: Dict,
        section_id: int,
        permission_group_id: Optional[int],
        user_segment_id: Optional[int]
    ) -> Dict:
        """Create a new article in the specific section (brand host)."""
        brand_base, _, brand_obj = self._brand_base_for_section(section_id)

        headers = dict(self.session.headers)
        if brand_obj and brand_obj.get("id"):
            headers["X-Zendesk-Brand-Id"] = str(brand_obj["id"])

        url = f"{brand_base}/api/v2/help_center/sections/{int(section_id)}/articles.json"
        r = self.session.post(url, json={"article": article_data}, headers=headers, timeout=90)

        if r.status_code >= 400:
            logging.error("API POST %s failed (%s): %s", url, r.status_code, r.text)
        r.raise_for_status()

        created = r.json().get('article', {}) or {}
        logging.info("Created article: %s - %s", created.get('id'), created.get('title'))
        return created

    def update_article(
        self,
        article_id: int,
        payload: Dict,
        update_labels_mode: str,
        existing_labels: Optional[List[str]],
        *,
        section_id: Optional[int] = None,
        existing_html_url: Optional[str] = None,
    ) -> Dict:
        """
        Update an existing article; now host-aware.

        Priority for choosing host:
          1) existing_html_url (from the found article) → use its host
          2) section_id → resolve the section’s brand host
          3) fallback: try account host lookup to discover html_url → host
          4) final fallback: account host
        """
        from urllib.parse import urlparse as _urlparse

        headers = dict(self.session.headers)

        # 1) If caller told us the article's html_url, use its host
        host = ""
        if existing_html_url:
            host = _urlparse(existing_html_url.strip().lower()).netloc

        # 2) Else if we know the section, resolve its brand host
        if not host and section_id:
            brand_base, _, brand_obj = self._brand_base_for_section(int(section_id))
            host = _urlparse(brand_base).netloc
            if brand_obj and brand_obj.get("id"):
                headers["X-Zendesk-Brand-Id"] = str(brand_obj["id"])

        # 3) Else try to look up the article to learn its host; fallback is account host
        if not host and not section_id:
            # Only probe account host if we truly have no section hint
            try:
                look = self.safe_request('GET', f"{self.base_url}/api/v2/help_center/articles/{int(article_id)}.json", quiet_404=True)
                art = (look.json() or {}).get('article', {}) or {}
                html = (art.get("html_url") or "").strip().lower()
                if html:
                    host = _urlparse(html).netloc
            except Exception:
                host = ""

        # Build final URL and Brand header if we recognized the host
        if host and host != self.base_url.replace("https://", ""):
            self._load_brands_map()
            brand_obj = self._host_to_brand.get(host) if self._host_to_brand else None
            if brand_obj and brand_obj.get("id"):
                headers["X-Zendesk-Brand-Id"] = str(brand_obj["id"])
            url = f"https://{host}/api/v2/help_center/articles/{int(article_id)}.json"
        else:
            url = f"{self.base_url}/api/v2/help_center/articles/{int(article_id)}.json"

        # Merge labels client-side if requested
        if update_labels_mode == "merge" and payload.get("article", {}).get("label_names") and existing_labels is not None:
            merged = set(x.strip() for x in (existing_labels or []) if str(x).strip())
            for x in payload["article"]["label_names"]:
                merged.add(str(x).strip())
            payload = {**payload}
            payload["article"] = {**payload["article"], "label_names": sorted(merged)}

        r = self.session.put(url, json=payload, headers=headers, timeout=90)
        if r.status_code >= 400:
            logging.error("API PUT %s failed (%s): %s", url, r.status_code, r.text)
        r.raise_for_status()

        updated = (r.json() or {}).get('article', {}) or {}
        logging.info("Updated article: %s - %s", updated.get('id'), updated.get('title'))
        return updated

    # -----------------------------
    # Public driver
    # -----------------------------
    def import_articles(
        self,
        file_path: str,
        section_id: Optional[int] = None,
        update_existing: bool = False,
        create_drafts: bool = True,
        progress_callback: Optional[Callable[[int, int, str], None]] = None,
        update_labels_mode: str = "merge",
        permission_group_id: Optional[int] = None,
        user_segment_id: Optional[int] = None,
        default_locale: Optional[str] = None,
    ) -> Dict:
        """
        Import articles from .xlsx/.xls or .docx into a specific section.
        - If update_existing=True: update by id, else exact title, else overlapping labels.
        - If no section_id is provided and rows don’t override it, picks first available section.
        - default_locale (UI) is used only for rows that do not carry a locale.
        """
        # Validate section early (only if provided)
        if section_id is not None:
            try:
                section_id = int(section_id)
            except (TypeError, ValueError):
                raise ValueError(f"section_id must be an integer, got {section_id!r}")
        logging.info("[IMPORT] Target section_id (from caller) = %s", section_id)

        if progress_callback:
            progress_callback(0, 0, "Parsing import file...")

        articles = self.parse_file(file_path)

        if not articles:
            raise ValueError("No articles found in the import file")

        # If no section_id provided anywhere, pick a fallback
        if not section_id:
            _, sections = self.get_sections_and_categories()
            if sections:
                section_id = int(sections[0]['id'])
                logging.info("Using default section: %s (ID: %s)", sections[0].get('name'), section_id)
            else:
                raise ValueError("No sections available and no section_id provided.")

        # Default permission & segment if not provided
        permission_group_id, user_segment_id = self._prepare_permissions(permission_group_id, user_segment_id)

        total = len(articles)
        if progress_callback:
            progress_callback(0, total, f"Found {total} articles to import")

        created, updated, failed = [], [], []

        for i, row in enumerate(articles, 1):
            try:
                # drafts override
                if create_drafts:
                    row['draft'] = True

                # per-row section override or fallback
                row_section_id = int(row['section_id']) if row.get('section_id') else section_id

                # Row locale wins; the UI default only fills gaps
                row['locale'] = (row.get('locale') or default_locale or 'en-us').strip().lower()

                candidate_id = row.get('id')
                title = (row.get('title') or '').strip()
                labels = row.get('label_names') or []
                if isinstance(labels, str):
                    labels = [s.strip() for s in labels.split(',') if s.strip()]

                if progress_callback:
                    preview = (title or f"#{candidate_id}" or "Untitled")[:60]
                    progress_callback(i, total, f"Processing {i}/{total}: {preview}...")

                # build payload
                payload = self._build_payload_from_row(
                    row, permission_group_id, user_segment_id, update_labels_mode
                )

                # update-in-place
                if update_existing:
                    existing = self._find_existing_article(row_section_id, candidate_id, title, labels)
                    if existing:
                        existing_labels = existing.get('label_names') or []
                        existing_html_url = existing.get('html_url')  # host hint
                        updated_article = self.update_article(
                            existing['id'],
                            payload,
                            update_labels_mode,
                            existing_labels,
                            section_id=row_section_id,           # section hint (brand host)
                            existing_html_url=existing_html_url  # direct host hint
                        )
                        updated.append(updated_article)
                        continue

                # create
                created_article = self.create_article(
                    payload["article"],
                    section_id=row_section_id,
                    permission_group_id=permission_group_id,
                    user_segment_id=user_segment_id
                )
                created.append(created_article)

            except Exception as e:
                logging.error("Failed to import %r: %s", row.get('title'), e)
                failed.append({"title": row.get('title'), "error": str(e)})

        return {
            "created_articles": created,
            "updated_articles": updated,
            "failed_articles": failed,
            "total_processed": total,
            "success_count": len(created) + len(updated),
            "created_count": len(created),
            "updated_count": len(updated),
            "fail_count": len(failed),
        }


def _to_bool(v) -> bool:
    """Excel gives us True/False, 1/0, 'TRUE', 'yes', 'no'... treat them all sensibly."""
    if isinstance(v, bool):
        return v
    if isinstance(v, (int, float)):
        return v != 0
    return str(v).strip().lower() in ("1", "true", "yes", "y", "on")


def _esc(s: str) -> str:
    return (s or "").replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
