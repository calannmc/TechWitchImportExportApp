# enhanced_zendesk_exporter.py

import io
import os
import re
import time
import zipfile
import logging
from datetime import datetime
from typing import Callable, Dict, List, Optional

import pandas as pd
import requests

from html_to_docx import html_to_docx_bytes

_EXCEL_ENGINE = "openpyxl"

ProgressCb = Optional[Callable[[int, int, int, str], None]]


class EnhancedZendeskExporter:
    """
    Brand-aware Help Center exporter.

    - If brand_id is provided:
        * Resolves brand host (custom domain if set, else <subdomain>.zendesk.com)
        * Queries the BRAND HOST for sections and articles
        * Sends X-Zendesk-Brand-Id where useful

    - If brand_id is None:
        * Falls back to account host and exports everything visible

    Output:
        - format_type="excel": single XLSX with one sheet "Articles"
        - format_type="word":  ZIP of .docx files (one per article) plus an index.xlsx
        - Files are written into export_dir (default ./exports)
    """

    def __init__(self, subdomain: str, email: str, api_token: str, export_dir: Optional[str] = None):
        self.subdomain = subdomain.strip()
        self.base_url = f"https://{self.subdomain}.zendesk.com"
        self.export_dir = os.path.abspath(export_dir or "exports")

        self.session = requests.Session()
        self.session.headers.update({
            "Authorization": _basic_auth(email, api_token),
            "Accept": "application/json",
        })

    # ---------------- Public API ----------------
    def export_articles(self, format_type: str = "excel", brand_id: Optional[int] = None,
                        progress_callback: ProgressCb = None) -> List[str]:
        brand_ctx = self._resolve_brand_context(brand_id)

        sections = self._list_sections(brand_ctx)
        if brand_id is not None and not sections:
            logging.info("No sections found for brand_id=%s (host=%s)", brand_id, brand_ctx["host"])

        rows = self._collect_rows(sections, brand_ctx, progress_callback)
        if progress_callback:
            progress_callback(len(sections), len(sections), len(rows), "Writing files…")
        return self._write_output(format_type=format_type, rows=rows)

    # ---------------- HTTP ----------------
    def _get(self, url: str, headers: Optional[Dict] = None, timeout: int = 90) -> requests.Response:
        """GET with a couple of retries on 429 / 5xx."""
        for attempt in range(4):
            r = self.session.get(url, headers=headers, timeout=timeout)
            if r.status_code == 429 or r.status_code >= 500:
                wait = int(r.headers.get("Retry-After") or (2 ** attempt))
                logging.warning("GET %s -> %s, retrying in %ss", url, r.status_code, wait)
                time.sleep(min(wait, 30))
                continue
            return r
        return r

    def _paged(self, url: str, key: str, headers: Optional[Dict] = None,
               raise_on_auth: bool = True, strict: bool = False) -> List[Dict]:
        """Follow next_page links. strict=True turns any HTTP error into an exception
        instead of quietly returning what was collected so far."""
        out: List[Dict] = []
        while url:
            r = self._get(url, headers=headers)
            if r.status_code in (401, 403) and raise_on_auth:
                raise RuntimeError(_auth_error(r))
            if r.status_code >= 400:
                logging.warning("GET %s failed (%s): %s", url, r.status_code, r.text[:200])
                if strict:
                    raise RuntimeError(f"Zendesk returned {r.status_code} for {url}. Check the subdomain.")
                break
            j = r.json() or {}
            out.extend(j.get(key, []) or [])
            url = j.get("next_page")
        return out

    # ---------------- Internals ----------------
    def _resolve_brand_context(self, brand_id: Optional[int]) -> Dict:
        """
        Decide which host to hit and what headers to send.
        Returns dict: { 'host': 'xxx.zendesk.com' or custom domain,
                        'headers': {...}, 'brand_id': <int or None> }
        """
        ctx = {"host": self.base_url.replace("https://", ""), "headers": {}, "brand_id": brand_id}

        if brand_id is None:
            return ctx

        r = self._get(f"{self.base_url}/api/v2/brands.json", timeout=60)
        if r.status_code in (401, 403):
            raise RuntimeError(_auth_error(r))
        r.raise_for_status()
        brands = (r.json() or {}).get("brands", []) or []
        b = next((x for x in brands if int(x.get("id", 0)) == int(brand_id)), None)
        if not b:
            raise ValueError(f"Brand id {brand_id} not found on account")

        custom = (b.get("host_mapping") or "").strip().lower()
        sub = (b.get("subdomain") or "").strip()
        ctx["host"] = custom or (f"{sub}.zendesk.com" if sub else ctx["host"])
        ctx["headers"] = {"X-Zendesk-Brand-Id": str(brand_id)}
        return ctx

    def _list_sections(self, ctx: Dict) -> List[Dict]:
        """
        List ALL sections for the chosen context.
        Prefer brand host; fall back to brand-scoped endpoint on account host if needed.
        """
        host = ctx["host"]
        headers = ctx["headers"]
        brand_id = ctx["brand_id"]

        if brand_id is not None:
            sections = self._paged(f"https://{host}/api/v2/help_center/sections.json?per_page=100", "sections", headers)
            if sections:
                return sections

            sections = self._paged(f"{self.base_url}/api/v2/help_center/brands/{brand_id}/sections.json?per_page=100", "sections", headers)
            if sections:
                return sections

            cats = self._paged(f"{self.base_url}/api/v2/help_center/brands/{brand_id}/categories.json?per_page=100", "categories", headers)
            if not cats:
                cats_all = self._paged(f"{self.base_url}/api/v2/help_center/categories.json?per_page=100", "categories", headers)
                cats = [c for c in cats_all if int(c.get("brand_id") or 0) == int(brand_id)]
            cat_ids = {int(c.get("id")) for c in cats if c.get("id")}
            if not cat_ids:
                return []

            collected: List[Dict] = []
            for cid in sorted(cat_ids):
                collected.extend(self._paged(f"{self.base_url}/api/v2/help_center/categories/{cid}/sections.json?per_page=100", "sections", headers))
            return collected

        # brand_id is None -> account-wide
        return self._paged(f"{self.base_url}/api/v2/help_center/sections.json?per_page=100", "sections", headers, strict=True)

    def _list_section_articles(self, section_id: int, ctx: Dict) -> List[Dict]:
        # Always prefer the brand host for articles to avoid cross-brand leakage
        return self._paged(
            f"https://{ctx['host']}/api/v2/help_center/sections/{int(section_id)}/articles.json?per_page=100",
            "articles", ctx["headers"], raise_on_auth=False)

    def _collect_rows(self, sections: List[Dict], ctx: Dict, progress_callback: ProgressCb = None) -> List[Dict]:
        rows: List[Dict] = []
        total = len(sections or [])
        for i, s in enumerate(sections or [], 1):
            sid = s.get("id")
            sname = s.get("name")
            cat_id = s.get("category_id")
            if progress_callback:
                progress_callback(i - 1, total, len(rows), f"Section {i}/{total}: {sname}")

            for a in self._list_section_articles(sid, ctx) or []:
                rows.append({
                    "article_id": a.get("id"),
                    "title": a.get("title"),
                    "locale": a.get("locale"),
                    "section_id": sid,
                    "section": sname,
                    "category_id": cat_id,
                    "labels": ",".join(a.get("label_names") or []),
                    "draft": bool(a.get("draft")),
                    "promoted": bool(a.get("promoted")),
                    "html_url": a.get("html_url"),
                    "created_at": a.get("created_at"),
                    "updated_at": a.get("updated_at"),
                    "body": a.get("body") or "",
                })
        return rows

    def _write_output(self, format_type: str, rows: List[Dict]) -> List[str]:
        out_dir = self.export_dir
        os.makedirs(out_dir, exist_ok=True)
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        out_files: List[str] = []
        columns = ["article_id", "title", "locale", "section_id", "section", "category_id", "labels",
                   "draft", "promoted", "html_url", "created_at", "updated_at", "body"]

        if not rows:
            logging.info("No rows to export; creating empty file for traceability.")

        fmt = (format_type or "excel").lower()
        if fmt == "excel":
            out_path = os.path.join(out_dir, f"zendesk_export_{ts}.xlsx")
            df = pd.DataFrame(rows, columns=columns)
            with pd.ExcelWriter(out_path, engine=_EXCEL_ENGINE) as writer:
                df.to_excel(writer, index=False, sheet_name="Articles")
            out_files.append(out_path)

        elif fmt == "word":
            zip_path = os.path.join(out_dir, f"zendesk_export_{ts}.zip")
            used_names = set()
            with zipfile.ZipFile(zip_path, mode="w", compression=zipfile.ZIP_DEFLATED) as zf:
                for r in rows:
                    aid = r.get("article_id")
                    title = (r.get("title") or "").strip() or f"Article {aid}"
                    base = _safe_filename(title) or f"article_{aid}"
                    name = f"{base}.docx"
                    if name in used_names:
                        name = f"{base}_{aid}.docx"
                    used_names.add(name)
                    meta = {
                        "Article ID": aid, "Locale": r.get("locale"), "Section": r.get("section"),
                        "Section ID": r.get("section_id"), "Labels": r.get("labels"),
                        "Draft": r.get("draft"), "URL": r.get("html_url"),
                    }
                    zf.writestr(name, html_to_docx_bytes(title, r.get("body") or "", meta))

                # Index sheet so the ids/sections are not lost
                df = pd.DataFrame([{k: v for k, v in r.items() if k != "body"} for r in rows],
                                  columns=[c for c in columns if c != "body"])
                buf = io.BytesIO()
                with pd.ExcelWriter(buf, engine=_EXCEL_ENGINE) as writer:
                    df.to_excel(writer, index=False, sheet_name="Articles")
                zf.writestr("index.xlsx", buf.getvalue())
            out_files.append(zip_path)

        else:
            raise ValueError(f"Unsupported export format: {format_type}")

        logging.info("Export complete: %s", out_files)
        return out_files


def _basic_auth(email: str, api_token: str) -> str:
    import base64 as _b64
    creds = _b64.b64encode(f"{email}/token:{api_token}".encode()).decode()
    return f"Basic {creds}"

def _auth_error(r: requests.Response) -> str:
    if r.status_code == 401:
        return "Zendesk rejected the credentials (401). Check the email, the API token and that token access is enabled."
    return "Zendesk refused the request (403). The user needs Guide agent or admin rights."

def _safe_filename(name: str) -> str:
    name = re.sub(r"[^\w\-]+", "_", (name or "").strip())
    name = re.sub(r"_+", "_", name).strip("_")
    return name[:120] if name else ""
