import os
import sys
import uuid
import socket
import logging
import secrets
import threading
from datetime import datetime, timezone
from typing import Dict, Any, Optional, List

import requests
from flask import (
    Flask, request, jsonify, render_template, redirect,
    url_for, send_from_directory, flash
)
from werkzeug.utils import secure_filename

from zendesk_importer import ZendeskImporter
from enhanced_zendesk_exporter import EnhancedZendeskExporter

APP_NAME = "TechWitch Zendesk Helper"
APP_VERSION = "1.1.0"

# -----------------------------------------------------------------------------
# Paths
#
# Two cases:
#   * running from source: everything sits next to app.py
#   * running as a PyInstaller exe: templates/static are unpacked into a temp
#     folder (sys._MEIPASS) but uploads/exports must live next to the exe so
#     the user can find them and they survive restarts.
# -----------------------------------------------------------------------------
FROZEN = getattr(sys, "frozen", False)
if FROZEN:
    RESOURCE_DIR = getattr(sys, "_MEIPASS", os.path.dirname(sys.executable))
    BASE_DIR = os.path.dirname(os.path.abspath(sys.executable))
else:
    RESOURCE_DIR = os.path.abspath(os.path.dirname(__file__))
    BASE_DIR = RESOURCE_DIR

UPLOAD_DIR = os.path.join(BASE_DIR, "uploads")
EXPORT_DIR = os.path.join(BASE_DIR, "exports")
os.makedirs(UPLOAD_DIR, exist_ok=True)
os.makedirs(EXPORT_DIR, exist_ok=True)

# -----------------------------------------------------------------------------
# Flask setup
# -----------------------------------------------------------------------------
app = Flask(
    __name__,
    static_folder=os.path.join(RESOURCE_DIR, "static"),
    template_folder=os.path.join(RESOURCE_DIR, "templates"),
)
app.secret_key = os.environ.get("FLASK_SECRET_KEY") or secrets.token_hex(32)  # for flash()
app.config["MAX_CONTENT_LENGTH"] = 200 * 1024 * 1024  # 200 MB of uploads per request

logging.basicConfig(
    level=os.environ.get("LOG_LEVEL", "INFO"),
    format="%(levelname)s:%(name)s:%(message)s"
)
log = logging.getLogger("app")

# In-memory job registries for the simple progress UIs
EXPORT_JOBS: Dict[str, Dict[str, Any]] = {}
IMPORT_JOBS: Dict[str, Dict[str, Any]] = {}
_JOBS_LOCK = threading.Lock()

ALLOWED_IMPORT_EXT = {".xlsx", ".xls", ".docx"}

# -----------------------------------------------------------------------------
# Helpers
# -----------------------------------------------------------------------------
def _bool(s: Optional[str], default: bool = False) -> bool:
    if s is None:
        return default
    return str(s).strip().lower() in ("1", "true", "yes", "on")

def _load_creds_from_form_or_env(src: Dict[str, Any]):
    subdomain = (src.get("subdomain") or os.environ.get("ZENDESK_SUBDOMAIN") or "").strip()
    email     = (src.get("email")     or os.environ.get("ZENDESK_EMAIL")     or "").strip()
    api_token = (src.get("api_token") or os.environ.get("ZENDESK_API_TOKEN") or "").strip()
    if not (subdomain and email and api_token):
        raise RuntimeError("Missing Zendesk credentials (subdomain/email/api_token).")
    # People paste the full URL. Reduce "https://acme.zendesk.com/" to "acme".
    subdomain = subdomain.replace("https://", "").replace("http://", "")
    subdomain = subdomain.split("/")[0].replace(".zendesk.com", "")
    return subdomain, email, api_token

def _job_update(reg: Dict[str, Dict[str, Any]], job_id: str, **updates):
    with _JOBS_LOCK:
        j = reg.get(job_id, {})
        j.update(updates)
        reg[job_id] = j

def _job_get(reg: Dict[str, Dict[str, Any]], job_id: str) -> Dict[str, Any]:
    with _JOBS_LOCK:
        return dict(reg.get(job_id, {"status": "error", "error": "Unknown job"}))

def _basic_auth(email: str, api_token: str) -> str:
    import base64
    creds = base64.b64encode(f"{email}/token:{api_token}".encode()).decode()
    return f"Basic {creds}"

def _zendesk_session(email: str, api_token: str) -> "requests.Session":
    s = requests.Session()
    s.headers.update({
        "Authorization": _basic_auth(email, api_token),
        "Accept": "application/json",
        "User-Agent": f"{APP_NAME}/{APP_VERSION}",
    })
    return s

def _api_error(r: "requests.Response") -> str:
    """Turn a Zendesk error response into something a person can act on."""
    if r.status_code == 401:
        return "Zendesk rejected the credentials (401). Check the email, the API token and that token access is enabled in Admin Center."
    if r.status_code == 403:
        return "Zendesk refused the request (403). The user needs Guide admin or agent rights for this brand."
    if r.status_code == 404:
        return f"Not found (404): {r.url}"
    if r.status_code == 429:
        return "Zendesk rate limit hit (429). Wait a minute and try again."
    try:
        j = r.json()
        return f"{r.status_code}: {j.get('error') or j.get('description') or r.text[:300]}"
    except Exception:
        return f"{r.status_code}: {r.text[:300]}"

# -----------------------------------------------------------------------------
# UI pages
# -----------------------------------------------------------------------------
@app.context_processor
def _inject_globals():
    return {"app_name": APP_NAME, "app_version": APP_VERSION, "frozen": FROZEN}

@app.get("/")
def index():
    return render_template("index.html")

@app.get("/health")
def health():
    return jsonify({"status": "ok", "version": APP_VERSION,
                    "time": datetime.now(timezone.utc).isoformat()})

@app.post("/shutdown")
def shutdown():
    """Used by the 'Quit' button so the exe can be closed from the browser."""
    def _die():
        import time
        time.sleep(0.5)
        os._exit(0)
    threading.Thread(target=_die, daemon=True).start()
    return jsonify({"ok": True, "message": "Server stopping. You can close this tab."})

# -----------------------------------------------------------------------------
# API used by the front-end (credentials are posted in JSON)
# -----------------------------------------------------------------------------
@app.post("/api/brands")
def api_brands():
    try:
        data = request.get_json(force=True) or {}
        subdomain, email, token = _load_creds_from_form_or_env(data)
        base_url = f"https://{subdomain}.zendesk.com"
        s = _zendesk_session(email, token)

        r = s.get(f"{base_url}/api/v2/brands.json", timeout=60)
        if r.status_code >= 400:
            return jsonify({"error": _api_error(r)}), 400
        brands = (r.json() or {}).get("brands", []) or []
        out = [{"id": b.get("id"), "name": b.get("name"), "active": b.get("active", True)} for b in brands]
        return jsonify(out)
    except requests.ConnectionError:
        return jsonify({"error": f"Could not reach https://{subdomain}.zendesk.com. Check the subdomain and your connection."}), 400
    except Exception as e:
        return jsonify({"error": str(e)}), 400

@app.post("/api/sections_by_brand")
def api_sections_by_brand():
    """
    Robust brand-scoped sections:
    1) Try account-host brand endpoint
    2) Query the brand's host directly (with X-Zendesk-Brand-Id)
    3) Fallback via category->sections
    """
    try:
        data = request.get_json(force=True) or {}
        subdomain, email, token = _load_creds_from_form_or_env(data)
        raw_brand_id = data.get("brand_id")
        if raw_brand_id is None or str(raw_brand_id).strip() == "":
            return jsonify({"error": "brand_id required"}), 400
        brand_id = int(raw_brand_id)

        account_base = f"https://{subdomain}.zendesk.com"
        s = _zendesk_session(email, token)

        def _paged(url: str, key: str, headers: Optional[dict] = None) -> List[Dict[str, Any]]:
            out: List[Dict[str, Any]] = []
            while url:
                r = s.get(url, headers=headers or {}, timeout=60)
                if r.status_code >= 400:
                    return [{"__error__": True, "status": r.status_code, "text": r.text}]
                j = r.json() or {}
                out.extend(j.get(key, []) or [])
                url = j.get("next_page")
            return out

        def _failed(lst) -> bool:
            return bool(lst and isinstance(lst[0], dict) and lst[0].get("__error__"))

        # Load brands to find host
        brands_r = s.get(f"{account_base}/api/v2/brands.json", timeout=60)
        if brands_r.status_code >= 400:
            return jsonify({"error": _api_error(brands_r)}), 400
        brands = (brands_r.json() or {}).get("brands", []) or []
        b = next((x for x in brands if int(x.get("id", 0)) == brand_id), None)
        if not b:
            return jsonify({"error": f"brand {brand_id} not found on account"}), 404

        custom_host = (b.get("host_mapping") or "").strip().lower()
        sub_host    = (b.get("subdomain") or "").strip()
        brand_host  = custom_host or (f"{sub_host}.zendesk.com" if sub_host else "")

        # 1) official brand-scoped endpoint on account host
        sections = _paged(f"{account_base}/api/v2/help_center/brands/{brand_id}/sections.json?per_page=100", "sections")

        # 2) query the brand host directly (natural scoping)
        if _failed(sections) or not sections:
            if not brand_host:
                brand_host = account_base.replace("https://", "")
            headers = {"X-Zendesk-Brand-Id": str(brand_id)}
            sections2 = _paged(f"https://{brand_host}/api/v2/help_center/sections.json?per_page=100", "sections", headers=headers)

            if _failed(sections2):
                # 3) fallback via categories
                cats1 = _paged(f"{account_base}/api/v2/help_center/brands/{brand_id}/categories.json?per_page=100", "categories")
                if _failed(cats1):
                    cats_all = _paged(f"{account_base}/api/v2/help_center/categories.json?per_page=100", "categories")
                    cats = [] if _failed(cats_all) else [c for c in cats_all if int(c.get("brand_id") or 0) == brand_id]
                else:
                    cats = cats1

                cat_ids = {int(c.get("id")) for c in (cats or []) if c.get("id")}
                if not cat_ids:
                    return jsonify([])

                filtered: List[Dict[str, Any]] = []
                for cid in sorted(cat_ids):
                    per = _paged(f"{account_base}/api/v2/help_center/categories/{cid}/sections.json?per_page=100", "sections")
                    if _failed(per):
                        continue
                    filtered.extend(per)
                sections = filtered
            else:
                sections = sections2

        return jsonify([
            {"id": x.get("id"), "name": x.get("name"), "category_id": x.get("category_id")}
            for x in (sections or [])
        ])

    except Exception as e:
        log.exception("sections_by_brand failed")
        return jsonify({"error": str(e)}), 400

@app.post("/api/create_cat_and_section")
def api_create_cat_and_section():
    """
    Creates a Category + Section on a given brand. Body JSON:
      { subdomain, email, api_token, brand_id, category_name, section_name, locale }
    """
    try:
        data = request.get_json(force=True) or {}
        subdomain, email, token = _load_creds_from_form_or_env(data)
        brand_id = int(data["brand_id"])
        category_name = (data.get("category_name") or "Knowledge Base").strip()
        section_name  = (data.get("section_name")  or "Getting Started").strip()
        locale        = (data.get("locale")        or "en-us").strip().lower()

        base_url = f"https://{subdomain}.zendesk.com"
        s = _zendesk_session(email, token)

        rc = s.post(f"{base_url}/api/v2/help_center/categories.json",
                    json={"category": {"name": category_name, "brand_id": brand_id, "locale": locale}},
                    timeout=60)
        if rc.status_code >= 400:
            return jsonify({"error": "Category create failed. " + _api_error(rc)}), 400
        cat = (rc.json() or {}).get("category") or {}
        cat_id = cat.get("id")

        rs = s.post(f"{base_url}/api/v2/help_center/sections.json",
                    json={"section": {"name": section_name, "category_id": cat_id, "brand_id": brand_id, "locale": locale}},
                    timeout=60)
        if rs.status_code >= 400:
            return jsonify({"error": "Section create failed. " + _api_error(rs)}), 400
        sec = (rs.json() or {}).get("section") or {}

        return jsonify({
            "category_id": cat_id,
            "category_name": cat.get("name"),
            "section_id": sec.get("id"),
            "section_name": sec.get("name"),
        })
    except Exception as e:
        return jsonify({"error": str(e)}), 400

# -----------------------------------------------------------------------------
# Start EXPORT job -> /export page polls /progress/<job_id>
# -----------------------------------------------------------------------------
@app.post("/start_export")
def start_export():
    try:
        form = request.form
        subdomain, email, token = _load_creds_from_form_or_env(form)
        export_type = (form.get("export_type") or "excel").strip().lower()
        if export_type not in ("excel", "word"):
            raise ValueError(f"Unknown export type: {export_type}")
        brand_id = form.get("brand_id")
        brand_id = int(brand_id) if brand_id else None

        job_id = uuid.uuid4().hex[:12]
        EXPORT_JOBS[job_id] = {"status": "starting", "message": "Preparing…", "files": [],
                               "sections_total": 0, "sections_done": 0, "articles": 0}

        def _progress(sections_done: int, sections_total: int, articles: int, msg: str):
            _job_update(EXPORT_JOBS, job_id, sections_done=sections_done,
                        sections_total=sections_total, articles=articles, message=msg)

        def _run():
            try:
                _job_update(EXPORT_JOBS, job_id, status="running", message="Listing sections…")
                exporter = EnhancedZendeskExporter(subdomain, email, token, export_dir=EXPORT_DIR)
                files = exporter.export_articles(format_type=export_type, brand_id=brand_id,
                                                 progress_callback=_progress)

                # Build links WITHOUT url_for (no app context in this thread)
                file_objs = [
                    {"name": os.path.basename(fp), "url": f"/downloads/{os.path.basename(fp)}"}
                    for fp in files
                ]
                _job_update(EXPORT_JOBS, job_id, status="completed", message="Done", files=file_objs)
            except Exception as e:
                log.exception("Export job failed")
                _job_update(EXPORT_JOBS, job_id, status="error", error=str(e))

        threading.Thread(target=_run, daemon=True).start()
        return render_template("export.html", job_id=job_id)

    except Exception as e:
        flash(f"Export failed to start: {e}", "error")
        return redirect(url_for("index"))

@app.get("/progress/<job_id>")
def get_progress(job_id: str):
    return jsonify(_job_get(EXPORT_JOBS, job_id))

@app.get("/downloads/<path:filename>")
def download_export(filename: str):
    return send_from_directory(EXPORT_DIR, filename, as_attachment=True)

# -----------------------------------------------------------------------------
# Start IMPORT job -> /import page polls /import_progress/<job_id>
# -----------------------------------------------------------------------------
@app.post("/start_import")
def start_import():
    try:
        # Creds + options are standard form fields (hidden inputs set by JS)
        subdomain, email, token = _load_creds_from_form_or_env(request.form)

        # Options
        section_id = request.form.get("section_id")
        section_id = int(section_id) if section_id else None
        update_existing     = _bool(request.form.get("update_existing"), False)
        create_drafts       = _bool(request.form.get("create_drafts"), True)
        update_labels_mode  = (request.form.get("update_labels_mode") or "merge").strip().lower()
        default_locale      = (request.form.get("default_locale") or "").strip().lower() or None

        # Files: input name is "import_file" (multiple)
        files = request.files.getlist("import_file") or []
        saved_paths = []
        skipped = []
        for f in files:
            if not f or not f.filename:
                continue
            fn = secure_filename(f.filename)
            ext = os.path.splitext(fn)[1].lower()
            if ext not in ALLOWED_IMPORT_EXT:
                skipped.append(f.filename)
                continue
            # Prefix with a short id so two uploads with the same name never clobber each other
            full = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex[:8]}_{fn}")
            f.save(full)
            saved_paths.append((full, fn))

        if not saved_paths:
            msg = "Please choose at least one .xlsx / .xls / .docx file to import."
            if skipped:
                msg += f" Skipped: {', '.join(skipped)}"
            flash(msg, "error")
            return redirect(url_for("index"))

        job_id = uuid.uuid4().hex[:12]
        IMPORT_JOBS[job_id] = {
            "status": "starting",
            "message": "Parsing files…",
            "processed_articles": 0,
            "total_articles": 0,
            "created_articles": 0,
            "failed_articles": [],
            "events": [],
            "skipped_files": skipped,
        }

        def _run():
            try:
                importer = ZendeskImporter(subdomain, email, token)

                total_created   = 0
                total_failed    = []
                total_updated   = 0
                total_processed = 0
                total_expected  = 0
                events          = []

                # Count rows up front so the bar has a real total
                for path, _ in saved_paths:
                    try:
                        total_expected += len(importer.parse_file(path))
                    except Exception as e:
                        log.warning("Could not pre-count %s: %s", path, e)
                _job_update(IMPORT_JOBS, job_id, total_articles=total_expected, status="running")

                def _progress(i: int, total: int, msg: str):
                    _job_update(IMPORT_JOBS, job_id,
                                processed_articles=total_processed + i,
                                message=msg)

                for path, display_name in saved_paths:
                    _job_update(IMPORT_JOBS, job_id, status="running",
                                message=f"Importing {display_name}…")

                    try:
                        result = importer.import_articles(
                            file_path=path,
                            section_id=section_id,
                            update_existing=update_existing,
                            create_drafts=create_drafts,
                            progress_callback=_progress,
                            update_labels_mode=update_labels_mode,
                            permission_group_id=None,
                            user_segment_id=None,
                            default_locale=default_locale,
                        )
                    except Exception as e:
                        # One bad file should not kill the whole run
                        log.exception("Import of %s failed", display_name)
                        events.append({"type": "failed", "id": None, "title": display_name, "error": str(e)})
                        total_failed.append({"title": display_name, "error": str(e)})
                        _job_update(IMPORT_JOBS, job_id, events=list(events))
                        continue

                    total_created   += result.get("created_count", 0)
                    total_updated   += result.get("updated_count", 0)
                    total_processed += result.get("total_processed", 0)
                    total_failed.extend(result.get("failed_articles", []))

                    for a in result.get("created_articles", []):
                        events.append({"type": "created", "id": a.get("id"), "title": a.get("title"),
                                       "locale": a.get("locale"), "draft": bool(a.get("draft")),
                                       "section_id": a.get("section_id"), "html_url": a.get("html_url")})
                    for a in result.get("updated_articles", []):
                        events.append({"type": "updated", "id": a.get("id"), "title": a.get("title"),
                                       "locale": a.get("locale"), "draft": bool(a.get("draft")),
                                       "section_id": a.get("section_id"), "html_url": a.get("html_url")})
                    for fa in result.get("failed_articles", []):
                        events.append({"type": "failed", "id": None, "title": fa.get("title"), "error": fa.get("error")})

                    _job_update(IMPORT_JOBS, job_id,
                                processed_articles=total_processed,
                                created_articles=total_created,
                                message=f"Done {display_name}",
                                events=list(events))

                _job_update(IMPORT_JOBS, job_id, status="completed",
                            processed_articles=max(total_processed, total_expected),
                            message=f"Completed: {total_created} created, {total_updated} updated, {len(total_failed)} failed",
                            failed_articles=total_failed,
                            events=list(events))

            except Exception as e:
                log.exception("Import job failed")
                _job_update(IMPORT_JOBS, job_id, status="error", error=str(e))
            finally:
                # Uploaded copies are not needed once the job is done
                for path, _ in saved_paths:
                    try:
                        os.remove(path)
                    except OSError:
                        pass

        threading.Thread(target=_run, daemon=True).start()
        return render_template("import.html", job_id=job_id)

    except Exception as e:
        log.exception("Failed to start import")
        flash(f"Import failed to start: {e}", "error")
        return redirect(url_for("index"))

@app.get("/import_progress/<job_id>")
def get_import_progress(job_id: str):
    return jsonify(_job_get(IMPORT_JOBS, job_id))

# -----------------------------------------------------------------------------
# JSON endpoints for scripting (no UI)
# -----------------------------------------------------------------------------
@app.post("/export")
def run_export_json():
    try:
        data = request.get_json(force=True) or {}
        subdomain, email, token = _load_creds_from_form_or_env(data)
        exporter = EnhancedZendeskExporter(subdomain, email, token, export_dir=EXPORT_DIR)
        files = exporter.export_articles(format_type=data.get("format", "excel"), brand_id=data.get("brand_id"))
        return jsonify({"ok": True, "files": files})
    except Exception as e:
        log.exception("Export failed")
        return jsonify({"ok": False, "error": str(e)}), 500

@app.post("/import")
def run_import_json():
    """Multipart endpoint expecting a 'file' field."""
    try:
        subdomain, email, token = _load_creds_from_form_or_env(request.form)
        f = request.files.get("file")
        if not f:
            return jsonify({"ok": False, "error": "Missing file"}), 400
        path = os.path.join(UPLOAD_DIR, f"{uuid.uuid4().hex[:8]}_{secure_filename(f.filename)}")
        f.save(path)

        section_id = request.form.get("section_id")
        section_id = int(section_id) if section_id else None

        importer = ZendeskImporter(subdomain, email, token)
        try:
            result = importer.import_articles(
                file_path=path,
                section_id=section_id,
                update_existing=_bool(request.form.get("update_existing"), False),
                create_drafts=_bool(request.form.get("create_drafts"), True),
                progress_callback=None,
                update_labels_mode=(request.form.get("update_labels_mode") or "merge").strip().lower(),
                permission_group_id=None,
                user_segment_id=None,
                default_locale=(request.form.get("default_locale") or "").strip().lower() or None,
            )
        finally:
            try:
                os.remove(path)
            except OSError:
                pass
        return jsonify({"ok": True, "result": result})
    except Exception as e:
        log.exception("Bulk import failed")
        return jsonify({"ok": False, "error": str(e)}), 500

# -----------------------------------------------------------------------------
# Main
# -----------------------------------------------------------------------------
def _free_port(preferred: int) -> int:
    """Use the preferred port if it is free, otherwise let the OS pick one."""
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        try:
            s.bind(("127.0.0.1", preferred))
            return preferred
        except OSError:
            pass
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        s.bind(("127.0.0.1", 0))
        return s.getsockname()[1]

def main():
    import webbrowser, time

    port = _free_port(int(os.environ.get("PORT", "5000")))
    url = f"http://127.0.0.1:{port}"

    print("=" * 60, flush=True)
    print(f" {APP_NAME} v{APP_VERSION}")
    print(f" Running at {url}")
    print(f" Exports go to: {EXPORT_DIR}")
    print(" Keep this window open. Close it (or use Quit in the app) to stop.", flush=True)
    print("=" * 60, flush=True)

    if not _bool(os.environ.get("NO_BROWSER"), False):
        def _open():
            time.sleep(0.8)
            webbrowser.open(url)
        threading.Thread(target=_open, daemon=True).start()

    try:
        from waitress import serve
        serve(app, host="127.0.0.1", port=port, threads=8)
    except ImportError:
        app.run(host="127.0.0.1", port=port, debug=False)

if __name__ == "__main__":
    main()
