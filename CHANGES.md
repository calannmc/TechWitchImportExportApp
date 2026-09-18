# What changed in 1.1.0

Problems found in the stable zip, and what was done about each.

## Would break the exe

- **Exports went to the wrong folder when run as an exe.** The exporter wrote to `exports` relative to the current working directory, but the download route served from a folder relative to `app.py`. Inside a PyInstaller exe that second path is a temp folder that disappears on exit, and if the exe was launched from a shortcut the two folders differed, so Download gave 404. Now both use one `EXPORT_DIR` next to the exe (or next to `app.py` when run from source), and the exporter is told where to write.
- **Templates and static files were looked up relative to `app.py`.** In a one-file exe they live in `sys._MEIPASS`. Fixed with a `RESOURCE_DIR` that switches on `sys.frozen`.
- **`mammoth` was missing from `requirements.txt`** but the Word importer needs it. Without it the code silently fell back to plain-text import and lost all formatting. Added, along with `beautifulsoup4` for the new Word export.
- **`requirements.txt` listed several packages twice with conflicting pins** (Flask 3.1.2 and 3.0.3, pandas 2.3.2 and 2.2.2, lxml 6.0.1 and 5.3.0). pip takes the last one it sees, so the venv did not match what the exe had been built with. Now one list with ranges.
- **The PyInstaller spec had no hidden imports** for mammoth, python-docx's default template or waitress, and had UPX on, which trips some antivirus scanners. Spec rewritten; `pip`-installed `.venv`, `build` and `dist` folders no longer ship in the zip (they were 300 MB of the 369 MB).
- **Port 5000 hard-coded.** If something else had it (common on macOS, where AirPlay uses 5000) the app crashed. It now falls back to a free port and prints which one.
- **No way to stop the exe** other than closing the console. Added a Quit button that calls `/shutdown`.
- Runs under waitress instead of Flask's development server, which is what the Procfile and the exe were meant to use anyway.

## Wrong behaviour

- **Locale precedence was backwards.** The page said the row's `locale` wins, the code made the page's Default locale overwrite every row. Now the row wins and the default fills gaps (your choice).
- **Word import always set `locale` to `en-us`** in the parser, which then beat the page's default. Removed.
- **"Word" export produced a zip of `.html` files, not Word documents.** Now one real `.docx` per article (title, headings, bold/italic, links as real hyperlinks, nested lists, tables, images as placeholder lines with the URL) plus an `index.xlsx` so ids, sections and labels are not lost. The importer round-trips these files: title comes back from the Title style, headings keep their levels, nested bullets stay nested.
- **Word import title detection** looked for a bold paragraph before it looked for a heading, and left the title paragraph in the body so it appeared twice. Now: Title style, then first heading, then bold line, then first paragraph, and the winning paragraph is removed from the body. HTML entities in the title (`&lt;`) are decoded.
- **Excel `draft` / `promoted` columns containing text** (`FALSE`, `no`, `0`) were treated as True because `bool("FALSE")` is True. Proper parsing now.
- **Section and brand caches were class-level dictionaries** shared by every `ZendeskImporter` instance in the process, so switching accounts within one run could reuse another account's section lookups. Moved to the instance.
- **One bad file aborted the whole import.** Now it is reported as failed and the remaining files still run.
- **Uploaded files were kept forever** in `uploads/` (the zip contained your Charity Excellence and Autism Hub documents). They are now deleted once the job finishes, and two uploads with the same name no longer overwrite each other.
- **Unsupported file types** were passed straight through to the importer and failed with a stack trace. Now filtered out and listed as skipped.
- **Account-wide export quietly produced an empty file** when the sections request failed. It now reports the error.
- Subdomain field accepts `https://acme.zendesk.com/` as well as `acme`.
- 401 / 403 / 429 from Zendesk now come back as plain sentences instead of raw JSON.
- Retries with back-off on 429 and 5xx during export.
- `datetime.utcnow()` (deprecated) replaced.
- Flask secret key is random per run instead of `dev-secret`.

## UI

- `export.html` and `import.html` had a `<style>` block above `{% extends %}` (Jinja ignores content outside blocks in a child template) and used Bootstrap classes (`progress-bar`, `list-group`, `badge bg-success`, `table-light`) with no Bootstrap loaded. The results table on the import page had white text on a white-bordered table. Both pages rebuilt on the app's own stylesheet.
- `base.html` carried a second copy of the stylesheet inline with different colour variables, fighting `style.css`. One stylesheet now.
- Export progress shows sections done and articles found instead of a fixed 75% bar.
- Import progress updates per article, not per file.
- Failed Load brands shows the reason on the page instead of an `alert()`.
- Optional "remember subdomain and email in this browser" (token never stored).
- `index_bk.html` removed.
