# TechWitch Zendesk Helper

Export Zendesk Help Centre articles to Excel or Word, and import them back from Excel or Word. Runs on your own machine, in your browser, on http://127.0.0.1:5000. Credentials are sent to Zendesk only and never written to disk.

## Run it

**Windows, no Python needed:** download `TechWitchImportExportApp.exe` from [Releases](../../releases/latest) and double-click it. A console window opens, then your browser. Keep the console open while you work. Press Quit in the app (top right) or close the console to stop. Exports land in an `exports` folder next to the exe.

**Windows, from source:** double-click `run.bat`. First run creates a `.venv` and installs the dependencies (a minute or so). After that it starts straight away.

**Mac / Linux:** `./run.sh`

Python 3.10 or newer for the source route.

## Build the exe

Two ways.

1. Locally on Windows: double-click `build_exe.bat`. Output: `dist\TechWitchImportExportApp.exe`.
2. On GitHub: push this folder to a repo. The workflow in `.github/workflows/build.yml` builds the exe on a Windows runner on every push to main, uploads it as a build artifact, and attaches it to a release when you push a tag like `v1.1.0`.

The exe is around 55 MB because pandas and numpy come along. First start takes a few seconds while it unpacks. Each release carries a `SHA256SUMS.txt`; check a download with `Get-FileHash TechWitchImportExportApp.exe -Algorithm SHA256` in PowerShell.

If Windows SmartScreen complains, click More info, then Run anyway. The exe is not code-signed.

## Using it

1. Enter subdomain (just `acme`, or paste the full URL, both work), email and API token. Press Load brands.
2. **Export:** pick Excel or Word, optionally limit to a brand, Start Export. Excel gives one row per article with the HTML body in the `body` column. Word gives a zip with one `.docx` per article plus `index.xlsx` listing ids, sections and labels.
3. **Import:** pick brand and section, choose files, Start Import. Progress and per-article results show on the next page with links to each article.

Import file rules:

- Excel columns: `id`, `title`, `body`, `locale`, `labels`, `section_id`, `draft`, `promoted`. Column names are case-insensitive and a few aliases are accepted (`article_id`, `label_names`, `content`). An export file imports as-is.
- A `section_id` on a row overrides the section chosen on the page.
- A row's `locale` wins. The Default locale on the page is used only for rows without one, and for Word files.
- Word files become one article each. The document Title style (which the exporter writes), else the first heading, else the first bold line, becomes the article title. Everything else becomes the body with headings, lists, links, tables and bold/italic kept.
- Update if existing matches by `id`, then exact title, then overlapping labels. Labels merge or replace as chosen.
- Create as drafts forces draft on. Set it to No to honour the row's `draft` column, or publish.

## Files

| File | What it is |
|---|---|
| `app.py` | Flask app and routes |
| `enhanced_zendesk_exporter.py` | Export logic (brand aware) |
| `zendesk_importer.py` | Import logic and file parsing |
| `html_to_docx.py` | HTML body to .docx converter used by the Word export |
| `templates/`, `static/` | UI |
| `requirements.txt` | Runtime deps. `requirements-build.txt` adds PyInstaller |
| `TechWitchImportExportApp.spec` | PyInstaller build spec |
| `run.bat`, `run.sh`, `build_exe.bat` | Launchers |
| `Procfile` | For hosting on Heroku-style platforms with waitress |

## Environment variables (optional)

`PORT` (default 5000, falls back to a free port if taken), `NO_BROWSER=1` to skip opening the browser, `ZENDESK_SUBDOMAIN` / `ZENDESK_EMAIL` / `ZENDESK_API_TOKEN` to prefill credentials server-side, `LOG_LEVEL`.
