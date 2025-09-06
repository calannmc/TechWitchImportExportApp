# TechWitch Import/Export App

A desktop tool for importing and exporting Zendesk Help Center articles.

👉 **[Download the latest release here](../../releases/latest)**

---

## Features
- Import articles from Excel/Word into specific Zendesk sections
- Export articles by brand into Excel or Word/HTML
- Runs locally at `http://127.0.0.1:5000`
- No Python required — just run the `.exe`

---

## How to Use
1. Download the ZIP from [Releases](../../releases/latest).
2. Extract it and run `TechWitchImportExportApp.exe`.
3. Your browser will open automatically (or visit `http://127.0.0.1:5000`).

---

## Verification
Each release includes a SHA-256 checksum.  
You can verify your download with PowerShell:

```powershell
Get-FileHash TechWitchImportExportApp-1.0.0.zip -Algorithm SHA256
