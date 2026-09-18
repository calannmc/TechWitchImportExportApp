# -*- mode: python ; coding: utf-8 -*-
# PyInstaller spec for TechWitch Zendesk Helper.
# Build:  pyinstaller --noconfirm --clean TechWitchImportExportApp.spec

import os
from PyInstaller.utils.hooks import collect_data_files, collect_submodules

block_cipher = None

hidden = []
hidden += collect_submodules("mammoth")
hidden += collect_submodules("docx")
hidden += collect_submodules("openpyxl")
hidden += ["waitress", "xlrd", "bs4", "html_to_docx", "zendesk_importer", "enhanced_zendesk_exporter"]

datas = [("templates", "templates"), ("static", "static")]
datas += collect_data_files("docx")     # default .docx template python-docx needs
datas += collect_data_files("mammoth")

icon = "static/favicon.ico" if os.path.exists("static/favicon.ico") else None

a = Analysis(
    ["app.py"],
    pathex=["."],
    binaries=[],
    datas=datas,
    hiddenimports=hidden,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    # Big optional pandas/numpy extras that the app never touches
    excludes=["matplotlib", "scipy", "IPython", "jupyter", "notebook", "pytest",
              "tkinter", "PyQt5", "PySide2", "numpy.f2py", "pandas.tests", "numpy.tests"],
    noarchive=False,
    optimize=0,
)
pyz = PYZ(a.pure)

exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.datas,
    [],
    name="TechWitchImportExportApp",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,          # UPX causes false positives in some AV scanners; leave off
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,       # keeps a window so errors are visible and the app can be closed
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon=icon,
)
