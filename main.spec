# -*- mode: python ; coding: utf-8 -*-
from PyInstaller.utils.hooks import collect_all

datas = []
binaries = []
hiddenimports = [
    "core",
    "core.config",
    "core.comparator",
    "core.excel_writer",
    "core.http_client",
    "core.product_matcher",
    "core.menu",
    "scrapers",
    "scrapers.amazon_scraper",
    "scrapers.aqua_scraper",
    "scrapers.furaipan_scraper",
    "scrapers.rakuten_scraper",
    "scrapers.yahoo_scraper",
    "scrapers.yodobashi_scraper",
]

for pkg in ["curl_cffi", "openpyxl", "pandas", "bs4", "dotenv"]:
    d, b, h = collect_all(pkg)
    datas += d
    binaries += b
    hiddenimports += h

a = Analysis(
    ["main.py"],
    pathex=["."],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[],
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
    name="main",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
)
