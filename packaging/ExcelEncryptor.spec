# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller yapılandırması.

  Windows : tek dosya  ExcelEncryptor.exe   (konsol penceresi yok)
  macOS   : Excel Encryptor.app             (Apple Silicon)

Çalıştırma:  pyinstaller packaging/ExcelEncryptor.spec --noconfirm
"""

import os
import sys

ROOT = os.path.abspath(os.path.join(SPECPATH, ".."))
SRC = os.path.join(ROOT, "src")
IS_MAC = sys.platform == "darwin"

icon = os.path.join(ROOT, "assets", "icon.icns" if IS_MAC else "icon.ico")
if not os.path.exists(icon):
    icon = None

a = Analysis(
    [os.path.join(SRC, "app.py")],
    pathex=[SRC],
    binaries=[],
    datas=[],
    hiddenimports=["tksheet", "openpyxl", "defusedxml"],
    hookspath=[],
    runtime_hooks=[],
    # Saldırı yüzeyini ve boyutu küçültmek için gereksiz modülleri dışarıda bırak
    excludes=[
        "numpy", "pandas", "matplotlib", "scipy", "PIL", "pytest",
        "setuptools", "pip", "test", "unittest", "pydoc", "doctest",
        "http.server", "xmlrpc", "ftplib", "telnetlib", "smtplib",
    ],
    noarchive=False,
)
pyz = PYZ(a.pure)

if IS_MAC:
    exe = EXE(
        pyz, a.scripts, [],
        exclude_binaries=True,
        name="Excel Encryptor",
        console=False,
        icon=icon,
    )
    coll = COLLECT(exe, a.binaries, a.datas, name="Excel Encryptor")
    app = BUNDLE(
        coll,
        name="Excel Encryptor.app",
        icon=icon,
        bundle_identifier="dev.asimogg.excelencryptor",
        info_plist={
            "CFBundleName": "Excel Encryptor",
            "CFBundleDisplayName": "Excel Encryptor",
            "CFBundleShortVersionString": "2.0",
            "CFBundleVersion": "2.0",
            "NSHighResolutionCapable": True,
            "LSMinimumSystemVersion": "12.0",
            # Sadece kullanıcının seçtiği dosyalara erişir; ağ kullanmaz.
            "NSHumanReadableCopyright": "MIT License",
        },
    )
else:
    exe = EXE(
        pyz, a.scripts, a.binaries, a.datas, [],
        name="ExcelEncryptor",
        console=False,          # pencereli uygulama — konsol açılmaz
        upx=False,              # UPX sıkıştırması antivirüs yanlış-pozitifi üretir
        icon=icon,
    )
