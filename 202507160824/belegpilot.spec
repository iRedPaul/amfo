# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

# Basis-Konfiguration
block_cipher = None

# Sammle alle notwendigen Module
hiddenimports = [
    # Core Module
    'core.hotfolder_manager',
    'core.config_manager',
    'core.file_watcher',
    'core.pdf_processor',
    'core.logging_config',
    'core.ocr_processor',
    'core.export_processor',
    'core.xml_field_processor',
    'core.function_parser',
    'core.license_manager',
    'core.oauth2_manager',
    
    # GUI Module
    'gui.main_window',
    'gui.hotfolder_dialog',
    'gui.expression_editor_base',
    'gui.settings_dialog',
    'gui.database_config_dialog',
    'gui.ocr_zone_editor',
    
    # Models
    'models.hotfolder_config',
    'models.export_config',
    
    # Externe Abhängigkeiten
    'watchdog',
    'watchdog.observers',
    'watchdog.events',
    'PyPDF2',
    'fitz',  # PyMuPDF
    'PIL',
    'PIL.Image',
    'pytesseract',
    'ocrmypdf',
    'reportlab',
    'xml.etree',
    'xml.etree.ElementTree',
    'keyring',
    'cryptography',
    'pyodbc',
    'openpyxl',
    'python-docx',
    'lxml',
    'dateutil',
    'requests',
    'smtplib',
    'email',
    'email.mime',
    'email.mime.multipart',
    'email.mime.text',
    'email.mime.base',
    
    # Tkinter Module
    'tkinter',
    'tkinter.ttk',
    'tkinter.messagebox',
    'tkinter.filedialog',
    'tkinter.simpledialog',
    
    # Standard Module
    'json',
    'uuid',
    'datetime',
    'logging',
    'tempfile',
    'subprocess',
    'threading',
    'pathlib',
    'shutil',
    'glob',
    'time',
    'os',
    'sys',
    're',
]

# Daten-Dateien
datas = [    
    # Icon
    ('gui/assets/icon.ico', '.'),
    
]

# Binäre Dateien
binaries = []

# Sammle alle PyMuPDF Daten
tmp_ret = collect_all('fitz')
datas += tmp_ret[0]
binaries += tmp_ret[1]
hiddenimports += tmp_ret[2]

# Sammle OCRmyPDF Daten
tmp_ret = collect_all('ocrmypdf')
datas += tmp_ret[0]
binaries += tmp_ret[1]
hiddenimports += tmp_ret[2]

# Sammle reportlab Daten
tmp_ret = collect_data_files('reportlab')
datas += tmp_ret

# Analyse-Einstellungen
a = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib',
        'numpy',
        'pandas',
        'scipy',
        'IPython',
        'jupyter',
        'notebook',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)

# Erstelle PYZ-Archiv
pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

# EXE-Einstellungen für SINGLE FILE
exe = EXE(
    pyz,
    a.scripts,
    a.binaries,
    a.zipfiles,
    a.datas,
    [],
    name='belegpilot',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=False,  # Kein Konsolenfenster für GUI-Anwendung
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='gui/assets/icon.ico',  # Icon hinzugefügt
    version_file=None,
    uac_admin=True,  # Admin-Rechte anfordern
    uac_uiaccess=False,
)