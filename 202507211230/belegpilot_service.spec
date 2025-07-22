# -*- mode: python ; coding: utf-8 -*-

import sys
import os
from PyInstaller.utils.hooks import collect_all, collect_data_files, collect_submodules

# Basis-Konfiguration
block_cipher = None

# Sammle alle notwendigen Module
hiddenimports = [
    # Windows Service spezifisch
    'win32timezone',
    'win32service',
    'win32serviceutil',
    'win32event',
    'win32evtlog',
    'win32evtlogutil',
    'servicemanager',
    'win32com',
    'win32com.client',
    'pythoncom',
    
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
    ('gui/assets/icon.ico', 'gui/assets'),
]

# Prüfe welche Verzeichnisse existieren und füge sie hinzu
if os.path.exists('config'):
    datas.append(('config', 'config'))

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
    ['windows_service.py'],
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
        'tkinter',  # Service braucht kein GUI
        'tkinter.ttk',
        'tkinter.messagebox',
        'tkinter.filedialog',
        'tkinter.simpledialog',
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
    name='belegpilot_service',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=True,
    upx_exclude=[],
    runtime_tmpdir=None,
    console=True,  # Service braucht Konsole für Status-Ausgaben
    disable_windowed_traceback=False,
    argv_emulation=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='gui/assets/icon.ico',
    version_file=None,
    uac_admin=True,  # Admin-Rechte für Service-Installation
    uac_uiaccess=False,
)
