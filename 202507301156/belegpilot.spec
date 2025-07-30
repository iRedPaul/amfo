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
    'pypdf',
    'fitz',
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
    'win32timezone'
]

# Daten-Dateien
datas = [    
    # Icon (ICO)
    ('gui/assets/icon.ico', 'gui/assets'),
    ('gui/assets/icon.png', 'gui/assets'),
    ('gui/assets/banner.png', 'gui/assets')
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

# --- ERSTE EXE: Der Windows-Dienst ---
a_service = Analysis(
    ['belegpilot_service.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy', 'IPython', 'jupyter', 'notebook',
        'tkinter', 'gui' # GUI-Module werden für den Dienst nicht benötigt
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz_service = PYZ(a_service.pure, a_service.zipped_data, cipher=block_cipher)

exe_service = EXE(
    pyz_service,
    a_service.scripts,
    [],
    exclude_binaries=True,
    name='belegpilot_service',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='gui/assets/icon.ico',
    uac_admin=True
)

coll_service = COLLECT(
    exe_service,
    a_service.binaries,
    a_service.zipfiles,
    a_service.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='belegpilot_service'
)

# --- ZWEITE EXE: Das Programm (GUI) ---
a_gui = Analysis(
    ['main.py'],
    pathex=[],
    binaries=binaries,
    datas=datas,
    hiddenimports=hiddenimports,
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=[
        'matplotlib', 'numpy', 'pandas', 'scipy', 'IPython', 'jupyter', 'notebook',
    ],
    win_no_prefer_redirects=False,
    win_private_assemblies=False,
    cipher=block_cipher,
    noarchive=False,
)
pyz_gui = PYZ(a_gui.pure, a_gui.zipped_data, cipher=block_cipher)

exe_gui = EXE(
    pyz_gui,
    a_gui.scripts,
    [],
    exclude_binaries=True,
    name='belegpilot',
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=False,
    disable_windowed_traceback=False,
    target_arch=None,
    codesign_identity=None,
    entitlements_file=None,
    icon='gui/assets/icon.ico',
    uac_admin=True
)

coll_gui = COLLECT(
    exe_gui,
    a_gui.binaries,
    a_gui.zipfiles,
    a_gui.datas,
    strip=False,
    upx=False,
    upx_exclude=[],
    name='belegpilot'
)