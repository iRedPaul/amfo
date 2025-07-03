# -*- coding: utf-8 -*-
"""
Nuitka Build-Konfiguration für Hotfolder PDF Processor
"""

# Versionsinformationen
COMPANY_NAME = "Ihre Firma GmbH"
PRODUCT_NAME = "Hotfolder PDF Processor"
VERSION = "1.0.0"
COPYRIGHT = f"© 2024 {COMPANY_NAME}"

# Gemeinsame Nuitka-Optionen
COMMON_OPTIONS = [
    # Standalone-Anwendung
    '--standalone',
    
    # Python-Version
    '--python-flag=no_site',
    '--python-flag=no_warnings',
    
    # Optimierungen
    '--assume-yes-for-downloads',
    '--enable-plugin=tk-inter',
    # '--enable-plugin=multiprocessing',  # Entfernt - ist automatisch aktiviert
    # '--enable-plugin=numpy',  # Entfernt - deprecated
    
    # Windows-spezifisch
    '--windows-company-name=' + COMPANY_NAME,
    '--windows-product-name=' + PRODUCT_NAME,
    '--windows-file-version=' + VERSION,
    '--windows-product-version=' + VERSION,
    '--windows-file-description=Automatische PDF-Verarbeitung',
    
    # Folge allen Imports
    '--follow-imports',
    
    # Include Pfade
    '--include-package=gui',
    '--include-package=core',
    '--include-package=models',
    
    # Standard-Module
    '--include-package=tkinter',
    '--include-package=PIL',
    '--include-package=watchdog',
    '--include-package=PyPDF2',
    '--include-package=reportlab',
    '--include-package=pytesseract',
    '--include-package=pdf2image',
    '--include-package=lxml',
    '--include-package=win32com',
    '--include-package=win32service',
    '--include-package=win32serviceutil',
    '--include-package=servicemanager',
    
    # Optionale Module (falls vorhanden)
    '--include-package=ocrmypdf',
    '--include-package=pikepdf',
    '--include-package=cryptography',
    '--include-package=keyring',
    '--include-package=fitz',
    
    # Daten-Dateien
    '--include-data-file=settings.json=settings.json',
    '--include-data-file=counters.json=counters.json',
    '--include-data-file=hotfolders.json=hotfolders.json',
    
    # Poppler-Verzeichnis
    '--include-data-dir=poppler=poppler',
    
    # Optimierungen
    '--remove-output',
    '--no-pyi-file',
]

# Icon falls vorhanden
import os
if os.path.exists('icon.ico'):
    COMMON_OPTIONS.extend([
        '--windows-icon-from-ico=icon.ico',
        '--include-data-file=icon.ico=icon.ico'
    ])

# Build-Definitionen
NUITKA_BUILDS = [
    {
        'name': 'Hauptanwendung',
        'options': COMMON_OPTIONS + [
            '--output-filename=Hotfolder PDF Processor.exe',
            '--windows-console-mode=disable',  # GUI-Anwendung
            '--onefile',  # Single-File-Executable
            'main.py'
        ]
    },
    {
        'name': 'Windows Service',
        'options': COMMON_OPTIONS + [
            '--output-filename=HotfolderPDFService.exe',
            '--windows-console-mode=attach',  # Service-Konsole
            '--onefile',  # Single-File-Executable
            'windows_service.py'
        ]
    }
]