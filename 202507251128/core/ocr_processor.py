"""
OCR-Verarbeitung für PDFs
"""
import os
import re
from typing import Dict, List, Tuple, Optional, Any
from pathlib import Path
import pytesseract
from pdf2image import convert_from_path
from PIL import Image
import tempfile
import logging
import time
import json

# Logger für dieses Modul
logger = logging.getLogger(__name__)


class OCRProcessor:
    """Führt OCR auf PDF-Dateien aus und extrahiert Text"""

    def __init__(self):
        # Versuche Tesseract zu finden
        self._setup_tesseract()
        
        # Setze Poppler-Pfad
        self._setup_poppler()

    def _setup_tesseract(self):
        """Konfiguriert Tesseract OCR"""
        # Windows-spezifische Einstellungen für versteckte Konsolen
        if os.name == 'nt':
            os.environ['TESSERACT_DISABLE_DEBUG_CONSOLE'] = '1'
        
        # Basis-Verzeichnis für dependencies
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        dependencies_dir = os.path.join(base_dir, 'dependencies')
        
        # Prüfe dependencies Ordner zuerst
        tesseract_path = os.path.join(dependencies_dir, 'Tesseract-OCR', 'tesseract.exe')
        if os.path.exists(tesseract_path):
            pytesseract.pytesseract.tesseract_cmd = tesseract_path
            
            # Konfiguriere pytesseract für Windows ohne Konsolen-Fenster
            if os.name == 'nt':
                try:
                    import subprocess
                    
                    # Erstelle STARTUPINFO für versteckte Fenster
                    si = subprocess.STARTUPINFO()
                    si.dwFlags |= subprocess.STARTF_USESHOWWINDOW
                    si.wShowWindow = subprocess.SW_HIDE
                    
                    # Monkey-patch pytesseract.run_tesseract
                    original_run_tesseract = pytesseract.pytesseract.run_tesseract
                    
                    def run_tesseract_no_console(*args, **kwargs):
                        kwargs['startupinfo'] = si
                        return original_run_tesseract(*args, **kwargs)
                    
                    pytesseract.pytesseract.run_tesseract = run_tesseract_no_console
                    logger.debug("Tesseract für versteckte Konsolen konfiguriert")
                except Exception as e:
                    logger.debug(f"Konnte Tesseract-Konsolen nicht verstecken: {e}")
            
            logger.debug(f"Tesseract gefunden: {tesseract_path}")
            return
        
        # Prüfe Installation im Programmverzeichnis (nach Installation)
        program_paths = [
            os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), 'belegpilot', 'dependencies', 'Tesseract-OCR', 'tesseract.exe'),
            os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'), 'belegpilot', 'dependencies', 'Tesseract-OCR', 'tesseract.exe')
        ]
        
        for path in program_paths:
            if os.path.exists(path):
                pytesseract.pytesseract.tesseract_cmd = path
                logger.debug(f"Tesseract gefunden in Installationsverzeichnis: {path}")
                return
        
        # Standard-Pfade für Tesseract auf Windows
        tesseract_paths = [
            r"C:\Program Files\Tesseract-OCR\tesseract.exe",
            r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
            r"C:\Users\%USERNAME%\AppData\Local\Tesseract-OCR\tesseract.exe"
        ]

        for path in tesseract_paths:
            expanded_path = os.path.expandvars(path)
            if os.path.exists(expanded_path):
                pytesseract.pytesseract.tesseract_cmd = expanded_path
                logger.debug(f"Tesseract gefunden: {expanded_path}")
                return

        # Wenn nicht gefunden, hoffen wir dass es im PATH ist
        logger.warning("Tesseract nicht gefunden. Bitte im dependencies Ordner platzieren.")

    def _setup_poppler(self):
        """Konfiguriert Poppler-Pfad"""
        self.poppler_path = None
        
        # Basis-Verzeichnis für dependencies
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        dependencies_dir = os.path.join(base_dir, 'dependencies')
        
        # Prüfe dependencies Ordner zuerst
        poppler_path = os.path.join(dependencies_dir, 'poppler', 'bin')
        if os.path.exists(poppler_path):
            self.poppler_path = poppler_path
            logger.debug(f"Poppler gefunden: {poppler_path}")
            return
        
        # Prüfe Installation im Programmverzeichnis (nach Installation)
        program_paths = [
            os.path.join(os.environ.get('ProgramFiles', 'C:\\Program Files'), 'belegpilot', 'dependencies', 'poppler', 'bin'),
            os.path.join(os.environ.get('ProgramFiles(x86)', 'C:\\Program Files (x86)'), 'belegpilot', 'dependencies', 'poppler', 'bin')
        ]
        
        for path in program_paths:
            if os.path.exists(path):
                self.poppler_path = path
                logger.debug(f"Poppler gefunden in Installationsverzeichnis: {path}")
                return
        
        logger.error("Poppler nicht gefunden! Bitte im dependencies Ordner platzieren.")

    def _get_poppler_path(self):
        """Gibt den Poppler-Pfad zurück oder None"""
        return self.poppler_path

    def _preprocess_image_for_ocr(self, image: Image.Image) -> Image.Image:
        """Bereitet ein Bild für eine bessere OCR-Erkennung vor."""
        # Konvertiere zu Graustufen für bessere Kontraste
        return image.convert('L')

    def extract_text_from_pdf(self, pdf_path: str, language: str = 'deu') -> str:
        """
        Extrahiert Text aus einer PDF-Datei mittels OCR mit Vorverarbeitung
        """
        try:
            # WICHTIG: Normalisiere den Pfad für Windows
            pdf_path = os.path.normpath(pdf_path)
            
            # Prüfe ob Datei existiert
            if not os.path.exists(pdf_path):
                logger.error(f"PDF-Datei nicht gefunden: {pdf_path}")
                return ""
            
            # Warte kurz, falls Datei noch geschrieben wird
            max_retries = 3
            for i in range(max_retries):
                try:
                    # Versuche die Datei zu öffnen um sicherzustellen, dass sie nicht gesperrt ist
                    with open(pdf_path, 'rb') as f:
                        # Lese nur den Header um zu prüfen ob es eine gültige PDF ist
                        header = f.read(4)
                        if header != b'%PDF':
                            logger.error(f"Keine gültige PDF-Datei: {pdf_path}")
                            return ""
                    break
                except Exception as e:
                    if i < max_retries - 1:
                        logger.warning(f"Datei noch gesperrt, warte... ({i+1}/{max_retries})")
                        time.sleep(1)
                    else:
                        logger.error(f"Datei konnte nicht geöffnet werden: {e}")
                        return ""
            
            logger.debug(f"Starte OCR für: {pdf_path}")

            # Konvertiere PDF zu Bildern
            with tempfile.TemporaryDirectory() as temp_dir:
                poppler_path = self._get_poppler_path()
                
                # Prüfe ob Poppler existiert
                if not poppler_path or not os.path.exists(poppler_path):
                    logger.error(f"Poppler nicht gefunden. Bitte im dependencies Ordner platzieren.")
                    return ""
                
                # Konvertiere mit explizitem poppler_path Parameter
                images = convert_from_path(
                    pdf_path, 
                    dpi=300, 
                    poppler_path=poppler_path,
                    output_folder=temp_dir
                )

                all_text = []
                for i, image in enumerate(images):
                    logger.debug(f"OCR auf Seite {i+1}/{len(images)}")
                    # Wende Bildvorverarbeitung an
                    preprocessed_image = self._preprocess_image_for_ocr(image)
                    # OCR auf jeder Seite
                    text = pytesseract.image_to_string(preprocessed_image, lang=language)
                    all_text.append(f"--- Seite {i+1} ---\n{text}")

                result_text = "\n\n".join(all_text)
                logger.info(f"OCR abgeschlossen für {os.path.basename(pdf_path)}: {len(result_text)} Zeichen extrahiert")
                return result_text

        except Exception as e:
            logger.error(f"Fehler bei OCR für {pdf_path}: {e}", exc_info=True)
            return ""

    def extract_text_from_zone(self, pdf_path: str, page_num: int,
                              zone: Tuple[int, int, int, int],
                              language: str = 'deu') -> str:
        """
        Extrahiert Text aus einer bestimmten Zone einer PDF-Seite mit Bildvorverarbeitung.
        """
        try:
            # WICHTIG: Normalisiere den Pfad für Windows
            pdf_path = os.path.normpath(pdf_path)
            
            # Prüfe ob Datei existiert
            if not os.path.exists(pdf_path):
                logger.error(f"PDF-Datei nicht gefunden für Zone-OCR: {pdf_path}")
                return ""
            
            logger.debug(f"OCR-Zone-Extraktion: {pdf_path}, Seite {page_num}, Zone {zone}")

            # Konvertiere spezifische Seite
            poppler_path = self._get_poppler_path()
            
            # Prüfe ob Poppler existiert
            if not poppler_path or not os.path.exists(poppler_path):
                logger.error(f"Poppler nicht gefunden. Bitte im dependencies Ordner platzieren.")
                return ""
            
            images = convert_from_path(
                pdf_path, 
                dpi=300,
                first_page=page_num, 
                last_page=page_num, 
                poppler_path=poppler_path
            )

            if not images:
                logger.warning(f"Keine Bilder aus PDF-Seite {page_num} konvertiert")
                return ""

            image = images[0]

            # Schneide Zone aus
            x, y, w, h = zone
            cropped = image.crop((x, y, x + w, y + h))

            # Wende Bildvorverarbeitung an
            preprocessed_cropped = self._preprocess_image_for_ocr(cropped)

            # OCR auf vorverarbeiteter Zone mit optimierter Konfiguration
            custom_config = r'--oem 3 --psm 6'
            text = pytesseract.image_to_string(preprocessed_cropped, lang=language, config=custom_config)
            result = text.strip()

            logger.debug(f"Zone-OCR Ergebnis: '{result[:50]}...' ({len(result)} Zeichen)")
            return result

        except Exception as e:
            logger.error(f"Fehler bei Zone OCR für {os.path.basename(pdf_path)}, Seite {page_num}: {e}", exc_info=True)
            return ""