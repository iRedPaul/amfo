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

# Logger für dieses Modul
logger = logging.getLogger(__name__)


class OCRProcessor:
    """Führt OCR auf PDF-Dateien aus und extrahiert Text"""
    
    def __init__(self):
        # Versuche Tesseract zu finden
        self._setup_tesseract()
        
    def _setup_tesseract(self):
        """Konfiguriert Tesseract OCR"""
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
        logger.warning("Tesseract nicht in Standard-Pfaden gefunden. Stelle sicher, dass es installiert ist.")
    
    def extract_text_from_pdf(self, pdf_path: str, language: str = 'deu') -> str:
        """
        Extrahiert Text aus einer PDF-Datei mittels OCR
        
        Args:
            pdf_path: Pfad zur PDF-Datei
            language: OCR-Sprache (deu für Deutsch, eng für Englisch)
            
        Returns:
            Extrahierter Text
        """
        try:
            logger.debug(f"Starte OCR für: {pdf_path}")
            
            # Konvertiere PDF zu Bildern
            with tempfile.TemporaryDirectory() as temp_dir:
                poppler_path = os.path.join(os.path.dirname(__file__), '..', 'poppler', 'bin')
                images = convert_from_path(pdf_path, dpi=300, poppler_path=poppler_path)
                
                all_text = []
                for i, image in enumerate(images):
                    logger.debug(f"OCR auf Seite {i+1}/{len(images)}")
                    # OCR auf jeder Seite
                    text = pytesseract.image_to_string(image, lang=language)
                    all_text.append(f"--- Seite {i+1} ---\n{text}")
                
                result_text = "\n\n".join(all_text)
                logger.info(f"OCR abgeschlossen für {pdf_path}: {len(result_text)} Zeichen extrahiert")
                return result_text
                
        except Exception as e:
            logger.error(f"Fehler bei OCR für {pdf_path}: {e}", exc_info=True)
            return ""
    
    def extract_text_from_zone(self, pdf_path: str, page_num: int, 
                             zone: Tuple[int, int, int, int], 
                             language: str = 'deu') -> str:
        """
        Extrahiert Text aus einer bestimmten Zone einer PDF-Seite
        
        Args:
            pdf_path: Pfad zur PDF-Datei
            page_num: Seitennummer (1-basiert)
            zone: Tuple (x, y, width, height) in Pixeln
            language: OCR-Sprache
            
        Returns:
            Extrahierter Text aus der Zone
        """
        try:
            logger.debug(f"OCR-Zone-Extraktion: {pdf_path}, Seite {page_num}, Zone {zone}")
            
            # Konvertiere spezifische Seite
            poppler_path = os.path.join(os.path.dirname(__file__), '..', 'poppler', 'bin')
            images = convert_from_path(pdf_path, dpi=300, 
                                     first_page=page_num, last_page=page_num, poppler_path=poppler_path)
            
            if not images:
                logger.warning(f"Keine Bilder aus PDF-Seite {page_num} konvertiert")
                return ""
            
            image = images[0]
            
            # Schneide Zone aus
            x, y, w, h = zone
            cropped = image.crop((x, y, x + w, y + h))
            
            # OCR auf Zone
            text = pytesseract.image_to_string(cropped, lang=language)
            result = text.strip()
            
            logger.debug(f"Zone-OCR Ergebnis: '{result[:50]}...' ({len(result)} Zeichen)")
            return result
            
        except Exception as e:
            logger.error(f"Fehler bei Zone OCR für {pdf_path}, Seite {page_num}: {e}", exc_info=True)
            return ""
    
    def find_text_by_pattern(self, text: str, pattern: str) -> List[str]:
        """
        Findet Text mittels regulärem Ausdruck
        
        Args:
            text: Zu durchsuchender Text
            pattern: Regulärer Ausdruck
            
        Returns:
            Liste der gefundenen Übereinstimmungen
        """
        try:
            matches = re.findall(pattern, text, re.MULTILINE | re.IGNORECASE)
            logger.debug(f"Pattern '{pattern}' gefunden: {len(matches)} Treffer")
            return matches
        except re.error as e:
            logger.error(f"Fehler im regulären Ausdruck '{pattern}': {e}")
            return []
    
    def extract_invoice_data(self, text: str) -> Dict[str, str]:
        """
        Extrahiert typische Rechnungsdaten aus Text
        
        Returns:
            Dictionary mit gefundenen Daten
        """
        data = {}
        
        # Rechnungsnummer
        patterns = {
            'invoice_number': [
                r'Rechnungsnummer[:\s]+([A-Z0-9\-/]+)',
                r'Rechnung\s*Nr[.:\s]+([A-Z0-9\-/]+)',
                r'Invoice\s*Number[:\s]+([A-Z0-9\-/]+)'
            ],
            'date': [
                r'Datum[:\s]+(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})',
                r'Rechnungsdatum[:\s]+(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})',
                r'Date[:\s]+(\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4})'
            ],
            'amount': [
                r'Gesamtbetrag[:\s]+([0-9.,]+)\s*€',
                r'Total[:\s]+([0-9.,]+)\s*€',
                r'Summe[:\s]+([0-9.,]+)\s*€',
                r'€\s*([0-9.,]+)',
                r'EUR\s*([0-9.,]+)'
            ],
            'tax_id': [
                r'USt[\-\.]?IdNr[.:\s]+([A-Z]{2}[A-Z0-9]+)',
                r'VAT[:\s]+([A-Z]{2}[A-Z0-9]+)',
                r'UID[:\s]+([A-Z]{2}[A-Z0-9]+)'
            ]
        }
        
        for field, pattern_list in patterns.items():
            for pattern in pattern_list:
                matches = re.search(pattern, text, re.IGNORECASE)
                if matches:
                    data[field] = matches.group(1).strip()
                    logger.debug(f"Rechnungsfeld '{field}' gefunden: {data[field]}")
                    break
        
        logger.info(f"Rechnungsdaten extrahiert: {len(data)} Felder gefunden")
        return data
    
    def extract_all_numbers(self, text: str) -> List[str]:
        """Extrahiert alle Zahlen aus dem Text"""
        numbers = re.findall(r'\b\d+[.,]?\d*\b', text)
        logger.debug(f"Zahlen extrahiert: {len(numbers)} gefunden")
        return numbers
    
    def extract_all_dates(self, text: str) -> List[str]:
        """Extrahiert alle Datumsangaben aus dem Text"""
        date_patterns = [
            r'\d{1,2}[.\-/]\d{1,2}[.\-/]\d{2,4}',
            r'\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2}',
            r'\d{1,2}\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[a-z]*\s+\d{2,4}'
        ]
        
        dates = []
        for pattern in date_patterns:
            dates.extend(re.findall(pattern, text, re.IGNORECASE))
        
        logger.debug(f"Datumsangaben extrahiert: {len(dates)} gefunden")
        return dates