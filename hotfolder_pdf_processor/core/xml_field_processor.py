"""
XML-Feldverarbeitung für SSDS-Format
"""
import os
import re
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Any, Tuple
from datetime import datetime
from pathlib import Path
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.ocr_processor import OCRProcessor


class FieldMapping:
    """Definiert wie ein XML-Feld befüllt werden soll"""
    
    def __init__(self, field_name: str, source_type: str = "ocr_text", 
                 pattern: str = "", zone: Optional[Tuple[int, int, int, int]] = None,
                 page_num: int = 1):
        self.field_name = field_name
        self.source_type = source_type  # "ocr_text", "ocr_zone", "fixed", "date", "filename"
        self.pattern = pattern  # Regulärer Ausdruck
        self.zone = zone  # (x, y, width, height) für OCR-Zone
        self.page_num = page_num
    
    def to_dict(self) -> dict:
        """Konvertiert zu Dictionary für Speicherung"""
        return {
            "field_name": self.field_name,
            "source_type": self.source_type,
            "pattern": self.pattern,
            "zone": self.zone,
            "page_num": self.page_num
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'FieldMapping':
        """Erstellt aus Dictionary"""
        return cls(
            field_name=data["field_name"],
            source_type=data.get("source_type", "ocr_text"),
            pattern=data.get("pattern", ""),
            zone=data.get("zone"),
            page_num=data.get("page_num", 1)
        )


class XMLFieldProcessor:
    """Verarbeitet und aktualisiert XML-Felder basierend auf OCR-Daten"""
    
    def __init__(self):
        self.ocr_processor = OCRProcessor()
        self._ocr_cache = {}  # Cache für OCR-Ergebnisse
    
    def process_xml_with_mappings(self, xml_path: str, pdf_path: str, 
                                  mappings: List[FieldMapping]) -> bool:
        """
        Verarbeitet eine XML-Datei mit den definierten Feld-Mappings
        
        Args:
            xml_path: Pfad zur XML-Datei
            pdf_path: Pfad zur zugehörigen PDF-Datei
            mappings: Liste der Feld-Mappings
            
        Returns:
            True wenn erfolgreich
        """
        try:
            # Parse XML
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # Finde Document-Element
            doc_elem = root.find(".//Document")
            if doc_elem is None:
                print("Kein Document-Element in XML gefunden")
                return False
            
            # Finde Fields-Element
            fields_elem = doc_elem.find("Fields")
            if fields_elem is None:
                print("Kein Fields-Element in XML gefunden")
                return False
            
            # Cache für OCR-Text
            ocr_full_text = None
            
            # Verarbeite jedes Mapping
            for mapping in mappings:
                try:
                    value = self._extract_value(pdf_path, mapping, ocr_full_text)
                    
                    if value:
                        # Finde oder erstelle Feld
                        field_elem = fields_elem.find(mapping.field_name)
                        if field_elem is None:
                            # Erstelle neues Feld
                            field_elem = ET.SubElement(fields_elem, mapping.field_name)
                        
                        # Setze Wert
                        field_elem.text = value
                        print(f"Feld '{mapping.field_name}' gesetzt auf: {value}")
                    
                except Exception as e:
                    print(f"Fehler bei Mapping für Feld '{mapping.field_name}': {e}")
            
            # Speichere XML
            self._indent_xml(root)
            tree.write(xml_path, encoding='utf-8', xml_declaration=True)
            
            return True
            
        except Exception as e:
            print(f"Fehler bei XML-Verarbeitung: {e}")
            return False
    
    def _extract_value(self, pdf_path: str, mapping: FieldMapping, 
                      ocr_full_text: Optional[str] = None) -> Optional[str]:
        """Extrahiert einen Wert basierend auf dem Mapping"""
        
        if mapping.source_type == "fixed":
            return ""
        
        elif mapping.source_type == "date":
            user_format = mapping.pattern.strip()
            import re as _re
            # Merke, ob d oder m (einstellig) im Format vorkommt
            single_day = False
            single_month = False
            def mark_token(match):
                token = match.group(0)
                if token == "yyyy": return "%Y"
                if token == "yy": return "%y"
                if token == "mm": return "%m"
                if token == "dd": return "%d"
                if token == "m":
                    nonlocal single_month
                    single_month = True
                    return "%m"
                if token == "d":
                    nonlocal single_day
                    single_day = True
                    return "%d"
                return token
            py_format = _re.sub(r"yyyy|yy|mm|dd|m|d", mark_token, user_format)
            try:
                result = datetime.now().strftime(py_format)
                # Entferne führende Null bei Tag/Monat, falls gewünscht
                if single_day or single_month:
                    # Ersetze nur die passenden Stellen
                    # Finde alle Positionen von %d und %m im Formatstring
                    now = datetime.now()
                    day = str(now.day)
                    month = str(now.month)
                    # Ersetze jeweils nur die erste passende Null
                    if single_day:
                        result = _re.sub(r'\b0?'+day+r'\b', day, result, count=1)
                    if single_month:
                        result = _re.sub(r'\b0?'+month+r'\b', month, result, count=1)
                return result
            except Exception as e:
                print(f"Ungültiges Datumsformat: {user_format} → {py_format}: {e}")
                return ""
        
        elif mapping.source_type == "filename":
            filename = os.path.basename(pdf_path)
            if mapping.pattern:
                matches = re.search(mapping.pattern, filename)
                if matches:
                    return matches.group(1) if matches.groups() else matches.group(0)
            return os.path.splitext(filename)[0]
        
        elif mapping.source_type == "ocr_text":
            # Hole OCR-Text (cached)
            if pdf_path not in self._ocr_cache:
                print(f"Führe OCR aus auf: {pdf_path}")
                self._ocr_cache[pdf_path] = self.ocr_processor.extract_text_from_pdf(pdf_path)
            
            ocr_text = self._ocr_cache[pdf_path]
            
            if mapping.pattern:
                matches = re.search(mapping.pattern, ocr_text, re.MULTILINE | re.IGNORECASE)
                if matches:
                    return matches.group(1) if matches.groups() else matches.group(0)
            
            return None
        
        elif mapping.source_type == "ocr_zone":
            if mapping.zone:
                text = self.ocr_processor.extract_text_from_zone(
                    pdf_path, mapping.page_num, mapping.zone
                )
                
                if mapping.pattern and text:
                    matches = re.search(mapping.pattern, text, re.MULTILINE | re.IGNORECASE)
                    if matches:
                        return matches.group(1) if matches.groups() else matches.group(0)
                
                return text.strip() if text else None
        
        return None
    
    def _indent_xml(self, elem, level=0):
        """Formatiert XML mit Einrückungen"""
        i = "\n" + level * "  "
        if len(elem):
            if not elem.text or not elem.text.strip():
                elem.text = i + "  "
            if not elem.tail or not elem.tail.strip():
                elem.tail = i
            for elem in elem:
                self._indent_xml(elem, level + 1)
            if not elem.tail or not elem.tail.strip():
                elem.tail = i
        else:
            if level and (not elem.tail or not elem.tail.strip()):
                elem.tail = i
    
    def get_available_fields(self, xml_path: str) -> List[str]:
        """Gibt alle verfügbaren Felder in einer XML-Datei zurück"""
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            fields = []
            fields_elem = root.find(".//Fields")
            if fields_elem is not None:
                for field in fields_elem:
                    fields.append(field.tag)
            
            return sorted(fields)
            
        except Exception as e:
            print(f"Fehler beim Lesen der XML-Felder: {e}")
            return []
    
    def clear_ocr_cache(self):
        """Leert den OCR-Cache"""
        self._ocr_cache.clear()