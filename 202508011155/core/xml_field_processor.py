"""
XML-Feldverarbeitung mit erweiterter Funktionsunterstützung
"""
import os
import re
import xml.etree.ElementTree as ET
from typing import Dict, List, Optional, Any, Tuple
from datetime import datetime
from pathlib import Path
import sys
import logging

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.ocr_processor import OCRProcessor
from core.function_parser import FunctionParser, VariableExtractor

# Logger für dieses Modul
logger = logging.getLogger(__name__)


class FieldMapping:
    """Definiert wie ein XML-Feld befüllt werden soll"""
    
    def __init__(self, field_name: str, source_type: str = "expression", 
                 expression: str = "", zone: Optional[Tuple[int, int, int, int]] = None,
                 page_num: int = 1, zones: Optional[List[Dict]] = None):
        self.field_name = field_name
        self.source_type = source_type  # "expression", "ocr_zone"
        self.expression = expression  # Funktionsausdruck
        self.zone = zone  # (x, y, width, height) für OCR-Zone (legacy)
        self.page_num = page_num
        self.zones = zones or []  # Liste von OCR-Zonen für Multi-Zone Support
    
    def to_dict(self) -> dict:
        """Konvertiert zu Dictionary für Speicherung"""
        return {
            "field_name": self.field_name,
            "source_type": self.source_type,
            "expression": self.expression,
            "zone": self.zone,
            "page_num": self.page_num,
            "zones": self.zones
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'FieldMapping':
        """Erstellt aus Dictionary"""
        return cls(
            field_name=data["field_name"],
            source_type=data.get("source_type", "expression"),
            expression=data.get("expression", ""),
            zone=data.get("zone"),
            page_num=data.get("page_num", 1),
            zones=data.get("zones", [])
        )


class XMLFieldProcessor:
    """Verarbeitet und aktualisiert XML-Felder basierend auf Funktionen und Variablen"""
    
    def __init__(self):
        self.ocr_processor = OCRProcessor()
        self.function_parser = FunctionParser()
        self._ocr_cache = {}  # Cache für OCR-Ergebnisse
        self._zone_cache = {}  # Cache für OCR-Zonen
    
    def process_xml_with_mappings(self, xml_path: str = "", pdf_path: str = "", 
                                  mappings: List[FieldMapping] = [], 
                                  ocr_zones: List[Dict] = [],
                                  input_path: str = "",
                                  original_pdf_path: str = "") -> bool:
        """
        Verarbeitet eine XML-Datei mit den definierten Feld-Mappings
        
        Args:
            xml_path: Pfad zur XML-Datei
            pdf_path: Pfad zur zugehörigen PDF-Datei (temporär)
            mappings: Liste der Feld-Mappings
            ocr_zones: Liste der OCR-Zonen vom Hotfolder
            input_path: Pfad zum Hotfolder (für Level-Variablen)
            original_pdf_path: Original-Pfad der PDF (für Level-Variablen)
            
        Returns:
            True wenn erfolgreich
        """
        try:
            # NEU: Prüfe auf zirkuläre Abhängigkeiten
            has_cycle, error_msg = self._check_circular_dependencies(mappings)
            if has_cycle:
                logger.error(f"Verarbeitung abgebrochen: {error_msg}")
                return False
            
            # Parse XML
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # Finde Document-Element
            doc_elem = root.find(".//Document")
            if doc_elem is None:
                logger.warning("Kein Document-Element in XML gefunden")
                return False
            
            # Finde Fields-Element
            fields_elem = doc_elem.find("Fields")
            if fields_elem is None:
                logger.warning("Kein Fields-Element in XML gefunden")
                return False
            
            # Sammle alle verfügbaren Variablen (ohne bereits evaluierte Felder)
            context = self._build_context(xml_path, pdf_path, mappings, ocr_zones, 
                                        input_path, original_pdf_path)
            
            # Dictionary für bereits evaluierte Felder
            evaluated_fields = {}
            
            # Verarbeite jedes Mapping in Reihenfolge
            for mapping in mappings:
                try:
                    # Füge bereits evaluierte Felder zum Kontext hinzu
                    context.update(evaluated_fields)
                    
                    # Extrahiere Wert basierend auf Expression
                    value = self._evaluate_mapping(mapping, context)
                    
                    # ÄNDERUNG: Erstelle auch Felder mit leerem Wert
                    # Speichere evaluierten Wert für nachfolgende Felder (auch wenn leer)
                    if value is not None:
                        evaluated_fields[mapping.field_name] = str(value)
                    else:
                        # Wenn None zurückgegeben wird, verwende leeren String
                        evaluated_fields[mapping.field_name] = ""
                    
                    # Finde oder erstelle Feld
                    field_elem = fields_elem.find(mapping.field_name)
                    if field_elem is None:
                        # Erstelle neues Feld
                        field_elem = ET.SubElement(fields_elem, mapping.field_name)
                    
                    # Setze Wert (kann auch leer sein)
                    if value is not None:
                        field_elem.text = str(value)
                    else:
                        field_elem.text = ""  # Explizit leeren Text setzen
                    
                    logger.info(f"Feld '{mapping.field_name}' gesetzt auf: '{value if value is not None else ''}'")
                    
                except Exception as e:
                    logger.error(f"Fehler bei Mapping für Feld '{mapping.field_name}': {e}")
            
            # Speichere XML
            self._indent_xml(root)
            tree.write(xml_path, encoding='utf-8', xml_declaration=True)
            
            # Leere Caches
            self._ocr_cache.clear()
            self._zone_cache.clear()
            
            return True
            
        except Exception as e:
            logger.error(f"Fehler bei XML-Verarbeitung: {e}")
            return False

    def _build_context(self, xml_path: str = "", pdf_path: str = "", 
                      mappings: List[FieldMapping] = [], 
                      ocr_zones: List[Dict] = [],
                      input_path: str = "",
                      original_pdf_path: str = "") -> Dict[str, Any]:
        """Baut den Kontext mit allen verfügbaren Variablen auf"""
        context = {}
        
        # Standard-Variablen
        context.update(VariableExtractor.get_standard_variables())
        
        # Datei-Variablen (vom temporären Pfad)
        context.update(VariableExtractor.get_file_variables(pdf_path))
        
        # NEU: Level-Variablen hinzufügen
        if input_path and original_pdf_path:
            # Verwende den Original-Pfad für Level-Variablen
            level_vars = VariableExtractor.get_level_variables(original_pdf_path, input_path)
            context.update(level_vars)
            context['InputPath'] = input_path
            
            # Debug-Ausgabe
            logger.debug(f"Level-Variablen aus Original-Pfad '{original_pdf_path}':")
            for i in range(6):
                level_key = f'level{i}'
                if level_key in level_vars:
                    logger.debug(f"  {level_key}: '{level_vars[level_key]}'")
        else:
            # Wenn kein input_path oder original_pdf_path, setze Level-Variablen auf leer
            for i in range(6):
                context[f'level{i}'] = ""
        
        # XML-Variablen (bestehende Felder aus der XML-Datei)
        if xml_path and os.path.exists(xml_path):
            context.update(VariableExtractor.get_xml_variables(xml_path))
        
        # OCR-Text (lazy loading)
        if any(m.expression and 'OCR' in m.expression for m in mappings):
            if pdf_path not in self._ocr_cache:
                logger.info(f"Führe OCR aus auf: {pdf_path}")
                self._ocr_cache[pdf_path] = self.ocr_processor.extract_text_from_pdf(pdf_path)
            
            context['OCR_FullText'] = self._ocr_cache[pdf_path]
        
        # OCR-Zonen vom Hotfolder
        if ocr_zones:
            for zone_info in ocr_zones:
                zone_key = f"{zone_info['page_num']}_{zone_info['zone']}"
                if zone_key not in self._zone_cache:
                    logger.info(f"Führe OCR aus für Zone '{zone_info['name']}' auf Seite {zone_info['page_num']}")
                    zone_text = self.ocr_processor.extract_text_from_zone(
                        pdf_path, zone_info['page_num'], zone_info['zone']
                    )
                    self._zone_cache[zone_key] = zone_text
                
                # Zone-Name hat bereits OCR_ Präfix, verwende ihn direkt
                zone_name = zone_info.get('name', f'OCR_Zone_{len(self._zone_cache)+1}')
                context[zone_name] = self._zone_cache[zone_key]
                
                logger.debug(f"Zone '{zone_name}' enthält: '{self._zone_cache[zone_key][:50]}...'")
        
        # OCR-Zonen aus Mappings (für Multi-Zone Support)
        for mapping in mappings:
            if mapping.zones:
                for zone_info in mapping.zones:
                    zone_key = f"{zone_info['page_num']}_{zone_info['zone']}"
                    if zone_key not in self._zone_cache:
                        zone_text = self.ocr_processor.extract_text_from_zone(
                            pdf_path, zone_info['page_num'], zone_info['zone']
                        )
                        self._zone_cache[zone_key] = zone_text
                    
                    # Zone-Name hat bereits OCR_ Präfix
                    zone_name = zone_info.get('name', f'OCR_Zone_{len(self._zone_cache)+1}')
                    context[zone_name] = self._zone_cache[zone_key]
        
        return context
 
    def _evaluate_mapping(self, mapping: FieldMapping, 
                         context: Dict[str, Any]) -> Optional[str]:
        """Evaluiert ein Mapping und gibt den Wert zurück"""
        
        # ÄNDERUNG: Wenn keine Expression definiert ist, gebe leeren String zurück
        if not mapping.expression:
            return ""
        
        if mapping.source_type == "ocr_zone":
            # Multi-Zone Support
            if mapping.zones:
                # Wenn eine Expression definiert ist, verwende sie
                if mapping.expression:
                    return self.function_parser.parse_and_evaluate(mapping.expression, context)
                else:
                    # Kombiniere alle Zonen-Texte
                    zone_texts = []
                    for zone_info in mapping.zones:
                        zone_name = zone_info.get('name')
                        if zone_name and zone_name in context:
                            zone_texts.append(context[zone_name])
                    return " ".join(zone_texts)
            
            # Einzelne Zone (falls noch verwendet)
            elif mapping.zone:
                zone_name = f"OCR_Zone_{mapping.field_name}"
                zone_var = f"<{zone_name}>"
                
                # Wenn eine Expression definiert ist, verwende sie
                if mapping.expression:
                    # Ersetze Platzhalter in der Expression
                    expression = mapping.expression.replace("<ZONE>", zone_var)
                    return self.function_parser.parse_and_evaluate(expression, context)
                else:
                    # Nur Zone-Text zurückgeben
                    return context.get(zone_name, "")
        elif mapping.source_type == "skript":
            return self.function_parser.parse_and_evaluate(f'SCRIPTING({mapping.expression})', context)
        elif mapping.source_type == "sql":
            return self.function_parser.parse_and_evaluate(f'SQL({mapping.expression})', context)
        else:
            # Normale Expression-Verarbeitung
            return self.function_parser.parse_and_evaluate(mapping.expression, context)
        
        return None
    
    def _indent_xml(self, elem, level=0):
        """Formatiert XML mit Einrückungen"""
        i = "\n" + level * "  "
        if len(elem):
            if not elem.text or not elem.text.strip():
                elem.text = i + "  "
            if not elem.tail or not elem.tail.strip():
                elem.tail = i
            for child in elem:
                self._indent_xml(child, level + 1)
            if not child.tail or not child.tail.strip():
                child.tail = i
        else:
            # ÄNDERUNG: Für leere Elemente setze explizit einen Leerzeichen-Text
            # Dies verhindert selbstschließende Tags
            if elem.text is None or elem.text == "":
                elem.text = "\n" + "  " * (level + 1)
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
            logger.error(f"Fehler beim Lesen der XML-Felder: {e}")
            return []
    
    def clear_ocr_cache(self):
        """Leert den OCR-Cache"""
        self._ocr_cache.clear()
        self._zone_cache.clear()
    
    def get_available_variables(self, xml_path: str = "", 
                               pdf_path: str = "",
                               input_path: str = "") -> Dict[str, List[str]]:
        """Gibt alle verfügbaren Variablen gruppiert zurück"""
        variables = {
            "Standard": [],
            "OCR": ["OCR_FullText"],
            "Datei": [],
            "XML": [],
            "Datum": [],
            "Level": []  # NEU: Level-Variablen
        }
        
        # Standard-Variablen
        std_vars = VariableExtractor.get_standard_variables()
        variables["Standard"] = list(std_vars.keys())
        
        # Datei-Variablen
        if pdf_path:
            file_vars = VariableExtractor.get_file_variables(pdf_path)
            variables["Datei"] = list(file_vars.keys())
        
        # NEU: Level-Variablen
        if pdf_path and input_path:
            level_vars = VariableExtractor.get_level_variables(pdf_path, input_path)
            variables["Level"] = [k for k in level_vars.keys() if level_vars[k]]  # Nur nicht-leere Level
        else:
            # Zeige alle möglichen Level-Variablen
            variables["Level"] = [f"level{i}" for i in range(6)]
        
        # XML-Variablen
        if xml_path and os.path.exists(xml_path):
            xml_vars = VariableExtractor.get_xml_variables(xml_path)
            # Zeige sowohl mit als auch ohne XML_ Prefix
            variables["XML"] = []
            for k in xml_vars.keys():
                if k.startswith("XML_"):
                    variables["XML"].append(k[4:])  # Ohne Prefix
                    variables["XML"].append(k)      # Mit Prefix
        
        return variables
    
    def get_available_functions(self) -> Dict[str, List[Dict[str, str]]]:
        """Gibt alle verfügbaren Funktionen mit Beschreibungen zurück"""
        return {
            "String": [
                {"name": "FORMAT", "syntax": 'FORMAT("<VAR>","<FORMATSTRING>")', 
                 "desc": "Formatiert eine Zeichenkette"},
                {"name": "TRIM", "syntax": 'TRIM("<VAR>")', 
                 "desc": "Entfernt Leerzeichen"},
                {"name": "LEFT", "syntax": 'LEFT("<VAR>",<LENGTH>)', 
                 "desc": "Gibt die ersten n Zeichen zurück"},
                {"name": "RIGHT", "syntax": 'RIGHT("<VAR>",<LENGTH>)', 
                 "desc": "Gibt die letzten n Zeichen zurück"},
                {"name": "MID", "syntax": 'MID("<VAR>",<START>,<LENGTH>)', 
                 "desc": "Gibt Zeichen aus der Mitte zurück"},
                {"name": "TOUPPER", "syntax": 'TOUPPER("<VAR>")', 
                 "desc": "Konvertiert zu Großbuchstaben"},
                {"name": "TOLOWER", "syntax": 'TOLOWER("<VAR>")', 
                 "desc": "Konvertiert zu Kleinbuchstaben"},
                {"name": "LEN", "syntax": 'LEN("<VAR>")', 
                 "desc": "Gibt die Länge zurück"},
                {"name": "INDEXOF", "syntax": 'INDEXOF(<START>,"<STRING>","<SEARCH>",<CASE>)', 
                 "desc": "Findet Position eines Zeichens"},
            ],
            "Datum": [
                {"name": "FORMATDATE", "syntax": 'FORMATDATE("<FORMAT>")', 
                 "desc": "Formatiert das aktuelle Datum"},
            ],
            "Numerisch": [
                {"name": "AUTOINCREMENT", "syntax": 'AUTOINCREMENT("<VAR>",<STEP>)', 
                 "desc": "Zählt einen Wert hoch"},
            ],
            "Bedingungen": [
                {"name": "IF", "syntax": 'IF("<VAR>","<OP>","<VALUE>","<TRUE>","<FALSE>")', 
                 "desc": "Bedingte Auswertung"},
            ],
            "RegEx": [
                {"name": "REGEXP.MATCH", "syntax": 'REGEXP.MATCH("<VAR>","<PATTERN>",<INDEX>)', 
                 "desc": "Findet Muster mit regulären Ausdrücken"},
                {"name": "REGEXP.REPLACE", "syntax": 'REGEXP.REPLACE("<VAR>","<PATTERN>","<REPLACE>")', 
                 "desc": "Ersetzt Muster mit regulären Ausdrücken"},
            ],
            "Scripting": [
                {"name": "SCRIPTING", "syntax": 'SCRIPTING("<PATH>","<VAR1>","<VAR2>",...)', 
                 "desc": "Führt externes Script aus"},
            ],
        }
        
    def _check_circular_dependencies(self, mappings: List[FieldMapping]) -> Tuple[bool, Optional[str]]:
        """
        Prüft auf zirkuläre Abhängigkeiten zwischen Feldern
        
        Returns:
            Tuple[bool, Optional[str]]: (Hat Zyklus, Fehlermeldung)
        """
        import re
        
        # Baue Abhängigkeitsgraph
        dependencies = {}
        for mapping in mappings:
            field_name = mapping.field_name
            # Extrahiere alle Variablen-Referenzen aus der Expression
            referenced_fields = re.findall(r'<([^>]+)>', mapping.expression)
            # Filtere nur Feld-Referenzen (keine System-Variablen)
            field_refs = []
            for ref in referenced_fields:
                # Prüfe ob es ein definiertes Feld ist
                if any(m.field_name == ref for m in mappings):
                    field_refs.append(ref)
            dependencies[field_name] = field_refs
        
        # Tiefensuche für Zyklus-Erkennung
        def has_cycle(field: str, visited: set, rec_stack: set, path: list) -> Tuple[bool, Optional[str]]:
            visited.add(field)
            rec_stack.add(field)
            path.append(field)
            
            for dep in dependencies.get(field, []):
                if dep not in visited:
                    has_cycle_result, cycle_path = has_cycle(dep, visited, rec_stack, path.copy())
                    if has_cycle_result:
                        return True, cycle_path
                elif dep in rec_stack:
                    # Zyklus gefunden
                    cycle_start = path.index(dep)
                    cycle = path[cycle_start:] + [dep]
                    return True, " → ".join(cycle)
            
            rec_stack.remove(field)
            return False, None
        
        # Prüfe alle Felder
        visited = set()
        for field in dependencies:
            if field not in visited:
                has_cycle_result, cycle_path = has_cycle(field, visited, set(), [])
                if has_cycle_result:
                    return True, f"Zirkuläre Abhängigkeit gefunden: {cycle_path}"
        
        return False, None