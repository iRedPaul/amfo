"""
PDF-Verarbeitungsfunktionen
"""
import os
import shutil
from typing import Optional, Dict, Any
from pathlib import Path
import PyPDF2
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.colors import Color
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
import sys
import xml.etree.ElementTree as ET
import subprocess
import tempfile
import uuid

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.hotfolder_config import ProcessingAction, DocumentPair, HotfolderConfig
from core.xml_field_processor import XMLFieldProcessor, FieldMapping
from core.ocr_processor import OCRProcessor
from core.export_processor import ExportProcessor


class PDFProcessor:
    """Verarbeitet PDF-Dateien basierend auf den konfigurierten Aktionen"""
    
    def __init__(self):
        self.xml_processor = XMLFieldProcessor()
        self.ocr_processor = OCRProcessor()
        self.export_processor = ExportProcessor()
        self.function_parser = None  # Lazy load
        self._ocr_cache = {}  # Cache für OCR-Ergebnisse
        self._zone_cache = {}  # Cache für OCR-Zonen
        self.supported_actions = {
            ProcessingAction.COMPRESS: self._compress_pdf,
            ProcessingAction.SPLIT: self._split_pdf,
            ProcessingAction.OCR: self._perform_ocr,
            ProcessingAction.PDF_A: self._convert_to_pdf_a
        }
        
        # Erstelle zentralen temporären Arbeitsordner
        self.temp_base_dir = os.path.join(tempfile.gettempdir(), "hotfolder_pdf_processor")
        os.makedirs(self.temp_base_dir, exist_ok=True)
    
    def process_document(self, doc_pair: DocumentPair, hotfolder: HotfolderConfig) -> bool:
        """
        Verarbeitet ein Dokument oder Dokumentenpaar
        
        Returns:
            bool: True wenn erfolgreich, False bei Fehler
        """
        # Erstelle eindeutigen temporären Arbeitsordner für diese Verarbeitung
        work_dir = os.path.join(self.temp_base_dir, f"work_{uuid.uuid4().hex}")
        os.makedirs(work_dir, exist_ok=True)
        
        try:
            # Verschiebe Dateien in temporären Arbeitsordner
            temp_pdf_path = os.path.join(work_dir, os.path.basename(doc_pair.pdf_path))
            shutil.move(doc_pair.pdf_path, temp_pdf_path)
            
            temp_xml_path = None
            if doc_pair.has_xml:
                temp_xml_path = os.path.join(work_dir, os.path.basename(doc_pair.xml_path))
                shutil.move(doc_pair.xml_path, temp_xml_path)
            
            # Erstelle temporäres DocumentPair mit neuen Pfaden
            temp_doc_pair = DocumentPair(
                pdf_path=temp_pdf_path,
                xml_path=temp_xml_path
            )
            
            # Prüfe ob Split aktiviert ist
            is_split = ProcessingAction.SPLIT in hotfolder.actions
            
            # Arbeite mit der temporären PDF
            pdf_to_process = temp_pdf_path
            
            # Führe PDF-Aktionen aus
            for action in hotfolder.actions:
                if action in self.supported_actions:
                    params = hotfolder.action_params.get(action.value, {})
                    
                    # Füge Output-Pfad für Split-Aktion hinzu
                    if action == ProcessingAction.SPLIT:
                        params['output_path'] = hotfolder.output_path
                    
                    success = self.supported_actions[action](pdf_to_process, params)
                    if not success:
                        print(f"Aktion {action.value} fehlgeschlagen für {os.path.basename(doc_pair.pdf_path)}")
                        raise Exception(f"Aktion {action.value} fehlgeschlagen")
            
            # Führe Exporte durch wenn konfiguriert
            if hasattr(hotfolder, 'export_configs') and hotfolder.export_configs:
                # Konvertiere OCR-Zonen zu Dictionary-Format
                ocr_zones = []
                if hasattr(hotfolder, 'ocr_zones'):
                    ocr_zones = [zone.to_dict() for zone in hotfolder.ocr_zones]
                
                # Führe Exporte durch
                export_results = self.export_processor.process_exports(
                    pdf_to_process,
                    temp_xml_path,
                    hotfolder.export_configs,
                    ocr_zones,
                    hotfolder.xml_field_mappings
                )
                
                # Prüfe ob alle Exporte erfolgreich waren
                all_successful = all(success for success, _ in export_results)
                if not all_successful:
                    # Mindestens ein Export fehlgeschlagen
                    failed_exports = [msg for success, msg in export_results if not success]
                    print(f"Export-Fehler: {', '.join(failed_exports)}")
                    
                    # Wenn alle Exporte fehlschlagen, behandle als Fehler
                    if not any(success for success, _ in export_results):
                        raise Exception("Alle Exporte fehlgeschlagen")
            
            # Legacy: Verschiebe verarbeitete Dateien zum Output wenn keine Exporte definiert
            if not hasattr(hotfolder, 'export_configs') or not hotfolder.export_configs:
                if is_split:
                    # Bei Split wurden die Dateien bereits im Output-Ordner erstellt
                    # Verschiebe nur die XML wenn vorhanden
                    if temp_doc_pair.has_xml:
                        base_name = os.path.splitext(os.path.basename(temp_doc_pair.pdf_path))[0]
                        xml_output = os.path.join(hotfolder.output_path, f"{base_name}.xml")
                        shutil.copy2(temp_xml_path, xml_output)
                        
                        # Verarbeite XML-Feld-Mappings
                        if hotfolder.xml_field_mappings:
                            mappings = [FieldMapping.from_dict(m) for m in hotfolder.xml_field_mappings]
                            # Verwende die erste Seite für XML-Verarbeitung
                            first_page_pdf = os.path.join(hotfolder.output_path, f"{base_name}_pages", "page_001.pdf")
                            if os.path.exists(first_page_pdf):
                                # Konvertiere OCR-Zonen zu Dictionary-Format
                                ocr_zones = []
                                if hasattr(hotfolder, 'ocr_zones'):
                                    ocr_zones = [zone.to_dict() for zone in hotfolder.ocr_zones]
                                
                                self.xml_processor.process_xml_with_mappings(
                                    xml_output, first_page_pdf, mappings, ocr_zones
                                )
                else:
                    # Normal: Verschiebe PDF zum Output
                    output_path = self._get_output_path(temp_doc_pair, hotfolder)
                    shutil.move(pdf_to_process, output_path)
                    
                    # Verarbeite XML wenn vorhanden
                    if temp_doc_pair.has_xml and hotfolder.process_pairs:
                        # Kopiere XML zum Output
                        xml_output = output_path.replace('.pdf', '.xml')
                        shutil.copy2(temp_xml_path, xml_output)
                        
                        # Verarbeite XML-Feld-Mappings
                        if hotfolder.xml_field_mappings:
                            mappings = [FieldMapping.from_dict(m) for m in hotfolder.xml_field_mappings]
                            
                            # Konvertiere OCR-Zonen zu Dictionary-Format
                            ocr_zones = []
                            if hasattr(hotfolder, 'ocr_zones'):
                                ocr_zones = [zone.to_dict() for zone in hotfolder.ocr_zones]
                            
                            self.xml_processor.process_xml_with_mappings(
                                xml_output, output_path, mappings, ocr_zones
                            )
                        else:
                            # Legacy-Verarbeitung
                            self._process_xml_data(temp_xml_path, output_path, hotfolder)
            
            # Leere OCR-Cache nach Verarbeitung
            self.xml_processor.clear_ocr_cache()
            self._ocr_cache.clear()
            self._zone_cache.clear()
            
            print(f"Erfolgreich verarbeitet: {temp_doc_pair.base_name}")
            return True
            
        except Exception as e:
            print(f"Fehler bei der Verarbeitung von {os.path.basename(doc_pair.pdf_path)}: {e}")
            
            # Bestimme Fehlerpfad
            error_path = self._get_error_path(doc_pair, hotfolder)
            os.makedirs(error_path, exist_ok=True)
            
            # Verschiebe Dateien zum Fehlerpfad
            try:
                if os.path.exists(temp_pdf_path):
                    error_pdf = os.path.join(error_path, os.path.basename(doc_pair.pdf_path))
                    # Füge Zeitstempel hinzu wenn Datei bereits existiert
                    if os.path.exists(error_pdf):
                        from datetime import datetime
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        base, ext = os.path.splitext(error_pdf)
                        error_pdf = f"{base}_{timestamp}{ext}"
                    shutil.move(temp_pdf_path, error_pdf)
                    
                if temp_xml_path and os.path.exists(temp_xml_path):
                    error_xml = os.path.join(error_path, os.path.basename(doc_pair.xml_path))
                    if os.path.exists(error_xml):
                        from datetime import datetime
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        base, ext = os.path.splitext(error_xml)
                        error_xml = f"{base}_{timestamp}{ext}"
                    shutil.move(temp_xml_path, error_xml)
                    
                print(f"Dateien in Fehlerpfad verschoben: {error_path}")
                
            except Exception as move_error:
                print(f"Fehler beim Verschieben in Fehlerpfad: {move_error}")
                # Als letzter Ausweg: Verschiebe zurück zum Input
                try:
                    if os.path.exists(temp_pdf_path):
                        shutil.move(temp_pdf_path, doc_pair.pdf_path)
                    if temp_xml_path and os.path.exists(temp_xml_path):
                        shutil.move(temp_xml_path, doc_pair.xml_path)
                except:
                    pass
            
            return False
        finally:
            # Aufräumen: Lösche temporären Arbeitsordner
            try:
                if os.path.exists(work_dir):
                    shutil.rmtree(work_dir)
            except Exception as cleanup_error:
                print(f"Fehler beim Aufräumen des temporären Ordners: {cleanup_error}")
    
    def _get_error_path(self, doc_pair: DocumentPair, hotfolder: HotfolderConfig) -> str:
        """Bestimmt den Fehlerpfad für fehlgeschlagene Verarbeitungen"""
        # Baue Kontext für Variablen
        context = {}
        if self.function_parser is None:
            from core.function_parser import FunctionParser, VariableExtractor
            self.function_parser = FunctionParser()
            self.variable_extractor = VariableExtractor()
        
        context.update(self.variable_extractor.get_standard_variables())
        context.update(self.variable_extractor.get_file_variables(doc_pair.pdf_path))
        context['InputPath'] = hotfolder.input_path
        context['OutputPath'] = hotfolder.output_path
        
        # Verwende Export-Processor für Fehlerpfad-Bestimmung
        error_path_expr = hotfolder.error_path if hasattr(hotfolder, 'error_path') else ""
        return self.export_processor.get_error_path(error_path_expr, context)
    
    def _get_output_path(self, doc_pair: DocumentPair, hotfolder: HotfolderConfig) -> str:
        """Bestimmt den Output-Pfad für die verarbeitete Datei"""
        # Lazy load function parser
        if not self.function_parser:
            from core.function_parser import FunctionParser, VariableExtractor
            self.function_parser = FunctionParser()
            self.variable_extractor = VariableExtractor()
        
        # Baue Kontext für Dateiname-Expression
        context = self.variable_extractor.get_standard_variables()
        context.update(self.variable_extractor.get_file_variables(doc_pair.pdf_path))
        
        # Füge OCR-Variablen hinzu wenn OCR aktiviert
        if ProcessingAction.OCR in hotfolder.actions:
            if doc_pair.pdf_path not in self._ocr_cache:
                self._ocr_cache[doc_pair.pdf_path] = self.ocr_processor.extract_text_from_pdf(doc_pair.pdf_path)
            context['OCR_FullText'] = self._ocr_cache[doc_pair.pdf_path]
        
        # Füge OCR-Zonen-Variablen hinzu
        if hasattr(hotfolder, 'ocr_zones') and hotfolder.ocr_zones:
            for i, zone in enumerate(hotfolder.ocr_zones):
                zone_key = f"{zone.page_num}_{zone.zone}"
                if zone_key not in self._zone_cache:
                    zone_text = self.ocr_processor.extract_text_from_zone(
                        doc_pair.pdf_path, zone.page_num, zone.zone
                    )
                    self._zone_cache[zone_key] = zone_text
                
                # Variable mit benutzerdefiniertem Namen
                context[zone.name] = self._zone_cache[zone_key]
                # Legacy-Support
                context[f'ZONE_{i+1}'] = self._zone_cache[zone_key]
        
        # Evaluiere Expression
        expression = hotfolder.output_filename_expression if hasattr(hotfolder, 'output_filename_expression') else "<FileName>"
        filename_base = self.function_parser.parse_and_evaluate(expression, context)
        
        # Stelle sicher, dass Dateiname gültig ist
        # Entferne ungültige Zeichen
        invalid_chars = '<>:"|?*'
        for char in invalid_chars:
            filename_base = filename_base.replace(char, '_')
        
        # Füge .pdf Endung hinzu wenn nicht vorhanden
        if not filename_base.lower().endswith('.pdf'):
            filename_base += '.pdf'
        
        output_path = os.path.join(hotfolder.output_path, filename_base)
        
        # Verhindere Überschreiben
        if os.path.exists(output_path):
            base, ext = os.path.splitext(output_path)
            counter = 1
            while os.path.exists(f"{base}_{counter}{ext}"):
                counter += 1
            output_path = f"{base}_{counter}{ext}"
        
        return output_path
    
    def _process_xml_data(self, xml_path: str, pdf_path: str, hotfolder: HotfolderConfig) -> None:
        """Verarbeitet XML-Daten und wendet sie auf das PDF an"""
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # Beispiel: Extrahiere Metadaten aus XML
            metadata = {}
            for child in root:
                metadata[child.tag] = child.text
            
            # Füge Metadaten zum PDF hinzu
            if metadata:
                self._update_metadata(pdf_path, {"xml_metadata": metadata})
                
        except Exception as e:
            print(f"Fehler beim Verarbeiten der XML-Datei: {e}")
    
    def _update_metadata(self, pdf_path: str, metadata: Dict[str, Any]) -> None:
        """Aktualisiert PDF-Metadaten"""
        try:
            reader = PyPDF2.PdfReader(pdf_path)
            writer = PyPDF2.PdfWriter()
            
            # Kopiere alle Seiten
            for page in reader.pages:
                writer.add_page(page)
            
            # Füge Metadaten hinzu
            writer.add_metadata(metadata)
            
            # Speichere als temporäre Datei
            temp_path = pdf_path + ".metadata"
            with open(temp_path, 'wb') as output_file:
                writer.write(output_file)
            
            # Ersetze Original
            shutil.move(temp_path, pdf_path)
            
        except Exception as e:
            print(f"Fehler beim Aktualisieren der Metadaten: {e}")
    
    # PDF-Verarbeitungsfunktionen
    def _compress_pdf(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Komprimiert eine PDF-Datei"""
        try:
            # Einfache Komprimierung durch Neuschreiben
            reader = PyPDF2.PdfReader(pdf_path)
            writer = PyPDF2.PdfWriter()
            
            for page in reader.pages:
                page.compress_content_streams()  # Komprimiere Inhalte
                writer.add_page(page)
            
            temp_path = pdf_path + ".compressed"
            with open(temp_path, 'wb') as output_file:
                writer.write(output_file)
            
            # Prüfe ob komprimierte Datei kleiner ist
            original_size = os.path.getsize(pdf_path)
            compressed_size = os.path.getsize(temp_path)
            
            if compressed_size < original_size:
                shutil.move(temp_path, pdf_path)
                print(f"PDF komprimiert: {original_size} -> {compressed_size} bytes")
            else:
                os.remove(temp_path)
                print("Komprimierung hat keine Größenreduzierung gebracht")
            
            return True
            
        except Exception as e:
            print(f"Fehler bei der Komprimierung: {e}")
            return False
    
    def _split_pdf(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Teilt eine PDF in einzelne Seiten auf"""
        try:
            reader = PyPDF2.PdfReader(pdf_path)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            
            # Hole Output-Pfad aus params (wird von process_document gesetzt)
            output_path = params.get('output_path', os.path.dirname(pdf_path))
            
            # Erstelle Ordner für Split-Dateien im Output
            split_dir = os.path.join(output_path, f"{base_name}_pages")
            os.makedirs(split_dir, exist_ok=True)
            
            for i, page in enumerate(reader.pages):
                writer = PyPDF2.PdfWriter()
                writer.add_page(page)
                
                output_file = os.path.join(split_dir, f"page_{i+1:03d}.pdf")
                with open(output_file, 'wb') as f:
                    writer.write(f)
            
            print(f"PDF aufgeteilt in {len(reader.pages)} Seiten: {split_dir}")
            return True
            
        except Exception as e:
            print(f"Fehler beim Aufteilen: {e}")
            return False
    
    def _perform_ocr(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Führt OCR auf der PDF aus und erstellt eine durchsuchbare PDF"""
        try:
            language = params.get("language", "deu")  # Standard: Deutsch
            
            # Verwende OCRmyPDF wenn verfügbar (bessere Qualität)
            try:
                import ocrmypdf
                
                print(f"Führe OCR aus mit ocrmypdf...")
                ocrmypdf.ocr(
                    pdf_path,
                    pdf_path,
                    language=language,
                    force_ocr=True,
                    optimize=1
                )
                print(f"OCR erfolgreich durchgeführt")
                return True
                
            except ImportError:
                # Fallback: Tesseract direkt verwenden
                print("ocrmypdf nicht installiert, verwende Tesseract direkt...")
                
                # Prüfe ob Tesseract verfügbar ist
                try:
                    subprocess.run(["tesseract", "--version"], 
                                 capture_output=True, check=True)
                except:
                    print("Tesseract nicht gefunden! Bitte installieren Sie Tesseract OCR.")
                    return False
                
                # Konvertiere zu durchsuchbarer PDF mit Tesseract
                with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as temp_out:
                    temp_out_path = temp_out.name
                
                cmd = [
                    "tesseract",
                    pdf_path,
                    temp_out_path[:-4],  # Ohne .pdf Endung
                    "-l", language,
                    "pdf"
                ]
                
                result = subprocess.run(cmd, capture_output=True, text=True)
                
                if result.returncode == 0 and os.path.exists(temp_out_path):
                    shutil.move(temp_out_path, pdf_path)
                    print("OCR erfolgreich durchgeführt")
                    return True
                else:
                    print(f"OCR fehlgeschlagen: {result.stderr}")
                    return False
                    
        except Exception as e:
            print(f"Fehler bei OCR: {e}")
            return False
    
    def _convert_to_pdf_a(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Konvertiert PDF zu PDF/A"""
        try:
            # Versuche ocrmypdf für PDF/A Konvertierung
            try:
                import ocrmypdf
                
                print(f"Konvertiere zu PDF/A...")
                ocrmypdf.ocr(
                    pdf_path,
                    pdf_path,
                    pdfa_image_compression="jpeg",
                    output_type="pdfa",
                    tesseract_timeout=0,  # Kein OCR, nur PDF/A
                    skip_text=True  # Überspringe Text-Layer
                )
                print(f"PDF/A Konvertierung erfolgreich")
                return True
                
            except ImportError:
                print("ocrmypdf nicht installiert. Für PDF/A Konvertierung bitte installieren: pip install ocrmypdf")
                
                # Einfacher Fallback: Füge PDF/A Metadaten hinzu
                try:
                    reader = PyPDF2.PdfReader(pdf_path)
                    writer = PyPDF2.PdfWriter()
                    
                    # Kopiere alle Seiten
                    for page in reader.pages:
                        writer.add_page(page)
                    
                    # Füge PDF/A konforme Metadaten hinzu
                    metadata = {
                        '/Title': params.get('title', 'Dokument'),
                        '/Producer': 'Hotfolder PDF Processor',
                        '/Creator': 'Hotfolder PDF Processor',
                        '/CreationDate': 'D:20240101000000',
                        '/ModDate': 'D:20240101000000'
                    }
                    writer.add_metadata(metadata)
                    
                    temp_path = pdf_path + ".pdfa"
                    with open(temp_path, 'wb') as output_file:
                        writer.write(output_file)
                    
                    shutil.move(temp_path, pdf_path)
                    print("Basis PDF/A Metadaten hinzugefügt (für volle Konformität ocrmypdf installieren)")
                    return True
                    
                except Exception as e:
                    print(f"Fehler bei PDF/A Konvertierung: {e}")
                    return False
                    
        except Exception as e:
            print(f"Fehler bei PDF/A Konvertierung: {e}")
            return False
    
    def cleanup_temp_dir(self):
        """Räumt den temporären Arbeitsordner auf"""
        try:
            if os.path.exists(self.temp_base_dir):
                # Lösche nur alte Arbeitsordner (älter als 1 Tag)
                import time
                now = time.time()
                
                for work_dir in os.listdir(self.temp_base_dir):
                    dir_path = os.path.join(self.temp_base_dir, work_dir)
                    if os.path.isdir(dir_path):
                        # Prüfe Alter des Ordners
                        dir_age = now - os.path.getmtime(dir_path)
                        if dir_age > 86400:  # 24 Stunden
                            try:
                                shutil.rmtree(dir_path)
                                print(f"Alter temporärer Ordner gelöscht: {work_dir}")
                            except Exception as e:
                                print(f"Fehler beim Löschen von {work_dir}: {e}")
        except Exception as e:
            print(f"Fehler beim Aufräumen des temporären Verzeichnisses: {e}")