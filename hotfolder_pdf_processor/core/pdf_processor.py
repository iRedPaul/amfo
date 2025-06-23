"""
PDF-Verarbeitungsfunktionen
"""
import os
import shutil
from typing import Optional, Dict, Any, List
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
from datetime import datetime

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.hotfolder_config import ProcessingAction, DocumentPair, HotfolderConfig
from core.xml_field_processor import XMLFieldProcessor, FieldMapping
from core.ocr_processor import OCRProcessor
from core.export_processor import ExportProcessor
from core.export_action import ExportConfig


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
        
        # Merke Original-Pfade für Fehlerbehandlung
        original_pdf_path = doc_pair.pdf_path
        original_xml_path = doc_pair.xml_path
        
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
            
            # Arbeite mit der temporären PDF
            pdf_to_process = temp_pdf_path
            
            # Führe PDF-Aktionen aus
            for action in hotfolder.actions:
                if action in self.supported_actions:
                    params = hotfolder.action_params.get(action.value, {})
                    
                    success = self.supported_actions[action](pdf_to_process, params)
                    if not success:
                        print(f"Aktion {action.value} fehlgeschlagen für {os.path.basename(doc_pair.pdf_path)}")
                        raise Exception(f"Aktion {action.value} fehlgeschlagen")
            
            # Extrahiere OCR-Text wenn OCR durchgeführt wurde
            if ProcessingAction.OCR in hotfolder.actions:
                ocr_text = self.ocr_processor.extract_text_from_pdf(pdf_to_process)
                self._ocr_cache[pdf_to_process] = ocr_text
                self.export_processor.set_ocr_cache(pdf_to_process, ocr_text)
            
            # WICHTIG: Verarbeite XML-Felder VOR den Exporten
            evaluated_fields = {}
            if temp_doc_pair.has_xml and hotfolder.process_pairs and hotfolder.xml_field_mappings:
                # Erstelle temporäre XML im Arbeitsordner
                temp_xml_work = os.path.join(work_dir, "temp_fields.xml")
                shutil.copy2(temp_xml_path, temp_xml_work)
                
                # Verarbeite XML-Felder
                mappings = [FieldMapping.from_dict(m) for m in hotfolder.xml_field_mappings]
                ocr_zones = []
                if hasattr(hotfolder, 'ocr_zones'):
                    ocr_zones = [zone.to_dict() for zone in hotfolder.ocr_zones]
                
                # Verarbeite die Felder
                self.xml_processor.process_xml_with_mappings(
                    temp_xml_work, pdf_to_process, mappings, ocr_zones
                )
                
                # Extrahiere die evaluierten Felder für die Export-Expressions
                try:
                    tree = ET.parse(temp_xml_work)
                    root = tree.getroot()
                    fields_elem = root.find(".//Fields")
                    if fields_elem is not None:
                        for field in fields_elem:
                            if field.text:
                                evaluated_fields[field.tag] = field.text
                except Exception as e:
                    print(f"Fehler beim Extrahieren der XML-Felder: {e}")
                
                # Lösche temporäre XML
                os.remove(temp_xml_work)
                
                # Verwende die finale XML für Exporte
                if temp_xml_path:
                    shutil.copy2(temp_xml_work, temp_xml_path)
            
            # NEU: Verarbeite Export-Konfigurationen
            exports = []
            if hasattr(hotfolder, 'exports') and hotfolder.exports:
                # Verwende konfigurierte Exporte
                exports = [ExportConfig.from_dict(e) for e in hotfolder.exports]
            else:
                # Fallback: Erstelle Standard-Export aus alten Feldern
                from core.export_action import ExportType, MetaDataFormat, MetaDataConfig
                
                default_export = ExportConfig(
                    id="legacy",
                    name="Standard-Ausgabe",
                    enabled=True,
                    export_type=ExportType.FILE,
                    output_path_expression=hotfolder.output_path,
                    filename_expression=getattr(hotfolder, 'output_filename_expression', '<FileName>'),
                    create_path_if_not_exists=True,
                    error_output_path=hotfolder.error_path,
                    metadata_config=MetaDataConfig(format=MetaDataFormat.NONE)
                )
                exports = [default_export]
            
            # Verarbeite alle Exporte
            processed_files = self.export_processor.process_exports(
                temp_doc_pair, hotfolder, pdf_to_process, temp_xml_path, 
                exports, evaluated_fields
            )
            
            # Leere Caches nach Verarbeitung
            self.xml_processor.clear_ocr_cache()
            self._ocr_cache.clear()
            self._zone_cache.clear()
            self.export_processor.cleanup()
            
            print(f"Erfolgreich verarbeitet: {temp_doc_pair.base_name}")
            return True
            
        except Exception as e:
            print(f"Fehler bei der Verarbeitung von {os.path.basename(original_pdf_path)}: {e}")
            
            # Bei Fehler: Verschiebe zurück zum Input oder in Fehler-Ordner
            # Dies wird jetzt von den Export-Konfigurationen gehandhabt
            self._move_back_to_input(temp_pdf_path, original_pdf_path, 
                                   temp_xml_path, original_xml_path)
            
            return False
        finally:
            # Aufräumen: Lösche temporären Arbeitsordner
            try:
                if os.path.exists(work_dir):
                    shutil.rmtree(work_dir)
            except Exception as cleanup_error:
                print(f"Fehler beim Aufräumen des temporären Ordners: {cleanup_error}")
    
    def _move_back_to_input(self, temp_pdf_path: str, original_pdf_path: str,
                           temp_xml_path: Optional[str], original_xml_path: Optional[str]):
        """Verschiebt Dateien zurück zum Input-Ordner bei Fehlern"""
        try:
            if os.path.exists(temp_pdf_path):
                shutil.move(temp_pdf_path, original_pdf_path)
            if temp_xml_path and os.path.exists(temp_xml_path):
                shutil.move(temp_xml_path, original_xml_path)
        except Exception as move_error:
            print(f"Fehler beim Zurückverschieben der Dateien: {move_error}")
    
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
            
            # Erstelle Ordner für Split-Dateien im temporären Verzeichnis
            split_dir = os.path.join(os.path.dirname(pdf_path), f"{base_name}_pages")
            os.makedirs(split_dir, exist_ok=True)
            
            for i, page in enumerate(reader.pages):
                writer = PyPDF2.PdfWriter()
                writer.add_page(page)
                
                output_file = os.path.join(split_dir, f"page_{i+1:03d}.pdf")
                with open(output_file, 'wb') as f:
                    writer.write(f)
            
            print(f"PDF aufgeteilt in {len(reader.pages)} Seiten: {split_dir}")
            
            # Bei Split: Ersetze Original-PDF durch erste Seite für weitere Verarbeitung
            # Die einzelnen Seiten werden separat exportiert
            first_page_path = os.path.join(split_dir, "page_001.pdf")
            if os.path.exists(first_page_path):
                shutil.copy2(first_page_path, pdf_path)
            
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