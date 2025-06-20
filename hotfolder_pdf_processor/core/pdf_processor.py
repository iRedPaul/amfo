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

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.hotfolder_config import ProcessingAction, DocumentPair, HotfolderConfig
from core.xml_field_processor import XMLFieldProcessor, FieldMapping
from core.ocr_processor import OCRProcessor


class PDFProcessor:
    """Verarbeitet PDF-Dateien basierend auf den konfigurierten Aktionen"""
    
    def __init__(self):
        self.xml_processor = XMLFieldProcessor()
        self.ocr_processor = OCRProcessor()
        self.function_parser = None  # Lazy load
        self._ocr_cache = {}  # Cache für OCR-Ergebnisse
        self._zone_cache = {}  # Cache für OCR-Zonen
        self.supported_actions = {
            ProcessingAction.COMPRESS: self._compress_pdf,
            ProcessingAction.SPLIT: self._split_pdf,
            ProcessingAction.OCR: self._perform_ocr,
            ProcessingAction.PDF_A: self._convert_to_pdf_a
        }
    
    def process_document(self, doc_pair: DocumentPair, hotfolder: HotfolderConfig) -> bool:
        """
        Verarbeitet ein Dokument oder Dokumentenpaar
        
        Returns:
            bool: True wenn erfolgreich, False bei Fehler
        """
        try:
            # Prüfe ob Split aktiviert ist
            is_split = ProcessingAction.SPLIT in hotfolder.actions
            
            if is_split:
                # Bei Split direkt mit Original arbeiten
                pdf_to_process = doc_pair.pdf_path
            else:
                # Erstelle temporäre Arbeitskopie
                pdf_to_process = self._create_temp_copy(doc_pair.pdf_path)
            
            # Führe PDF-Aktionen aus
            for action in hotfolder.actions:
                if action in self.supported_actions:
                    params = hotfolder.action_params.get(action.value, {})
                    
                    # Füge Output-Pfad für Split-Aktion hinzu
                    if action == ProcessingAction.SPLIT:
                        params['output_path'] = hotfolder.output_path
                    
                    success = self.supported_actions[action](pdf_to_process, params)
                    if not success:
                        print(f"Aktion {action.value} fehlgeschlagen für {doc_pair.pdf_path}")
                        if not is_split and os.path.exists(pdf_to_process):
                            os.remove(pdf_to_process)
                        return False
            
            # Verschiebe verarbeitete Dateien zum Output
            if is_split:
                # Bei Split wurden die Dateien bereits im Output-Ordner erstellt
                # Verschiebe nur die XML wenn vorhanden
                if doc_pair.has_xml:
                    base_name = os.path.splitext(os.path.basename(doc_pair.pdf_path))[0]
                    xml_output = os.path.join(hotfolder.output_path, f"{base_name}.xml")
                    shutil.copy2(doc_pair.xml_path, xml_output)
                    
                    # Verarbeite XML-Feld-Mappings
                    if hotfolder.xml_field_mappings:
                        mappings = [FieldMapping.from_dict(m) for m in hotfolder.xml_field_mappings]
                        # Verwende die erste Seite für XML-Verarbeitung
                        first_page_pdf = os.path.join(hotfolder.output_path, f"{base_name}_pages", "page_001.pdf")
                        if os.path.exists(first_page_pdf):
                            self.xml_processor.process_xml_with_mappings(
                                xml_output, first_page_pdf, mappings
                            )
            else:
                # Normal: Verschiebe PDF zum Output
                output_path = self._get_output_path(doc_pair, hotfolder)
                shutil.move(pdf_to_process, output_path)
                
                # Verarbeite XML wenn vorhanden
                if doc_pair.has_xml and hotfolder.process_pairs:
                    # Kopiere XML zum Output
                    xml_output = output_path.replace('.pdf', '.xml')
                    shutil.copy2(doc_pair.xml_path, xml_output)
                    
                    # Verarbeite XML-Feld-Mappings
                    if hotfolder.xml_field_mappings:
                        mappings = [FieldMapping.from_dict(m) for m in hotfolder.xml_field_mappings]
                        self.xml_processor.process_xml_with_mappings(
                            xml_output, output_path, mappings
                        )
                    else:
                        # Legacy-Verarbeitung
                        self._process_xml_data(doc_pair.xml_path, output_path, hotfolder)
            
            # Lösche Originaldateien aus Input-Ordner
            os.remove(doc_pair.pdf_path)
            if doc_pair.has_xml:
                os.remove(doc_pair.xml_path)
            
            # Leere OCR-Cache nach Verarbeitung
            self.xml_processor.clear_ocr_cache()
            self._ocr_cache.clear()
            self._zone_cache.clear()
            
            print(f"Erfolgreich verarbeitet: {doc_pair.base_name}")
            return True
            
        except Exception as e:
            print(f"Fehler bei der Verarbeitung von {doc_pair.pdf_path}: {e}")
            # Aufräumen bei Fehler
            if 'pdf_to_process' in locals() and os.path.exists(pdf_to_process) and pdf_to_process != doc_pair.pdf_path:
                os.remove(pdf_to_process)
            return False
    
    def _create_temp_copy(self, pdf_path: str) -> str:
        """Erstellt eine temporäre Kopie der PDF-Datei"""
        temp_path = pdf_path + ".tmp"
        shutil.copy2(pdf_path, temp_path)
        return temp_path
    
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