"""
PDF-Verarbeitungsfunktionen - Vereinfachte Version mit nur 3 Export-Formaten
"""
import os
import shutil
from typing import Optional, Dict, Any, List, Tuple
from pathlib import Path
from pypdf import PdfWriter, PdfReader
from PIL import Image
import sys
import xml.etree.ElementTree as ET
import subprocess
import tempfile
import uuid
import logging
from datetime import datetime
import fitz  # PyMuPDF für bessere PDF-Analyse
import ocrmypdf
from ocrmypdf import PdfContext
from core.logging_config import setup_logging
import json

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.xml_field_processor import XMLFieldProcessor, FieldMapping
from core.ocr_processor import OCRProcessor
from core.export_processor import ExportProcessor
from models.hotfolder_config import HotfolderConfig, ProcessingAction, DocumentPair
from models.export_config import ExportSettings

logger = logging.getLogger(__name__)

class PDFProcessor:
    """Vereinfachter PDF-Prozessor mit nur 3 Export-Formaten"""
    
    def __init__(self):
        self.xml_processor = XMLFieldProcessor()
        self.ocr_processor = OCRProcessor()
        self.export_processor = ExportProcessor()
        self.function_parser = None
        self.variable_extractor = None
        self._ocr_cache = {}
        self._zone_cache = {}
        
        self.supported_actions = {
            ProcessingAction.COMPRESS: self._compress_pdf
        }
        
        # Erstelle zentralen temporären Arbeitsordner
        self.temp_base_dir = os.path.join(tempfile.gettempdir(), "hotfolder_pdf_processor")
        os.makedirs(self.temp_base_dir, exist_ok=True)
        
        # Lade Einstellungen
        self.settings = self._load_settings()
        
        # Prüfe Abhängigkeiten beim Start
        self._check_dependencies()
    
    def _load_settings(self) -> ExportSettings:
        """Lädt Einstellungen aus settings.json"""
        try:
            settings_file = "config/settings.json"
            
            if os.path.exists(settings_file):
                with open(settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return ExportSettings.from_dict(data)
            else:
                # Erstelle Default-Settings
                settings = ExportSettings()
                self._save_settings(settings)
                return settings
        except Exception as e:
            logger.error(f"Fehler beim Laden der Einstellungen: {e}")
            return ExportSettings()

    def _save_settings(self, settings: ExportSettings):
        """Speichert Einstellungen"""
        try:
            settings_file = "config/settings.json"
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings.to_dict(), f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Fehler beim Speichern der Einstellungen: {e}")
    
    def _check_dependencies(self):
        """Prüft ob alle benötigten Abhängigkeiten verfügbar sind"""
        warnings = []
        
        # Basis-Verzeichnis für dependencies
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        dependencies_dir = os.path.join(base_dir, 'dependencies')
        
        # Prüfe Tesseract (für PDF/A-Export)
        if not self._is_tesseract_available():
            warnings.append("Tesseract nicht gefunden - PDF/A (Durchsuchbar) Export eingeschränkt")
            warnings.append(f"Bitte Tesseract im dependencies Ordner platzieren: {dependencies_dir}")
        
        # Prüfe OCRmyPDF (für PDF/A-Export)
        try:
            import ocrmypdf
        except ImportError:
            warnings.append("OCRmyPDF nicht installiert - PDF/A (Durchsuchbar) Export nicht möglich")
        
        if warnings:
            logger.warning("Konfigurationswarnungen:\n" + "\n".join(warnings))
    
    def _is_tesseract_available(self) -> bool:
        """Prüft ob Tesseract verfügbar ist"""
        # Vereinfachte Prüfung - nur die wichtigsten Pfade
        base_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
        tesseract_paths = [
            os.path.join(base_dir, 'dependencies', 'Tesseract-OCR', 'tesseract.exe'),
            'tesseract'  # System PATH
        ]
        
        for path in tesseract_paths:
            try:
                if os.path.exists(path) or subprocess.run([path, '--version'], capture_output=True).returncode == 0:
                    self._tesseract_path = path
                    if os.path.exists(path):
                        os.environ['TESSERACT_PATH'] = os.path.dirname(path)
                    return True
            except:
                continue
        
        return False
    
    def process_document(self, doc_pair: DocumentPair, hotfolder: HotfolderConfig) -> bool:
        """
        Verarbeitet ein Dokument mit vereinfachter Logik
        """
        work_dir = os.path.join(self.temp_base_dir, f"work_{uuid.uuid4().hex}")
        os.makedirs(work_dir, exist_ok=True)
        
        # Variable für verarbeitete XML
        processed_xml_path = None
        
        try:
            # Verschiebe Dateien in temporären Arbeitsordner
            temp_pdf_path = os.path.join(work_dir, os.path.basename(doc_pair.pdf_path))
            shutil.move(doc_pair.pdf_path, temp_pdf_path)
            
            temp_xml_path = None
            if doc_pair.has_xml and doc_pair.xml_path is not None:
                temp_xml_path = os.path.join(work_dir, os.path.basename(doc_pair.xml_path))
                shutil.move(doc_pair.xml_path, temp_xml_path)
            
            # Qualitätskontrolle vor Verarbeitung
            if not self._validate_pdf(temp_pdf_path):
                raise Exception("PDF-Validierung fehlgeschlagen - Datei möglicherweise beschädigt")
            
            # Analysiere PDF für optimale Verarbeitung
            pdf_info = self._analyze_pdf(temp_pdf_path)
            logger.info(f"PDF-Analyse: {pdf_info}")
            
            # XML-Feld-Mappings anwenden - auch ohne XML-Datei verarbeiten
            if hotfolder.xml_field_mappings:
                mappings = [FieldMapping.from_dict(m) for m in hotfolder.xml_field_mappings]
                ocr_zones = [
                    zone if isinstance(zone, dict) else zone.to_dict()
                    for zone in hotfolder.ocr_zones
                ]
                
                # Wenn keine XML vorhanden, erstelle eine temporäre XML
                if not temp_xml_path:
                    # Erstelle eine minimale XML-Datei für die Feldverarbeitung
                    temp_xml_path = os.path.join(work_dir, "temp_fields.xml")
                    with open(temp_xml_path, 'w', encoding='utf-8') as f:
                        f.write('<?xml version="1.0" encoding="UTF-8"?>\n')
                        f.write('<root>\n')
                        f.write('  <Document>\n')
                        f.write('    <Fields>\n')
                        # Erstelle leere Felder für alle definierten Mappings
                        for mapping in mappings:
                            f.write(f'      <{mapping.field_name}></{mapping.field_name}>\n')
                        f.write('    </Fields>\n')
                        f.write('  </Document>\n')
                        f.write('</root>\n')
                
                # Verarbeite XML-Felder
                success = self.xml_processor.process_xml_with_mappings(
                    temp_xml_path, temp_pdf_path, mappings, ocr_zones, 
                    input_path=hotfolder.input_path,
                    original_pdf_path=doc_pair.pdf_path
                )
                
                if success:
                    logger.info(f"XML-Felder erfolgreich verarbeitet")
                    processed_xml_path = temp_xml_path
                else:
                    logger.error("XML-Feldverarbeitung fehlgeschlagen")
            
            # Führe nur noch unterstützte PDF-Aktionen aus (nur COMPRESS)
            compression_enabled = False
            for action in hotfolder.actions:
                if action in self.supported_actions:
                    params = hotfolder.action_params.get(action.value, {})
                    
                    # Füge PDF-Info zu Parametern hinzu für intelligente Verarbeitung
                    params['pdf_info'] = pdf_info
                    
                    logger.info(f"Führe Aktion aus: {action.value}")
                    success = self.supported_actions[action](temp_pdf_path, params)
                    
                    if not success:
                        raise Exception(f"Aktion {action.value} fehlgeschlagen")
                    
                    if action == ProcessingAction.COMPRESS:
                        compression_enabled = True
                    
                    # Qualitätskontrolle nach jeder Aktion
                    if not self._validate_pdf(temp_pdf_path):
                        raise Exception(f"PDF-Validierung nach {action.value} fehlgeschlagen")
            
            # Führe Exporte durch ODER verschiebe in Error wenn keine Exporte
            if hasattr(hotfolder, 'export_configs') and hotfolder.export_configs:
                ocr_zones = [
                    zone if isinstance(zone, dict) else zone.to_dict()
                    for zone in hotfolder.ocr_zones
                ]
                
                export_results = self.export_processor.process_exports(
                    temp_pdf_path,
                    processed_xml_path or temp_xml_path,
                    hotfolder.export_configs,
                    ocr_zones,
                    hotfolder.xml_field_mappings,
                    original_pdf_path=doc_pair.pdf_path,
                    input_path=hotfolder.input_path,
                    compression_enabled=compression_enabled
                )
                
                all_successful = all(success for success, _ in export_results)
                if not all_successful:
                    failed_exports = [msg for success, msg in export_results if not success]
                    raise Exception(f"Export-Fehler: {', '.join(failed_exports)}")
            else:
                # Keine Exporte konfiguriert - verschiebe in Error-Ordner
                logger.info("Keine Exporte konfiguriert - verschiebe Dateien in Error-Ordner")
                
                error_path = self._get_error_path(doc_pair, hotfolder)
                os.makedirs(error_path, exist_ok=True)
                
                # Verschiebe PDF
                error_pdf = os.path.join(error_path, os.path.basename(doc_pair.pdf_path))
                if os.path.exists(error_pdf):
                    timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                    base, ext = os.path.splitext(error_pdf)
                    error_pdf = f"{base}_{timestamp}{ext}"
                shutil.move(temp_pdf_path, error_pdf)
                
                # Verschiebe XML wenn vorhanden
                if temp_xml_path and os.path.exists(temp_xml_path):
                    error_xml = os.path.join(error_path, os.path.basename(doc_pair.xml_path or "temp_fields.xml"))
                    if os.path.exists(error_xml):
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        base, ext = os.path.splitext(error_xml)
                        error_xml = f"{base}_{timestamp}{ext}"
                    shutil.move(temp_xml_path, error_xml)
                
                logger.info(f"Dateien ohne Export in Error-Ordner verschoben: {error_path}")
            
            # Abschließende Qualitätskontrolle
            final_info = self._analyze_pdf(temp_pdf_path)
            logger.info(f"Finale PDF-Analyse: {final_info}")
            
            # Leere Caches
            self._ocr_cache.clear()
            self._zone_cache.clear()
            
            logger.info(f"Erfolgreich verarbeitet: {os.path.basename(doc_pair.pdf_path)}")
            return True
            
        except Exception as e:
            logger.error(f"Fehler bei der Verarbeitung: {e}")
            
            # Verschiebe in Fehlerpfad
            error_path = self._get_error_path(doc_pair, hotfolder)
            os.makedirs(error_path, exist_ok=True)
            
            try:
                if os.path.exists(temp_pdf_path):
                    error_pdf = os.path.join(error_path, os.path.basename(doc_pair.pdf_path))
                    if os.path.exists(error_pdf):
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        base, ext = os.path.splitext(error_pdf)
                        error_pdf = f"{base}_{timestamp}{ext}"
                    shutil.move(temp_pdf_path, error_pdf)
                    
                if temp_xml_path and os.path.exists(temp_xml_path):
                    error_xml = os.path.join(error_path, os.path.basename(doc_pair.xml_path or "temp_fields.xml"))
                    if os.path.exists(error_xml):
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        base, ext = os.path.splitext(error_xml)
                        error_xml = f"{base}_{timestamp}{ext}"
                    shutil.move(temp_xml_path, error_xml)
                    
                logger.info(f"Dateien in Fehlerpfad verschoben: {error_path}")
                
            except Exception as move_error:
                logger.error(f"Fehler beim Verschieben in Fehlerpfad: {move_error}")
            
            return False
            
        finally:
            # Aufräumen
            try:
                if os.path.exists(work_dir):
                    shutil.rmtree(work_dir)
            except Exception as cleanup_error:
                logger.error(f"Fehler beim Aufräumen: {cleanup_error}")
           
    def _validate_pdf(self, pdf_path: str) -> bool:
        """Validiert ob PDF gültig und nicht beschädigt ist"""
        try:
            # Versuche PDF mit PyMuPDF zu öffnen
            doc = fitz.open(pdf_path)
            page_count = doc.page_count
            
            # Prüfe ob mindestens eine Seite vorhanden
            if page_count == 0:
                doc.close()
                return False
            
            # Versuche erste Seite zu laden
            page = doc[0]
            _ = page.get_pixmap()
            
            doc.close()
            return True
            
        except Exception as e:
            logger.error(f"PDF-Validierung fehlgeschlagen: {e}")
            return False
    
    def _analyze_pdf(self, pdf_path: str) -> Dict[str, Any]:
        """Vereinfachte PDF-Analyse"""
        try:
            doc = fitz.open(pdf_path)
            
            info = {
                "pages": doc.page_count,
                "file_size_mb": os.path.getsize(pdf_path) / (1024 * 1024),
                "has_images": False,
                "has_text": False,
                "image_count": 0
            }
            
            # Prüfe erste 3 Seiten für Performance
            for i in range(min(3, doc.page_count)):
                page = doc[i]
                
                # Text prüfen
                text = page.get_text()
                if len(text.strip()) > 50:
                    info["has_text"] = True
                
                # Bilder prüfen
                image_list = page.get_images()
                info["image_count"] += len(image_list)
                if len(image_list) > 0:
                    info["has_images"] = True
            
            doc.close()
            return info
            
        except Exception as e:
            logger.error(f"PDF-Analyse fehlgeschlagen: {e}")
            return {
                "pages": 0,
                "file_size_mb": 0,
                "has_images": False,
                "has_text": False,
                "image_count": 0
            }
    
    def _compress_pdf(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """PDF-Komprimierung mit pypdf"""
        try:
            original_size = os.path.getsize(pdf_path)
            
            # Hole direkte Parameter
            compression_level = params.get('compression_level', 6)
            image_quality = params.get('image_quality', 70)
            
            logger.info(f"Verwende Komprimierung: Level {compression_level}, Bildqualität {image_quality}%")
            
            # Komprimierung durchführen
            success = self._compress_with_pypdf(pdf_path, compression_level, image_quality, params.get('pdf_info', {}))
            
            if success:
                compressed_size = os.path.getsize(pdf_path)
                reduction_percent = (1 - compressed_size/original_size) * 100
                logger.info(f"Komprimierung erfolgreich: {reduction_percent:.1f}% Reduktion")
                return True
            
            return False
            
        except Exception as e:
            logger.error(f"Komprimierung fehlgeschlagen: {e}")
            return False
    
    def _compress_with_pypdf(self, pdf_path: str, compression_level: int, image_quality: int, pdf_info: Dict[str, Any]) -> bool:
        """Komprimierung mit pypdf inklusive Bildkomprimierung"""
        try:
            temp_output = pdf_path + '.compressed'
            
            # Öffne PDF mit pypdf
            writer = PdfWriter(clone_from=pdf_path)
            
            # Komprimiere alle Seiten
            compress_images = pdf_info.get('has_images', False)
            
            logger.info(f"Komprimiere mit Level {compression_level}, Bildqualität: {image_quality}")
            
            for i, page in enumerate(writer.pages):
                try:
                    # Komprimiere Inhaltsströme nur wenn Level > 0
                    if compression_level > 0:
                        page.compress_content_streams(level=compression_level)
                    
                    # Komprimiere Bilder wenn vorhanden
                    if compress_images:
                        for img in page.images:
                            try:
                                # Ersetze Bild mit reduzierter Qualität
                                img.replace(img.image, quality=image_quality)
                            except Exception as img_error:
                                logger.debug(f"Bildkomprimierung übersprungen für Bild auf Seite {i+1}: {img_error}")
                    
                    # Konsolidiere Objekte
                    page.scale_by(1.0)  # Trick um Objekte zu konsolidieren
                    
                except Exception as e:
                    logger.warning(f"Warnung bei Seite {i+1}: {e}")
            
            # Zusätzliche Optimierungen
            try:
                writer.compress_identical_objects()  # Ohne Parameter
                writer.remove_duplication()
            except Exception as e:
                logger.debug(f"Optimierung übersprungen: {e}")
            
            # Schreibe komprimierte PDF
            with open(temp_output, 'wb') as output_file:
                writer.write(output_file)
            
            # Prüfe Ergebnis
            if os.path.exists(temp_output) and os.path.getsize(temp_output) > 0:
                # Prüfe ob komprimierte Datei kleiner ist
                original_size = os.path.getsize(pdf_path)
                compressed_size = os.path.getsize(temp_output)
                
                if compressed_size >= original_size:
                    logger.warning(f"Komprimierte Datei ist nicht kleiner ({compressed_size} >= {original_size})")
                    # Bei pypdf trotzdem verwenden, da optimiert
                
                # Validiere komprimierte PDF
                if self._validate_pdf(temp_output):
                    shutil.move(temp_output, pdf_path)
                    return True
                else:
                    logger.error("Komprimierte PDF ist ungültig")
                    if os.path.exists(temp_output):
                        os.remove(temp_output)
                    return False
            
            return False
            
        except Exception as e:
            logger.error(f"pypdf-Komprimierung fehlgeschlagen: {e}")
            # Aufräumen bei Fehler
            temp_output = pdf_path + '.compressed'
            if os.path.exists(temp_output):
                try:
                    os.remove(temp_output)
                except:
                    pass
            return False
    
    def _build_context(self, pdf_path: str, xml_path: Optional[str], 
                       xml_field_mappings: List[Dict], ocr_zones: List[Dict],
                       original_pdf_path: str = None, input_path: str = None) -> Dict[str, Any]:
        """Baut den Kontext für Variablen-Evaluation auf"""
        context = {}
        
        # Basis-Variablen
        if original_pdf_path:
            context['FileName'] = os.path.splitext(os.path.basename(original_pdf_path))[0]
            context['FileExtension'] = os.path.splitext(original_pdf_path)[1]
            context['FilePath'] = original_pdf_path
            context['FullFileName'] = os.path.basename(original_pdf_path)
        else:
            context['FileName'] = os.path.splitext(os.path.basename(pdf_path))[0]
            context['FileExtension'] = '.pdf'
            context['FilePath'] = pdf_path
            context['FullFileName'] = os.path.basename(pdf_path)
        
        # Erweiterte Dateiinformationen
        context['FileSize'] = str(os.path.getsize(pdf_path)) if os.path.exists(pdf_path) else '0'
        context['FileSizeMB'] = f"{os.path.getsize(pdf_path) / (1024*1024):.2f}" if os.path.exists(pdf_path) else '0'
        
        # Datum und Zeit
        now = datetime.now()
        context['Date'] = now.strftime('%d.%m.%Y')
        context['DateDE'] = now.strftime('%d.%m.%Y')
        context['DateISO'] = now.strftime('%Y-%m-%d')
        context['Time'] = now.strftime('%H:%M:%S')
        context['TimeShort'] = now.strftime('%H-%M-%S')
        context['DateTime'] = now.strftime('%d.%m.%Y %H:%M:%S')
        context['DateTimeISO'] = now.strftime('%Y-%m-%d_%H-%M-%S')
        context['Year'] = now.strftime('%Y')
        context['Month'] = now.strftime('%m')
        context['MonthName'] = now.strftime('%B')
        context['Day'] = now.strftime('%d')
        context['Hour'] = now.strftime('%H')
        context['Minute'] = now.strftime('%M')
        context['Second'] = now.strftime('%S')
        context['Weekday'] = now.strftime('%A')
        context['WeekdayShort'] = now.strftime('%a')
        context['WeekNumber'] = now.strftime('%V')
        context['Timestamp'] = str(int(now.timestamp()))
        
        # Level-Variablen
        if input_path and original_pdf_path:
            # Verwende function_parser für Level-Variablen
            if self.variable_extractor is None:
                from core.function_parser import VariableExtractor
                self.variable_extractor = VariableExtractor()
            
            level_vars = self.variable_extractor.get_level_variables(original_pdf_path, input_path)
            context.update(level_vars)
            context['InputPath'] = input_path
        else:
            # Leere Level-Variablen
            for i in range(6):
                context[f'level{i}'] = ""
        
        # OCR-Text falls vorhanden
        if hasattr(self, '_ocr_cache') and pdf_path in self._ocr_cache:
            context['OCR_FullText'] = self._ocr_cache[pdf_path]
        else:
            # Versuche OCR-Text zu extrahieren
            try:
                full_text = self.ocr_processor.extract_text_from_pdf(pdf_path)
                context['OCR_FullText'] = full_text
                self._ocr_cache[pdf_path] = full_text
            except:
                context['OCR_FullText'] = ""
        
        # OCR-Zonen
        if ocr_zones:
            for zone_dict in ocr_zones:
                zone_name = zone_dict.get('name', 'Unnamed')
                
                # Prüfe ob Zone bereits im Cache ist
                cache_key = f"{pdf_path}_{zone_name}"
                if hasattr(self, '_zone_cache') and cache_key in self._zone_cache:
                    context[zone_name] = self._zone_cache[cache_key]
                else:
                    # Extrahiere Text aus Zone
                    try:
                        page_num = zone_dict.get('page_num', 1)
                        zone_coords = zone_dict.get('zone', (0, 0, 100, 100))
                        
                        zone_text = self.ocr_processor.extract_text_from_zone(
                            pdf_path, page_num, zone_coords
                        )
                        
                        context[zone_name] = zone_text
                        if hasattr(self, '_zone_cache'):
                            self._zone_cache[cache_key] = zone_text
                    except:
                        context[zone_name] = ""
        
        # XML-Felder
        if xml_field_mappings:
            # Wenn XML vorhanden, lade Werte daraus
            if xml_path and os.path.exists(xml_path):
                try:
                    tree = ET.parse(xml_path)
                    root = tree.getroot()
                    
                    fields_elem = root.find(".//Fields")
                    if fields_elem is not None:
                        for field in fields_elem:
                            if field.text:
                                context[field.tag] = field.text
                except Exception as e:
                    logger.error(f"XML-Parsing fehlgeschlagen: {e}")
            
            # Füge alle definierten Felder zum Kontext hinzu (mit leeren Werten wenn nicht vorhanden)
            for mapping in xml_field_mappings:
                field_name = mapping.get('field_name', '')
                if field_name and field_name not in context:
                    context[field_name] = ""  # Leerer String als Default
        
        return context
  
    def _get_error_path(self, doc_pair: DocumentPair, hotfolder: HotfolderConfig) -> str:
        """Bestimmt den Fehlerpfad"""
        context = {}
        if self.function_parser is None:
            from core.function_parser import FunctionParser, VariableExtractor
            self.function_parser = FunctionParser()
            self.variable_extractor = VariableExtractor()
        
        context.update(self.variable_extractor.get_standard_variables())
        context.update(self.variable_extractor.get_file_variables(doc_pair.pdf_path))
        context.update(self.variable_extractor.get_level_variables(doc_pair.pdf_path, hotfolder.input_path))
        context['InputPath'] = hotfolder.input_path
        
        error_path_expr = hotfolder.error_path if hasattr(hotfolder, 'error_path') else ""
        return self.export_processor.get_error_path(error_path_expr, context)
    
    def cleanup_temp_dir(self):
        """Räumt temporäre Dateien auf"""
        try:
            if os.path.exists(self.temp_base_dir):
                import time
                now = time.time()
                
                for work_dir in os.listdir(self.temp_base_dir):
                    dir_path = os.path.join(self.temp_base_dir, work_dir)
                    if os.path.isdir(dir_path):
                        dir_age = now - os.path.getmtime(dir_path)
                        if dir_age > 86400:  # 24 Stunden
                            try:
                                shutil.rmtree(dir_path)
                                logger.info(f"Temporärer Ordner gelöscht: {work_dir}")
                            except Exception as e:
                                logger.error(f"Fehler beim Löschen von {work_dir}: {e}")
        except Exception as e:
            logger.error(f"Fehler beim Aufräumen: {e}")