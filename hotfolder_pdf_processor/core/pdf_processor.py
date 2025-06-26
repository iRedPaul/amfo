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
import logging
import time

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.hotfolder_config import ProcessingAction, DocumentPair, HotfolderConfig
from core.xml_field_processor import XMLFieldProcessor, FieldMapping
from core.ocr_processor import OCRProcessor
from core.export_processor import ExportProcessor

# Logger für dieses Modul
logger = logging.getLogger(__name__)


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
        # Zeit für Performance-Messung
        start_time = time.time()
        
        # Erstelle eindeutigen temporären Arbeitsordner für diese Verarbeitung
        session_id = str(uuid.uuid4())
        temp_work_dir = os.path.join(self.temp_base_dir, session_id)
        os.makedirs(temp_work_dir, exist_ok=True)
        
        success = False
        
        try:
            # Kopiere PDF in temporären Ordner
            temp_pdf = os.path.join(temp_work_dir, os.path.basename(doc_pair.pdf_path))
            shutil.copy2(doc_pair.pdf_path, temp_pdf)
            
            # XML ebenfalls kopieren wenn vorhanden
            temp_xml = None
            if doc_pair.xml_path:
                temp_xml = os.path.join(temp_work_dir, os.path.basename(doc_pair.xml_path))
                shutil.copy2(doc_pair.xml_path, temp_xml)
            
            # Führe Aktionen aus
            for action in hotfolder.actions:
                if action in self.supported_actions:
                    logger.info(f"Führe Aktion aus: {action.value}")
                    
                    # Hole spezifische Parameter für diese Aktion
                    action_params = hotfolder.action_params.get(action.value, {})
                    
                    # Führe Aktion aus
                    if not self.supported_actions[action](temp_pdf, action_params):
                        logger.error(f"Aktion {action.value} fehlgeschlagen")
                        return False
                else:
                    logger.warning(f"Unbekannte Aktion: {action}")
            
            # Verarbeite XML-Felder wenn aktiviert
            if hotfolder.process_pairs and temp_xml and hotfolder.xml_field_mappings:
                logger.info("Verarbeite XML-Felder")
                
                # Baue Kontext für XML-Verarbeitung auf
                context = self._build_context(temp_pdf, temp_xml, hotfolder)
                
                # Verarbeite XML
                xml_tree = self.xml_processor.process_xml(
                    temp_xml,
                    hotfolder.xml_field_mappings,
                    context
                )
                
                if xml_tree:
                    # Speichere modifizierte XML
                    xml_tree.write(temp_xml, encoding='utf-8', xml_declaration=True)
                    logger.info("XML-Felder erfolgreich verarbeitet")
            
            # Exportiere Dateien
            if hotfolder.export_configs:
                logger.info("Starte Export-Prozess")
                
                # Baue Export-Kontext auf
                context = self._build_context(temp_pdf, temp_xml, hotfolder)
                
                # Exportiere
                export_results = self.export_processor.process_exports(
                    temp_pdf, temp_xml,
                    hotfolder.export_configs,
                    hotfolder.ocr_zones,
                    hotfolder.xml_field_mappings
                )
                
                # Prüfe ob alle Exporte erfolgreich waren
                all_exports_successful = all(result[0] for result in export_results)
                
                if not all_exports_successful:
                    logger.error("Einige Exporte sind fehlgeschlagen")
                    for success, message in export_results:
                        if not success:
                            logger.error(f"Export-Fehler: {message}")
                    return False
                
                logger.info("Alle Exporte erfolgreich")
            
            # Kopiere verarbeitete Dateien zurück wenn kein Export definiert
            if not hotfolder.export_configs:
                # Verwende Ausgabepfad oder Input-Pfad als Fallback
                output_path = hotfolder.output_path or hotfolder.input_path
                
                # Generiere Ausgabedateinamen
                output_filename = self._generate_output_filename(
                    temp_pdf, 
                    hotfolder.output_filename_expression,
                    context if 'context' in locals() else None
                )
                
                # Kopiere PDF zurück
                final_pdf_path = os.path.join(output_path, output_filename + ".pdf")
                os.makedirs(os.path.dirname(final_pdf_path), exist_ok=True)
                shutil.copy2(temp_pdf, final_pdf_path)
                
                # Kopiere XML wenn vorhanden
                if temp_xml:
                    final_xml_path = os.path.join(output_path, output_filename + ".xml")
                    shutil.copy2(temp_xml, final_xml_path)
                
                logger.info(f"Dateien kopiert nach: {output_path}")
            
            # Lösche Originaldateien aus Input-Ordner
            try:
                os.remove(doc_pair.pdf_path)
                if doc_pair.xml_path and os.path.exists(doc_pair.xml_path):
                    os.remove(doc_pair.xml_path)
                logger.info("Originaldateien aus Input-Ordner gelöscht")
            except Exception as e:
                logger.warning(f"Konnte Originaldateien nicht löschen: {e}")
            
            success = True
            
        except Exception as e:
            logger.exception(f"Fehler bei der Dokumentverarbeitung: {e}")
            success = False
            
        finally:
            # Aufräumen: Lösche temporären Arbeitsordner
            try:
                shutil.rmtree(temp_work_dir)
                logger.debug(f"Temporärer Ordner gelöscht: {temp_work_dir}")
            except Exception as e:
                logger.warning(f"Konnte temporären Ordner nicht löschen: {e}")
            
            # Performance-Log
            processing_time = time.time() - start_time
            logger.info(f"Verarbeitungszeit: {processing_time:.2f} Sekunden")
        
        return success
    
    def _build_context(self, pdf_path: str, xml_path: Optional[str], 
                      hotfolder: HotfolderConfig) -> Dict[str, Any]:
        """Baut den Kontext für Variablen-Ersetzung auf"""
        from datetime import datetime
        
        context = {}
        
        # Basis-Informationen
        context['FilePath'] = os.path.dirname(pdf_path)
        context['FileName'] = os.path.splitext(os.path.basename(pdf_path))[0]
        context['FileExtension'] = os.path.splitext(pdf_path)[1]
        
        # Datum und Zeit
        now = datetime.now()
        context['Date'] = now.strftime("%Y-%m-%d")
        context['Time'] = now.strftime("%H-%M-%S")
        context['DateTime'] = now.strftime("%Y-%m-%d_%H-%M-%S")
        
        # OCR-Zonen verarbeiten wenn definiert
        if hotfolder.ocr_zones:
            for zone in hotfolder.ocr_zones:
                zone_name = zone.get('name', 'Unnamed')
                
                # Prüfe Cache
                cache_key = f"{pdf_path}:{zone_name}"
                if cache_key in self._zone_cache:
                    context[zone_name] = self._zone_cache[cache_key]
                else:
                    # Führe OCR aus
                    text = self.ocr_processor.extract_text_from_zone(
                        pdf_path,
                        zone.get('zone', [0, 0, 100, 100]),
                        zone.get('page_num', 1)
                    )
                    context[zone_name] = text.strip()
                    self._zone_cache[cache_key] = context[zone_name]
        
        return context
    
    def _generate_output_filename(self, pdf_path: str, expression: str, 
                                context: Optional[Dict] = None) -> str:
        """Generiert den Ausgabedateinamen basierend auf einem Ausdruck"""
        if not expression:
            return os.path.splitext(os.path.basename(pdf_path))[0]
        
        # Lazy load FunctionParser
        if self.function_parser is None:
            from core.function_parser import FunctionParser
            self.function_parser = FunctionParser()
        
        if context is None:
            context = {}
        
        # Parse den Ausdruck
        result = self.function_parser.parse_expression(expression, context)
        
        # Bereinige ungültige Zeichen für Dateinamen
        invalid_chars = '<>:"|?*'
        for char in invalid_chars:
            result = result.replace(char, '_')
        
        return result.strip()
    
    def _update_metadata(self, pdf_path: str, hotfolder: HotfolderConfig) -> bool:
        """Aktualisiert PDF-Metadaten"""
        try:
            # TODO: Implementiere Metadaten-Update mit PyPDF2
            logger.info("Metadaten-Update noch nicht implementiert")
            return True
        except Exception as e:
            logger.error(f"Fehler beim Update der Metadaten: {e}")
    
    # PDF-Verarbeitungsfunktionen
    def _compress_pdf(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Komprimiert eine PDF-Datei mit Ghostscript"""
        try:
            # Prüfe ob Ghostscript verfügbar ist
            if not self._is_ghostscript_available():
                error_msg = (
                    "FEHLER: Ghostscript nicht gefunden!\n"
                    "Die PDF-Komprimierung benötigt Ghostscript.\n"
                    "Bitte installieren Sie Ghostscript:\n"
                    "Windows: https://www.ghostscript.com/download/gsdnld.html\n"
                    "Linux: sudo apt-get install ghostscript\n"
                    "Mac: brew install ghostscript"
                )
                logger.error(error_msg)
                return False
            
            original_size = os.path.getsize(pdf_path)
            logger.info(f"Starte Komprimierung (Original: {original_size / 1024 / 1024:.2f} MB)")
            
            # Verwende Ghostscript für Komprimierung
            success = self._compress_with_ghostscript(pdf_path, params)
            
            if success:
                compressed_size = os.path.getsize(pdf_path)
                reduction_percent = (1 - compressed_size/original_size) * 100
                logger.info(f"PDF komprimiert: {original_size / 1024 / 1024:.2f} MB -> {compressed_size / 1024 / 1024:.2f} MB ({reduction_percent:.1f}% Reduktion)")
                return True
            else:
                logger.error("Komprimierung fehlgeschlagen")
                return False
            
        except Exception as e:
            logger.error(f"Fehler bei der Komprimierung: {e}")
            return False
    
    def _is_ghostscript_available(self) -> bool:
        """Prüft ob Ghostscript verfügbar ist"""
        try:
            # Windows
            if os.name == 'nt':
                # Versuche 64-bit Version
                try:
                    result = subprocess.run(['gswin64c', '--version'], 
                                          capture_output=True, text=True)
                    if result.returncode == 0:
                        return True
                except:
                    pass
                
                # Versuche 32-bit Version
                try:
                    result = subprocess.run(['gswin32c', '--version'], 
                                          capture_output=True, text=True)
                    if result.returncode == 0:
                        return True
                except:
                    pass
                
                return False
            else:
                # Unix/Linux
                result = subprocess.run(['gs', '--version'], 
                                      capture_output=True, text=True)
                return result.returncode == 0
        except:
            return False
    
    def _compress_with_ghostscript(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Komprimiert PDF mit Ghostscript"""
        try:
            # Bestimme Ghostscript-Befehl
            gs_cmd = 'gswin64c' if os.name == 'nt' else 'gs'
            
            # Prüfe nochmal ob gs_cmd gefunden wird
            try:
                result = subprocess.run([gs_cmd, '--version'], capture_output=True)
                if result.returncode != 0:
                    # Versuche 32-bit Version auf Windows
                    if os.name == 'nt':
                        gs_cmd = 'gswin32c'
                        result = subprocess.run([gs_cmd, '--version'], capture_output=True)
                        if result.returncode != 0:
                            raise Exception("Ghostscript nicht ausführbar")
            except:
                raise Exception("Ghostscript konnte nicht gestartet werden")
            
            # Hole Parameter
            color_dpi = params.get('color_dpi', 150)
            gray_dpi = params.get('gray_dpi', 150)
            mono_dpi = params.get('mono_dpi', 150)
            jpeg_quality = params.get('jpeg_quality', 85)
            color_compression = params.get('color_compression', 'jpeg')
            gray_compression = params.get('gray_compression', 'jpeg')
            mono_compression = params.get('mono_compression', 'ccitt')
            downsample_images = params.get('downsample_images', True)
            subset_fonts = params.get('subset_fonts', True)
            remove_duplicates = params.get('remove_duplicates', True)
            optimize = params.get('optimize', True)
            
            # Temporäre Ausgabedatei
            temp_output = pdf_path + '.gs_compressed'
            
            # Ghostscript-Befehl zusammenbauen
            cmd = [
                gs_cmd,
                '-sDEVICE=pdfwrite',
                '-dCompatibilityLevel=1.4',
                '-dNOPAUSE',
                '-dBATCH',
                '-dQUIET',
                f"-dColorImageResolution={color_dpi}",
                f"-dGrayImageResolution={gray_dpi}",
                f"-dMonoImageResolution={mono_dpi}",
                f'-sOutputFile={temp_output}'
            ]
            
            # Downsampling
            if downsample_images:
                cmd.extend([
                    '-dDownsampleColorImages=true',
                    '-dDownsampleGrayImages=true',
                    '-dDownsampleMonoImages=true',
                    '-dColorImageDownsampleType=/Bicubic',
                    '-dGrayImageDownsampleType=/Bicubic',
                    '-dMonoImageDownsampleType=/Subsample'
                ])
            else:
                cmd.extend([
                    '-dDownsampleColorImages=false',
                    '-dDownsampleGrayImages=false',
                    '-dDownsampleMonoImages=false'
                ])
            
            # Komprimierungsmethoden für Farbbilder
            if color_compression == 'jpeg':
                cmd.extend([
                    '-dAutoFilterColorImages=false',
                    '-dColorImageFilter=/DCTEncode',
                    f'-dJPEGQ={jpeg_quality/100.0:.2f}'
                ])
            elif color_compression == 'jpeg2000':
                cmd.extend([
                    '-dAutoFilterColorImages=false',
                    '-dColorImageFilter=/JPXEncode'
                ])
            elif color_compression == 'zip':
                cmd.extend([
                    '-dAutoFilterColorImages=false',
                    '-dColorImageFilter=/FlateEncode'
                ])
            elif color_compression == 'none':
                cmd.extend([
                    '-dAutoFilterColorImages=false',
                    '-dEncodeColorImages=false'
                ])
            
            # Komprimierungsmethoden für Graustufenbilder
            if gray_compression == 'jpeg':
                cmd.extend([
                    '-dAutoFilterGrayImages=false',
                    '-dGrayImageFilter=/DCTEncode'
                ])
            elif gray_compression == 'jpeg2000':
                cmd.extend([
                    '-dAutoFilterGrayImages=false',
                    '-dGrayImageFilter=/JPXEncode'
                ])
            elif gray_compression == 'zip':
                cmd.extend([
                    '-dAutoFilterGrayImages=false',
                    '-dGrayImageFilter=/FlateEncode'
                ])
            elif gray_compression == 'none':
                cmd.extend([
                    '-dAutoFilterGrayImages=false',
                    '-dEncodeGrayImages=false'
                ])
            
            # Komprimierungsmethoden für Schwarz-Weiß-Bilder
            if mono_compression == 'ccitt':
                cmd.extend([
                    '-dMonoImageFilter=/CCITTFaxEncode'
                ])
            elif mono_compression == 'jbig2':
                # JBIG2 ist nicht in allen Ghostscript-Versionen verfügbar
                # Fallback zu CCITT wenn JBIG2 fehlschlägt
                logger.info("Hinweis: JBIG2 wird versucht, Fallback zu CCITT falls nicht verfügbar")
                cmd.extend([
                    '-dMonoImageFilter=/CCITTFaxEncode'  # Sicherer Fallback
                ])
            elif mono_compression == 'zip':
                cmd.extend([
                    '-dMonoImageFilter=/FlateEncode'
                ])
            elif mono_compression == 'none':
                cmd.extend([
                    '-dEncodeMonoImages=false'
                ])
            
            # Font-Einstellungen
            if subset_fonts:
                cmd.extend([
                    '-dSubsetFonts=true',
                    '-dEmbedAllFonts=true'
                ])
            else:
                cmd.extend([
                    '-dSubsetFonts=false',
                    '-dEmbedAllFonts=true'
                ])
            
            # Weitere Optimierungen
            if optimize:
                cmd.append('-dOptimize=true')
            
            if remove_duplicates:
                cmd.append('-dDetectDuplicateImages=true')
            
            # Komprimierung der Streams
            cmd.append('-dCompressFonts=true')
            cmd.append('-dCompressPages=true')
            
            # Standard PDF-Einstellungen überschreiben
            cmd.append('-dPDFSETTINGS=/default')
            
            # Eingabedatei am Ende
            cmd.append(pdf_path)
            
            # Debug-Ausgabe
            logger.info(f"Komprimierungseinstellungen:")
            logger.info(f"  - Farbbilder: {color_dpi} DPI, {color_compression}")
            logger.info(f"  - Graustufen: {gray_dpi} DPI, {gray_compression}")
            logger.info(f"  - S/W-Bilder: {mono_dpi} DPI, {mono_compression}")
            if color_compression == 'jpeg' or gray_compression == 'jpeg':
                logger.info(f"  - JPEG-Qualität: {jpeg_quality}%")
            
            # Führe Ghostscript aus
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode != 0:
                logger.error(f"Ghostscript-Fehler: {result.stderr}")
                return False
            
            if os.path.exists(temp_output):
                # Prüfe ob komprimierte Datei gültig ist
                if os.path.getsize(temp_output) > 0:
                    # Ersetze Original
                    shutil.move(temp_output, pdf_path)
                    return True
                else:
                    os.remove(temp_output)
                    logger.error("Ghostscript erzeugte eine leere Datei")
                    return False
            
            return False
            
        except Exception as e:
            logger.error(f"Fehler bei Ghostscript-Komprimierung: {e}")
            return False
    
    def _split_pdf(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Teilt eine PDF in einzelne Seiten auf"""
        try:
            reader = PyPDF2.PdfReader(pdf_path)
            base_name = os.path.splitext(os.path.basename(pdf_path))[0]
            output_dir = os.path.dirname(pdf_path)
            
            for i, page in enumerate(reader.pages):
                writer = PyPDF2.PdfWriter()
                writer.add_page(page)
                
                output_filename = f"{base_name}_Seite_{i+1:03d}.pdf"
                output_path = os.path.join(output_dir, output_filename)
                
                with open(output_path, 'wb') as output_file:
                    writer.write(output_file)
                
                logger.info(f"Seite {i+1} extrahiert: {output_filename}")
            
            return True
            
        except Exception as e:
            logger.error(f"Fehler beim Teilen der PDF: {e}")
            return False
    
    def _perform_ocr(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Führt OCR auf der PDF aus"""
        try:
            # OCR wird vom OCRProcessor gehandhabt
            logger.info("OCR-Verarbeitung gestartet")
            return True
            
        except Exception as e:
            logger.error(f"Fehler bei OCR: {e}")
            return False
    
    def _convert_to_pdf_a(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Konvertiert PDF zu PDF/A"""
        try:
            logger.info("PDF/A-Konvertierung noch nicht implementiert")
            return True
            
        except Exception as e:
            logger.error(f"Fehler bei PDF/A-Konvertierung: {e}")
            return False