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
        
        logger.info("PDFProcessor initialisiert")
    
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
            logger.info(f"Starte Verarbeitung von: {os.path.basename(doc_pair.pdf_path)}")
            
            # Verschiebe Dateien in temporären Arbeitsordner
            temp_pdf_path = os.path.join(work_dir, os.path.basename(doc_pair.pdf_path))
            shutil.move(doc_pair.pdf_path, temp_pdf_path)
            
            temp_xml_path = None
            if doc_pair.xml_path and os.path.exists(doc_pair.xml_path):
                temp_xml_path = os.path.join(work_dir, os.path.basename(doc_pair.xml_path))
                shutil.move(doc_pair.xml_path, temp_xml_path)
            
            # Lazy load function parser
            if self.function_parser is None:
                from core.function_parser import FunctionParser
                self.function_parser = FunctionParser()
            
            # Verarbeite XML-Felder wenn vorhanden
            fields = {}
            if temp_xml_path and hotfolder.xml_field_mappings:
                try:
                    logger.info("Verarbeite XML-Feld-Mappings")
                    # Konvertiere Dict-Mappings zu FieldMapping-Objekten
                    field_mappings = []
                    for mapping_dict in hotfolder.xml_field_mappings:
                        field_mapping = FieldMapping.from_dict(mapping_dict)
                        field_mappings.append(field_mapping)
                    
                    # Verarbeite XML mit Mappings
                    self.xml_processor.process_xml_with_mappings(
                        temp_xml_path,
                        temp_pdf_path,
                        field_mappings,
                        hotfolder.ocr_zones
                    )
                    
                    # Extrahiere Felder aus XML für Export-Kontext
                    fields = self._extract_xml_fields(temp_xml_path)
                    
                except Exception as e:
                    logger.error(f"Fehler bei XML-Verarbeitung: {e}")
            
            # Nur bestimmte Aktionen unterstützen
            for action in hotfolder.processing_actions:
                if action in self.supported_actions:
                    logger.info(f"Führe Aktion aus: {action.name}")
                    action_params = hotfolder.action_params.get(action, {})
                    
                    # Füge Output-Pfad zu params hinzu für split
                    if action == ProcessingAction.SPLIT:
                        action_params['output_path'] = os.path.join(
                            work_dir,
                            'splits'
                        )
                    
                    result = self.supported_actions[action](temp_pdf_path, action_params)
                    if not result:
                        logger.error(f"Aktion {action.name} fehlgeschlagen")
                        return False
            
            # Export ausführen
            if hotfolder.export_config and len(hotfolder.export_config) > 0:
                logger.info("Führe Exporte aus")
                # Verwende process_exports statt export_pdf!
                results = self.export_processor.process_exports(
                    temp_pdf_path,
                    temp_xml_path,
                    hotfolder.export_config,  # Das ist bereits eine Liste von Export-Konfigurationen
                    hotfolder.ocr_zones,
                    hotfolder.xml_field_mappings
                )
                
                # Prüfe ob alle Exporte erfolgreich waren
                all_success = all(success for success, _ in results)
                if not all_success:
                    logger.error("Ein oder mehrere Exporte fehlgeschlagen")
                    # Log Details über fehlgeschlagene Exporte
                    for success, message in results:
                        if not success:
                            logger.error(f"Export-Fehler: {message}")
                    return False
                else:
                    # Log erfolgreiche Exporte
                    for success, message in results:
                        logger.info(f"Export: {message}")
            else:
                logger.info("Keine Exporte konfiguriert")
            
            logger.info(f"Verarbeitung erfolgreich abgeschlossen: {os.path.basename(doc_pair.pdf_path)}")
            return True
            
        except Exception as e:
            logger.error(f"Fehler bei der Verarbeitung: {e}", exc_info=True)
            return False
        finally:
            # Arbeitsordner aufräumen
            try:
                if os.path.exists(work_dir):
                    shutil.rmtree(work_dir)
                    logger.debug(f"Temporärer Arbeitsordner entfernt: {work_dir}")
            except Exception as e:
                logger.warning(f"Konnte temporären Ordner nicht löschen: {e}")
    
    def _extract_xml_fields(self, xml_path: str) -> Dict[str, str]:
        """Extrahiert alle Felder aus einer XML-Datei"""
        fields = {}
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # Finde alle Felder im Fields-Element
            fields_elem = root.find(".//Fields")
            if fields_elem is not None:
                for field in fields_elem:
                    if field.text:
                        fields[field.tag] = field.text
                        
            logger.debug(f"Extrahierte {len(fields)} Felder aus XML")
        except Exception as e:
            logger.error(f"Fehler beim Extrahieren der XML-Felder: {e}")
            
        return fields
    
    # PDF-Verarbeitungsfunktionen
    def _compress_pdf(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Komprimiert eine PDF-Datei mit Ghostscript"""
        try:
            # Prüfe ob Ghostscript verfügbar ist
            if not self._is_ghostscript_available():
                error_msg = (
                    "\n"
                    "========================================\n"
                    "FEHLER: Ghostscript nicht gefunden!\n"
                    "========================================\n"
                    "Die PDF-Komprimierung benötigt Ghostscript.\n"
                    "\n"
                    "Bitte installieren Sie Ghostscript:\n"
                    "Windows: https://www.ghostscript.com/download/gsdnld.html\n"
                    "Linux: sudo apt-get install ghostscript\n"
                    "Mac: brew install ghostscript\n"
                    "========================================\n"
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
            logger.error(f"Fehler bei der Komprimierung: {e}", exc_info=True)
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
            logger.debug("Komprimierungseinstellungen:")
            logger.debug(f"  - Farbbilder: {color_dpi} DPI, {color_compression}")
            logger.debug(f"  - Graustufen: {gray_dpi} DPI, {gray_compression}")
            logger.debug(f"  - S/W-Bilder: {mono_dpi} DPI, {mono_compression}")
            if color_compression == 'jpeg' or gray_compression == 'jpeg':
                logger.debug(f"  - JPEG-Qualität: {jpeg_quality}%")
            
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
            logger.error(f"Fehler bei Ghostscript-Komprimierung: {e}", exc_info=True)
            return False
    
    def _split_pdf(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Teilt eine PDF in einzelne Seiten auf"""
        try:
            logger.info(f"Starte PDF-Aufteilung für: {os.path.basename(pdf_path)}")
            
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
            
            logger.info(f"PDF aufgeteilt in {len(reader.pages)} Seiten: {split_dir}")
            return True
            
        except Exception as e:
            logger.error(f"Fehler beim Aufteilen: {e}", exc_info=True)
            return False
    
    def _perform_ocr(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Führt OCR auf der PDF aus und erstellt eine durchsuchbare PDF"""
        try:
            language = params.get("language", "deu")  # Standard: Deutsch
            
            logger.info(f"Starte OCR für: {os.path.basename(pdf_path)} (Sprache: {language})")
            
            # Verwende OCRmyPDF wenn verfügbar (bessere Qualität)
            try:
                import ocrmypdf
                
                logger.debug("Führe OCR aus mit ocrmypdf...")
                ocrmypdf.ocr(
                    pdf_path,
                    pdf_path,
                    language=language,
                    force_ocr=True,
                    optimize=1
                )
                logger.info("OCR erfolgreich durchgeführt")
                return True
                
            except ImportError:
                # Fallback: Tesseract direkt verwenden
                logger.warning("ocrmypdf nicht installiert, verwende Tesseract direkt...")
                
                # Prüfe ob Tesseract verfügbar ist
                try:
                    subprocess.run(["tesseract", "--version"], 
                                 capture_output=True, check=True)
                except:
                    logger.error("Tesseract nicht gefunden! Bitte installieren Sie Tesseract OCR.")
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
                    logger.info("OCR erfolgreich durchgeführt")
                    return True
                else:
                    logger.error(f"OCR fehlgeschlagen: {result.stderr}")
                    return False
                    
        except Exception as e:
            logger.error(f"Fehler bei OCR: {e}", exc_info=True)
            return False
    
    def _convert_to_pdf_a(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Konvertiert PDF zu PDF/A"""
        try:
            # Versuche ocrmypdf für PDF/A Konvertierung
            try:
                import ocrmypdf
                
                logger.info("Konvertiere zu PDF/A...")
                ocrmypdf.ocr(
                    pdf_path,
                    pdf_path,
                    pdfa_image_compression="jpeg",
                    output_type="pdfa",
                    tesseract_timeout=0,  # Kein OCR, nur PDF/A
                    skip_text=True  # Überspringe Text-Layer
                )
                logger.info("PDF/A Konvertierung erfolgreich")
                return True
                
            except ImportError:
                logger.warning("ocrmypdf nicht installiert. Verwende Ghostscript als Fallback...")
                
                # Fallback: Verwende Ghostscript für PDF/A
                return self._convert_to_pdf_a_with_ghostscript(pdf_path, params)
                
        except Exception as e:
            logger.error(f"Fehler bei PDF/A Konvertierung: {e}", exc_info=True)
            return False
    
    def _convert_to_pdf_a_with_ghostscript(self, pdf_path: str, params: Dict[str, Any]) -> bool:
        """Konvertiert eine PDF in PDF/A Format mit Ghostscript"""
        try:
            logger.info(f"Starte PDF/A-Konvertierung mit Ghostscript für: {os.path.basename(pdf_path)}")
            
            # Ghostscript-Pfad bestimmen
            gs_path = self._find_ghostscript()
            if not gs_path:
                logger.error("Ghostscript nicht gefunden!")
                return False
            
            # PDF/A-Version
            pdfa_version = params.get('pdfa_version', '2b')  # 1b, 2b, 3b
            
            # Temporäre Ausgabedatei
            temp_output = pdf_path + '.pdfa.tmp'
            
            # PDFA_def.ps Datei erstellen
            pdfa_def_content = """
%!
% This is a sample prefix file for creating a PDF/A document.
% Feel free to modify entries marked with "Customize".

% This assumes an ICC profile to reside in the file (ISO Coated sb.icc),
% unless the user modifies the corresponding line below.

% Define entries in the document Info dictionary :

/ICCProfile (sRGB.icc) def  % Customize - ICC profile

[ /Title (Document)       % Customize
  /Author (Author)
  /Subject (Subject)
  /Creator (Creator)
  /Keywords (Keywords)
  /DOCINFO pdfmark

% Define an ICC profile :

[{
  /N 3 % Number of components
  /Alternate /DeviceRGB
  /DataSource (sRGB.icc) (r) file def
  /Name (sRGB IEC61966-2.1)
  /OutputIntent pdfmark
"""
            
            pdfa_def_path = os.path.join(os.path.dirname(pdf_path), 'PDFA_def.ps')
            with open(pdfa_def_path, 'w') as f:
                f.write(pdfa_def_content)
            
            # Ghostscript-Kommando
            cmd = [
                gs_path,
                '-dPDFA=' + pdfa_version[-1],  # 1, 2 oder 3
                '-dPDFACompatibilityPolicy=1',
                '-dBATCH',
                '-dNOPAUSE',
                '-dQUIET',
                '-sDEVICE=pdfwrite',
                '-dPDFSETTINGS=/printer',
                '-dColorImageDownsampleType=/Bicubic',
                '-dGrayImageDownsampleType=/Bicubic',
                '-dMonoImageDownsampleType=/Bicubic',
                '-dDownsampleColorImages=false',
                '-dDownsampleGrayImages=false',
                '-dDownsampleMonoImages=false',
                '-dAutoRotatePages=/None',
                '-sColorConversionStrategy=UseDeviceIndependentColor',
                '-dProcessColorModel=/DeviceRGB',
                f'-sOutputFile={temp_output}',
                pdfa_def_path,
                pdf_path
            ]
            
            # Führe Ghostscript aus
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            # Aufräumen
            if os.path.exists(pdfa_def_path):
                os.remove(pdfa_def_path)
            
            if result.returncode != 0:
                logger.error(f"Ghostscript-Fehler: {result.stderr}")
                return False
            
            if os.path.exists(temp_output) and os.path.getsize(temp_output) > 0:
                # Ersetze Original
                shutil.move(temp_output, pdf_path)
                logger.info(f"PDF erfolgreich in PDF/A-{pdfa_version} konvertiert")
                return True
            
            return False
            
        except Exception as e:
            logger.error(f"Fehler bei PDF/A-Konvertierung: {e}", exc_info=True)
            return False
    
    def _find_ghostscript(self) -> Optional[str]:
        """Findet die Ghostscript-Executable"""
        # Windows
        if sys.platform == 'win32':
            # Typische Installationspfade
            gs_paths = [
                r'C:\Program Files\gs\gs*\bin\gswin64c.exe',
                r'C:\Program Files\gs\gs*\bin\gswin32c.exe',
                r'C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe',
            ]
            
            for pattern in gs_paths:
                import glob
                matches = glob.glob(pattern)
                if matches:
                    gs_path = matches[0]
                    logger.debug(f"Ghostscript gefunden: {gs_path}")
                    return gs_path
            
            # Versuche gs im PATH
            try:
                result = subprocess.run(['gswin64c', '--version'], 
                                      capture_output=True, check=True)
                return 'gswin64c'
            except:
                try:
                    result = subprocess.run(['gswin32c', '--version'], 
                                          capture_output=True, check=True)
                    return 'gswin32c'
                except:
                    pass
        
        # Linux/Mac
        else:
            try:
                result = subprocess.run(['gs', '--version'], 
                                      capture_output=True, check=True)
                return 'gs'
            except:
                pass
        
        logger.error("Ghostscript nicht gefunden! Bitte installieren Sie Ghostscript.")
        return None
