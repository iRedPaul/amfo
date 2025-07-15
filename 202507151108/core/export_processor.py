"""
Export-Prozessor - Vereinfachte Version mit nur 3 Formaten
"""
import os
import shutil
import tempfile
import json
import smtplib
import ssl
import subprocess
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import sys
import xml.etree.ElementTree as ET
from datetime import datetime
import logging
import ocrmypdf
import re
import fitz

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.export_config import ExportConfig, ExportFormat, ExportMethod, EmailConfig, ExportSettings, AuthMethod
from core.function_parser import FunctionParser, VariableExtractor
from core.ocr_processor import OCRProcessor
from core.oauth2_manager import OAuth2Manager, get_token_storage

logger = logging.getLogger(__name__)


class ExportProcessor:
    """Vereinfachter Export-Prozessor mit nur 3 Formaten"""

    def __init__(self):
        self.function_parser = FunctionParser()
        self.variable_extractor = VariableExtractor()
        self.ocr_processor = OCRProcessor()
        self._export_settings = None
        self._ocr_cache = {}

    def _get_unique_filename(self, filepath: str) -> str:
        """Generiert eindeutigen Dateinamen mit Nummerierung (_1, _2, etc.)"""
        if not os.path.exists(filepath):
            return filepath

        directory = os.path.dirname(filepath)
        filename = os.path.basename(filepath)
        name, ext = os.path.splitext(filename)

        # Zähler für eindeutigen Dateinamen
        counter = 1
        while True:
            new_filename = f"{name}_{counter}{ext}"
            new_filepath = os.path.join(directory, new_filename)
            
            if not os.path.exists(new_filepath):
                return new_filepath
                
            counter += 1
            
            # Sicherheitsgrenze
            if counter > 9999:
                # Fallback auf Zeitstempel wenn zu viele Dateien
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                new_filename = f"{name}_{timestamp}{ext}"
                return os.path.join(directory, new_filename)

    def process_exports(self, pdf_path: str, xml_path: Optional[str],
                        export_configs: List[Dict], ocr_zones: List[Dict] = None,
                        xml_field_mappings: List[Dict] = None,
                        original_pdf_path: str = None,
                        input_path: str = None, 
                        has_stamps: bool = False,
                        compression_enabled: bool = False) -> List[Tuple[bool, str]]:
        """
        Führt alle konfigurierten Exporte durch
        """
        results = []

        # Validiere PDF vor Export
        if not self._validate_pdf(pdf_path):
            return [(False, "PDF-Validierung fehlgeschlagen - Export abgebrochen")]

        # Baue Kontext auf
        context = self._build_context(pdf_path, xml_path, ocr_zones, 
                                    xml_field_mappings, input_path)
        
        # Überschreibe mit Original-Pfad-Informationen wenn vorhanden
        if original_pdf_path and input_path:
            level_vars = self.variable_extractor.get_level_variables(original_pdf_path, input_path)
            context.update(level_vars)
            
            original_filename = os.path.splitext(os.path.basename(original_pdf_path))[0]
            original_fullname = os.path.basename(original_pdf_path)
            
            # Prüfe ob temporärer Name eine UUID ist
            temp_filename = os.path.splitext(os.path.basename(pdf_path))[0]
            uuid_pattern = r'^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$'
            if re.match(uuid_pattern, temp_filename):
                context['FileName'] = original_filename
                context['FullFileName'] = original_fullname

        # Verarbeite jeden Export
        for export_dict in export_configs:
            try:
                if isinstance(export_dict, ExportConfig):
                    export = export_dict
                else:
                    export = ExportConfig.from_dict(export_dict)

                if not export.enabled:
                    continue

                # Log Export-Start
                logger.info(f"Starte Export '{export.name}' ({export.export_format.value})")

                success, message = self._process_single_export(
                    pdf_path, xml_path, export, context, has_stamps, compression_enabled
                )
                results.append((success, message))

                if success:
                    logger.info(f"Export '{export.name}' erfolgreich: {message}")
                else:
                    logger.error(f"Export '{export.name}' fehlgeschlagen: {message}")

            except Exception as e:
                logger.exception(f"Kritischer Fehler bei Export '{export_dict.get('name', 'Unbekannt')}'")
                results.append((False, f"Export-Fehler: {str(e)}"))

        return results

    def _validate_pdf(self, pdf_path: str) -> bool:
        """Validiert PDF vor Export"""
        try:
            doc = fitz.open(pdf_path)
            if doc.page_count == 0:
                doc.close()
                return False
            
            # Teste erste Seite
            page = doc[0]
            _ = page.get_pixmap(alpha=False)
            
            doc.close()
            return True
            
        except Exception as e:
            logger.error(f"PDF-Validierung fehlgeschlagen: {e}")
            return False

    def _build_context(self, pdf_path: str, xml_path: Optional[str],
                       ocr_zones: List[Dict] = None,
                       xml_field_mappings: List[Dict] = None,
                       input_path: str = None) -> Dict[str, Any]:
        """Baut erweiterten Kontext für Variablen auf"""
        context = {}

        # Basis-Dateiinformationen
        context['FilePath'] = os.path.dirname(pdf_path)
        context['FileName'] = os.path.splitext(os.path.basename(pdf_path))[0]
        context['FileExtension'] = os.path.splitext(pdf_path)[1]
        context['FullFileName'] = os.path.basename(pdf_path)
        context['FullPath'] = pdf_path
        context['FileSize'] = str(os.path.getsize(pdf_path)) if os.path.exists(pdf_path) else '0'
        context['FileSizeMB'] = f"{os.path.getsize(pdf_path) / (1024*1024):.2f}" if os.path.exists(pdf_path) else '0'

        # Erweiterte Zeitvariablen
        now = datetime.now()
        context.update({
            'Date': now.strftime('%Y-%m-%d'),
            'DateDE': now.strftime('%d.%m.%Y'),
            'Time': now.strftime('%H-%M-%S'),
            'TimeColon': now.strftime('%H:%M:%S'),
            'DateTime': now.strftime('%Y-%m-%d_%H-%M-%S'),
            'DateTimeDE': now.strftime('%d.%m.%Y %H:%M:%S'),
            'Year': now.strftime('%Y'),
            'Month': now.strftime('%m'),
            'MonthName': now.strftime('%B'),
            'Day': now.strftime('%d'),
            'Hour': now.strftime('%H'),
            'Minute': now.strftime('%M'),
            'Second': now.strftime('%S'),
            'Weekday': now.strftime('%A'),
            'WeekdayShort': now.strftime('%a'),
            'WeekNumber': now.strftime('%V'),
            'Timestamp': str(int(now.timestamp()))
        })

        # Level-Variablen
        if input_path:
            level_vars = self.variable_extractor.get_level_variables(pdf_path, input_path)
            context.update(level_vars)
            context['InputPath'] = input_path
        else:
            for i in range(6):
                context[f'level{i}'] = ""

        # OCR-Volltext
        if ocr_zones:
            if pdf_path not in self._ocr_cache:
                full_text = self.ocr_processor.extract_text_from_pdf(pdf_path)
                self._ocr_cache[pdf_path] = full_text
            else:
                full_text = self._ocr_cache[pdf_path]

            context['OCR_FullText'] = full_text

            # OCR-Zonen
            for zone_dict in ocr_zones:
                zone_name = zone_dict.get('name', 'Unnamed')
                page_num = zone_dict.get('page_num', zone_dict.get('page', 1))
                
                if 'zone' in zone_dict:
                    zone_coords = zone_dict['zone']
                else:
                    zone_coords = (
                        zone_dict.get('x', 0),
                        zone_dict.get('y', 0),
                        zone_dict.get('width', 100),
                        zone_dict.get('height', 100)
                    )

                zone_text = self.ocr_processor.extract_text_from_zone(
                    pdf_path, page_num, zone_coords
                )
                
                context[zone_name] = zone_text
                if not zone_name.startswith('OCR_'):
                    context[f'OCR_{zone_name}'] = zone_text

        # XML-Felder
        if xml_field_mappings and xml_path and os.path.exists(xml_path):
            try:
                tree = ET.parse(xml_path)
                root = tree.getroot()
                
                for mapping in xml_field_mappings:
                    field_name = mapping.get('field_name', '')
                    if field_name:
                        field_elem = root.find(f".//Fields/{field_name}")
                        if field_elem is not None and field_elem.text:
                            context[field_name] = field_elem.text
            except Exception as e:
                logger.error(f"XML-Parsing fehlgeschlagen: {e}")

        context['OutputPath'] = os.path.dirname(pdf_path)

        return context

    def _process_single_export(self, pdf_path: str, xml_path: Optional[str],
                              export: ExportConfig, context: Dict[str, Any],
                              has_stamps: bool = False, compression_enabled: bool = False) -> Tuple[bool, str]:
        """Verarbeitet einzelnen Export mit Fehlerbehandlung"""
        try:
            if export.export_method == ExportMethod.FILE:
                return self._export_to_file(pdf_path, xml_path, export, context, has_stamps, compression_enabled)
            elif export.export_method == ExportMethod.EMAIL:
                return self._export_to_email(pdf_path, xml_path, export, context, has_stamps, compression_enabled)
            else:
                return False, f"Export-Methode {export.export_method} nicht unterstützt"

        except Exception as e:
            logger.exception(f"Export-Fehler bei {export.name}")
            return False, f"Fehler: {str(e)}"

    def _export_to_file(self, pdf_path: str, xml_path: Optional[str],
                        export: ExportConfig, context: Dict[str, Any],
                        has_stamps: bool = False, compression_enabled: bool = False) -> Tuple[bool, str]:
        """Datei-Export mit nur 3 Formaten"""
        # Evaluiere Pfad und Dateiname
        export_path = self.function_parser.parse_and_evaluate(
            export.export_path_expression, context
        ) if export.export_path_expression else ""

        if not export_path:
            export_path = self.get_error_path("", context)

        export_filename = self.function_parser.parse_and_evaluate(
            export.export_filename_expression, context
        )
        export_filename = self.sanitize_filename(export_filename)

        # Erstelle Export-Pfad
        os.makedirs(export_path, exist_ok=True)

        # Format-spezifische Verarbeitung
        if export.export_format == ExportFormat.PDF:
            return self._export_pdf(pdf_path, export_path, export_filename, 
                                   export.format_params, has_stamps, compression_enabled)
            
        elif export.export_format == ExportFormat.SEARCHABLE_PDF_A:
            return self._export_pdf_a(pdf_path, export_path, export_filename, export.format_params)
            
        elif export.export_format == ExportFormat.XML:
            return self._export_xml(xml_path, export_path, export_filename)
            
        else:
            return False, f"Format {export.export_format} nicht implementiert"

    def _export_pdf(self, pdf_path: str, export_path: str, filename: str, 
                    params: Dict[str, Any], has_stamps: bool = False, 
                    compression_enabled: bool = False) -> Tuple[bool, str]:
        """Exportiert PDF (Original) - nur verschieben es sei denn Stempel oder Komprimierung"""
        try:
            output_file = os.path.join(export_path, f"{filename}.pdf")
            output_file = self._get_unique_filename(output_file)
            
            # Prüfe ob Nachbearbeitung nötig ist
            needs_processing = has_stamps or compression_enabled or params.get('update_metadata', False)
            
            if not needs_processing:
                # Einfach kopieren
                shutil.copy2(pdf_path, output_file)
                return True, f"PDF (Original) exportiert: {os.path.basename(output_file)}"
            else:
                # Nachbearbeitung erforderlich
                # Kopiere erst einmal
                shutil.copy2(pdf_path, output_file)
                
                # Optional: Metadaten aktualisieren
                if params.get('update_metadata', False):
                    self._update_pdf_metadata(output_file, params.get('metadata', {}))
                
                # Hinweis über Nachbearbeitung
                processing_info = []
                if has_stamps:
                    processing_info.append("Stempel")
                if compression_enabled:
                    processing_info.append("Komprimierung")
                if params.get('update_metadata', False):
                    processing_info.append("Metadaten")
                    
                info_text = f"PDF exportiert mit {', '.join(processing_info)}"
                return True, f"{info_text}: {os.path.basename(output_file)}"
            
        except Exception as e:
            logger.error(f"PDF-Export fehlgeschlagen: {e}")
            return False, f"PDF-Export-Fehler: {str(e)}"

    def _export_pdf_a(self, pdf_path: str, export_path: str, filename: str,
                      params: Dict[str, Any]) -> Tuple[bool, str]:
        """Exportiert als durchsuchbares PDF/A"""
        try:
            output_file = os.path.join(export_path, f"{filename}.pdf")
            output_file = self._get_unique_filename(output_file)
            
            logger.info(f"Starte PDF/A-Export (durchsuchbar)")
            
            # Prüfe zuerst, ob die PDF bereits Text hat
            has_text = self._check_pdf_has_text(pdf_path)
            logger.info(f"PDF hat {'bereits' if has_text else 'keinen'} Text")
            
            if not has_text:
                # OCR ist erforderlich - verwende OCRmyPDF
                args = {
                    "output_type": "pdfa-2",
                    "pdfa_image_compression": "lossless",
                    "optimize": 0,
                    "oversample": 0,
                    "jbig2_lossy": False,
                    "force_ocr": True,
                    "skip_text": False,
                    "language": params.get('language', 'deu'),
                    "rotate_pages": True,
                    "deskew": False,
                    "clean": False,
                    "tesseract_timeout": params.get('timeout', 600),
                    "invalidate_digital_signatures": True,
                }
                
                # Führe OCR aus
                result = ocrmypdf.ocr(pdf_path, output_file, **args)
                
                if result == ocrmypdf.ExitCode.ok:
                    return True, f"PDF/A (Durchsuchbar) exportiert: {os.path.basename(output_file)}"
                else:
                    logger.error(f"OCRmyPDF fehlgeschlagen: {result}")
                    # Fallback: Kopiere Original
                    shutil.copy2(pdf_path, output_file)
                    return True, f"PDF exportiert (ohne PDF/A): {os.path.basename(output_file)}"
                    
            else:
                # Hat bereits Text - verwende OCRmyPDF ohne OCR für PDF/A-Konvertierung
                args = {
                    "output_type": "pdfa-2",
                    "pdfa_image_compression": "lossless",
                    "optimize": 0,
                    "skip_text": True,
                    "tesseract_timeout": 0,
                    "force_ocr": False,
                }
                result = ocrmypdf.ocr(pdf_path, output_file, **args)
                if result == ocrmypdf.ExitCode.ok:
                    return True, f"PDF/A (Durchsuchbar) exportiert: {os.path.basename(output_file)}"
                else:
                    shutil.copy2(pdf_path, output_file)
                    return True, f"PDF exportiert (ohne PDF/A): {os.path.basename(output_file)}"

        except Exception as e:
            logger.exception("PDF/A-Export fehlgeschlagen")
            # Letzter Fallback: Kopiere die Datei
            output_file = os.path.join(export_path, f"{filename}.pdf")
            output_file = self._get_unique_filename(output_file)
            shutil.copy2(pdf_path, output_file)
            return True, f"PDF exportiert (ohne Konvertierung): {os.path.basename(output_file)}"

    def _check_pdf_has_text(self, pdf_path: str) -> bool:
        """Prüft ob eine PDF bereits Text enthält"""
        try:
            doc = fitz.open(pdf_path)
            
            # Prüfe die ersten paar Seiten
            pages_to_check = min(3, doc.page_count)
            total_text = ""
            
            for i in range(pages_to_check):
                page = doc[i]
                text = page.get_text()
                total_text += text
            
            doc.close()
            
            # Wenn mehr als 50 Zeichen Text gefunden wurden, hat die PDF Text
            return len(total_text.strip()) > 50
            
        except Exception as e:
            logger.error(f"Fehler beim Prüfen auf Text: {e}")
            return False

    def _export_xml(self, xml_path: Optional[str], export_path: str, 
                    filename: str) -> Tuple[bool, str]:
        """Exportiert XML-Datei"""
        if xml_path and os.path.exists(xml_path):
            output_file = os.path.join(export_path, f"{filename}.xml")
            output_file = self._get_unique_filename(output_file)
            shutil.copy2(xml_path, output_file)
            return True, f"XML exportiert: {os.path.basename(output_file)}"
        else:
            return False, "Keine XML-Datei vorhanden"

    def _export_to_email(self, pdf_path: str, xml_path: Optional[str],
                        export: ExportConfig, context: Dict[str, Any],
                        has_stamps: bool = False, compression_enabled: bool = False) -> Tuple[bool, str]:
        """E-Mail-Export mit den 3 Formaten"""
        if not export.email_config:
            return False, "Keine E-Mail-Konfiguration vorhanden"

        settings = self._get_export_settings()
        if not settings.smtp_server:
            return False, "SMTP-Server nicht konfiguriert"

        try:
            # Erstelle temporären Anhang
            with tempfile.TemporaryDirectory() as temp_dir:
                # Evaluiere Dateiname
                export_filename = self.function_parser.parse_and_evaluate(
                    export.export_filename_expression, context
                )
                export_filename = self.sanitize_filename(export_filename)

                # Erstelle Anhang im gewünschten Format
                if export.export_format == ExportFormat.PDF:
                    attachment_path = os.path.join(temp_dir, f"{export_filename}.pdf")
                    success, message = self._export_pdf(pdf_path, temp_dir, export_filename, 
                                                       export.format_params, has_stamps, compression_enabled)
                elif export.export_format == ExportFormat.SEARCHABLE_PDF_A:
                    attachment_path = os.path.join(temp_dir, f"{export_filename}.pdf")
                    success, message = self._export_pdf_a(pdf_path, temp_dir, export_filename, 
                                                         export.format_params)
                elif export.export_format == ExportFormat.XML:
                    attachment_path = os.path.join(temp_dir, f"{export_filename}.xml")
                    success, message = self._export_xml(xml_path, temp_dir, export_filename)
                else:
                    return False, f"E-Mail-Format {export.export_format} nicht unterstützt"

                if not success:
                    return False, f"Anhang-Erstellung fehlgeschlagen: {message}"

                # Finde erstellte Datei
                files = list(Path(temp_dir).glob(f"{export_filename}.*"))
                if not files:
                    return False, "Anhang konnte nicht erstellt werden"
                attachment_path = str(files[0])

                # Sende E-Mail
                return self._send_email(export.email_config, attachment_path,
                                      export_filename, context, settings)

        except Exception as e:
            logger.exception("E-Mail-Export fehlgeschlagen")
            return False, f"E-Mail-Fehler: {str(e)}"

    def _send_email(self, email_config: EmailConfig, attachment_path: str,
                   attachment_name: str, context: Dict[str, Any],
                   settings: ExportSettings) -> Tuple[bool, str]:
        """Sendet E-Mail"""
        try:
            # Erstelle Nachricht
            msg = MIMEMultipart('mixed')

            # Evaluiere E-Mail-Felder
            recipient = self.function_parser.parse_and_evaluate(
                email_config.recipient, context
            )
            subject = self.function_parser.parse_and_evaluate(
                email_config.subject_expression, context
            )
            body = self.function_parser.parse_and_evaluate(
                email_config.body_expression, context
            )

            # Header
            msg['From'] = settings.smtp_from_address
            msg['To'] = recipient
            msg['Subject'] = subject
            msg['Date'] = datetime.now().strftime('%a, %d %b %Y %H:%M:%S %z')

            # CC/BCC
            recipients = [recipient]
            if email_config.cc:
                cc = self.function_parser.parse_and_evaluate(email_config.cc, context)
                if cc:
                    msg['Cc'] = cc
                    recipients.extend(cc.split(','))
            
            if email_config.bcc:
                bcc = self.function_parser.parse_and_evaluate(email_config.bcc, context)
                if bcc:
                    recipients.extend(bcc.split(','))

            # Body
            msg_body = MIMEMultipart('alternative')
            msg_body.attach(MIMEText(body, 'plain', 'utf-8'))
            msg.attach(msg_body)

            # Anhang
            with open(attachment_path, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)

                # Content-Disposition mit korrektem Dateinamen
                ext = Path(attachment_path).suffix
                part.add_header(
                    'Content-Disposition',
                    'attachment',
                    filename=f"{attachment_name}{ext}"
                )
                msg.attach(part)

            # SMTP-Verbindung
            if settings.smtp_use_ssl and settings.smtp_port == 465:
                context_ssl = ssl.create_default_context()
                server = smtplib.SMTP_SSL(settings.smtp_server, settings.smtp_port, context=context_ssl)
            else:
                server = smtplib.SMTP(settings.smtp_server, settings.smtp_port)
                server.ehlo()
                if settings.smtp_use_tls:
                    context_ssl = ssl.create_default_context()
                    server.starttls(context=context_ssl)
                    server.ehlo()

            # Authentifizierung
            if settings.smtp_username and settings.smtp_password:
                server.login(settings.smtp_username, settings.smtp_password)

            # Sende E-Mail
            server.send_message(msg, to_addrs=recipients)
            server.quit()

            logger.info(f"E-Mail erfolgreich gesendet an {recipient}")
            return True, f"E-Mail gesendet an {recipient}"

        except Exception as e:
            logger.exception("E-Mail-Versand fehlgeschlagen")
            return False, f"E-Mail-Fehler: {str(e)}"

    def _update_pdf_metadata(self, pdf_path: str, metadata: Dict[str, str]):
        """Aktualisiert PDF-Metadaten"""
        try:
            doc = fitz.open(pdf_path)
            
            # Aktualisiere Metadaten
            current_metadata = doc.metadata or {}
            
            for key, value in metadata.items():
                if key.lower() in ['title', 'author', 'subject', 'keywords', 'creator', 'producer']:
                    current_metadata[key.lower()] = value
                else:
                    current_metadata[key] = value
            
            doc.set_metadata(current_metadata)
            doc.save(pdf_path + ".tmp", garbage=4, deflate=True, clean=True)
            doc.close()
            
            # Ersetze Original
            os.replace(pdf_path + ".tmp", pdf_path)
            
            logger.info("PDF-Metadaten erfolgreich aktualisiert")
            
        except Exception as e:
            logger.error(f"Metadaten-Update fehlgeschlagen: {e}")
            if os.path.exists(pdf_path + ".tmp"):
                os.remove(pdf_path + ".tmp")

    def _get_export_settings(self) -> ExportSettings:
        """Lädt Export-Einstellungen"""
        if self._export_settings is None:
            try:
                settings_file = "config/settings.json"
                
                if os.path.exists(settings_file):
                    with open(settings_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        self._export_settings = ExportSettings.from_dict(data)
                else:
                    self._export_settings = ExportSettings()
                    # Speichere Default-Settings
                    with open(settings_file, 'w', encoding='utf-8') as f:
                        json.dump(self._export_settings.to_dict(), f, indent=2, ensure_ascii=False)
            except Exception as e:
                logger.error(f"Fehler beim Laden der Einstellungen: {e}")
                self._export_settings = ExportSettings()

        return self._export_settings

    def get_error_path(self, error_path_expression: str, context: Dict[str, Any]) -> str:
        """Bestimmt Fehlerpfad"""
        if error_path_expression:
            return self.function_parser.parse_and_evaluate(error_path_expression, context)

        settings = self._get_export_settings()
        if settings.default_error_path:
            return settings.default_error_path

        # Fallback
        appdata = os.getenv('APPDATA')
        if appdata:
            default_error_path = os.path.join(appdata, 'HotfolderPDFProcessor', 'errors')
            os.makedirs(default_error_path, exist_ok=True)
            return default_error_path

        return os.path.join(context.get('FilePath', '.'), 'errors')

    def sanitize_filename(self, filename: str) -> str:
        """Bereinigt Dateinamen für Windows/Unix"""
        # Entferne ungültige Zeichen
        invalid_chars = '<>:"|?*\x00'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        
        # Entferne führende/nachfolgende Punkte und Leerzeichen
        filename = filename.strip('. ')
        
        # Maximale Länge
        if len(filename) > 200:
            filename = filename[:200]
        
        # Fallback wenn leer
        if not filename:
            filename = "export"
        
        return filename