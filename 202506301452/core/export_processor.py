"""
Export-Prozessor für verschiedene Formate und Methoden
"""
import os
import shutil
import tempfile
import json
import smtplib
import ssl
from pathlib import Path
from typing import Dict, List, Optional, Any, Tuple
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import sys
import xml.etree.ElementTree as ET
import csv
from datetime import datetime
import logging
import ocrmypdf
import re

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.export_config import ExportConfig, ExportFormat, ExportMethod, EmailConfig, ExportSettings, AuthMethod
from core.function_parser import FunctionParser, VariableExtractor
from core.ocr_processor import OCRProcessor
from core.oauth2_manager import OAuth2Manager, get_token_storage

# Logger für dieses Modul
logger = logging.getLogger(__name__)


class ExportProcessor:
    """Verarbeitet Exporte in verschiedene Formate und über verschiedene Methoden"""

    def __init__(self):
        self.function_parser = FunctionParser()
        self.variable_extractor = VariableExtractor()
        self.ocr_processor = OCRProcessor()
        self._export_settings = None
        self._ocr_cache = {}

    def _get_unique_filename(self, filepath: str) -> str:
        """
        Gibt einen eindeutigen Dateinamen zurück, falls die Datei bereits existiert.
        Hängt einen Counter oder Timestamp an.
        """
        if not os.path.exists(filepath):
            return filepath

        # Trenne Pfad, Dateiname und Erweiterung
        directory = os.path.dirname(filepath)
        filename = os.path.basename(filepath)
        name, ext = os.path.splitext(filename)

        # Versuche zuerst mit Counter
        counter = 1
        while counter < 1000:  # Maximal 1000 Versuche mit Counter
            new_filename = f"{name}_{counter}{ext}"
            new_filepath = os.path.join(directory, new_filename)
            if not os.path.exists(new_filepath):
                return new_filepath
            counter += 1

        # Falls alle Counter belegt sind, verwende Timestamp
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S_%f")[:-3]  # Mit Millisekunden
        new_filename = f"{name}_{timestamp}{ext}"
        new_filepath = os.path.join(directory, new_filename)

        # Falls auch das existiert (sehr unwahrscheinlich), füge noch eine Zufallszahl hinzu
        if os.path.exists(new_filepath):
            import random
            rand_num = random.randint(1000, 9999)
            new_filename = f"{name}_{timestamp}_{rand_num}{ext}"
            new_filepath = os.path.join(directory, new_filename)

        return new_filepath

    def process_exports(self, pdf_path: str, xml_path: Optional[str],
                        export_configs: List[Dict], ocr_zones: List[Dict] = None,
                        xml_field_mappings: List[Dict] = None,
                        original_pdf_path: str = None,
                        input_path: str = None) -> List[Tuple[bool, str]]:
        """
        Führt alle konfigurierten Exporte durch

        Args:
            pdf_path: Pfad zur zu exportierenden PDF (temporärer Pfad)
            xml_path: Optionaler Pfad zur XML-Datei
            export_configs: Liste der Export-Konfigurationen
            ocr_zones: OCR-Zonen
            xml_field_mappings: XML-Feld-Mappings
            original_pdf_path: Original-Pfad der PDF (vor Verschiebung in temp)
            input_path: Pfad zum Hotfolder

        Returns:
            Liste von (success, message) Tupeln für jeden Export
        """
        results = []

        # Baue Kontext für Variablen auf
        # Verwende den temporären Pfad für OCR und Datei-Operationen
        context = self._build_context(pdf_path, xml_path, ocr_zones, 
                                    xml_field_mappings, input_path)
        
        # NEU: Wenn original_pdf_path vorhanden, überschreibe Level-Variablen und Dateinamen
        if original_pdf_path and input_path:
            # Extrahiere Level-Variablen vom Original-Pfad
            level_vars = self.variable_extractor.get_level_variables(original_pdf_path, input_path)
            context.update(level_vars)
            
            # Überschreibe Dateinamen-bezogene Variablen mit Original-Werten
            # damit die Dateinamen nicht die UUID enthalten
            original_filename = os.path.splitext(os.path.basename(original_pdf_path))[0]
            original_fullname = os.path.basename(original_pdf_path)
            
            # Entferne temporäre UUID aus dem Dateinamen falls vorhanden
            temp_filename = os.path.splitext(os.path.basename(pdf_path))[0]
            
            # Wenn der temporäre Name eine UUID ist (Format: 8-4-4-4-12 Zeichen)
            import re
            uuid_pattern = r'^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}$'
            if re.match(uuid_pattern, temp_filename):
                # Verwende Original-Dateinamen
                context['FileName'] = original_filename
                context['FullFileName'] = original_fullname
            # Ansonsten behalte die aktuellen Werte (falls Datei bereits umbenannt wurde)

        for export_dict in export_configs:
            try:
                if isinstance(export_dict, ExportConfig):
                    export = export_dict
                else:
                    export = ExportConfig.from_dict(export_dict)

                if not export.enabled:
                    continue

                success, message = self._process_single_export(pdf_path, xml_path, export, context)
                results.append((success, message))

                logger.info(f"Export '{export.name}': {'Erfolgreich' if success else 'Fehlgeschlagen'} - {message}")

            except Exception as e:
                logger.exception(f"Fehler bei Export-Verarbeitung: {e}")
                results.append((False, f"Export-Fehler: {str(e)}"))

        return results

    def _build_context(self, pdf_path: str, xml_path: Optional[str],
                       ocr_zones: List[Dict] = None,
                       xml_field_mappings: List[Dict] = None,
                       input_path: str = None) -> Dict[str, Any]:
        """Baut den Kontext für Variablen-Ersetzung auf"""
        context = {}

        # Datei-Informationen (verwende aktuellen Pfad für Datei-Operationen)
        context['FilePath'] = os.path.dirname(pdf_path)
        context['FileName'] = os.path.splitext(os.path.basename(pdf_path))[0]
        context['FileExtension'] = os.path.splitext(pdf_path)[1]
        context['FullFileName'] = os.path.basename(pdf_path)
        context['FullPath'] = pdf_path
        context['FileSize'] = str(os.path.getsize(pdf_path)) if os.path.exists(pdf_path) else '0'

        # Datum und Zeit
        now = datetime.now()
        context['Date'] = now.strftime('%Y-%m-%d')
        context['Time'] = now.strftime('%H-%M-%S')
        context['DateTime'] = now.strftime('%Y-%m-%d_%H-%M-%S')
        context['Year'] = now.strftime('%Y')
        context['Month'] = now.strftime('%m')
        context['Day'] = now.strftime('%d')
        context['Hour'] = now.strftime('%H')
        context['Minute'] = now.strftime('%M')
        context['Second'] = now.strftime('%S')
        context['Weekday'] = now.strftime('%A')
        context['WeekNumber'] = now.strftime('%V')

        # NEU: Level-Variablen hinzufügen
        if input_path:
            # Verwende den aktuellen pdf_path für Level-Variablen
            # da dieser die korrekte Ordnerstruktur enthält
            level_vars = self.variable_extractor.get_level_variables(pdf_path, input_path)
            context.update(level_vars)
            context['InputPath'] = input_path
        else:
            # Wenn kein input_path übergeben wurde, setze Level-Variablen auf leer
            for i in range(6):
                context[f'level{i}'] = ""

        # OCR ausführen wenn Zonen definiert sind
        if ocr_zones:
            # Volltext-OCR (falls noch nicht im Cache)
            # WICHTIG: Verwende den aktuellen pdf_path (temp Pfad) für OCR
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
                
                # Unterstütze beide Formate für Koordinaten
                if 'zone' in zone_dict:
                    # Neues Format: zone als Liste/Tupel
                    zone_coords = zone_dict['zone']
                else:
                    # Altes Format: einzelne Koordinaten
                    zone_coords = (
                        zone_dict.get('x', 0),
                        zone_dict.get('y', 0),
                        zone_dict.get('width', 100),
                        zone_dict.get('height', 100)
                    )

                # WICHTIG: Verwende den aktuellen pdf_path für Zone OCR
                zone_text = self.ocr_processor.extract_text_from_zone(
                    pdf_path, page_num, zone_coords
                )
                # Zone-Name sollte bereits OCR_ Präfix haben
                context[zone_name] = zone_text
                # Zusätzlich ohne Präfix für Kompatibilität
                if not zone_name.startswith('OCR_'):
                    context[f'OCR_{zone_name}'] = zone_text

        # XML-Feld-Mappings als Variablen
        if xml_field_mappings and xml_path:
            for mapping in xml_field_mappings:
                field_name = mapping.get('field_name', '')
                if field_name:
                    # Versuche Wert aus XML zu lesen
                    try:
                        tree = ET.parse(xml_path)
                        root = tree.getroot()
                        field_elem = root.find(f".//Fields/{field_name}")
                        if field_elem is not None and field_elem.text:
                            context[field_name] = field_elem.text
                    except:
                        pass

        # Spezielle Export-Variablen
        context['OutputPath'] = os.path.dirname(pdf_path)

        return context

    def _process_single_export(self, pdf_path: str, xml_path: Optional[str],
                              export: ExportConfig, context: Dict[str, Any]) -> Tuple[bool, str]:
        """Verarbeitet einen einzelnen Export"""
        try:
            # Strukturiertes Logging für Export-Start
            logger.info("Export gestartet", extra={
                'export_id': export.id,
                'export_name': export.name,
                'export_method': export.export_method.value,
                'export_format': export.export_format.value,
                'pdf_path': os.path.basename(pdf_path)
            })

            # Export-Methode bestimmt den Prozess
            if export.export_method == ExportMethod.FILE:
                return self._export_to_file(pdf_path, xml_path, export, context)
            elif export.export_method == ExportMethod.EMAIL:
                return self._export_to_email(pdf_path, xml_path, export, context)
            else:
                return False, f"Export-Methode {export.export_method} nicht implementiert"

        except Exception as e:
            logger.exception(f"Fehler bei Export {export.name}")
            return False, f"Fehler: {str(e)}"

    def _export_to_file(self, pdf_path: str, xml_path: Optional[str],
                        export: ExportConfig, context: Dict[str, Any]) -> Tuple[bool, str]:
        """Exportiert als Datei"""
        # Evaluiere Pfad und Dateiname
        export_path = self.function_parser.parse_and_evaluate(
            export.export_path_expression, context
        ) if export.export_path_expression else ""

        # Wenn kein Export-Pfad definiert, verwende Fehlerordner
        if not export_path:
            export_path = self.get_error_path("", context)

        export_filename = self.function_parser.parse_and_evaluate(
            export.export_filename_expression, context
        )
        export_filename = self.sanitize_filename(export_filename)

        # Stelle sicher, dass Pfad existiert
        os.makedirs(export_path, exist_ok=True)

        # Konvertiere je nach Format
        if export.export_format == ExportFormat.PDF:
            # Einfache PDF-Kopie
            output_file = os.path.join(export_path, f"{export_filename}.pdf")
            output_file = self._get_unique_filename(output_file)  # Eindeutigen Namen sicherstellen
            shutil.copy2(pdf_path, output_file)
            return True, f"PDF exportiert nach {output_file}"

        elif export.export_format == ExportFormat.PDF_A:
            # PDF/A Konvertierung
            output_file = os.path.join(export_path, f"{export_filename}.pdf")
            output_file = self._get_unique_filename(output_file)  # Eindeutigen Namen sicherstellen
            success = self._convert_to_pdf_a(pdf_path, output_file, False)
            return success, f"PDF/A {'exportiert' if success else 'Fehler'}: {output_file}"

        elif export.export_format == ExportFormat.SEARCHABLE_PDF_A:
            # Durchsuchbares PDF/A
            output_file = os.path.join(export_path, f"{export_filename}.pdf")
            output_file = self._get_unique_filename(output_file)  # Eindeutigen Namen sicherstellen
            success = self._convert_to_pdf_a(pdf_path, output_file, True)
            return success, f"Durchsuchbares PDF/A {'exportiert' if success else 'Fehler'}: {output_file}"

        elif export.export_format in [ExportFormat.PNG, ExportFormat.JPG, ExportFormat.TIFF]:
            # Bild-Export
            return self._export_to_images(pdf_path, export_path, export_filename,
                                          export.export_format, export.format_params)

        elif export.export_format == ExportFormat.XML:
            # XML-Export
            if xml_path and os.path.exists(xml_path):
                output_file = os.path.join(export_path, f"{export_filename}.xml")
                output_file = self._get_unique_filename(output_file)  # Eindeutigen Namen sicherstellen
                shutil.copy2(xml_path, output_file)
                return True, f"XML exportiert nach {output_file}"
            else:
                return False, "Keine XML-Datei vorhanden"

        elif export.export_format == ExportFormat.JSON:
            # JSON-Export
            return self._export_to_json(pdf_path, xml_path, export_path,
                                      export_filename, context)

        elif export.export_format == ExportFormat.TXT:
            # Text-Export (OCR)
            output_file = os.path.join(export_path, f"{export_filename}.txt")
            output_file = self._get_unique_filename(output_file)  # Eindeutigen Namen sicherstellen
            with open(output_file, 'w', encoding='utf-8') as f:
                f.write(context.get('OCR_FullText', ''))
            return True, f"Text exportiert nach {output_file}"

        elif export.export_format == ExportFormat.CSV:
            # CSV-Export
            return self._export_to_csv(pdf_path, xml_path, export_path,
                                     export_filename, context)

        else:
            return False, f"Format {export.export_format} nicht implementiert"

    def _export_to_email(self, pdf_path: str, xml_path: Optional[str],
                        export: ExportConfig, context: Dict[str, Any]) -> Tuple[bool, str]:
        """Exportiert per E-Mail"""
        if not export.email_config:
            return False, "Keine E-Mail-Konfiguration vorhanden"

        # Lade E-Mail-Einstellungen
        settings = self._get_export_settings()
        if not settings.smtp_server:
            return False, "Keine SMTP-Server-Einstellungen konfiguriert"

        try:
            # Erstelle temporäre Export-Datei
            with tempfile.TemporaryDirectory() as temp_dir:
                # Evaluiere Dateiname
                export_filename = self.function_parser.parse_and_evaluate(
                    export.export_filename_expression, context
                )
                export_filename = self.sanitize_filename(export_filename)

                # Konvertiere je nach Format
                if export.export_format == ExportFormat.PDF:
                    attachment_path = os.path.join(temp_dir, f"{export_filename}.pdf")
                    shutil.copy2(pdf_path, attachment_path)
                else:
                    # Verwende File-Export-Logik für andere Formate
                    fake_export = ExportConfig(
                        id="temp",
                        name="temp",
                        export_method=ExportMethod.FILE,
                        export_format=export.export_format,
                        export_path_expression=temp_dir,
                        export_filename_expression=export_filename,
                        format_params=export.format_params
                    )
                    success, message = self._export_to_file(pdf_path, xml_path, fake_export, context)
                    if not success:
                        return False, f"Fehler beim Erstellen des Anhangs: {message}"

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
        """Sendet eine E-Mail mit Anhang"""
        try:
            # Erstelle Nachricht
            msg = MIMEMultipart()

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

            msg['From'] = settings.smtp_from_address
            msg['To'] = recipient
            msg['Subject'] = subject

            if email_config.cc:
                cc = self.function_parser.parse_and_evaluate(email_config.cc, context)
                if cc:
                    msg['Cc'] = cc
            if email_config.bcc:
                bcc = self.function_parser.parse_and_evaluate(email_config.bcc, context)
                if bcc:
                    msg['Bcc'] = bcc

            # Nachrichtentext
            msg.attach(MIMEText(body, 'plain', 'utf-8'))

            # Anhang
            with open(attachment_path, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)

                # Bestimme Dateiendung
                ext = Path(attachment_path).suffix
                part.add_header(
                    'Content-Disposition',
                    f'attachment; filename="{attachment_name}{ext}"'
                )
                msg.attach(part)

            # Prüfe Auth-Methode
            access_token = None
            if settings.smtp_auth_method == AuthMethod.OAUTH2:
                # OAuth2-Authentifizierung vorbereiten
                access_token = self._get_oauth2_access_token(settings)
                if not access_token:
                    return False, "OAuth2-Token konnte nicht abgerufen werden"

            # Verbinde zum Server
            if settings.smtp_use_ssl and settings.smtp_port == 465:
                # SSL direkt verwenden (Port 465)
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL(settings.smtp_server, settings.smtp_port, context=context)
            else:
                # TLS über STARTTLS (Port 587 oder andere)
                server = smtplib.SMTP(settings.smtp_server, settings.smtp_port)
                if settings.smtp_use_tls:
                    server.starttls()

            # Anmelden
            if settings.smtp_auth_method == AuthMethod.OAUTH2 and access_token:
                # OAuth2-Anmeldung
                oauth2_manager = OAuth2Manager(settings.oauth2_provider)
                auth_string = oauth2_manager.create_oauth2_sasl_string(
                    settings.smtp_from_address,
                    access_token
                )
                server.docmd('AUTH', 'XOAUTH2 ' + auth_string)
            else:
                # Standard-Anmeldung
                if settings.smtp_username and settings.smtp_password:
                    server.login(settings.smtp_username, settings.smtp_password)

            # Sende E-Mail
            recipients = [recipient]
            if email_config.cc:
                recipients.extend(email_config.cc.split(','))
            if email_config.bcc:
                recipients.extend(email_config.bcc.split(','))

            server.send_message(msg, to_addrs=recipients)
            server.quit()

            # Strukturiertes Logging bei Erfolg
            logger.info("E-Mail erfolgreich gesendet", extra={
                'recipient': recipient,
                'subject': subject[:50] + '...' if len(subject) > 50 else subject,
                'attachment': attachment_name,
                'size_kb': os.path.getsize(attachment_path) / 1024
            })

            return True, f"E-Mail gesendet an {recipient}"

        except Exception as e:
            logger.exception("E-Mail-Versand fehlgeschlagen")
            return False, f"E-Mail-Versand fehlgeschlagen: {str(e)}"

    def _get_oauth2_access_token(self, settings: ExportSettings) -> Optional[str]:
        """Holt oder erneuert den OAuth2 Access Token"""
        try:
            if not settings.oauth2_provider or not settings.oauth2_refresh_token:
                return None

            # OAuth2 Manager
            oauth2_manager = OAuth2Manager(settings.oauth2_provider)
            token_storage = get_token_storage()

            # Hole gespeicherte Tokens
            tokens = token_storage.get_tokens(settings.oauth2_provider, settings.smtp_from_address)
            if not tokens:
                # Verwende Tokens aus Settings als Fallback
                tokens = {
                    'access_token': settings.oauth2_access_token,
                    'refresh_token': settings.oauth2_refresh_token,
                    'token_expiry': settings.oauth2_token_expiry
                }

            # Prüfe ob Token erneuert werden muss
            if oauth2_manager.is_token_expired(tokens.get('token_expiry', '')):
                # Token erneuern
                oauth2_manager.set_client_credentials(
                    settings.oauth2_client_id,
                    settings.oauth2_client_secret
                )

                success, new_tokens = oauth2_manager.refresh_access_token(
                    tokens.get('refresh_token', '')
                )

                if success:
                    # Speichere neue Tokens
                    token_storage.set_tokens(
                        settings.oauth2_provider,
                        settings.smtp_from_address,
                        new_tokens
                    )

                    # Update Settings
                    settings.oauth2_access_token = new_tokens['access_token']
                    settings.oauth2_refresh_token = new_tokens.get('refresh_token', settings.oauth2_refresh_token)
                    settings.oauth2_token_expiry = new_tokens['token_expiry']

                    # Speichere Settings
                    self._save_export_settings(settings)

                    return new_tokens['access_token']
                else:
                    logger.error(f"OAuth2 Token-Erneuerung fehlgeschlagen: {new_tokens.get('error', 'Unbekannter Fehler')}")
                    return None

            return tokens.get('access_token')

        except Exception as e:
            logger.exception(f"Fehler beim Abrufen des OAuth2-Tokens: {e}")
            return None

    def _convert_to_pdf_a(self, input_pdf: str, output_pdf: str,
                          searchable: bool, pdfa_version: str = "2b") -> bool:
        """
        Konvertiert eine PDF-Datei in eine "perfekte" PDF/A-Datei, optional mit durchsuchbarem Text.
        Verwendet ocrmypdf für eine robuste Konvertierung und OCR.
        """
        try:
            logger.info(f"Starte {'durchsuchbare ' if searchable else ''}PDF/A-{pdfa_version} Konvertierung für: {os.path.basename(input_pdf)}")

            # Argumente für ocrmypdf
            args = {
                "output_type": "pdfa",
                "pdfa_image_compression": "jpeg",
                "jpeg_quality": 85,
                "deskew": True,
                # "clean": True,  # ENTFERNT: Diese Option benötigt unpaper
                "rotate_pages": True,
                "optimize": 1,  # GEÄNDERT von 3 auf 1: Optimize-Level 2 und 3 benötigen pngquant
            }

            if searchable:
                args.update({
                    "language": "deu",
                    "force_ocr": True,
                    "skip_text": False,
                    "tesseract_timeout": 600,
                })
            else:
                args.update({
                    "tesseract_timeout": 0,
                    "skip_text": True,
                })

            # Ausführung von ocrmypdf
            result = ocrmypdf.ocr(input_pdf, output_pdf, **args)

            if result == ocrmypdf.ExitCode.already_done_ocr:
                logger.warning(f"Datei wurde bereits mit OCR verarbeitet und übersprungen: {os.path.basename(input_pdf)}")
                shutil.copy2(input_pdf, output_pdf)
            elif result != ocrmypdf.ExitCode.ok:
                logger.error(f"ocrmypdf Konvertierung fehlgeschlagen mit Exit-Code: {result.name} ({result.value})")
                return False

            logger.info(f"PDF/A Konvertierung erfolgreich abgeschlossen: {os.path.basename(output_pdf)}")
            return True

        except ocrmypdf.exceptions.EncryptedPdfError:
            logger.error(f"Fehler: Die PDF-Datei ist verschlüsselt: {os.path.basename(input_pdf)}")
            return False
        except ocrmypdf.exceptions.InputFileError as e:
            logger.error(f"Fehler: Ungültige PDF-Eingabedatei: {os.path.basename(input_pdf)}. Details: {e}")
            return False
        except Exception as e:
            logger.exception(f"Unerwarteter Fehler bei der PDF/A Konvertierung für {os.path.basename(input_pdf)}")
            return False

    def _export_to_images(self, pdf_path: str, export_path: str,
                         base_filename: str, format: ExportFormat,
                         params: Dict[str, Any]) -> Tuple[bool, str]:
        """Exportiert PDF als Bilder"""
        try:
            from pdf2image import convert_from_path

            # Parameter
            dpi = params.get('dpi', 300)
            quality = params.get('quality', 95)

            # Konvertiere zu Bildern
            poppler_path = os.path.join(os.path.dirname(__file__), '..', 'poppler', 'bin')
            images = convert_from_path(pdf_path, dpi=dpi, poppler_path=poppler_path)

            # Speichere Bilder
            output_files = []
            for i, image in enumerate(images):
                if len(images) > 1:
                    filename = f"{base_filename}_page_{i+1:03d}"
                else:
                    filename = base_filename

                if format == ExportFormat.PNG:
                    output_file = os.path.join(export_path, f"{filename}.png")
                    output_file = self._get_unique_filename(output_file)  # Eindeutigen Namen sicherstellen
                    image.save(output_file, 'PNG')
                elif format == ExportFormat.JPG:
                    output_file = os.path.join(export_path, f"{filename}.jpg")
                    output_file = self._get_unique_filename(output_file)  # Eindeutigen Namen sicherstellen
                    image.save(output_file, 'JPEG', quality=quality)
                elif format == ExportFormat.TIFF:
                    output_file = os.path.join(export_path, f"{filename}.tiff")
                    if params.get('multipage', True) and len(images) > 1:
                        # Mehrseitiges TIFF
                        if i == 0:
                            output_file = self._get_unique_filename(output_file)  # Eindeutigen Namen sicherstellen
                            images[0].save(
                                output_file,
                                save_all=True,
                                append_images=images[1:],
                                compression="tiff_lzw"
                            )
                            output_files = [output_file]
                            break
                    else:
                        output_file = self._get_unique_filename(output_file)  # Eindeutigen Namen sicherstellen
                        image.save(output_file, 'TIFF', compression="tiff_lzw")

                output_files.append(output_file)

            return True, f"{len(output_files)} Bild(er) exportiert"

        except Exception as e:
            logger.exception("Bild-Export fehlgeschlagen")
            return False, f"Bild-Export-Fehler: {str(e)}"

    def _export_to_json(self, pdf_path: str, xml_path: Optional[str],
                       export_path: str, filename: str,
                       context: Dict[str, Any]) -> Tuple[bool, str]:
        """Exportiert als JSON"""
        try:
            data = {
                "document": {
                    "filename": os.path.basename(pdf_path),
                    "path": pdf_path,
                    "size": os.path.getsize(pdf_path),
                    "export_date": context.get('DateTime', '')
                },
                "metadata": {},
                "ocr_text": context.get('OCR_FullText', ''),
                "zones": {},
                "levels": {}  # NEU: Level-Informationen
            }

            # NEU: Füge Level-Variablen hinzu
            for i in range(6):
                level_key = f'level{i}'
                if level_key in context and context[level_key]:
                    data["levels"][level_key] = context[level_key]

            # XML-Daten hinzufügen
            if xml_path and os.path.exists(xml_path):
                try:
                    tree = ET.parse(xml_path)
                    root = tree.getroot()

                    # Metadaten
                    metadata = root.find(".//Metadata")
                    if metadata is not None:
                        for elem in metadata:
                            data["metadata"][elem.tag] = elem.text

                    # Felder
                    fields = root.find(".//Fields")
                    if fields is not None:
                        data["fields"] = {}
                        for field in fields:
                            data["fields"][field.tag] = field.text
                except:
                    pass

            # OCR-Zonen
            for key, value in context.items():
                if key.startswith('OCR_Zone_') or key.startswith('Zone_'):
                    data["zones"][key] = value

            # Speichere JSON
            output_file = os.path.join(export_path, f"{filename}.json")
            output_file = self._get_unique_filename(output_file)  # Eindeutigen Namen sicherstellen

            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)

            return True, f"JSON exportiert nach {output_file}"

        except Exception as e:
            logger.exception("JSON-Export fehlgeschlagen")
            return False, f"JSON-Export-Fehler: {str(e)}"

    def _export_to_csv(self, pdf_path: str, xml_path: Optional[str],
                      export_path: str, filename: str,
                      context: Dict[str, Any]) -> Tuple[bool, str]:
        """Exportiert als CSV"""
        try:
            rows = []
            row = {
                'Dateiname': os.path.basename(pdf_path),
                'Pfad': os.path.dirname(pdf_path),
                'Export_Datum': context.get('Date', ''),
                'Export_Zeit': context.get('Time', '')
            }

            # NEU: Füge Level-Variablen hinzu
            for i in range(6):
                level_key = f'level{i}'
                row[f'Level_{i}'] = context.get(level_key, '')

            # XML-Felder
            if xml_path and os.path.exists(xml_path):
                try:
                    tree = ET.parse(xml_path)
                    root = tree.getroot()
                    fields_elem = root.find(".//Fields")
                    if fields_elem is not None:
                        for field in fields_elem:
                            row[f"XML_{field.tag}"] = field.text or ""
                except:
                    pass

            # OCR-Zonen
            for key, value in context.items():
                if key.startswith('OCR_Zone_') or key.startswith('Zone_'):
                    # Kürze Text für CSV
                    text = str(value).replace('\n', ' ').replace('\r', '')
                    if len(text) > 100:
                        text = text[:100] + "..."
                    row[key] = text

            rows.append(row)

            # Schreibe CSV
            output_file = os.path.join(export_path, f"{filename}.csv")
            output_file = self._get_unique_filename(output_file)  # Eindeutigen Namen sicherstellen

            if rows:
                with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=rows[0].keys(), delimiter=';')
                    writer.writeheader()
                    writer.writerows(rows)

            return True, f"CSV exportiert nach {output_file}"

        except Exception as e:
            logger.exception("CSV-Export fehlgeschlagen")
            return False, f"CSV-Export-Fehler: {str(e)}"

    def _get_export_settings(self) -> ExportSettings:
        """Lädt die Export-Einstellungen"""
        if self._export_settings is None:
            try:
                settings_file = "settings.json"
                if os.path.exists(settings_file):
                    with open(settings_file, 'r', encoding='utf-8') as f:
                        data = json.load(f)
                        self._export_settings = ExportSettings.from_dict(data)
                else:
                    self._export_settings = ExportSettings()
            except Exception as e:
                logger.exception(f"Fehler beim Laden der Export-Einstellungen: {e}")
                self._export_settings = ExportSettings()

        return self._export_settings

    def _save_export_settings(self, settings: ExportSettings):
        """Speichert die Export-Einstellungen"""
        try:
            settings_file = "settings.json"
            with open(settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings.to_dict(), f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Fehler beim Speichern der Einstellungen: {e}")

    def get_error_path(self, error_path_expression: str, context: Dict[str, Any]) -> str:
        """Bestimmt den Fehlerpfad"""
        # Wenn Ausdruck definiert, evaluiere ihn
        if error_path_expression:
            return self.function_parser.parse_and_evaluate(error_path_expression, context)

        # Sonst verwende Standard aus Einstellungen
        settings = self._get_export_settings()
        if settings.default_error_path:
            return settings.default_error_path

        # Fallback: AppData-Ordner
        appdata = os.getenv('APPDATA')
        if appdata:
            default_error_path = os.path.join(appdata, 'HotfolderPDFProcessor', 'errors')
            os.makedirs(default_error_path, exist_ok=True)
            return default_error_path

        # Letzter Fallback: Input-Ordner/errors
        return os.path.join(context.get('FilePath', '.'), 'errors')

    def sanitize_filename(self, filename):
        """Entfernt ungültige Zeichen aus Dateinamen für Windows."""
        return re.sub(r'[\\/:*?"<>|]', '_', filename)