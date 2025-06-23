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

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.export_config import ExportConfig, ExportFormat, ExportMethod, EmailConfig, ExportSettings
from core.function_parser import FunctionParser, VariableExtractor
from core.ocr_processor import OCRProcessor


class ExportProcessor:
    """Verarbeitet Exporte in verschiedene Formate und über verschiedene Methoden"""
    
    def __init__(self):
        self.function_parser = FunctionParser()
        self.variable_extractor = VariableExtractor()
        self.ocr_processor = OCRProcessor()
        self._export_settings = None
        self._ocr_cache = {}
    
    def process_exports(self, pdf_path: str, xml_path: Optional[str], 
                       export_configs: List[Dict], ocr_zones: List[Dict] = None,
                       xml_field_mappings: List[Dict] = None) -> List[Tuple[bool, str]]:
        """
        Führt alle konfigurierten Exporte durch
        
        Returns:
            Liste von (success, message) Tupeln für jeden Export
        """
        results = []
        
        # Baue Kontext für Variablen auf
        context = self._build_context(pdf_path, xml_path, ocr_zones, xml_field_mappings)
        
        for export_dict in export_configs:
            try:
                export = ExportConfig.from_dict(export_dict)
                
                if not export.enabled:
                    continue
                
                # Führe Export durch
                success, message = self._process_single_export(
                    pdf_path, xml_path, export, context
                )
                
                results.append((success, f"{export.name}: {message}"))
                
            except Exception as e:
                results.append((False, f"Export-Fehler: {str(e)}"))
        
        # Leere Cache
        self._ocr_cache.clear()
        
        return results
    
    def _build_context(self, pdf_path: str, xml_path: Optional[str], 
                      ocr_zones: List[Dict] = None, 
                      xml_field_mappings: List[Dict] = None) -> Dict[str, Any]:
        """Baut den Kontext für Variablen auf"""
        context = {}
        
        # Standard-Variablen
        context.update(self.variable_extractor.get_standard_variables())
        
        # Datei-Variablen
        context.update(self.variable_extractor.get_file_variables(pdf_path))
        
        # XML-Variablen
        if xml_path and os.path.exists(xml_path):
            context.update(self.variable_extractor.get_xml_variables(xml_path))
        
        # OCR-Text (wenn benötigt)
        if pdf_path not in self._ocr_cache:
            self._ocr_cache[pdf_path] = self.ocr_processor.extract_text_from_pdf(pdf_path)
        context['OCR_FullText'] = self._ocr_cache[pdf_path]
        
        # OCR-Zonen
        if ocr_zones:
            for i, zone_info in enumerate(ocr_zones):
                zone_text = self.ocr_processor.extract_text_from_zone(
                    pdf_path, zone_info['page_num'], zone_info['zone']
                )
                zone_name = zone_info.get('name', f'Zone_{i+1}')
                context[zone_name] = zone_text
                context[f'OCR_{zone_name}'] = zone_text
        
        # XML-Feld-Mappings als Variablen
        if xml_field_mappings:
            for mapping in xml_field_mappings:
                field_name = mapping.get('field_name', '')
                if field_name and xml_path:
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
            # Export-Methode bestimmt den Prozess
            if export.export_method == ExportMethod.FILE:
                return self._export_to_file(pdf_path, xml_path, export, context)
            elif export.export_method == ExportMethod.EMAIL:
                return self._export_to_email(pdf_path, xml_path, export, context)
            else:
                return False, f"Export-Methode {export.export_method} nicht implementiert"
                
        except Exception as e:
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
        
        # Stelle sicher, dass Pfad existiert
        os.makedirs(export_path, exist_ok=True)
        
        # Konvertiere je nach Format
        if export.export_format == ExportFormat.PDF:
            # Einfache PDF-Kopie
            output_file = os.path.join(export_path, f"{export_filename}.pdf")
            shutil.copy2(pdf_path, output_file)
            return True, f"PDF exportiert nach {output_file}"
            
        elif export.export_format == ExportFormat.PDF_A:
            # PDF/A Konvertierung
            output_file = os.path.join(export_path, f"{export_filename}.pdf")
            success = self._convert_to_pdf_a(pdf_path, output_file, False)
            return success, f"PDF/A {'exportiert' if success else 'Fehler'}: {output_file}"
            
        elif export.export_format == ExportFormat.SEARCHABLE_PDF_A:
            # Durchsuchbares PDF/A
            output_file = os.path.join(export_path, f"{export_filename}.pdf")
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
            return False, f"E-Mail-Fehler: {str(e)}"
    
    def _send_email(self, email_config: EmailConfig, attachment_path: str, 
                   attachment_name: str, context: Dict[str, Any], 
                   settings: ExportSettings) -> Tuple[bool, str]:
        """Sendet eine E-Mail mit Anhang"""
        try:
            # Erstelle Nachricht
            msg = MIMEMultipart()
            
            # Evaluiere E-Mail-Felder
            subject = self.function_parser.parse_and_evaluate(
                email_config.subject_expression, context
            )
            body = self.function_parser.parse_and_evaluate(
                email_config.body_expression, context
            )
            
            msg['From'] = settings.smtp_from_address
            msg['To'] = email_config.recipient
            msg['Subject'] = subject
            
            if email_config.cc:
                msg['Cc'] = email_config.cc
            if email_config.bcc:
                msg['Bcc'] = email_config.bcc
            
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
            
            # Verbinde zum Server - WICHTIG: SSL/TLS korrekt handhaben
            if settings.smtp_use_ssl and settings.smtp_port == 465:
                # SSL direkt verwenden (Port 465)
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL(settings.smtp_server, settings.smtp_port, context=context)
            else:
                # TLS über STARTTLS (Port 587 oder andere)
                server = smtplib.SMTP(settings.smtp_server, settings.smtp_port)
                if settings.smtp_use_tls:
                    server.starttls()
            
            # Anmelden wenn Credentials vorhanden
            if settings.smtp_username and settings.smtp_password:
                server.login(settings.smtp_username, settings.smtp_password)
            
            # Sende E-Mail
            recipients = [email_config.recipient]
            if email_config.cc:
                recipients.extend(email_config.cc.split(','))
            if email_config.bcc:
                recipients.extend(email_config.bcc.split(','))
            
            server.send_message(msg, to_addrs=recipients)
            server.quit()
            
            return True, f"E-Mail gesendet an {email_config.recipient}"
            
        except Exception as e:
            return False, f"E-Mail-Versand fehlgeschlagen: {str(e)}"
    
    def _convert_to_pdf_a(self, input_pdf: str, output_pdf: str, 
                         searchable: bool) -> bool:
        """Konvertiert zu PDF/A"""
        try:
            import ocrmypdf
            
            ocrmypdf.ocr(
                input_pdf,
                output_pdf,
                pdfa_image_compression="jpeg",
                output_type="pdfa",
                tesseract_timeout=300 if searchable else 0,
                skip_text=not searchable,
                force_ocr=searchable,
                language="deu" if searchable else None
            )
            return True
            
        except ImportError:
            # Fallback: Kopiere einfach
            shutil.copy2(input_pdf, output_pdf)
            return True
        except Exception:
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
                    image.save(output_file, 'PNG')
                elif format == ExportFormat.JPG:
                    output_file = os.path.join(export_path, f"{filename}.jpg")
                    image.save(output_file, 'JPEG', quality=quality)
                elif format == ExportFormat.TIFF:
                    output_file = os.path.join(export_path, f"{filename}.tiff")
                    if params.get('multipage', True) and len(images) > 1:
                        # Mehrseitiges TIFF
                        if i == 0:
                            images[0].save(
                                output_file,
                                save_all=True,
                                append_images=images[1:],
                                compression="tiff_lzw"
                            )
                            output_files = [output_file]
                            break
                    else:
                        image.save(output_file, 'TIFF', compression="tiff_lzw")
                
                output_files.append(output_file)
            
            return True, f"{len(output_files)} Bild(er) exportiert"
            
        except Exception as e:
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
                "zones": {}
            }
            
            # Füge XML-Daten hinzu wenn vorhanden
            if xml_path and os.path.exists(xml_path):
                try:
                    tree = ET.parse(xml_path)
                    root = tree.getroot()
                    fields_elem = root.find(".//Fields")
                    if fields_elem is not None:
                        data["xml_fields"] = {
                            field.tag: field.text for field in fields_elem
                        }
                except:
                    pass
            
            # Füge OCR-Zonen hinzu
            for key, value in context.items():
                if key.startswith('OCR_Zone_') or key.startswith('Zone_'):
                    data["zones"][key] = value
            
            # Speichere JSON
            output_file = os.path.join(export_path, f"{filename}.json")
            with open(output_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            
            return True, f"JSON exportiert nach {output_file}"
            
        except Exception as e:
            return False, f"JSON-Export-Fehler: {str(e)}"
    
    def _export_to_csv(self, pdf_path: str, xml_path: Optional[str], 
                      export_path: str, filename: str, 
                      context: Dict[str, Any]) -> Tuple[bool, str]:
        """Exportiert als CSV"""
        try:
            # Sammle alle Daten
            rows = []
            
            # Basis-Informationen
            row = {
                'Dateiname': os.path.basename(pdf_path),
                'Dateipfad': pdf_path,
                'Dateigröße': os.path.getsize(pdf_path),
                'Export-Datum': context.get('DateTime', '')
            }
            
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
            
            if rows:
                with open(output_file, 'w', newline='', encoding='utf-8-sig') as f:
                    writer = csv.DictWriter(f, fieldnames=rows[0].keys(), delimiter=';')
                    writer.writeheader()
                    writer.writerows(rows)
            
            return True, f"CSV exportiert nach {output_file}"
            
        except Exception as e:
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
            except:
                self._export_settings = ExportSettings()
        
        return self._export_settings
    
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