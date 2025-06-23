"""
Export-Processor für die Verarbeitung von Exporten
"""
import os
import shutil
import csv
import json
import xml.etree.ElementTree as ET
from xml.dom import minidom
import subprocess
import tempfile
from typing import Dict, List, Optional, Any, Tuple
from pathlib import Path
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import sys
from datetime import datetime

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.export_action import ExportConfig, ExportType, MetaDataFormat
from core.function_parser import FunctionParser, VariableExtractor
from models.hotfolder_config import DocumentPair, HotfolderConfig


class ExportProcessor:
    """Verarbeitet Export-Konfigurationen"""
    
    def __init__(self):
        self.function_parser = FunctionParser()
        self.variable_extractor = VariableExtractor()
        self._ocr_cache = {}
        self._processed_files = []  # Liste der erstellten Dateien für Cleanup
    
    def process_exports(self, doc_pair: DocumentPair, hotfolder: HotfolderConfig,
                       pdf_path: str, xml_path: Optional[str],
                       exports: List[ExportConfig],
                       evaluated_fields: Dict[str, str] = None) -> List[str]:
        """
        Führt alle konfigurierten Exporte aus
        
        Returns:
            Liste der erstellten Export-Dateien
        """
        self._processed_files = []
        
        for export in exports:
            if not export.enabled:
                continue
            
            try:
                # Prüfe Bedingung wenn definiert
                if export.condition_expression:
                    context = self._build_context(doc_pair, hotfolder, pdf_path, 
                                                xml_path, evaluated_fields)
                    condition_result = self.function_parser.parse_and_evaluate(
                        export.condition_expression, context
                    )
                    if condition_result.lower() != "true":
                        print(f"Export '{export.name}' übersprungen (Bedingung nicht erfüllt)")
                        continue
                
                # Führe Export aus basierend auf Typ
                if export.export_type == ExportType.FILE:
                    self._process_file_export(export, doc_pair, hotfolder, 
                                            pdf_path, xml_path, evaluated_fields)
                elif export.export_type == ExportType.EMAIL:
                    self._process_email_export(export, doc_pair, hotfolder,
                                             pdf_path, xml_path, evaluated_fields)
                elif export.export_type == ExportType.SCRIPT:
                    self._process_script_export(export, doc_pair, hotfolder,
                                              pdf_path, xml_path, evaluated_fields)
                # Weitere Export-Typen können hier ergänzt werden
                
            except Exception as e:
                print(f"Fehler bei Export '{export.name}': {e}")
                # Bei Fehler: Verschiebe in Fehler-Ordner wenn definiert
                if export.error_output_path:
                    self._move_to_error_folder(pdf_path, xml_path, 
                                             export.error_output_path, str(e))
        
        return self._processed_files
    
    def _build_context(self, doc_pair: DocumentPair, hotfolder: HotfolderConfig,
                      pdf_path: str, xml_path: Optional[str],
                      evaluated_fields: Dict[str, str] = None) -> Dict[str, Any]:
        """Baut den Kontext für Variablen-Evaluation auf"""
        context = self.variable_extractor.get_standard_variables()
        context.update(self.variable_extractor.get_file_variables(pdf_path))
        
        # Füge evaluierte XML-Felder hinzu
        if evaluated_fields:
            context.update(evaluated_fields)
        
        # Füge spezielle Export-Variablen hinzu
        context['OutputPath'] = hotfolder.output_path
        context['InputPath'] = hotfolder.input_path
        context['ErrorPath'] = hotfolder.error_path
        
        # OCR-Text wenn verfügbar
        if hasattr(hotfolder, 'ocr_zones') and hotfolder.ocr_zones:
            # OCR wurde bereits durchgeführt, Text sollte im Cache sein
            if pdf_path in self._ocr_cache:
                context['OCR_FullText'] = self._ocr_cache[pdf_path]
        
        return context
    
    def _process_file_export(self, export: ExportConfig, doc_pair: DocumentPair,
                           hotfolder: HotfolderConfig, pdf_path: str,
                           xml_path: Optional[str],
                           evaluated_fields: Dict[str, str] = None):
        """Verarbeitet Datei-Export"""
        context = self._build_context(doc_pair, hotfolder, pdf_path, 
                                    xml_path, evaluated_fields)
        
        # Evaluiere Ausgabepfad
        output_path = self.function_parser.parse_and_evaluate(
            export.output_path_expression, context
        )
        
        # Erstelle Pfad wenn nicht vorhanden
        if export.create_path_if_not_exists:
            os.makedirs(output_path, exist_ok=True)
        
        # Evaluiere Dateiname
        filename = self.function_parser.parse_and_evaluate(
            export.filename_expression, context
        )
        
        # Stelle sicher dass Dateiname gültig ist
        filename = self._sanitize_filename(filename)
        if not filename.lower().endswith('.pdf'):
            filename += '.pdf'
        
        full_output_path = os.path.join(output_path, filename)
        
        # Prüfe ob an existierende Datei angehängt werden soll
        if export.append_to_existing_file and os.path.exists(full_output_path):
            self._append_to_pdf(pdf_path, full_output_path, 
                              export.append_position == "start",
                              export.max_file_size_mb,
                              export.max_pages_per_file)
        else:
            # Kopiere PDF
            shutil.copy2(pdf_path, full_output_path)
        
        self._processed_files.append(full_output_path)
        print(f"Export '{export.name}': PDF exportiert nach {full_output_path}")
        
        # Erstelle OCR-Textdatei wenn gewünscht
        if export.include_ocr_text_file:
            self._create_ocr_text_file(pdf_path, full_output_path, context)
        
        # Verarbeite XML wenn vorhanden
        if xml_path and os.path.exists(xml_path):
            xml_output = full_output_path.replace('.pdf', '.xml')
            shutil.copy2(xml_path, xml_output)
            self._processed_files.append(xml_output)
        
        # Erstelle Metadaten-Datei
        if export.metadata_config.format != MetaDataFormat.NONE:
            self._create_metadata_file(export, full_output_path, context, 
                                     evaluated_fields)
        
        # Führe Programme aus
        for program in export.programs:
            self._execute_program(program, full_output_path, context)
    
    def _process_email_export(self, export: ExportConfig, doc_pair: DocumentPair,
                            hotfolder: HotfolderConfig, pdf_path: str,
                            xml_path: Optional[str],
                            evaluated_fields: Dict[str, str] = None):
        """Verarbeitet E-Mail Export"""
        if not export.email_config:
            print(f"Export '{export.name}': Keine E-Mail-Konfiguration vorhanden")
            return
        
        context = self._build_context(doc_pair, hotfolder, pdf_path, 
                                    xml_path, evaluated_fields)
        
        # Evaluiere E-Mail-Felder
        to_address = self.function_parser.parse_and_evaluate(
            export.email_config.to_expression, context
        )
        cc_address = self.function_parser.parse_and_evaluate(
            export.email_config.cc_expression, context
        ) if export.email_config.cc_expression else ""
        
        subject = self.function_parser.parse_and_evaluate(
            export.email_config.subject_expression, context
        )
        body = self.function_parser.parse_and_evaluate(
            export.email_config.body_expression, context
        )
        
        # Erstelle E-Mail
        msg = MIMEMultipart()
        msg['From'] = export.email_config.from_address
        msg['To'] = to_address
        if cc_address:
            msg['Cc'] = cc_address
        msg['Subject'] = subject
        
        # Füge Body hinzu
        msg.attach(MIMEText(body, 'plain'))
        
        # Füge Anhänge hinzu
        attachments = []
        
        if export.email_config.attach_pdf:
            attachments.append((os.path.basename(pdf_path), pdf_path))
        
        if export.email_config.attach_xml and xml_path:
            attachments.append((os.path.basename(xml_path), xml_path))
        
        # Füge Anhänge zur E-Mail hinzu
        for filename, filepath in attachments:
            with open(filepath, 'rb') as f:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(f.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition', f'attachment; filename= {filename}')
                msg.attach(part)
        
        # Sende E-Mail
        try:
            with smtplib.SMTP(export.email_config.smtp_server, 
                             export.email_config.smtp_port) as server:
                if export.email_config.use_ssl:
                    server.starttls()
                if export.email_config.username:
                    server.login(export.email_config.username, 
                               export.email_config.password)
                
                all_recipients = [to_address]
                if cc_address:
                    all_recipients.extend(cc_address.split(','))
                
                server.send_message(msg)
                print(f"Export '{export.name}': E-Mail gesendet an {to_address}")
                
        except Exception as e:
            print(f"Fehler beim E-Mail-Versand: {e}")
            raise
    
    def _process_script_export(self, export: ExportConfig, doc_pair: DocumentPair,
                             hotfolder: HotfolderConfig, pdf_path: str,
                             xml_path: Optional[str],
                             evaluated_fields: Dict[str, str] = None):
        """Führt externe Programme aus"""
        context = self._build_context(doc_pair, hotfolder, pdf_path, 
                                    xml_path, evaluated_fields)
        
        for program in export.programs:
            self._execute_program(program, pdf_path, context)
    
    def _execute_program(self, program, file_path: str, context: Dict[str, Any]):
        """Führt ein externes Programm aus"""
        try:
            # Erweitere Kontext mit speziellen Variablen
            program_context = context.copy()
            program_context['ProcessedFile'] = file_path
            program_context['ProcessedFileName'] = os.path.basename(file_path)
            program_context['ProcessedFilePath'] = os.path.dirname(file_path)
            
            # Evaluiere Parameter
            params = self.function_parser.parse_and_evaluate(
                program.parameters, program_context
            )
            
            # Erstelle Kommando
            cmd = [program.path]
            if params:
                # Teile Parameter korrekt auf (berücksichtige Quotes)
                import shlex
                cmd.extend(shlex.split(params))
            
            # Führe Programm aus
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode != 0:
                print(f"Programm-Fehler: {result.stderr}")
            else:
                print(f"Programm erfolgreich ausgeführt: {program.path}")
                
        except Exception as e:
            print(f"Fehler bei Programmausführung: {e}")
    
    def _create_metadata_file(self, export: ExportConfig, pdf_path: str,
                            context: Dict[str, Any],
                            evaluated_fields: Dict[str, str] = None):
        """Erstellt Metadaten-Datei"""
        metadata = export.metadata_config
        
        # Evaluiere Ausgabepfad für Metadaten
        meta_output_path = self.function_parser.parse_and_evaluate(
            metadata.output_path_expression, context
        )
        os.makedirs(meta_output_path, exist_ok=True)
        
        # Evaluiere Dateiname
        if metadata.use_document_filename:
            meta_filename = os.path.splitext(os.path.basename(pdf_path))[0]
        else:
            meta_filename = self.function_parser.parse_and_evaluate(
                metadata.filename_expression, context
            )
        
        # Erstelle Metadaten-Inhalt
        meta_data = {
            'FileName': os.path.basename(pdf_path),
            'ProcessDate': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'FilePath': os.path.dirname(pdf_path)
        }
        
        # Füge XML-Felder hinzu
        if metadata.include_xml_fields and evaluated_fields:
            meta_data.update(evaluated_fields)
        
        # Füge OCR-Text hinzu wenn gewünscht
        if metadata.include_ocr_text and pdf_path in self._ocr_cache:
            meta_data['OCR_Text'] = self._ocr_cache[pdf_path]
        
        # Schreibe Metadaten-Datei
        if metadata.format == MetaDataFormat.CSV:
            self._write_csv_metadata(meta_data, meta_output_path, 
                                   meta_filename, metadata)
        elif metadata.format == MetaDataFormat.XML:
            self._write_xml_metadata(meta_data, meta_output_path, 
                                   meta_filename, metadata)
        elif metadata.format == MetaDataFormat.JSON:
            self._write_json_metadata(meta_data, meta_output_path, 
                                    meta_filename, metadata)
    
    def _write_csv_metadata(self, data: Dict[str, Any], output_path: str,
                          filename: str, config):
        """Schreibt CSV-Metadaten"""
        csv_file = os.path.join(output_path, f"{filename}.{config.csv_extension}")
        
        # Schreibe oder füge an existierende Datei an
        mode = 'a' if config.append_to_existing and os.path.exists(csv_file) else 'w'
        write_header = mode == 'w' or not os.path.exists(csv_file)
        
        with open(csv_file, mode, newline='', encoding='utf-8') as f:
            writer = csv.DictWriter(f, fieldnames=data.keys(),
                                  delimiter=config.csv_delimiter,
                                  quotechar=config.csv_text_qualifier)
            if write_header:
                writer.writeheader()
            writer.writerow(data)
        
        self._processed_files.append(csv_file)
        print(f"CSV-Metadaten erstellt: {csv_file}")
    
    def _write_xml_metadata(self, data: Dict[str, Any], output_path: str,
                          filename: str, config):
        """Schreibt XML-Metadaten"""
        xml_file = os.path.join(output_path, f"{filename}.{config.xml_extension}")
        
        # Erstelle XML-Struktur
        root = ET.Element("Document")
        metadata_elem = ET.SubElement(root, "Metadata")
        
        for key, value in data.items():
            elem = ET.SubElement(metadata_elem, key)
            elem.text = str(value)
        
        # Formatiere XML
        xml_str = minidom.parseString(ET.tostring(root)).toprettyxml(indent="  ")
        
        with open(xml_file, 'w', encoding='utf-8') as f:
            f.write(xml_str)
        
        self._processed_files.append(xml_file)
        print(f"XML-Metadaten erstellt: {xml_file}")
        
        # Führe XSLT-Transformation durch wenn konfiguriert
        if config.xml_transform_file and os.path.exists(config.xml_transform_file):
            self._apply_xslt_transform(xml_file, config.xml_transform_file)
    
    def _write_json_metadata(self, data: Dict[str, Any], output_path: str,
                           filename: str, config):
        """Schreibt JSON-Metadaten"""
        json_file = os.path.join(output_path, f"{filename}.json")
        
        with open(json_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, indent=2, ensure_ascii=False)
        
        self._processed_files.append(json_file)
        print(f"JSON-Metadaten erstellt: {json_file}")
    
    def _create_ocr_text_file(self, pdf_path: str, output_pdf_path: str,
                            context: Dict[str, Any]):
        """Erstellt separate OCR-Textdatei"""
        if pdf_path in self._ocr_cache:
            txt_file = output_pdf_path.replace('.pdf', '.txt')
            with open(txt_file, 'w', encoding='utf-8') as f:
                f.write(self._ocr_cache[pdf_path])
            self._processed_files.append(txt_file)
            print(f"OCR-Textdatei erstellt: {txt_file}")
    
    def _append_to_pdf(self, source_pdf: str, target_pdf: str, 
                      prepend: bool = False,
                      max_size_mb: Optional[int] = None,
                      max_pages: Optional[int] = None):
        """Fügt PDF an existierende Datei an"""
        import PyPDF2
        
        # Prüfe Größenbeschränkung
        if max_size_mb:
            current_size = os.path.getsize(target_pdf) / (1024 * 1024)
            if current_size >= max_size_mb:
                # Erstelle neue Datei mit Nummer
                base, ext = os.path.splitext(target_pdf)
                counter = 1
                while os.path.exists(f"{base}_{counter}{ext}"):
                    counter += 1
                target_pdf = f"{base}_{counter}{ext}"
                shutil.copy2(source_pdf, target_pdf)
                return
        
        # Lese PDFs
        merger = PyPDF2.PdfMerger()
        
        if prepend:
            merger.append(source_pdf)
            merger.append(target_pdf)
        else:
            merger.append(target_pdf)
            merger.append(source_pdf)
        
        # Schreibe temporäre Datei
        temp_file = target_pdf + ".tmp"
        merger.write(temp_file)
        merger.close()
        
        # Ersetze Original
        shutil.move(temp_file, target_pdf)
    
    def _sanitize_filename(self, filename: str) -> str:
        """Entfernt ungültige Zeichen aus Dateinamen"""
        invalid_chars = '<>:"|?*'
        for char in invalid_chars:
            filename = filename.replace(char, '_')
        return filename.strip()
    
    def _move_to_error_folder(self, pdf_path: str, xml_path: Optional[str],
                            error_path: str, error_msg: str):
        """Verschiebt Dateien in Fehler-Ordner"""
        try:
            os.makedirs(error_path, exist_ok=True)
            
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            error_pdf = os.path.join(error_path, 
                                   f"ERROR_{timestamp}_{os.path.basename(pdf_path)}")
            
            shutil.copy2(pdf_path, error_pdf)
            
            # Schreibe Fehlerinfo
            error_info = error_pdf.replace('.pdf', '_error.txt')
            with open(error_info, 'w', encoding='utf-8') as f:
                f.write(f"Fehler: {error_msg}\n")
                f.write(f"Zeitstempel: {timestamp}\n")
                f.write(f"Original-Datei: {pdf_path}\n")
            
            if xml_path and os.path.exists(xml_path):
                error_xml = os.path.join(error_path,
                                       f"ERROR_{timestamp}_{os.path.basename(xml_path)}")
                shutil.copy2(xml_path, error_xml)
                
        except Exception as e:
            print(f"Fehler beim Verschieben in Fehler-Ordner: {e}")
    
    def _apply_xslt_transform(self, xml_file: str, xslt_file: str):
        """Wendet XSLT-Transformation auf XML an"""
        try:
            import lxml.etree as etree
            
            # Lade XML und XSLT
            dom = etree.parse(xml_file)
            xslt = etree.parse(xslt_file)
            transform = etree.XSLT(xslt)
            
            # Führe Transformation durch
            newdom = transform(dom)
            
            # Speichere transformiertes XML
            transformed_file = xml_file.replace('.xml', '_transformed.xml')
            with open(transformed_file, 'wb') as f:
                f.write(etree.tostring(newdom, pretty_print=True, encoding='utf-8'))
            
            self._processed_files.append(transformed_file)
            print(f"XSLT-Transformation angewendet: {transformed_file}")
            
        except Exception as e:
            print(f"Fehler bei XSLT-Transformation: {e}")
    
    def set_ocr_cache(self, pdf_path: str, ocr_text: str):
        """Setzt OCR-Text im Cache"""
        self._ocr_cache[pdf_path] = ocr_text
    
    def cleanup(self):
        """Räumt temporäre Dateien auf"""
        self._ocr_cache.clear()
        self._processed_files.clear()') as f:
                f.write(f"Fehler: {error_msg}\n")
                f.write(f"Zeitstempel: {timestamp}\n")
                f.write(f"Original-Datei: {pdf_path}\n")
            
            if xml_path and os.path.exists(xml_path):
                error_xml = os.path.join(error_path,
                                       f"ERROR_{timestamp}_{os.path.basename(xml_path)}")
                shutil.copy2(xml_path, error_xml)
                
        except Exception as e:
            print(f"Fehler beim Verschieben in Fehler-Ordner: {e}")
    
    def _apply_xslt_transform(self, xml_file: str, xslt_file: str):
        """Wendet XSLT-Transformation auf XML an"""
        try:
            import lxml.etree as etree
            
            # Lade XML und XSLT
            dom = etree.parse(xml_file)
            xslt = etree.parse(xslt_file)
            transform = etree.XSLT(xslt)
            
            # Führe Transformation durch
            newdom = transform(dom)
            
            # Speichere transformiertes XML
            transformed_file = xml_file.replace('.xml', '_transformed.xml')
            with open(transformed_file, 'wb') as f:
                f.write(etree.tostring(newdom, pretty_print=True, encoding='utf-8'))
            
            self._processed_files.append(transformed_file)
            print(f"XSLT-Transformation angewendet: {transformed_file}")
            
        except Exception as e:
            print(f"Fehler bei XSLT-Transformation: {e}")
    
    def set_ocr_cache(self, pdf_path: str, ocr_text: str):
        """Setzt OCR-Text im Cache"""
        self._ocr_cache[pdf_path] = ocr_text
    
    def cleanup(self):
        """Räumt temporäre Dateien auf"""
        self._ocr_cache.clear()
        self._processed_files.clear()