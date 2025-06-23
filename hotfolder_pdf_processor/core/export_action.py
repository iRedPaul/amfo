"""
Export-Action für erweiterte Ausgabe-Funktionalität
"""
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any
from enum import Enum
import os


class ExportType(Enum):
    """Verfügbare Export-Typen"""
    FILE = "file"           # Dateiausgabe
    EMAIL = "email"         # E-Mail Versand
    FTP = "ftp"            # FTP Upload
    WEBDAV = "webdav"      # WebDAV Upload
    SHAREPOINT = "sharepoint"  # SharePoint Upload
    DROPBOX = "dropbox"    # Dropbox Upload
    GOOGLE_DRIVE = "google_drive"  # Google Drive Upload
    SCRIPT = "script"      # Externes Skript aufrufen


class MetaDataFormat(Enum):
    """Metadaten-Formate"""
    NONE = "none"
    CSV = "csv"
    XML = "xml"
    JSON = "json"


@dataclass
class MetaDataConfig:
    """Konfiguration für Meta-Datei Ausgabe"""
    format: MetaDataFormat = MetaDataFormat.NONE
    output_path_expression: str = "<OutputPath>"  # Mit Ausdrücken
    filename_expression: str = "<FileName>"
    use_document_filename: bool = True
    append_to_existing: bool = False
    include_ocr_text: bool = False
    include_ocr_zones: bool = False
    include_xml_fields: bool = True
    csv_delimiter: str = ";"
    csv_text_qualifier: str = '"'
    csv_record_delimiter: str = "\r\n"
    csv_extension: str = "csv"
    xml_extension: str = "xml"
    xml_transform_file: Optional[str] = None  # XSLT Transformation
    
    def to_dict(self) -> dict:
        return {
            "format": self.format.value,
            "output_path_expression": self.output_path_expression,
            "filename_expression": self.filename_expression,
            "use_document_filename": self.use_document_filename,
            "append_to_existing": self.append_to_existing,
            "include_ocr_text": self.include_ocr_text,
            "include_ocr_zones": self.include_ocr_zones,
            "include_xml_fields": self.include_xml_fields,
            "csv_delimiter": self.csv_delimiter,
            "csv_text_qualifier": self.csv_text_qualifier,
            "csv_record_delimiter": self.csv_record_delimiter,
            "csv_extension": self.csv_extension,
            "xml_extension": self.xml_extension,
            "xml_transform_file": self.xml_transform_file
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'MetaDataConfig':
        data = data.copy()
        if 'format' in data:
            data['format'] = MetaDataFormat(data['format'])
        return cls(**data)


@dataclass
class ExportProgram:
    """Externes Programm zur Ausführung"""
    path: str
    parameters: str
    run_as_64bit: bool = True
    run_for_each_document: bool = True
    
    def to_dict(self) -> dict:
        return {
            "path": self.path,
            "parameters": self.parameters,
            "run_as_64bit": self.run_as_64bit,
            "run_for_each_document": self.run_for_each_document
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'ExportProgram':
        return cls(**data)


@dataclass
class EmailConfig:
    """E-Mail Versand Konfiguration"""
    smtp_server: str = ""
    smtp_port: int = 587
    use_ssl: bool = True
    username: str = ""
    password: str = ""  # Sollte verschlüsselt gespeichert werden
    from_address: str = ""
    to_expression: str = ""  # Kann Ausdrücke enthalten
    cc_expression: str = ""
    bcc_expression: str = ""
    subject_expression: str = "Dokument: <FileName>"
    body_expression: str = "Anbei erhalten Sie das angeforderte Dokument."
    attach_pdf: bool = True
    attach_xml: bool = True
    attach_metadata: bool = False
    
    def to_dict(self) -> dict:
        return {
            "smtp_server": self.smtp_server,
            "smtp_port": self.smtp_port,
            "use_ssl": self.use_ssl,
            "username": self.username,
            "password": self.password,  # TODO: Verschlüsselung
            "from_address": self.from_address,
            "to_expression": self.to_expression,
            "cc_expression": self.cc_expression,
            "bcc_expression": self.bcc_expression,
            "subject_expression": self.subject_expression,
            "body_expression": self.body_expression,
            "attach_pdf": self.attach_pdf,
            "attach_xml": self.attach_xml,
            "attach_metadata": self.attach_metadata
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'EmailConfig':
        return cls(**data)


@dataclass
class ExportConfig:
    """Hauptkonfiguration für einen Export"""
    id: str
    name: str
    enabled: bool = True
    export_type: ExportType = ExportType.FILE
    
    # Dateiausgabe
    output_path_expression: str = "<OutputPath>"
    filename_expression: str = "<FileName>"
    create_path_if_not_exists: bool = True
    append_to_existing_file: bool = False
    max_file_size_mb: Optional[int] = None  # Max Dateigröße in MB
    max_pages_per_file: Optional[int] = None  # Max Seiten pro Datei
    append_position: str = "end"  # "end" oder "start"
    include_ocr_text_file: bool = False  # Separate OCR-Textdatei
    error_output_path: str = ""  # Fehlerausgabe-Ordner
    
    # Metadaten
    metadata_config: MetaDataConfig = field(default_factory=MetaDataConfig)
    
    # E-Mail
    email_config: Optional[EmailConfig] = None
    
    # Programme
    programs: List[ExportProgram] = field(default_factory=list)
    
    # Bedingungen (wenn implementiert)
    condition_expression: str = ""  # z.B. IF("<FileSize>", ">", "1000000", "true", "false")
    
    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "enabled": self.enabled,
            "export_type": self.export_type.value,
            "output_path_expression": self.output_path_expression,
            "filename_expression": self.filename_expression,
            "create_path_if_not_exists": self.create_path_if_not_exists,
            "append_to_existing_file": self.append_to_existing_file,
            "max_file_size_mb": self.max_file_size_mb,
            "max_pages_per_file": self.max_pages_per_file,
            "append_position": self.append_position,
            "include_ocr_text_file": self.include_ocr_text_file,
            "error_output_path": self.error_output_path,
            "metadata_config": self.metadata_config.to_dict(),
            "email_config": self.email_config.to_dict() if self.email_config else None,
            "programs": [p.to_dict() for p in self.programs],
            "condition_expression": self.condition_expression
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'ExportConfig':
        data = data.copy()
        
        # Konvertiere Enums
        if 'export_type' in data:
            data['export_type'] = ExportType(data['export_type'])
        
        # Konvertiere MetaDataConfig
        if 'metadata_config' in data:
            data['metadata_config'] = MetaDataConfig.from_dict(data['metadata_config'])
        else:
            data['metadata_config'] = MetaDataConfig()
        
        # Konvertiere EmailConfig
        if 'email_config' in data and data['email_config']:
            data['email_config'] = EmailConfig.from_dict(data['email_config'])
        
        # Konvertiere Programme
        if 'programs' in data:
            data['programs'] = [ExportProgram.from_dict(p) for p in data['programs']]
        
        return cls(**data)