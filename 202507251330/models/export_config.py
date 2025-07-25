"""
Export-Konfiguration und Einstellungen
"""
from dataclasses import dataclass, field
from typing import Dict, List, Optional, Any
from enum import Enum
import json
import os


class ExportFormat(Enum):
    """Verfügbare Export-Formate"""
    PDF = "pdf"
    SEARCHABLE_PDF_A = "searchable_pdf_a"
    XML = "xml"


class ExportMethod(Enum):
    """Export-Methoden"""
    FILE = "file"
    EMAIL = "email"
    FTP = "ftp"
    CLOUD = "cloud"


class AuthMethod(Enum):
    """Authentifizierungsmethoden"""
    BASIC = "basic"
    OAUTH2 = "oauth2"


@dataclass
class EmailConfig:
    """E-Mail-Konfiguration für Exporte"""
    recipient: str
    subject_expression: str
    body_expression: str
    cc: Optional[str] = ""
    bcc: Optional[str] = ""
    
    def to_dict(self) -> dict:
        return {
            "recipient": self.recipient,
            "subject_expression": self.subject_expression,
            "body_expression": self.body_expression,
            "cc": self.cc,
            "bcc": self.bcc
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'EmailConfig':
        return cls(**data)


@dataclass
class ExportConfig:
    """Konfiguration für einen einzelnen Export"""
    id: str
    name: str
    enabled: bool = True
    export_method: ExportMethod = ExportMethod.FILE
    export_format: ExportFormat = ExportFormat.PDF
    export_path_expression: str = ""
    export_filename_expression: str = "<FileName>"
    format_params: Dict[str, Any] = field(default_factory=dict)
    email_config: Optional[EmailConfig] = None
    
    def __post_init__(self):
        # Konvertiere Strings zu Enums wenn nötig
        if isinstance(self.export_method, str):
            self.export_method = ExportMethod(self.export_method)
        if isinstance(self.export_format, str):
            self.export_format = ExportFormat(self.export_format)
    
    def to_dict(self) -> dict:
        return {
            "id": self.id,
            "name": self.name,
            "enabled": self.enabled,
            "export_method": self.export_method.value,
            "export_format": self.export_format.value,
            "export_path_expression": self.export_path_expression,
            "export_filename_expression": self.export_filename_expression,
            "format_params": self.format_params,
            "email_config": self.email_config.to_dict() if self.email_config else None
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'ExportConfig':
        email_config = None
        if data.get("email_config"):
            email_config = EmailConfig.from_dict(data["email_config"])
        
        return cls(
            id=data["id"],
            name=data["name"],
            enabled=data.get("enabled", True),
            export_method=data.get("export_method", "file"),
            export_format=data.get("export_format", "pdf"),
            export_path_expression=data.get("export_path_expression", ""),
            export_filename_expression=data.get("export_filename_expression", "<FileName>"),
            format_params=data.get("format_params", {}),
            email_config=email_config
        )


@dataclass
class ExportSettings:
    """Globale Export-Einstellungen"""
    # E-Mail Einstellungen
    smtp_server: str = ""
    smtp_port: int = 587
    smtp_username: str = ""
    smtp_password: str = ""
    smtp_from_address: str = ""
    smtp_use_tls: bool = True
    smtp_use_ssl: bool = False
    smtp_auth_method: AuthMethod = AuthMethod.BASIC
    
    # OAuth2 Einstellungen
    oauth2_provider: str = ""
    oauth2_client_id: str = ""
    oauth2_client_secret: str = ""
    oauth2_access_token: str = ""
    oauth2_refresh_token: str = ""
    oauth2_token_expiry: str = ""
    
    # Pfad-Einstellungen
    default_export_path: str = ""
    default_error_path: str = ""
    
    # OCR-Einstellungen
    ocr_default_language: str = "deu"
    ocr_additional_languages: List[str] = field(default_factory=lambda: ["eng"])
    
    def __post_init__(self):
        if isinstance(self.smtp_auth_method, str):
            self.smtp_auth_method = AuthMethod(self.smtp_auth_method)
    
    def to_dict(self) -> dict:
        return {
            "smtp_server": self.smtp_server,
            "smtp_port": self.smtp_port,
            "smtp_username": self.smtp_username,
            "smtp_password": self.smtp_password,
            "smtp_from_address": self.smtp_from_address,
            "smtp_use_tls": self.smtp_use_tls,
            "smtp_use_ssl": self.smtp_use_ssl,
            "smtp_auth_method": self.smtp_auth_method.value,
            "oauth2_provider": self.oauth2_provider,
            "oauth2_client_id": self.oauth2_client_id,
            "oauth2_client_secret": self.oauth2_client_secret,
            "oauth2_access_token": self.oauth2_access_token,
            "oauth2_refresh_token": self.oauth2_refresh_token,
            "oauth2_token_expiry": self.oauth2_token_expiry,
            "default_export_path": self.default_export_path,
            "default_error_path": self.default_error_path,
            "ocr_default_language": self.ocr_default_language,
            "ocr_additional_languages": self.ocr_additional_languages
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'ExportSettings':
        # Kopiere die Daten und behandle spezielle Felder
        clean_data = {}
        
        # Behandle alle bekannten Felder
        field_names = {
            'smtp_server', 'smtp_port', 'smtp_username', 'smtp_password',
            'smtp_from_address', 'smtp_use_tls', 'smtp_use_ssl', 'smtp_auth_method',
            'oauth2_provider', 'oauth2_client_id', 'oauth2_client_secret',
            'oauth2_access_token', 'oauth2_refresh_token', 'oauth2_token_expiry',
            'default_export_path', 'default_error_path',
            'ocr_default_language', 'ocr_additional_languages'
        }
        
        for field_name in field_names:
            if field_name in data:
                clean_data[field_name] = data[field_name]
        
        # Erstelle Settings-Objekt
        settings = cls(**clean_data)
        
        # Ignoriere application_paths - wird nicht mehr verwendet
        
        return settings