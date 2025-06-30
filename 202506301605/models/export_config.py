"""
Datenmodell für Export-Konfigurationen
"""
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any
from enum import Enum
import json


class ExportFormat(Enum):
    """Verfügbare Export-Formate"""
    PDF = "pdf"
    PDF_A = "pdf_a"
    SEARCHABLE_PDF_A = "searchable_pdf_a"
    PNG = "png"
    JPG = "jpg"
    TIFF = "tiff"
    XML = "xml"
    JSON = "json"
    TXT = "txt"
    CSV = "csv"


class ExportMethod(Enum):
    """Export-Methoden"""
    FILE = "file"
    EMAIL = "email"
    FTP = "ftp"
    WEBDAV = "webdav"


class AuthMethod(Enum):
    """Authentifizierungsmethoden für E-Mail"""
    BASIC = "basic"
    OAUTH2 = "oauth2"


@dataclass
class EmailConfig:
    """E-Mail-Konfiguration für Exporte"""
    recipient: str = ""
    subject_expression: str = "Dokument: <FileName>"
    body_expression: str = "Anbei finden Sie das verarbeitete Dokument."
    cc: str = ""
    bcc: str = ""
    
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
        return cls(
            recipient=data.get("recipient", ""),
            subject_expression=data.get("subject_expression", "Dokument: <FileName>"),
            body_expression=data.get("body_expression", "Anbei finden Sie das verarbeitete Dokument."),
            cc=data.get("cc", ""),
            bcc=data.get("bcc", "")
        )


@dataclass
class ExportConfig:
    """Konfiguration für einen einzelnen Export"""
    id: str
    name: str
    enabled: bool = True
    export_method: ExportMethod = ExportMethod.FILE
    export_format: ExportFormat = ExportFormat.PDF
    export_path_expression: str = ""  # Leer statt <OutputPath>
    export_filename_expression: str = "<FileName>"
    
    # Format-spezifische Parameter
    format_params: Dict[str, Any] = field(default_factory=dict)
    
    # E-Mail-Konfiguration
    email_config: Optional[EmailConfig] = None
    
    # Weitere Export-Methoden können hier ergänzt werden
    
    def to_dict(self) -> dict:
        """Konvertiert die Konfiguration in ein Dictionary"""
        data = {
            "id": self.id,
            "name": self.name,
            "enabled": self.enabled,
            "export_method": self.export_method.value,
            "export_format": self.export_format.value,
            "export_path_expression": self.export_path_expression,
            "export_filename_expression": self.export_filename_expression,
            "format_params": self.format_params
        }
        
        if self.email_config:
            data["email_config"] = self.email_config.to_dict()
        
        return data
    
    @classmethod
    def from_dict(cls, data: dict) -> 'ExportConfig':
        """Erstellt eine ExportConfig aus einem Dictionary"""
        export_method = ExportMethod(data.get("export_method", ExportMethod.FILE.value))
        export_format = ExportFormat(data.get("export_format", ExportFormat.PDF.value))
        
        email_config = None
        if "email_config" in data:
            email_config = EmailConfig.from_dict(data["email_config"])
        
        return cls(
            id=data["id"],
            name=data["name"],
            enabled=data.get("enabled", True),
            export_method=export_method,
            export_format=export_format,
            export_path_expression=data.get("export_path_expression", ""),
            export_filename_expression=data.get("export_filename_expression", "<FileName>"),
            format_params=data.get("format_params", {}),
            email_config=email_config
        )


@dataclass
class ExportSettings:
    """Globale Export-Einstellungen"""
    default_error_path: str = ""
    smtp_server: str = ""
    smtp_port: int = 587
    smtp_use_ssl: bool = False  # Neu: SSL direkt verwenden (für Port 465)
    smtp_use_tls: bool = True   # TLS über STARTTLS (für Port 587)
    smtp_username: str = ""
    smtp_password: str = ""  # Sollte verschlüsselt gespeichert werden
    smtp_from_address: str = ""
    
    # OAuth2-Einstellungen
    smtp_auth_method: AuthMethod = AuthMethod.BASIC
    oauth2_provider: str = ""  # gmail, outlook, etc.
    oauth2_client_id: str = ""
    oauth2_client_secret: str = ""
    oauth2_refresh_token: str = ""
    oauth2_access_token: str = ""
    oauth2_token_expiry: str = ""
    
    def to_dict(self) -> dict:
        return {
            "default_error_path": self.default_error_path,
            "smtp_server": self.smtp_server,
            "smtp_port": self.smtp_port,
            "smtp_use_ssl": self.smtp_use_ssl,
            "smtp_use_tls": self.smtp_use_tls,
            "smtp_username": self.smtp_username,
            "smtp_password": self.smtp_password,
            "smtp_from_address": self.smtp_from_address,
            "smtp_auth_method": self.smtp_auth_method.value,
            "oauth2_provider": self.oauth2_provider,
            "oauth2_client_id": self.oauth2_client_id,
            "oauth2_client_secret": self.oauth2_client_secret,
            "oauth2_refresh_token": self.oauth2_refresh_token,
            "oauth2_access_token": self.oauth2_access_token,
            "oauth2_token_expiry": self.oauth2_token_expiry
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'ExportSettings':
        auth_method = AuthMethod(data.get("smtp_auth_method", AuthMethod.BASIC.value))
        
        return cls(
            default_error_path=data.get("default_error_path", ""),
            smtp_server=data.get("smtp_server", ""),
            smtp_port=data.get("smtp_port", 587),
            smtp_use_ssl=data.get("smtp_use_ssl", False),
            smtp_use_tls=data.get("smtp_use_tls", True),
            smtp_username=data.get("smtp_username", ""),
            smtp_password=data.get("smtp_password", ""),
            smtp_from_address=data.get("smtp_from_address", ""),
            smtp_auth_method=auth_method,
            oauth2_provider=data.get("oauth2_provider", ""),
            oauth2_client_id=data.get("oauth2_client_id", ""),
            oauth2_client_secret=data.get("oauth2_client_secret", ""),
            oauth2_refresh_token=data.get("oauth2_refresh_token", ""),
            oauth2_access_token=data.get("oauth2_access_token", ""),
            oauth2_token_expiry=data.get("oauth2_token_expiry", "")
        )