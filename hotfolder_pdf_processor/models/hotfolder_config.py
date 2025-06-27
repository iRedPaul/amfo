"""
Datenmodell für Hotfolder-Konfigurationen
"""
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Any
from enum import Enum
import json


class ProcessingAction(Enum):
    """Verfügbare Bearbeitungsaktionen für PDFs"""
    SPLIT = "split"
    COMPRESS = "compress"
    OCR = "ocr"
    PDF_A = "pdf_a"


@dataclass
class OCRZone:
    """Definiert eine OCR-Zone"""
    name: str
    zone: tuple  # (x, y, width, height)
    page_num: int
    
    def to_dict(self) -> dict:
        return {
            "name": self.name,
            "zone": list(self.zone),
            "page_num": self.page_num
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'OCRZone':
        return cls(
            name=data["name"],
            zone=tuple(data["zone"]),
            page_num=data["page_num"]
        )


@dataclass
class HotfolderConfig:
    """Konfiguration für einen einzelnen Hotfolder"""
    id: str
    name: str
    input_path: str
    output_path: str
    enabled: bool = True
    description: str = ""  # Neues Feld für Beschreibung
    process_pairs: bool = True  # PDF + XML Paare verarbeiten
    actions: List[ProcessingAction] = field(default_factory=list)
    action_params: Dict[str, Any] = field(default_factory=dict)
    file_patterns: List[str] = field(default_factory=lambda: ["*.pdf"])
    xml_field_mappings: List[Dict[str, Any]] = field(default_factory=list)
    output_filename_expression: str = "<FileName>"  # Neuer Export-Dateiname Ausdruck
    ocr_zones: List[OCRZone] = field(default_factory=list)  # OCR-Zonen
    
    # Neue Felder für Export-Funktionalität
    export_configs: List[Dict[str, Any]] = field(default_factory=list)  # Liste der Export-Konfigurationen
    error_path: str = ""  # Optionaler Fehlerpfad (leer = Standard verwenden)
    
    def to_dict(self) -> dict:
        """Konvertiert die Konfiguration in ein Dictionary"""
        return {
            "id": self.id,
            "name": self.name,
            "input_path": self.input_path,
            "output_path": self.output_path,
            "enabled": self.enabled,
            "description": self.description,
            "process_pairs": self.process_pairs,
            "actions": [action.value for action in self.actions],
            "action_params": self.action_params,
            "file_patterns": self.file_patterns,
            "xml_field_mappings": self.xml_field_mappings,
            "output_filename_expression": self.output_filename_expression,
            "ocr_zones": [zone.to_dict() for zone in self.ocr_zones],
            "export_configs": self.export_configs,
            "error_path": self.error_path
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'HotfolderConfig':
        """Erstellt eine HotfolderConfig aus einem Dictionary"""
        actions = [ProcessingAction(action) for action in data.get("actions", [])]
        ocr_zones = [OCRZone.from_dict(zone) for zone in data.get("ocr_zones", [])]
        
        return cls(
            id=data["id"],
            name=data["name"],
            input_path=data["input_path"],
            output_path=data["output_path"],
            enabled=data.get("enabled", True),
            description=data.get("description", ""),
            process_pairs=data.get("process_pairs", True),
            actions=actions,
            action_params=data.get("action_params", {}),
            file_patterns=data.get("file_patterns", ["*.pdf"]),
            xml_field_mappings=data.get("xml_field_mappings", []),
            output_filename_expression=data.get("output_filename_expression", "<FileName>"),
            ocr_zones=ocr_zones,
            export_configs=data.get("export_configs", []),
            error_path=data.get("error_path", "")
        )


@dataclass
class DocumentPair:
    """Repräsentiert ein Dokumentenpaar (PDF + XML)"""
    pdf_path: str
    xml_path: Optional[str] = None
    
    @property
    def has_xml(self) -> bool:
        """Prüft ob ein XML-Dokument vorhanden ist"""
        return self.xml_path is not None
    
    @property
    def base_name(self) -> str:
        """Gibt den Basis-Dateinamen ohne Erweiterung zurück"""
        import os
        return os.path.splitext(os.path.basename(self.pdf_path))[0]