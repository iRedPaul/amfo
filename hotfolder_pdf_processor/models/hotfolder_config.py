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
    PDF_A = "pdf_a"  # Neue PDF/A Konvertierung


@dataclass
class HotfolderConfig:
    """Konfiguration für einen einzelnen Hotfolder"""
    id: str
    name: str
    input_path: str
    output_path: str
    enabled: bool = True
    process_pairs: bool = True  # PDF + XML Paare verarbeiten
    actions: List[ProcessingAction] = field(default_factory=list)
    action_params: Dict[str, any] = field(default_factory=dict)
    file_patterns: List[str] = field(default_factory=lambda: ["*.pdf"])
    xml_field_mappings: List[Dict[str, Any]] = field(default_factory=list)  # Neue XML-Feld-Mappings
    
    def to_dict(self) -> dict:
        """Konvertiert die Konfiguration in ein Dictionary"""
        return {
            "id": self.id,
            "name": self.name,
            "input_path": self.input_path,
            "output_path": self.output_path,
            "enabled": self.enabled,
            "process_pairs": self.process_pairs,
            "actions": [action.value for action in self.actions],
            "action_params": self.action_params,
            "file_patterns": self.file_patterns,
            "xml_field_mappings": self.xml_field_mappings
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'HotfolderConfig':
        """Erstellt eine HotfolderConfig aus einem Dictionary"""
        actions = [ProcessingAction(action) for action in data.get("actions", [])]
        return cls(
            id=data["id"],
            name=data["name"],
            input_path=data["input_path"],
            output_path=data["output_path"],
            enabled=data.get("enabled", True),
            process_pairs=data.get("process_pairs", True),
            actions=actions,
            action_params=data.get("action_params", {}),
            file_patterns=data.get("file_patterns", ["*.pdf"]),
            xml_field_mappings=data.get("xml_field_mappings", [])
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