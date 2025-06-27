import json
import os
import logging
import threading
from typing import Dict, List, Optional, Any
from dataclasses import dataclass, field, asdict
import uuid
from datetime import datetime

logger = logging.getLogger(__name__)


@dataclass
class ExportConfig:
    """Konfiguration für einen Export"""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    name: str = ""
    enabled: bool = True
    export_method: str = "file"  # file, email, ftp, etc.
    export_format: str = "searchable_pdf_a"
    export_path_expression: str = ""
    export_filename_expression: str = "<FileName>"
    format_params: Dict[str, Any] = field(default_factory=dict)
    
    def to_dict(self) -> Dict[str, Any]:
        return asdict(self)
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'ExportConfig':
        return cls(**data)


@dataclass
class HotfolderConfig:
    """Konfiguration für einen Hotfolder"""
    id: str = field(default_factory=lambda: str(uuid.uuid4()))
    name: str = ""
    enabled: bool = True
    description: str = ""
    
    # Paths
    input_path: str = ""
    output_path: str = ""
    error_path: str = ""
    
    # Processing
    process_pairs: bool = False
    file_patterns: List[str] = field(default_factory=lambda: ["*.pdf"])
    output_filename_expression: str = "<FileName>"
    
    # OCR
    ocr_zones: List[Dict[str, Any]] = field(default_factory=list)
    
    # Fields
    xml_field_mappings: List[Dict[str, Any]] = field(default_factory=list)
    
    # Actions
    actions: List[str] = field(default_factory=list)
    action_params: Dict[str, Any] = field(default_factory=dict)
    
    # Exports
    export_configs: List[ExportConfig] = field(default_factory=list)
    
    def to_dict(self) -> Dict[str, Any]:
        data = asdict(self)
        data['export_configs'] = [ec.to_dict() for ec in self.export_configs]
        return data
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'HotfolderConfig':
        export_configs = data.pop('export_configs', [])
        hotfolder = cls(**data)
        hotfolder.export_configs = [ExportConfig.from_dict(ec) for ec in export_configs]
        return hotfolder


class ConfigManager:
    """Verwaltet die Konfiguration der Anwendung"""
    
    DEFAULT_CONFIG_FILE = "hotfolders.json"
    
    def __init__(self, config_file: str = None):
        self.config_file = config_file or self.DEFAULT_CONFIG_FILE
        self.hotfolders: List[HotfolderConfig] = []
        self.load_config()
    
    def load_config(self) -> None:
        """Lädt die Konfiguration aus der Datei"""
        if not os.path.exists(self.config_file):
            logger.info(f"Konfigurationsdatei {self.config_file} nicht gefunden, erstelle neue")
            self.create_default_config()
            self.save_config()
            return
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Lade Hotfolder mit neuer Struktur
            self.hotfolders = []
            for hf_data in data.get('hotfolders', []):
                # Konvertiere von alter zu neuer Struktur falls nötig
                converted_data = self._convert_from_storage(hf_data)
                hotfolder = HotfolderConfig.from_dict(converted_data)
                self.hotfolders.append(hotfolder)
            
            logger.info(f"{len(self.hotfolders)} Hotfolder geladen")
            
        except Exception as e:
            logger.exception(f"Fehler beim Laden der Konfiguration: {e}")
            self.create_default_config()
    
    def create_default_config(self):
        """Erstellt eine Standard-Konfiguration"""
        self.hotfolders = []
    
    def save_config(self) -> None:
        """Speichert die Konfiguration in die Datei"""
        data = {
            'version': '2.0',
            'hotfolders': []
        }
        
        for hotfolder in self.hotfolders:
            # Konvertiere zur Speicherstruktur
            hf_storage = self._convert_to_storage(hotfolder)
            data['hotfolders'].append(hf_storage)
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            logger.info("Konfiguration gespeichert")
        except Exception as e:
            logger.error(f"Fehler beim Speichern der Konfiguration: {e}")
    
    def add_hotfolder(self, hotfolder: HotfolderConfig) -> None:
        """Fügt einen neuen Hotfolder hinzu"""
        self.hotfolders.append(hotfolder)
        self.save_config()
        logger.info(f"Hotfolder '{hotfolder.name}' hinzugefügt")
    
    def update_hotfolder(self, hotfolder_id: str, updated_hotfolder: HotfolderConfig) -> None:
        """Aktualisiert einen bestehenden Hotfolder"""
        for i, hf in enumerate(self.hotfolders):
            if hf.id == hotfolder_id:
                self.hotfolders[i] = updated_hotfolder
                self.save_config()
                logger.info(f"Hotfolder '{updated_hotfolder.name}' aktualisiert")
                return
        logger.warning(f"Hotfolder mit ID {hotfolder_id} nicht gefunden")
    
    def delete_hotfolder(self, hotfolder_id: str) -> None:
        """Löscht einen Hotfolder"""
        self.hotfolders = [hf for hf in self.hotfolders if hf.id != hotfolder_id]
        self.save_config()
        logger.info(f"Hotfolder mit ID {hotfolder_id} gelöscht")
    
    def get_hotfolder(self, hotfolder_id: str) -> Optional[HotfolderConfig]:
        """Gibt einen spezifischen Hotfolder zurück"""
        for hf in self.hotfolders:
            if hf.id == hotfolder_id:
                return hf
        return None
    
    def get_hotfolders(self) -> List[HotfolderConfig]:
        """Gibt alle Hotfolder zurück"""
        return self.hotfolders
    
    def get_enabled_hotfolders(self) -> List[HotfolderConfig]:
        """Gibt nur die aktivierten Hotfolder zurück"""
        return [hf for hf in self.hotfolders if hf.enabled]
    
    def _convert_from_storage(self, stored_hf: Dict[str, Any]) -> Dict[str, Any]:
        """Konvertiert von der gespeicherten Struktur zu HotfolderConfig"""
        # Prüfe ob es bereits die neue Struktur ist
        if "input_path" in stored_hf:
            return stored_hf
        
        # Konvertiere alte Struktur
        return {
            "id": stored_hf.get("id"),
            "name": stored_hf.get("name"),
            "enabled": stored_hf.get("enabled", True),
            "description": stored_hf.get("description", ""),
            
            # Paths
            "input_path": stored_hf.get("paths", {}).get("input", ""),
            "output_path": stored_hf.get("paths", {}).get("output", ""),
            "error_path": stored_hf.get("paths", {}).get("error", ""),
            
            # Processing
            "process_pairs": stored_hf.get("processing", {}).get("process_pairs", False),
            "file_patterns": stored_hf.get("processing", {}).get("file_patterns", ["*.pdf"]),
            "output_filename_expression": stored_hf.get("processing", {}).get("output_filename_expression", "<FileName>"),
            
            # OCR
            "ocr_zones": stored_hf.get("ocr", {}).get("zones", []),
            
            # Fields
            "xml_field_mappings": stored_hf.get("fields", {}).get("mappings", []),
            
            # Actions
            "actions": stored_hf.get("actions", {}).get("list", []),
            "action_params": stored_hf.get("actions", {}).get("parameters", {}),
            
            # Exports
            "export_configs": [
                {
                    "id": export.get("id"),
                    "name": export.get("name"),
                    "enabled": export.get("enabled", True),
                    "export_method": export.get("method", "file"),
                    "export_format": export.get("format", "searchable_pdf_a"),
                    "export_path_expression": export.get("paths", {}).get("path_expression", ""),
                    "export_filename_expression": export.get("paths", {}).get("filename_expression", ""),
                    "format_params": export.get("format_params", {})
                }
                for export in stored_hf.get("exports", [])
            ]
        }
    
    def _convert_to_storage(self, hotfolder: HotfolderConfig) -> Dict[str, Any]:
        """Konvertiert von HotfolderConfig zur gespeicherten Struktur"""
        hf_dict = hotfolder.to_dict()
        
        return {
            "id": hf_dict.get("id"),
            "name": hf_dict.get("name"),
            "enabled": hf_dict.get("enabled", True),
            "description": hf_dict.get("description", ""),
            
            "paths": {
                "input": hf_dict.get("input_path", ""),
                "output": hf_dict.get("output_path", ""),
                "error": hf_dict.get("error_path", "")
            },
            
            "processing": {
                "process_pairs": hf_dict.get("process_pairs", False),
                "file_patterns": hf_dict.get("file_patterns", ["*.pdf"]),
                "output_filename_expression": hf_dict.get("output_filename_expression", "<FileName>")
            },
            
            "ocr": {
                "zones": hf_dict.get("ocr_zones", [])
            },
            
            "fields": {
                "mappings": hf_dict.get("xml_field_mappings", [])
            },
            
            "actions": {
                "list": hf_dict.get("actions", []),
                "parameters": hf_dict.get("action_params", {})
            },
            
            "exports": [
                {
                    "id": export.get("id"),
                    "name": export.get("name"),
                    "enabled": export.get("enabled", True),
                    "method": export.get("export_method", "file"),
                    "format": export.get("export_format", "searchable_pdf_a"),
                    "paths": {
                        "path_expression": export.get("export_path_expression", ""),
                        "filename_expression": export.get("export_filename_expression", "")
                    },
                    "format_params": export.get("format_params", {})
                }
                for export in hf_dict.get("export_configs", [])
            ]
        }


class SettingsManager:
    """Verwaltet die globalen Einstellungen der Anwendung"""
    
    DEFAULT_SETTINGS_FILE = "settings.json"
    
    def __init__(self, settings_file: str = None):
        self.settings_file = settings_file or self.DEFAULT_SETTINGS_FILE
        self.paths: Dict[str, str] = {}
        self.smtp: Dict[str, Any] = {}
        self.load_settings()
    
    def load_settings(self) -> None:
        """Lädt die Einstellungen aus der Datei"""
        if not os.path.exists(self.settings_file):
            logger.info(f"Einstellungsdatei {self.settings_file} nicht gefunden, erstelle neue")
            self.create_default_settings()
            self.save_settings()
            return
        
        try:
            with open(self.settings_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Lade Pfade
            self.paths = data.get("paths", {})
            
            # Lade SMTP-Einstellungen
            self.smtp = data.get("smtp", self._get_default_smtp())
            
            logger.info("Einstellungen geladen")
            
        except Exception as e:
            logger.exception(f"Fehler beim Laden der Einstellungen: {e}")
            self.create_default_settings()
    
    def create_default_settings(self):
        """Erstellt Standard-Einstellungen"""
        self.paths = {
            "default_error": ""
        }
        self.smtp = self._get_default_smtp()
    
    def _get_default_smtp(self) -> Dict[str, Any]:
        """Gibt die Standard-SMTP-Konfiguration zurück"""
        return {
            "server": {
                "host": "",
                "port": 587,
                "use_ssl": False,
                "use_tls": True
            },
            "auth": {
                "method": "basic",  # basic, oauth2
                "basic": {
                    "username": "",
                    "password": "",
                    "from_address": ""
                },
                "oauth2": {
                    "client_id": "",
                    "client_secret": "",
                    "tenant_id": "",
                    "from_address": ""
                }
            }
        }
    
    def save_settings(self) -> None:
        """Speichert die Einstellungen in die Datei"""
        data = {
            "paths": self.paths,
            "smtp": self.smtp
        }
        
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            logger.debug("Einstellungen gespeichert")
        except Exception as e:
            logger.error(f"Fehler beim Speichern der Einstellungen: {e}")
    
    def get_default_error_path(self) -> str:
        """Gibt den Standard-Fehlerpfad zurück"""
        return self.paths.get("default_error", "")
    
    def set_default_error_path(self, path: str) -> None:
        """Setzt den Standard-Fehlerpfad"""
        self.paths["default_error"] = path
        self.save_settings()
    
    def get_smtp_config(self) -> Dict[str, Any]:
        """Gibt die komplette SMTP-Konfiguration zurück"""
        return self.smtp.copy()
    
    def update_smtp_config(self, smtp_config: Dict[str, Any]) -> None:
        """Aktualisiert die SMTP-Konfiguration"""
        self.smtp = smtp_config
        self.save_settings()
    
    def get_smtp_server(self) -> str:
        return self.smtp.get("server", {}).get("host", "")
    
    def get_smtp_port(self) -> int:
        return self.smtp.get("server", {}).get("port", 587)
    
    def get_smtp_use_ssl(self) -> bool:
        return self.smtp.get("server", {}).get("use_ssl", False)
    
    def get_smtp_use_tls(self) -> bool:
        return self.smtp.get("server", {}).get("use_tls", True)
    
    def get_smtp_auth_method(self) -> str:
        return self.smtp.get("auth", {}).get("method", "basic")
    
    def get_smtp_username(self) -> str:
        return self.smtp.get("auth", {}).get("basic", {}).get("username", "")
    
    def get_smtp_password(self) -> str:
        return self.smtp.get("auth", {}).get("basic", {}).get("password", "")
    
    def get_smtp_from_address(self) -> str:
        return self.smtp.get("auth", {}).get("basic", {}).get("from_address", "")


class CounterManager:
    """Verwaltet die Zähler der Anwendung"""
    
    DEFAULT_COUNTERS_FILE = "counters.json"
    
    def __init__(self, counters_file: str = None):
        self.counters_file = counters_file or self.DEFAULT_COUNTERS_FILE
        self.counters: Dict[str, Dict[str, int]] = {
            "auto": {},
            "custom": {},
            "system": {}
        }
        self._lock = threading.Lock()  # Thread-Sicherheit für Zähler
        
        self.load_counters()
    
    def load_counters(self) -> None:
        """Lädt die Zähler aus der Datei"""
        if not os.path.exists(self.counters_file):
            logger.info(f"Zählerdatei {self.counters_file} nicht gefunden, erstelle neue")
            self.create_default_counters()
            self.save_counters()
            return
        
        try:
            with open(self.counters_file, 'r', encoding='utf-8') as f:
                content = f.read()
                if not content.strip():
                    logger.warning("Zählerdatei ist leer, initialisiere mit leeren Zählern")
                    self.create_default_counters()
                    self.save_counters()
                    return
                
                data = json.loads(content)
            
            # Lade Zähler
            self.counters = data.get("counters", {
                "auto": {},
                "custom": {},
                "system": {}
            })
            
            # Stelle sicher, dass alle Kategorien existieren
            for category in ["auto", "custom", "system"]:
                if category not in self.counters:
                    self.counters[category] = {}
            
            logger.info("Zähler geladen")
            
        except json.JSONDecodeError as e:
            logger.error(f"Fehler beim Laden der Zähler (JSON ungültig): {e}")
            self.create_default_counters()
            self.save_counters()
        except Exception as e:
            logger.exception(f"Fehler beim Laden der Zähler: {e}")
            self.create_default_counters()
    
    def create_default_counters(self):
        """Erstellt Standard-Zähler"""
        self.counters = {
            "auto": {},
            "custom": {},
            "system": {}
        }
    
    def save_counters(self) -> None:
        """Speichert die Zähler in die Datei"""
        with self._lock:
            data = {
                "counters": self.counters
            }
            
            try:
                # Atomic write: Schreibe erst in temporäre Datei
                temp_file = f"{self.counters_file}.tmp"
                with open(temp_file, 'w', encoding='utf-8') as f:
                    json.dump(data, f, indent=2, ensure_ascii=False)
                
                # Ersetze alte Datei
                if os.path.exists(self.counters_file):
                    os.remove(self.counters_file)
                os.rename(temp_file, self.counters_file)
                
                logger.debug("Zähler gespeichert")
            except Exception as e:
                logger.error(f"Fehler beim Speichern der Zähler: {e}")
                if os.path.exists(temp_file):
                    os.remove(temp_file)
    
    def get_counter(self, name: str, category: str = "auto") -> int:
        """Gibt den aktuellen Wert eines Zählers zurück"""
        with self._lock:
            return self.counters.get(category, {}).get(name, 0)
    
    def increment_counter(self, name: str, category: str = "auto", increment: int = 1) -> int:
        """Erhöht einen Zähler und gibt den neuen Wert zurück"""
        with self._lock:
            if category not in self.counters:
                self.counters[category] = {}
            
            current_value = self.counters[category].get(name, 0)
            new_value = current_value + increment
            self.counters[category][name] = new_value
            
            # Speichere nach jeder Änderung
            self.save_counters()
            
            logger.debug(f"Zähler {category}/{name} erhöht: {current_value} -> {new_value}")
            return new_value
    
    def set_counter(self, name: str, value: int, category: str = "auto") -> None:
        """Setzt einen Zähler auf einen bestimmten Wert"""
        with self._lock:
            if category not in self.counters:
                self.counters[category] = {}
            
            old_value = self.counters[category].get(name, 0)
            self.counters[category][name] = value
            self.save_counters()
            
            logger.debug(f"Zähler {category}/{name} gesetzt: {old_value} -> {value}")
    
    def reset_counter(self, name: str, value: int = 0, category: str = "auto") -> None:
        """Setzt einen Zähler auf einen bestimmten Wert zurück (Standard: 0)"""
        self.set_counter(name, value, category)
    
    def delete_counter(self, name: str, category: str = "auto") -> None:
        """Löscht einen Zähler"""
        with self._lock:
            if category in self.counters and name in self.counters[category]:
                del self.counters[category][name]
                self.save_counters()
                logger.debug(f"Zähler {category}/{name} gelöscht")
    
    def get_all_counters(self, category: Optional[str] = None) -> Dict[str, Any]:
        """Gibt alle Zähler (oder alle in einer Kategorie) zurück"""
        with self._lock:
            if category:
                return self.counters.get(category, {}).copy()
            else:
                return self.counters.copy()
    
    def get_auto_counter(self, name: str) -> int:
        """Kompatibilitätsmethode für auto_Counter1 etc."""
        if name.startswith("auto_"):
            return self.get_counter(name, "auto")
        else:
            return self.get_counter(f"auto_{name}", "auto")
    
    def increment_auto_counter(self, name: str, increment: int = 1) -> int:
        """Kompatibilitätsmethode für auto_Counter1 etc."""
        if name.startswith("auto_"):
            return self.increment_counter(name, "auto", increment)
        else:
            return self.increment_counter(f"auto_{name}", "auto", increment)
    
    def get_and_increment(self, counter_name: str, start_value: int = 1, step: int = 1) -> int:
        """
        Gibt den aktuellen Wert eines Zählers zurück und erhöht ihn.
        Wenn der Zähler nicht existiert, wird er mit start_value initialisiert.
        """
        with self._lock:
            category = "auto"
            
            # Stelle sicher, dass die Kategorie existiert
            if category not in self.counters:
                self.counters[category] = {}
            
            # Hole aktuellen Wert oder initialisiere mit start_value
            current_value = self.counters[category].get(counter_name, None)
            if current_value is None:
                # Zähler existiert nicht, initialisiere mit start_value
                current_value = start_value
            else:
                # Zähler existiert, gib aktuellen Wert zurück
                pass
            
            # Erhöhe den Zähler für den nächsten Aufruf
            new_value = current_value + step
            self.counters[category][counter_name] = new_value
            
            # Speichere nach jeder Änderung
            self.save_counters()
            
            logger.debug(f"Zähler {category}/{counter_name}: {current_value} zurückgegeben, nächster Wert: {new_value}")
            return current_value
    
    def list_counters(self, category: str = "auto") -> Dict[str, int]:
        """Listet alle Zähler in einer Kategorie auf"""
        with self._lock:
            return self.counters.get(category, {}).copy()
    
    def clear_all_counters(self, category: str = "auto") -> None:
        """Löscht alle Zähler in einer Kategorie"""
        with self._lock:
            if category in self.counters:
                self.counters[category] = {}
                self.save_counters()
                logger.info(f"Alle Zähler in Kategorie '{category}' gelöscht")


# Globale Instanzen für einfachen Zugriff
_config_manager = None
_settings_manager = None
_counter_manager = None

def get_config_manager() -> ConfigManager:
    """Gibt die globale ConfigManager-Instanz zurück"""
    global _config_manager
    if _config_manager is None:
        _config_manager = ConfigManager()
    return _config_manager

def get_settings_manager() -> SettingsManager:
    """Gibt die globale SettingsManager-Instanz zurück"""
    global _settings_manager
    if _settings_manager is None:
        _settings_manager = SettingsManager()
    return _settings_manager

def get_counter_manager() -> CounterManager:
    """Gibt die globale CounterManager-Instanz zurück"""
    global _counter_manager
    if _counter_manager is None:
        _counter_manager = CounterManager()
    return _counter_manager


if __name__ == "__main__":
    # Test
    config = get_config_manager()
    settings = get_settings_manager()
    counters = get_counter_manager()
    
    print(f"Geladene Hotfolder: {len(config.hotfolders)}")
    print(f"Standard-Fehlerpfad: {settings.get_default_error_path()}")
    print(f"Zähler-Statistiken: {counters.get_all_counters()}")
