"""
Vereinheitlichter Manager für Konfiguration, Einstellungen und Zähler
"""
import json
import os
from typing import List, Optional, Dict, Any
import logging
import threading
import sys

# Füge das Hauptverzeichnis zum Python-Pfad hinzu
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.hotfolder_config import HotfolderConfig

logger = logging.getLogger(__name__)


class ConfigManager:
    """Verwaltet die Konfiguration der Anwendung"""
    
    def __init__(self, config_file: str = "config.json"):
        self.config_file = config_file
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
                content = f.read()
                if not content.strip():
                    logger.warning("Konfigurationsdatei ist leer, initialisiere mit leerer Konfiguration")
                    self.create_default_config()
                    self.save_config()
                    return
                    
                data = json.loads(content)
            
            # Lade Hotfolders
            self.hotfolders = []
            for hf_data in data.get("hotfolders", []):
                hotfolder_dict = self._convert_from_storage(hf_data)
                self.hotfolders.append(HotfolderConfig.from_dict(hotfolder_dict))
            
            logger.info(f"Konfiguration geladen: {len(self.hotfolders)} Hotfolder")
            
        except json.JSONDecodeError as e:
            logger.error(f"Fehler beim Laden der Konfiguration (JSON ungültig): {e}")
            self.create_default_config()
            self.save_config()
        except Exception as e:
            logger.exception(f"Fehler beim Laden der Konfiguration: {e}")
            self.hotfolders = []
    
    def create_default_config(self):
        """Erstellt eine Standard-Konfiguration"""
        self.hotfolders = []
    
    def _convert_from_storage(self, stored_hf: Dict[str, Any]) -> Dict[str, Any]:
        """Konvertiert von gespeicherter Struktur zur internen HotfolderConfig-Struktur"""
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
    
    def save_config(self) -> None:
        """Speichert die Konfiguration in die Datei"""
        data = {
            "hotfolders": [
                self._convert_to_storage(hf) 
                for hf in self.hotfolders
            ]
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            logger.debug(f"Konfiguration gespeichert: {len(self.hotfolders)} Hotfolder")
        except Exception as e:
            logger.error(f"Fehler beim Speichern der Konfiguration: {e}")
    
    def add_hotfolder(self, hotfolder: HotfolderConfig) -> None:
        """Fügt einen neuen Hotfolder hinzu"""
        self.hotfolders.append(hotfolder)
        logger.info(f"Hotfolder hinzugefügt: {hotfolder.name}")
        self.save_config()
    
    def update_hotfolder(self, hotfolder: HotfolderConfig) -> None:
        """Aktualisiert einen bestehenden Hotfolder"""
        for i, hf in enumerate(self.hotfolders):
            if hf.id == hotfolder.id:
                self.hotfolders[i] = hotfolder
                logger.info(f"Hotfolder aktualisiert: {hotfolder.name}")
                self.save_config()
                return
        logger.warning(f"Hotfolder mit ID {hotfolder.id} nicht gefunden")
    
    def delete_hotfolder(self, hotfolder_id: str) -> None:
        """Löscht einen Hotfolder"""
        initial_count = len(self.hotfolders)
        self.hotfolders = [hf for hf in self.hotfolders if hf.id != hotfolder_id]
        
        if len(self.hotfolders) < initial_count:
            logger.info(f"Hotfolder mit ID {hotfolder_id} gelöscht")
            self.save_config()
        else:
            logger.warning(f"Hotfolder mit ID {hotfolder_id} nicht gefunden")
    
    def get_hotfolder(self, hotfolder_id: str) -> Optional[HotfolderConfig]:
        """Gibt einen spezifischen Hotfolder zurück"""
        for hf in self.hotfolders:
            if hf.id == hotfolder_id:
                return hf
        logger.debug(f"Hotfolder mit ID {hotfolder_id} nicht gefunden")
        return None
    
    def get_enabled_hotfolders(self) -> List[HotfolderConfig]:
        """Gibt alle aktivierten Hotfolder zurück"""
        enabled = [hf for hf in self.hotfolders if hf.enabled]
        logger.debug(f"{len(enabled)} von {len(self.hotfolders)} Hotfoldern sind aktiviert")
        return enabled
    
    def validate_paths(self, hotfolder: HotfolderConfig) -> tuple[bool, str]:
        """Validiert die Pfade eines Hotfolders"""
        # Prüfe Input-Pfad
        if not os.path.exists(hotfolder.input_path):
            try:
                os.makedirs(hotfolder.input_path, exist_ok=True)
                logger.info(f"Input-Pfad erstellt: {hotfolder.input_path}")
            except Exception as e:
                return False, f"Input-Pfad konnte nicht erstellt werden: {e}"
        
        # Prüfe Output-Pfad (falls verschieden von Input)
        if hotfolder.output_path and hotfolder.output_path != hotfolder.input_path:
            if not os.path.exists(hotfolder.output_path):
                try:
                    os.makedirs(hotfolder.output_path, exist_ok=True)
                    logger.info(f"Output-Pfad erstellt: {hotfolder.output_path}")
                except Exception as e:
                    return False, f"Output-Pfad konnte nicht erstellt werden: {e}"
        
        return True, "Pfade sind gültig"


class SettingsManager:
    """Verwaltet die globalen Einstellungen der Anwendung"""
    
    DEFAULT_SETTINGS_FILE = "settings.json"
    
    def __init__(self, settings_file: str = None):
        self.settings_file = settings_file or self.DEFAULT_SETTINGS_FILE
        self.paths: Dict[str, str] = {}
        self.smtp: Dict[str, Any] = {}
        self.processing: Dict[str, Any] = {}
        self.ui: Dict[str, Any] = {}
        self.logging_config: Dict[str, Any] = {}
        
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
                content = f.read()
                if not content.strip():
                    logger.warning("Einstellungsdatei ist leer, initialisiere mit Standardeinstellungen")
                    self.create_default_settings()
                    self.save_settings()
                    return
                
                data = json.loads(content)
            
            # Lade alle Bereiche
            self.paths = data.get("paths", self._get_default_paths())
            self.smtp = data.get("smtp", self._get_default_smtp())
            self.processing = data.get("processing", self._get_default_processing())
            self.ui = data.get("ui", self._get_default_ui())
            self.logging_config = data.get("logging", self._get_default_logging())
            
            logger.info("Einstellungen geladen")
            
        except json.JSONDecodeError as e:
            logger.error(f"Fehler beim Laden der Einstellungen (JSON ungültig): {e}")
            self.create_default_settings()
            self.save_settings()
        except Exception as e:
            logger.exception(f"Fehler beim Laden der Einstellungen: {e}")
            self.create_default_settings()
    
    def create_default_settings(self):
        """Erstellt Standardeinstellungen"""
        self.paths = self._get_default_paths()
        self.smtp = self._get_default_smtp()
        self.processing = self._get_default_processing()
        self.ui = self._get_default_ui()
        self.logging_config = self._get_default_logging()
    
    def _get_default_paths(self) -> Dict[str, str]:
        """Gibt Standard-Pfade zurück"""
        appdata = os.getenv('APPDATA')
        if appdata:
            default_error_path = os.path.join(appdata, 'HotfolderPDFProcessor', 'errors')
        else:
            default_error_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'errors')
        
        os.makedirs(default_error_path, exist_ok=True)
        
        return {
            "default_error": default_error_path
        }
    
    def _get_default_smtp(self) -> Dict[str, Any]:
        """Gibt Standard-SMTP-Einstellungen zurück"""
        return {
            "server": {
                "host": "",
                "port": 587,
                "use_ssl": False,
                "use_tls": True
            },
            "auth": {
                "method": "basic",
                "basic": {
                    "username": "",
                    "password": "",
                    "from_address": ""
                },
                "oauth2": {
                    "provider": "",
                    "client_id": "",
                    "client_secret": "",
                    "refresh_token": "",
                    "access_token": "",
                    "token_expiry": ""
                }
            }
        }
    
    def _get_default_processing(self) -> Dict[str, Any]:
        """Gibt Standard-Verarbeitungseinstellungen zurück"""
        return {
            "defaults": {},
            "limits": {}
        }
    
    def _get_default_ui(self) -> Dict[str, Any]:
        """Gibt Standard-UI-Einstellungen zurück"""
        return {
            "preferences": {},
            "recent": {}
        }
    
    def _get_default_logging(self) -> Dict[str, Any]:
        """Gibt Standard-Logging-Einstellungen zurück"""
        return {
            "level": "INFO",
            "path": None,
            "max_size": None,
            "keep_days": None
        }
    
    def save_settings(self) -> None:
        """Speichert die Einstellungen in die Datei"""
        data = {
            "paths": self.paths,
            "smtp": self.smtp,
            "processing": self.processing,
            "ui": self.ui,
            "logging": self.logging_config
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
    
    def reset_counter(self, name: str, category: str = "auto") -> None:
        """Setzt einen Zähler auf 0 zurück"""
        self.set_counter(name, 0, category)
    
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