import json
import os
import logging
import threading
from typing import Dict, List, Optional, Any, Tuple
from dataclasses import dataclass, field, asdict
import uuid
from datetime import datetime
from models.hotfolder_config import HotfolderConfig, ProcessingAction, OCRZone

logger = logging.getLogger(__name__)


def ensure_config_directory():
    """Stellt sicher, dass das config Verzeichnis existiert"""
    config_dir = "config"
    if not os.path.exists(config_dir):
        os.makedirs(config_dir)
        logger.info(f"Config-Verzeichnis erstellt: {config_dir}")


class ConfigManager:
    """Verwaltet die Konfiguration der Anwendung"""
    
    DEFAULT_CONFIG_FILE = "config/hotfolders.json"
    
    def __init__(self, config_file: Optional[str] = None):
        ensure_config_directory()
        if config_file is None:
            config_file = self.DEFAULT_CONFIG_FILE
        self.config_file = str(config_file)
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
            self.hotfolders = [HotfolderConfig.from_dict(hf_data) for hf_data in data.get('hotfolders', [])]
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
            'version': '1.0',
            'hotfolders': [hf.to_dict() for hf in self.hotfolders]
        }
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            logger.info("Konfiguration gespeichert")
            
            # Sende Reload-Signal an den Dienst
            try:
                from core.service_communication import ServiceCommunicationClient
                success, response = ServiceCommunicationClient.reload_config()
                if success:
                    logger.info(f"Dienst benachrichtigt: {response.get('message', 'OK')}")
                else:
                    logger.warning(f"Dienst konnte nicht benachrichtigt werden: {response}")
            except ImportError:
                logger.debug("Service-Kommunikation nicht verfügbar (läuft wahrscheinlich als Dienst)")
            except Exception as e:
                logger.warning(f"Fehler beim Benachrichtigen des Dienstes: {e}")
                
        except Exception as e:
            logger.error(f"Fehler beim Speichern der Konfiguration: {e}")
    
    def validate_paths(self, hotfolder: HotfolderConfig) -> Tuple[bool, str]:
        """Validiert die Pfade eines Hotfolders
        
        Returns:
            Tuple[bool, str]: (Gültig, Fehlermeldung oder Erfolgsmeldung)
        """
        try:
            # Prüfe Input-Pfad
            if not hotfolder.input_path:
                return False, "Input-Pfad ist erforderlich"
            
            # Normalisiere Pfade für Vergleich
            input_path = os.path.abspath(hotfolder.input_path)
            
            # Prüfe ob Input-Pfad existiert oder erstellt werden kann
            if not os.path.exists(input_path):
                try:
                    os.makedirs(input_path, exist_ok=True)
                    logger.info(f"Input-Pfad erstellt: {input_path}")
                except Exception as e:
                    return False, f"Input-Pfad konnte nicht erstellt werden: {e}"
            
            # Prüfe ob Input-Pfad ein Verzeichnis ist
            if not os.path.isdir(input_path):
                return False, "Input-Pfad muss ein Verzeichnis sein"
            
            # Prüfe Schreibrechte für Input-Pfad
            if not os.access(input_path, os.W_OK):
                return False, "Keine Schreibrechte für Input-Pfad"
            
            # Prüfe Export-Pfade wenn vorhanden
            if hasattr(hotfolder, 'export_configs') and hotfolder.export_configs:
                for export_config in hotfolder.export_configs:
                    if hasattr(export_config, 'export_path') and export_config.export_path:
                        export_path = export_config.export_path
                        
                        # Prüfe ob Export-Pfad ein Verzeichnis ist (wenn es existiert)
                        if os.path.exists(export_path) and not os.path.isdir(export_path):
                            return False, f"Export-Pfad muss ein Verzeichnis sein: {export_path}"
                        
                        # Versuche Export-Pfad zu erstellen wenn er nicht existiert
                        if not os.path.exists(export_path):
                            try:
                                os.makedirs(export_path, exist_ok=True)
                                logger.info(f"Export-Pfad erstellt: {export_path}")
                            except Exception as e:
                                return False, f"Export-Pfad konnte nicht erstellt werden: {export_path} - {e}"
            
            # Prüfe Fehler-Pfad wenn vorhanden
            if hasattr(hotfolder, 'error_path') and hotfolder.error_path:
                error_path = hotfolder.error_path
                
                # Prüfe ob Fehler-Pfad existiert oder erstellt werden kann
                if not os.path.exists(error_path):
                    try:
                        os.makedirs(error_path, exist_ok=True)
                        logger.info(f"Fehler-Pfad erstellt: {error_path}")
                    except Exception as e:
                        return False, f"Fehler-Pfad konnte nicht erstellt werden: {e}"
                
                # Prüfe ob Fehler-Pfad ein Verzeichnis ist
                if not os.path.isdir(error_path):
                    return False, "Fehler-Pfad muss ein Verzeichnis sein"
            
            return True, "Pfade sind gültig"
            
        except Exception as e:
            logger.exception("Fehler bei der Pfad-Validierung")
            return False, f"Fehler bei der Pfad-Validierung: {str(e)}"
    
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
        """Gibt einen Hotfolder anhand seiner ID zurück"""
        for hf in self.hotfolders:
            if hf.id == hotfolder_id:
                return hf
        return None
    
    def get_all_hotfolders(self) -> List[HotfolderConfig]:
        """Gibt alle Hotfolder zurück"""
        return self.hotfolders.copy()
    
    def get_enabled_hotfolders(self) -> List[HotfolderConfig]:
        """Gibt alle aktivierten Hotfolder zurück"""
        return [hf for hf in self.hotfolders if hf.enabled]
    
    def is_input_path_used(self, input_path: str, exclude_id: Optional[str] = None) -> Optional[str]:
        """Prüft ob ein Input-Pfad bereits von einem anderen Hotfolder verwendet wird.
        Gibt den Namen des Hotfolders zurück, der ihn verwendet."""
        input_norm = os.path.normcase(os.path.abspath(input_path))
        for hf in self.hotfolders:
            if exclude_id and hf.id == exclude_id:
                continue  # Skip self
            existing_norm = os.path.normcase(os.path.abspath(hf.input_path))
            if input_norm == existing_norm:
                return hf.name
        return None
    
    def export_hotfolder(self, hotfolder_id: str, export_path: str) -> tuple[bool, str]:
        """Exportiert einen einzelnen Hotfolder als JSON-Datei"""
        try:
            hotfolder = self.get_hotfolder(hotfolder_id)
            if not hotfolder:
                return False, "Hotfolder nicht gefunden"
            
            # Erstelle Export-Daten mit nur einem Hotfolder
            export_data = {
                'version': '1.0',
                'hotfolders': [hotfolder.to_dict()]
            }
            
            # Speichere in Datei
            with open(export_path, 'w', encoding='utf-8') as f:
                json.dump(export_data, f, indent=2, ensure_ascii=False)
            
            logger.info(f"Hotfolder '{hotfolder.name}' exportiert nach {export_path}")
            return True, "Hotfolder erfolgreich exportiert"
            
        except Exception as e:
            logger.exception(f"Fehler beim Exportieren des Hotfolders: {e}")
            return False, f"Fehler beim Exportieren: {str(e)}"
    
    def import_hotfolder(self, import_path: str, generate_new_id: bool = True) -> tuple[bool, str]:
        """Importiert einen Hotfolder aus einer JSON-Datei. Der Hotfolder wird standardmäßig deaktiviert."""
        try:
            # Lade JSON-Datei
            with open(import_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Validiere Struktur
            if 'hotfolders' not in data or not isinstance(data['hotfolders'], list):
                return False, "Ungültige Dateistruktur: 'hotfolders' Array fehlt"
            
            if len(data['hotfolders']) == 0:
                return False, "Keine Hotfolder in der Datei gefunden"
            
            # Nimm den ersten Hotfolder
            hotfolder_data = data['hotfolders'][0]
            
            # Erstelle HotfolderConfig aus den Daten
            imported_hotfolder = HotfolderConfig.from_dict(hotfolder_data)
            
            # Generiere neue ID wenn gewünscht (verhindert Duplikate)
            if generate_new_id:
                imported_hotfolder.id = str(uuid.uuid4())
            
            # Deaktiviere den Hotfolder beim Import
            imported_hotfolder.enabled = False
            
            # Prüfe ob ID bereits existiert
            if not generate_new_id:
                for hf in self.hotfolders:
                    if hf.id == imported_hotfolder.id:
                        return False, "Ein Hotfolder mit dieser ID existiert bereits"
            
            # Füge Hotfolder hinzu
            self.add_hotfolder(imported_hotfolder)
            
            logger.info(f"Hotfolder '{imported_hotfolder.name}' importiert (deaktiviert)")
            return True, f"Hotfolder '{imported_hotfolder.name}' erfolgreich importiert (deaktiviert)"
            
        except json.JSONDecodeError as e:
            logger.error(f"JSON-Fehler beim Importieren: {e}")
            return False, "Ungültige JSON-Datei"
        except Exception as e:
            logger.exception(f"Fehler beim Importieren des Hotfolders: {e}")
            return False, f"Fehler beim Importieren: {str(e)}"


class SettingsManager:
    """Verwaltet die globalen Einstellungen der Anwendung"""
    
    DEFAULT_SETTINGS_FILE = "config/settings.json"
    
    def __init__(self, settings_file: Optional[str] = None):
        ensure_config_directory()
        if settings_file is None:
            settings_file = self.DEFAULT_SETTINGS_FILE
        self.settings_file = str(settings_file)
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
                    "provider": "",
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
            
            # Benachrichtige den Dienst auch bei Settings-Änderungen
            try:
                from core.service_communication import ServiceCommunicationClient
                success, response = ServiceCommunicationClient.reload_config()
                if success:
                    logger.debug(f"Dienst über Settings-Änderung benachrichtigt")
            except ImportError:
                pass  # Service-Kommunikation nicht verfügbar
            except Exception:
                pass  # Fehler ignorieren
                
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
    
    DEFAULT_COUNTERS_FILE = "config/counters.json"
    
    def __init__(self, counters_file: Optional[str] = None):
        ensure_config_directory()
        if counters_file is None:
            counters_file = self.DEFAULT_COUNTERS_FILE
        self.counters_file = str(counters_file)
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