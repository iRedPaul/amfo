"""
Verwaltung der Anwendungskonfiguration
"""
import json
import os
from typing import List, Optional
from pathlib import Path
import sys

# Füge das Hauptverzeichnis zum Python-Pfad hinzu
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.hotfolder_config import HotfolderConfig
from core.logger import get_logger


class ConfigManager:
    """Verwaltet die Konfiguration der Anwendung"""
    
    def __init__(self, config_file: str = "config.json"):
        self.config_file = config_file
        self.hotfolders: List[HotfolderConfig] = []
        self.logger = get_logger('ConfigManager')
        self.load_config()
    
    def load_config(self) -> None:
        """Lädt die Konfiguration aus der Datei"""
        if not os.path.exists(self.config_file):
            self.logger.info(f"Konfigurationsdatei {self.config_file} nicht gefunden, erstelle neue")
            self.save_config()  # Erstelle leere Konfiguration
            return
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                content = f.read()
                if not content.strip():  # Datei ist leer
                    self.logger.warning("Konfigurationsdatei ist leer, erstelle neue")
                    self.hotfolders = []
                    self.save_config()
                    return
                    
                data = json.loads(content)
                self.hotfolders = [
                    HotfolderConfig.from_dict(hf) 
                    for hf in data.get("hotfolders", [])
                ]
                self.logger.info(f"Konfiguration geladen: {len(self.hotfolders)} Hotfolder")
        except json.JSONDecodeError as e:
            self.logger.error(f"Fehler beim Laden der Konfiguration (JSON ungültig): {e}")
            self.hotfolders = []
            self.save_config()  # Erstelle neue gültige Konfiguration
        except Exception as e:
            self.logger.error(f"Fehler beim Laden der Konfiguration: {e}")
            self.hotfolders = []
    
    def save_config(self) -> None:
        """Speichert die Konfiguration in die Datei"""
        data = {
            "hotfolders": [hf.to_dict() for hf in self.hotfolders]
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            self.logger.debug(f"Konfiguration gespeichert: {len(self.hotfolders)} Hotfolder")
        except Exception as e:
            self.logger.error(f"Fehler beim Speichern der Konfiguration: {e}")
    
    def add_hotfolder(self, hotfolder: HotfolderConfig) -> None:
        """Fügt einen neuen Hotfolder hinzu"""
        self.hotfolders.append(hotfolder)
        self.save_config()
        self.logger.info(f"Hotfolder hinzugefügt: {hotfolder.name}")
    
    def update_hotfolder(self, hotfolder: HotfolderConfig) -> None:
        """Aktualisiert einen bestehenden Hotfolder"""
        for i, hf in enumerate(self.hotfolders):
            if hf.id == hotfolder.id:
                self.hotfolders[i] = hotfolder
                self.save_config()
                self.logger.info(f"Hotfolder aktualisiert: {hotfolder.name}")
                return
    
    def delete_hotfolder(self, hotfolder_id: str) -> None:
        """Löscht einen Hotfolder"""
        hotfolder_name = None
        for hf in self.hotfolders:
            if hf.id == hotfolder_id:
                hotfolder_name = hf.name
                break
        
        self.hotfolders = [hf for hf in self.hotfolders if hf.id != hotfolder_id]
        self.save_config()
        
        if hotfolder_name:
            self.logger.info(f"Hotfolder gelöscht: {hotfolder_name}")
    
    def get_hotfolder(self, hotfolder_id: str) -> Optional[HotfolderConfig]:
        """Gibt einen spezifischen Hotfolder zurück"""
        for hf in self.hotfolders:
            if hf.id == hotfolder_id:
                return hf
        return None
    
    def get_enabled_hotfolders(self) -> List[HotfolderConfig]:
        """Gibt alle aktivierten Hotfolder zurück"""
        enabled = [hf for hf in self.hotfolders if hf.enabled]
        self.logger.debug(f"Aktivierte Hotfolder: {len(enabled)} von {len(self.hotfolders)}")
        return enabled
    
    def validate_paths(self, hotfolder: HotfolderConfig) -> tuple[bool, str]:
        """Validiert die Pfade eines Hotfolders"""
        # Prüfe Input-Pfad
        if not os.path.exists(hotfolder.input_path):
            try:
                os.makedirs(hotfolder.input_path, exist_ok=True)
                self.logger.info(f"Input-Pfad erstellt: {hotfolder.input_path}")
            except Exception as e:
                self.logger.error(f"Input-Pfad konnte nicht erstellt werden: {e}")
                return False, f"Input-Pfad konnte nicht erstellt werden: {e}"
        
        # Bei neuer Export-basierter Konfiguration gibt es keinen festen Output-Pfad mehr
        # Der Output-Pfad wird nur für Legacy-Kompatibilität verwendet
        # Daher keine Validierung des Output-Pfads gegen Input-Pfad
        
        # Wenn Export-Konfigurationen vorhanden sind, ist alles OK
        if hasattr(hotfolder, 'export_configs') and hotfolder.export_configs:
            return True, "OK"
        
        # Legacy: Prüfe Output-Pfad nur wenn keine Exporte definiert
        if not hasattr(hotfolder, 'export_configs') or not hotfolder.export_configs:
            # Für Legacy-Kompatibilität: Erstelle Output-Ordner im Input-Ordner
            if not hotfolder.output_path:
                hotfolder.output_path = os.path.join(hotfolder.input_path, "output")
            
            if not os.path.exists(hotfolder.output_path):
                try:
                    os.makedirs(hotfolder.output_path, exist_ok=True)
                    self.logger.info(f"Output-Pfad erstellt: {hotfolder.output_path}")
                except Exception as e:
                    self.logger.error(f"Output-Pfad konnte nicht erstellt werden: {e}")
                    return False, f"Output-Pfad konnte nicht erstellt werden: {e}"
        
        return True, "OK"