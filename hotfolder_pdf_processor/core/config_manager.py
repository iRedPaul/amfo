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


class ConfigManager:
    """Verwaltet die Konfiguration der Anwendung"""
    
    def __init__(self, config_file: str = "config.json"):
        self.config_file = config_file
        self.hotfolders: List[HotfolderConfig] = []
        self.load_config()
    
    def load_config(self) -> None:
        """Lädt die Konfiguration aus der Datei"""
        if not os.path.exists(self.config_file):
            self.save_config()  # Erstelle leere Konfiguration
            return
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                content = f.read()
                if not content.strip():  # Datei ist leer
                    self.hotfolders = []
                    self.save_config()
                    return
                    
                data = json.loads(content)
                self.hotfolders = [
                    HotfolderConfig.from_dict(hf) 
                    for hf in data.get("hotfolders", [])
                ]
        except json.JSONDecodeError as e:
            print(f"Fehler beim Laden der Konfiguration (JSON ungültig): {e}")
            self.hotfolders = []
            self.save_config()  # Erstelle neue gültige Konfiguration
        except Exception as e:
            print(f"Fehler beim Laden der Konfiguration: {e}")
            self.hotfolders = []
    
    def save_config(self) -> None:
        """Speichert die Konfiguration in die Datei"""
        data = {
            "hotfolders": [hf.to_dict() for hf in self.hotfolders]
        }
        
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Fehler beim Speichern der Konfiguration: {e}")
    
    def add_hotfolder(self, hotfolder: HotfolderConfig) -> None:
        """Fügt einen neuen Hotfolder hinzu"""
        self.hotfolders.append(hotfolder)
        self.save_config()
    
    def update_hotfolder(self, hotfolder: HotfolderConfig) -> None:
        """Aktualisiert einen bestehenden Hotfolder"""
        for i, hf in enumerate(self.hotfolders):
            if hf.id == hotfolder.id:
                self.hotfolders[i] = hotfolder
                self.save_config()
                return
    
    def delete_hotfolder(self, hotfolder_id: str) -> None:
        """Löscht einen Hotfolder"""
        self.hotfolders = [hf for hf in self.hotfolders if hf.id != hotfolder_id]
        self.save_config()
    
    def get_hotfolder(self, hotfolder_id: str) -> Optional[HotfolderConfig]:
        """Gibt einen spezifischen Hotfolder zurück"""
        for hf in self.hotfolders:
            if hf.id == hotfolder_id:
                return hf
        return None
    
    def get_enabled_hotfolders(self) -> List[HotfolderConfig]:
        """Gibt alle aktivierten Hotfolder zurück"""
        return [hf for hf in self.hotfolders if hf.enabled]
    
    def validate_paths(self, hotfolder: HotfolderConfig) -> tuple[bool, str]:
        """Validiert die Pfade eines Hotfolders"""
        # Prüfe Input-Pfad
        if not os.path.exists(hotfolder.input_path):
            try:
                os.makedirs(hotfolder.input_path, exist_ok=True)
            except Exception as e:
                return False, f"Input-Pfad konnte nicht erstellt werden: {e}"
        
        # Prüfe Output-Pfad
        if not os.path.exists(hotfolder.output_path):
            try:
                os.makedirs(hotfolder.output_path, exist_ok=True)
            except Exception as e:
                return False, f"Output-Pfad konnte nicht erstellt werden: {e}"
        
        # Prüfe ob Input und Output unterschiedlich sind
        if os.path.abspath(hotfolder.input_path) == os.path.abspath(hotfolder.output_path):
            return False, "Input- und Output-Ordner dürfen nicht identisch sein"
        
        return True, "OK"