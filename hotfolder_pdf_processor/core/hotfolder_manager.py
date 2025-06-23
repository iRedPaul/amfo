"""
Zentrale Verwaltung aller Hotfolder
"""
import threading
import time
import uuid
from typing import List, Optional, Dict
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.config_manager import ConfigManager
from core.file_watcher import FileWatcher
from models.hotfolder_config import HotfolderConfig


class HotfolderManager:
    """Verwaltet alle Hotfolder und deren Überwachung"""
    
    def __init__(self):
        self.config_manager = ConfigManager()
        self.file_watcher = FileWatcher()
        self._monitor_thread: Optional[threading.Thread] = None
        self._running = False
        self._lock = threading.Lock()
    
    def start(self):
        """Startet den Hotfolder-Manager"""
        with self._lock:
            if self._running:
                return
            
            self._running = True
            
            # Starte Überwachung für alle aktivierten Hotfolder
            for hotfolder in self.config_manager.get_enabled_hotfolders():
                self.file_watcher.start_watching(hotfolder)
                # Scanne existierende Dateien
                self.file_watcher.scan_existing_files(hotfolder)
            
            # Starte Monitor-Thread
            self._monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
            self._monitor_thread.start()
            
            print("Hotfolder-Manager gestartet")
    
    def stop(self):
        """Stoppt den Hotfolder-Manager"""
        with self._lock:
            if not self._running:
                return
            
            self._running = False
            self.file_watcher.stop_all()
            
            if self._monitor_thread:
                self._monitor_thread.join(timeout=5)
            
            print("Hotfolder-Manager gestoppt")
    
    def _monitor_loop(self):
        """Hauptschleife für die Dateiverarbeitung"""
        while self._running:
            try:
                # Verarbeite ausstehende Dateien
                self.file_watcher.process_pending_files()
                
                # Kurze Pause
                time.sleep(0.5)
                
            except Exception as e:
                print(f"Fehler im Monitor-Loop: {e}")
                time.sleep(1)
    
    def create_hotfolder(self, name: str, input_path: str, output_path: str,
                        actions: List[str], action_params: Optional[dict] = None, 
                        xml_field_mappings: Optional[List[Dict]] = None,
                        output_filename_expression: str = "<FileName>",
                        ocr_zones: Optional[List[Dict]] = None) -> tuple[bool, str]:
        """Erstellt einen neuen Hotfolder"""
        try:
            # Generiere eindeutige ID
            hotfolder_id = str(uuid.uuid4())
            
            # Erstelle Konfiguration
            from models.hotfolder_config import ProcessingAction, OCRZone
            
            # Konvertiere Action-Strings zu Enums
            action_enums = []
            for action in actions:
                try:
                    action_enums.append(ProcessingAction(action))
                except ValueError:
                    return False, f"Unbekannte Aktion: {action}"
            
            # Konvertiere OCR-Zonen
            ocr_zone_objects = []
            if ocr_zones:
                for zone_dict in ocr_zones:
                    ocr_zone = OCRZone(
                        name=zone_dict['name'],
                        zone=tuple(zone_dict['zone']),
                        page_num=zone_dict['page_num']
                    )
                    ocr_zone_objects.append(ocr_zone)
            
            hotfolder = HotfolderConfig(
                id=hotfolder_id,
                name=name,
                input_path=input_path,
                output_path=output_path,
                actions=action_enums,
                action_params=action_params or {},
                output_filename_expression=output_filename_expression,
                ocr_zones=ocr_zone_objects
            )
            
            # Validiere Pfade
            valid, message = self.config_manager.validate_paths(hotfolder)
            if not valid:
                return False, message
            
            # Speichere Konfiguration
            self.config_manager.add_hotfolder(hotfolder)
            
            # Starte Überwachung wenn aktiviert
            if hotfolder.enabled and self._running:
                self.file_watcher.start_watching(hotfolder)
                self.file_watcher.scan_existing_files(hotfolder)
            
            return True, hotfolder_id
            
        except Exception as e:
            return False, f"Fehler beim Erstellen des Hotfolders: {e}"
    
    def update_hotfolder(self, hotfolder_id: str, **kwargs) -> tuple[bool, str]:
        """Aktualisiert einen bestehenden Hotfolder"""
        try:
            hotfolder = self.config_manager.get_hotfolder(hotfolder_id)
            if not hotfolder:
                return False, "Hotfolder nicht gefunden"
            
            # Stoppe aktuelle Überwachung
            if self._running:
                self.file_watcher.stop_watching(hotfolder_id)
            
            # Aktualisiere Felder
            for key, value in kwargs.items():
                if hasattr(hotfolder, key):
                    if key == "actions" and isinstance(value, list):
                        # Konvertiere Action-Strings zu Enums
                        from models.hotfolder_config import ProcessingAction
                        action_enums = []
                        for action in value:
                            try:
                                action_enums.append(ProcessingAction(action))
                            except ValueError:
                                return False, f"Unbekannte Aktion: {action}"
                        setattr(hotfolder, key, action_enums)
                    elif key == "ocr_zones" and isinstance(value, list):
                        # Konvertiere OCR-Zonen zu Objekten
                        from models.hotfolder_config import OCRZone
                        ocr_zone_objects = []
                        for zone_dict in value:
                            ocr_zone = OCRZone(
                                name=zone_dict['name'],
                                zone=tuple(zone_dict['zone']),
                                page_num=zone_dict['page_num']
                            )
                            ocr_zone_objects.append(ocr_zone)
                        setattr(hotfolder, key, ocr_zone_objects)
                    elif key == "xml_field_mappings":
                        # Stelle sicher, dass XML-Feld-Mappings übernommen werden
                        setattr(hotfolder, key, value)
                    else:
                        setattr(hotfolder, key, value)
            
            # Validiere Pfade wenn geändert
            if "input_path" in kwargs or "output_path" in kwargs:
                valid, message = self.config_manager.validate_paths(hotfolder)
                if not valid:
                    return False, message
            
            # Speichere Änderungen
            self.config_manager.update_hotfolder(hotfolder)
            
            # Starte Überwachung neu wenn aktiviert
            if hotfolder.enabled and self._running:
                self.file_watcher.start_watching(hotfolder)
                self.file_watcher.scan_existing_files(hotfolder)
            
            return True, "Hotfolder aktualisiert"
            
        except Exception as e:
            return False, f"Fehler beim Aktualisieren: {e}"
    
    def delete_hotfolder(self, hotfolder_id: str) -> tuple[bool, str]:
        """Löscht einen Hotfolder"""
        try:
            # Stoppe Überwachung
            if self._running:
                self.file_watcher.stop_watching(hotfolder_id)
            
            # Lösche Konfiguration
            self.config_manager.delete_hotfolder(hotfolder_id)
            
            return True, "Hotfolder gelöscht"
            
        except Exception as e:
            return False, f"Fehler beim Löschen: {e}"
    
    def toggle_hotfolder(self, hotfolder_id: str, enabled: bool) -> tuple[bool, str]:
        """Aktiviert oder deaktiviert einen Hotfolder"""
        try:
            success, message = self.update_hotfolder(hotfolder_id, enabled=enabled)
            
            if success:
                if enabled:
                    return True, "Hotfolder aktiviert"
                else:
                    return True, "Hotfolder deaktiviert"
            else:
                return False, message
                
        except Exception as e:
            return False, f"Fehler beim Umschalten: {e}"
    
    def get_hotfolders(self) -> List[HotfolderConfig]:
        """Gibt alle konfigurierten Hotfolder zurück"""
        return self.config_manager.hotfolders
    
    def get_hotfolder(self, hotfolder_id: str) -> Optional[HotfolderConfig]:
        """Gibt einen spezifischen Hotfolder zurück"""
        return self.config_manager.get_hotfolder(hotfolder_id)