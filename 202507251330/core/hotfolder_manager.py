"""
Zentrale Verwaltung aller Hotfolder
"""
import threading
import time
import uuid
from typing import List, Optional, Dict
import sys
import os
import logging

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.config_manager import ConfigManager
from core.file_watcher import FileWatcher
from models.hotfolder_config import HotfolderConfig
from core.license_manager import get_license_manager
from core.service_communication import ServiceCommunicationServer

# Logger für dieses Modul
logger = logging.getLogger(__name__)


class HotfolderManager:
    """Verwaltet alle Hotfolder und deren Überwachung"""
    
    def __init__(self):
        self.config_manager = ConfigManager()
        self.file_watcher = FileWatcher()
        self._monitor_thread: Optional[threading.Thread] = None
        self._rescan_thread: Optional[threading.Thread] = None
        self._running = False
        self._lock = threading.Lock()
        self._rescan_interval = 300  # 5 Minuten
        self._last_rescan = time.time()
        
        # Service-Kommunikation
        self._service_comm = ServiceCommunicationServer(self._reload_configuration)
    
    def start(self):
        """Startet den Hotfolder-Manager"""
        with self._lock:
            if self._running:
                return
            
            self._running = True
            
            # Starte Service-Kommunikation
            self._service_comm.start()
            
            # Prüfe Lizenz
            license_manager = get_license_manager()
            license_deactivated_any = False
            
            if not license_manager.is_licensed():
                logger.warning("Keine gültige Lizenz - deaktiviere alle Hotfolder")
                # Deaktiviere alle Hotfolder wenn keine Lizenz
                for hotfolder in self.config_manager.hotfolders:
                    if hotfolder.enabled:
                        hotfolder.enabled = False
                        license_deactivated_any = True
                        logger.info(f"Hotfolder '{hotfolder.name}' wurde wegen fehlender Lizenz deaktiviert")
                
                # Speichere Änderungen wenn welche deaktiviert wurden
                if license_deactivated_any:
                    self.config_manager.save_config()
            else:
                # Mit gültiger Lizenz: Starte Überwachung für alle aktivierten Hotfolder
                for hotfolder in self.config_manager.get_enabled_hotfolders():
                    # Doppelte Prüfung - nur wirklich aktivierte Hotfolder starten
                    if hotfolder.enabled:
                        self.file_watcher.start_watching(hotfolder)
                        # Scanne existierende Dateien
                        self.file_watcher.scan_existing_files(hotfolder)
            
            # Starte Monitor-Thread
            self._monitor_thread = threading.Thread(target=self._monitor_loop, daemon=True)
            self._monitor_thread.start()
            
            # Starte Rescan-Thread
            self._rescan_thread = threading.Thread(target=self._rescan_loop, daemon=True)
            self._rescan_thread.start()
            
            logger.info("Hotfolder-Manager gestartet")
            
            # Führe sofortigen Scan durch wenn Lizenz vorhanden ist
            if license_manager.is_licensed():
                # Kleine Verzögerung für Stabilität
                time.sleep(0.5)
                
                # Triggere sofortigen Scan für alle aktivierten Hotfolder
                for hotfolder in self.config_manager.get_enabled_hotfolders():
                    if hotfolder.enabled:
                        logger.info(f"Führe initialen Scan durch für: {hotfolder.name}")
                        self.file_watcher.rescan_hotfolder(hotfolder)
            
            # Gib Flag zurück ob Hotfolder deaktiviert wurden
            return license_deactivated_any
    
    def stop(self):
        """Stoppt den Hotfolder-Manager"""
        with self._lock:
            if not self._running:
                return
            
            self._running = False
            
            # Stoppe Service-Kommunikation
            self._service_comm.stop()
            
            self.file_watcher.stop_all()
            
            if self._monitor_thread:
                self._monitor_thread.join(timeout=5)
            
            if self._rescan_thread:
                self._rescan_thread.join(timeout=5)
            
            logger.info("Hotfolder-Manager gestoppt")
    
    def _monitor_loop(self):
        """Hauptschleife für die Dateiverarbeitung"""
        while self._running:
            try:
                # Verarbeite ausstehende Dateien
                self.file_watcher.process_pending_files()
                
                # Kurze Pause
                time.sleep(0.5)
                
            except Exception as e:
                logger.error(f"Fehler im Monitor-Loop: {e}")
                time.sleep(1)
    
    def _rescan_loop(self):
        """Periodischer Rescan aller Hotfolder und Lizenzprüfung."""
        while self._running:
            try:
                current_time = time.time()

                if current_time - self._last_rescan >= self._rescan_interval:
                    logger.info(f"Führe periodischen Scan und Lizenzprüfung durch (alle {self._rescan_interval} Sekunden)...")

                    # Periodische Lizenzprüfung
                    license_manager = get_license_manager()
                    if not license_manager.is_licensed():
                        logger.warning("Lizenz ist abgelaufen oder ungültig. Deaktiviere alle Hotfolder.")
                        # Stoppe die Überwachung für alle laufenden Hotfolder
                        for hotfolder_id in list(self.file_watcher.observers.keys()):
                            self.file_watcher.stop_watching(hotfolder_id)

                        # Setze alle Hotfolder im Speicher auf 'disabled' (wird nicht gespeichert, um die GUI nicht zu überschreiben)
                        for hotfolder in self.config_manager.hotfolders:
                            hotfolder.enabled = False

                        logger.info("Alle Hotfolder-Überwachungen wurden aufgrund der ungültigen Lizenz gestoppt.")
                    else:
                        # Rescan nur durchführen, wenn die Lizenz gültig ist
                        for hotfolder in self.config_manager.get_enabled_hotfolders():
                            if hotfolder.id in self.file_watcher.handlers:
                                logger.debug(f"Rescanne Hotfolder: {hotfolder.name}")
                                self.file_watcher.rescan_hotfolder(hotfolder)

                    self._last_rescan = current_time
                    logger.debug("Periodischer Scan abgeschlossen.")

                # Prüfe alle 10 Sekunden, ob der nächste Scan fällig ist
                time.sleep(10)

            except Exception as e:
                logger.error(f"Fehler im Rescan-Loop: {e}")
                time.sleep(60)  # Bei Fehler länger warten

    def _reload_configuration(self):
        """Lädt die Konfiguration neu und aktualisiert die Überwachungen"""
        try:
            with self._lock:
                logger.info("Starte Config-Reload (ausgelöst durch GUI)...")
                
                # Merke aktuelle Hotfolder-IDs
                old_hotfolder_ids = {hf.id for hf in self.config_manager.hotfolders}
                
                # Lade neue Konfiguration
                self.config_manager.load_config()
                
                # Neue Hotfolder-IDs
                new_hotfolder_ids = {hf.id for hf in self.config_manager.hotfolders}
                
                # Finde Änderungen
                removed_ids = old_hotfolder_ids - new_hotfolder_ids
                added_ids = new_hotfolder_ids - old_hotfolder_ids
                existing_ids = old_hotfolder_ids & new_hotfolder_ids
                
                # Prüfe Lizenz für neue Konfiguration
                license_manager = get_license_manager()
                if not license_manager.is_licensed():
                    logger.warning("Keine gültige Lizenz - deaktiviere alle Hotfolder")
                    for hotfolder in self.config_manager.hotfolders:
                        if hotfolder.enabled:
                            hotfolder.enabled = False
                            logger.info(f"Hotfolder '{hotfolder.name}' wurde wegen fehlender Lizenz deaktiviert")
                    self.config_manager.save_config()
                
                # Entferne gelöschte Hotfolder
                for hotfolder_id in removed_ids:
                    logger.info(f"Entferne Überwachung für gelöschten Hotfolder: {hotfolder_id}")
                    self.file_watcher.stop_watching(hotfolder_id)
                
                # Aktualisiere bestehende Hotfolder
                for hotfolder_id in existing_ids:
                    hotfolder = self.config_manager.get_hotfolder(hotfolder_id)
                    if hotfolder:
                        # Stoppe alte Überwachung
                        self.file_watcher.stop_watching(hotfolder_id)
                        
                        # Starte neue Überwachung wenn aktiviert
                        if hotfolder.enabled:
                            logger.info(f"Aktualisiere Überwachung für: {hotfolder.name}")
                            self.file_watcher.start_watching(hotfolder)
                            self.file_watcher.scan_existing_files(hotfolder)
                        else:
                            logger.info(f"Hotfolder deaktiviert: {hotfolder.name}")
                
                # Füge neue Hotfolder hinzu
                for hotfolder_id in added_ids:
                    hotfolder = self.config_manager.get_hotfolder(hotfolder_id)
                    if hotfolder and hotfolder.enabled:
                        logger.info(f"Starte Überwachung für neuen Hotfolder: {hotfolder.name}")
                        self.file_watcher.start_watching(hotfolder)
                        self.file_watcher.scan_existing_files(hotfolder)
                
                logger.info("Config-Reload abgeschlossen")
                
        except Exception as e:
            logger.error(f"Fehler beim Config-Reload: {e}")
    
    def create_hotfolder(self, name: str, input_path: str,
                            description: str = "",
                            process_pairs: bool = True,
                            actions: Optional[List[str]] = None,
                            action_params: Optional[dict] = None,
                            xml_field_mappings: Optional[List[Dict]] = None,
                            output_filename_expression: str = "<FileName>",
                            ocr_zones: Optional[List[Dict]] = None,
                            export_configs: Optional[List[Dict]] = None,
                            stamp_configs: Optional[List[Dict]] = None,
                            error_path: str = "",
                            file_patterns: Optional[List[str]] = None) -> tuple[bool, str]:
        """Erstellt einen neuen Hotfolder"""
        try:
            # Prüfe auf doppelten Input-Ordner nur bei AKTIVIERTEN Hotfoldern (normalisiert, case-insensitive)
            new_input_norm = os.path.normcase(os.path.abspath(input_path))
            for hf in self.get_hotfolders():
                # Nur aktivierte Hotfolder prüfen
                if hf.enabled:
                    existing_input_norm = os.path.normcase(os.path.abspath(hf.input_path))
                    if new_input_norm == existing_input_norm:
                        return False, f"Es existiert bereits ein aktivierter Hotfolder mit dem gleichen Input-Ordner: {hf.input_path}"

            # Lizenzprüfung direkt hier durchführen
            is_licensed = get_license_manager().is_licensed()
            is_enabled = is_licensed  # Hotfolder ist nur aktiv, wenn eine Lizenz vorhanden ist

            # Generiere eindeutige ID
            hotfolder_id = str(uuid.uuid4())

            # Erstelle Konfiguration
            from models.hotfolder_config import ProcessingAction, OCRZone

            # Konvertiere Action-Strings zu Enums
            if actions is None:
                action_enums = []
            else:
                action_enums = []
                for action in actions:
                    try:
                        action_enums.append(ProcessingAction(action))
                    except ValueError:
                        return False, f"Unbekannte Aktion: {action}"

            # Konvertiere OCR-Zonen und stelle sicher, dass sie OCR_ Präfix haben
            ocr_zone_objects = []
            if ocr_zones:
                for zone_dict in ocr_zones:
                    # Stelle sicher, dass der Name ein OCR_ Präfix hat
                    zone_name = zone_dict['name']
                    if not zone_name.startswith('OCR_'):
                        zone_name = f'OCR_{zone_name}'

                    ocr_zone = OCRZone(
                        name=zone_name,
                        zone=tuple(zone_dict['zone']),
                        page_num=zone_dict['page_num']
                    )
                    ocr_zone_objects.append(ocr_zone)

            if xml_field_mappings is None:
                xml_field_mappings = []
            if export_configs is None:
                export_configs = []
            if stamp_configs is None:
                stamp_configs = []
            if ocr_zone_objects is None:
                ocr_zone_objects = []
            if file_patterns is None:
                file_patterns = ['*.pdf']
                
            hotfolder = HotfolderConfig(
                id=hotfolder_id,
                name=name,
                input_path=input_path,
                enabled=is_enabled,
                description=description,
                process_pairs=process_pairs,
                actions=action_enums,
                action_params=action_params or {},
                xml_field_mappings=xml_field_mappings,
                output_filename_expression=output_filename_expression,
                ocr_zones=ocr_zone_objects,
                export_configs=export_configs,
                stamp_configs=stamp_configs,
                error_path=error_path,
                file_patterns=file_patterns
            )

            # Validiere Pfade
            valid, message = self.config_manager.validate_paths(hotfolder)
            if not valid:
                return False, message

            # Speichere Konfiguration
            self.config_manager.add_hotfolder(hotfolder)

            # Starte Überwachung wenn aktiviert (wird nur passieren, wenn is_licensed true war)
            if hotfolder.enabled and self._running:
                self.file_watcher.start_watching(hotfolder)
                self.file_watcher.scan_existing_files(hotfolder)

            # --- NEU: Passe die Erfolgsmeldung an ---
            success_message = f"Hotfolder '{name}' erfolgreich erstellt."
            if not is_enabled:
                success_message += " (Deaktiviert, da keine gültige Lizenz vorhanden ist)"
                
            return True, success_message

        except Exception as e:
            logger.exception("Fehler beim Erstellen des Hotfolders")
            return False, f"Fehler beim Erstellen des Hotfolders: {e}"

    def update_hotfolder(self, hotfolder_id: str, **kwargs) -> tuple[bool, str]:
        """Aktualisiert einen bestehenden Hotfolder"""
        try:
            hotfolder = self.config_manager.get_hotfolder(hotfolder_id)
            if not hotfolder:
                return False, "Hotfolder nicht gefunden"
            
            # Prüfe auf doppelten Input-Ordner, falls input_path geändert wird
            if "input_path" in kwargs:
                new_input_path = kwargs["input_path"]
                new_input_norm = os.path.normcase(os.path.abspath(new_input_path))
                for hf in self.get_hotfolders():
                    if hf.id == hotfolder_id:
                        continue  # Sich selbst überspringen
                    # Nur aktivierte Hotfolder prüfen
                    if hf.enabled:
                        existing_input_norm = os.path.normcase(os.path.abspath(hf.input_path))
                        if new_input_norm == existing_input_norm:
                            return False, f"Es existiert bereits ein aktivierter Hotfolder mit dem gleichen Input-Ordner: {hf.input_path}"
            
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
                        # Konvertiere OCR-Zonen zu Objekten und stelle sicher, dass sie OCR_ Präfix haben
                        from models.hotfolder_config import OCRZone
                        ocr_zone_objects = []
                        for zone_dict in value:
                            # Stelle sicher, dass der Name ein OCR_ Präfix hat
                            zone_name = zone_dict['name']
                            if not zone_name.startswith('OCR_'):
                                zone_name = f'OCR_{zone_name}'
                            
                            ocr_zone = OCRZone(
                                name=zone_name,
                                zone=tuple(zone_dict['zone']),
                                page_num=zone_dict['page_num']
                            )
                            ocr_zone_objects.append(ocr_zone)
                        setattr(hotfolder, key, ocr_zone_objects)
                    elif key == "xml_field_mappings":
                        # Stelle sicher, dass XML-Feld-Mappings übernommen werden
                        setattr(hotfolder, key, value)
                    elif key == "export_configs":
                        # Stelle sicher, dass Export-Konfigurationen übernommen werden
                        setattr(hotfolder, key, value)
                    elif key == "stamp_configs":
                        # Stelle sicher, dass Stempel-Konfigurationen übernommen werden
                        setattr(hotfolder, key, value)
                    elif key == "error_path":
                        # Stelle sicher, dass Fehlerpfad übernommen wird
                        setattr(hotfolder, key, value)
                    elif key == "description":
                        # Stelle sicher, dass Beschreibung übernommen wird
                        setattr(hotfolder, key, value)
                    else:
                        setattr(hotfolder, key, value)
            
            # Validiere Pfade wenn geändert
            if "input_path" in kwargs or "output_path" in kwargs:
                valid, message = self.config_manager.validate_paths(hotfolder)
                if not valid:
                    return False, message
            
            # Speichere Änderungen (löst automatisch Config-Reload im Dienst aus)
            self.config_manager.update_hotfolder(hotfolder_id, hotfolder)
            
            # WICHTIG: Starte Überwachung NICHT hier neu, da der Config-Reload das bereits macht!
            # Das verhindert die doppelte Überwachung und die Warnung
            
            return True, "Hotfolder aktualisiert"
            
        except Exception as e:
            logger.exception("Fehler beim Aktualisieren")
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
            logger.exception("Fehler beim Löschen")
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
            logger.exception("Fehler beim Umschalten")
            return False, f"Fehler beim Umschalten: {e}"
    
    def get_hotfolders(self) -> List[HotfolderConfig]:
        """Gibt alle konfigurierten Hotfolder zurück"""
        return self.config_manager.hotfolders
    
    def get_hotfolder(self, hotfolder_id: str) -> Optional[HotfolderConfig]:
        """Gibt einen spezifischen Hotfolder zurück"""
        return self.config_manager.get_hotfolder(hotfolder_id)
    
    def set_rescan_interval(self, seconds: int):
        """Setzt das Rescan-Intervall in Sekunden"""
        if seconds >= 60:  # Mindestens 1 Minute
            self._rescan_interval = seconds
            logger.info(f"Rescan-Intervall geändert auf {seconds} Sekunden")
