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
from models.hotfolder_config import HotfolderConfig, ProcessingAction, OCRZone
from core.license_manager import get_license_manager
from core.service_communication import ServiceCommunicationServer

# Logger für dieses Modul
logger = logging.getLogger(__name__)


class HotfolderManager:
    """Verwaltet alle Hotfolder und deren Überwachung"""
    
    # Konstanten
    DEFAULT_RESCAN_INTERVAL = 300  # 5 Minuten
    MIN_RESCAN_INTERVAL = 60       # 1 Minute
    RESCAN_CHECK_INTERVAL = 10     # 10 Sekunden
    THREAD_JOIN_TIMEOUT = 5        # 5 Sekunden
    INITIAL_SCAN_DELAY = 0.5       # 0.5 Sekunden
    MONITOR_LOOP_DELAY = 0.5       # 0.5 Sekunden
    MAX_ERROR_WAIT_TIME = 300      # 5 Minuten max Wartezeit bei Fehlern
    
    def __init__(self):
        self.config_manager = ConfigManager()
        self.file_watcher = FileWatcher()
        self._monitor_thread: Optional[threading.Thread] = None
        self._rescan_thread: Optional[threading.Thread] = None
        self._running = False
        self._lock = threading.Lock()
        self._rescan_interval = self.DEFAULT_RESCAN_INTERVAL
        self._last_rescan = time.time()
        
        # Service-Kommunikation
        self._service_comm = ServiceCommunicationServer(self._reload_configuration)
        
        logger.info("HotfolderManager initialisiert")
    
    def start(self) -> bool:
        """
        Startet den Hotfolder-Manager
        
        Returns:
            bool: True wenn Hotfolder wegen fehlender Lizenz deaktiviert wurden
        """
        with self._lock:
            if self._running:
                logger.warning("HotfolderManager läuft bereits")
                return False
            
            logger.info("Starte HotfolderManager...")
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
                enabled_count = 0
                for hotfolder in self.config_manager.get_enabled_hotfolders():
                    # Doppelte Prüfung - nur wirklich aktivierte Hotfolder starten
                    if hotfolder.enabled:
                        try:
                            self.file_watcher.start_watching(hotfolder)
                            # Scanne existierende Dateien
                            self.file_watcher.scan_existing_files(hotfolder)
                            enabled_count += 1
                            logger.info(f"Überwachung gestartet für: {hotfolder.name}")
                        except Exception as e:
                            logger.error(f"Fehler beim Starten der Überwachung für {hotfolder.name}: {e}")
                
                logger.info(f"{enabled_count} Hotfolder-Überwachungen gestartet")
            
            # Starte Monitor-Thread
            self._monitor_thread = threading.Thread(
                target=self._monitor_loop, 
                daemon=True,
                name="HotfolderMonitor"
            )
            self._monitor_thread.start()
            
            # Starte Rescan-Thread
            self._rescan_thread = threading.Thread(
                target=self._rescan_loop, 
                daemon=True,
                name="HotfolderRescan"
            )
            self._rescan_thread.start()
            
            logger.info("HotfolderManager erfolgreich gestartet")
            
            # Führe sofortigen Scan durch wenn Lizenz vorhanden ist
            if license_manager.is_licensed():
                # Kleine Verzögerung für Stabilität
                time.sleep(self.INITIAL_SCAN_DELAY)
                
                # Triggere sofortigen Scan für alle aktivierten Hotfolder
                scan_count = 0
                for hotfolder in self.config_manager.get_enabled_hotfolders():
                    if hotfolder.enabled:
                        try:
                            logger.info(f"Führe initialen Scan durch für: {hotfolder.name}")
                            self.file_watcher.rescan_hotfolder(hotfolder)
                            scan_count += 1
                        except Exception as e:
                            logger.error(f"Fehler beim initialen Scan für {hotfolder.name}: {e}")
                
                logger.info(f"Initialer Scan für {scan_count} Hotfolder abgeschlossen")
            
            # Gib Flag zurück ob Hotfolder deaktiviert wurden
            return license_deactivated_any
    
    def stop(self):
        """Stoppt den Hotfolder-Manager mit graceful shutdown"""
        logger.info("Starte Shutdown des HotfolderManagers...")
        
        with self._lock:
            if not self._running:
                logger.warning("HotfolderManager läuft nicht")
                return
            
            self._running = False
        
        # Service-Kommunikation außerhalb des Locks stoppen
        try:
            self._service_comm.stop()
            logger.info("Service-Kommunikation gestoppt")
        except Exception as e:
            logger.error(f"Fehler beim Stoppen der Service-Kommunikation: {e}")
        
        # Stoppe alle Datei-Überwachungen
        try:
            self.file_watcher.stop_all()
            logger.info("Alle Datei-Überwachungen gestoppt")
        except Exception as e:
            logger.error(f"Fehler beim Stoppen der Datei-Überwachungen: {e}")
        
        # Warte auf Threads
        for thread, name in [(self._monitor_thread, "Monitor"), 
                             (self._rescan_thread, "Rescan")]:
            if thread and thread.is_alive():
                logger.info(f"Warte auf {name}-Thread...")
                thread.join(timeout=self.THREAD_JOIN_TIMEOUT)
                if thread.is_alive():
                    logger.warning(f"{name}-Thread reagiert nicht auf Shutdown-Signal")
                else:
                    logger.info(f"{name}-Thread beendet")
        
        logger.info("HotfolderManager erfolgreich gestoppt")
    
    def _monitor_loop(self):
        """Hauptschleife für die Dateiverarbeitung"""
        logger.info("Monitor-Thread gestartet")
        consecutive_errors = 0
        
        while self._running:
            try:
                # Verarbeite ausstehende Dateien
                processed_count = self.file_watcher.process_pending_files()
                
                if processed_count > 0:
                    logger.debug(f"{processed_count} Dateien verarbeitet")
                
                # Reset Fehlerzähler bei Erfolg
                consecutive_errors = 0
                
                # Kurze Pause
                time.sleep(self.MONITOR_LOOP_DELAY)
                
            except Exception as e:
                consecutive_errors += 1
                wait_time = min(60 * consecutive_errors, self.MAX_ERROR_WAIT_TIME)
                logger.error(f"Fehler im Monitor-Loop (#{consecutive_errors}): {e}", exc_info=True)
                time.sleep(wait_time)
        
        logger.info("Monitor-Thread beendet")
    
    def _rescan_loop(self):
        """Periodischer Rescan aller Hotfolder und Lizenzprüfung."""
        logger.info("Rescan-Thread gestartet")
        consecutive_errors = 0
        
        while self._running:
            try:
                current_time = time.time()

                if current_time - self._last_rescan >= self._rescan_interval:
                    logger.info(f"Starte periodischen Scan (Intervall: {self._rescan_interval}s)")
                    
                    with self._lock:
                        # Periodische Lizenzprüfung
                        license_manager = get_license_manager()
                        if not license_manager.is_licensed():
                            logger.warning("Lizenz ist abgelaufen oder ungültig. Deaktiviere alle Hotfolder.")
                            
                            # Stoppe die Überwachung für alle laufenden Hotfolder
                            stopped_count = 0
                            for hotfolder_id in list(self.file_watcher.observers.keys()):
                                try:
                                    self.file_watcher.stop_watching(hotfolder_id)
                                    stopped_count += 1
                                except Exception as e:
                                    logger.error(f"Fehler beim Stoppen der Überwachung {hotfolder_id}: {e}")
                            
                            # Deaktiviere alle Hotfolder und speichere die Änderung
                            deactivated_count = 0
                            for hotfolder in self.config_manager.hotfolders:
                                if hotfolder.enabled:
                                    hotfolder.enabled = False
                                    deactivated_count += 1
                            
                            # Speichere die Änderungen in der Konfiguration
                            if deactivated_count > 0:
                                self.config_manager.save_config()
                                logger.info(f"{stopped_count} Überwachungen gestoppt, "
                                          f"{deactivated_count} Hotfolder deaktiviert und gespeichert")
                        else:
                            # Rescan nur durchführen, wenn die Lizenz gültig ist
                            scan_count = 0
                            error_count = 0
                            
                            for hotfolder in self.config_manager.get_enabled_hotfolders():
                                if hotfolder.id in self.file_watcher.handlers:
                                    try:
                                        logger.debug(f"Rescanne Hotfolder: {hotfolder.name}")
                                        self.file_watcher.rescan_hotfolder(hotfolder)
                                        scan_count += 1
                                    except Exception as e:
                                        error_count += 1
                                        logger.error(f"Fehler beim Rescan von {hotfolder.name}: {e}")
                            
                            logger.info(f"Periodischer Scan abgeschlossen: {scan_count} erfolgreich, "
                                      f"{error_count} Fehler")
                    
                    self._last_rescan = current_time
                    
                # Reset Fehlerzähler bei Erfolg
                consecutive_errors = 0
                
                # Prüfe regelmäßig, ob der nächste Scan fällig ist
                time.sleep(self.RESCAN_CHECK_INTERVAL)

            except Exception as e:
                consecutive_errors += 1
                wait_time = min(60 * consecutive_errors, self.MAX_ERROR_WAIT_TIME)
                logger.error(f"Fehler im Rescan-Loop (#{consecutive_errors}): {e}", exc_info=True)
                time.sleep(wait_time)
        
        logger.info("Rescan-Thread beendet")

    def _reload_configuration(self):
            """Lädt die Konfiguration neu und aktualisiert die Überwachungen"""
            try:
                with self._lock:
                    logger.info("Starte Config-Reload (ausgelöst durch GUI)...")
                    
                    # Performance-Messung
                    start_time = time.time()
                    
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
                        
                        # WICHTIG: Stoppe zuerst alle laufenden Überwachungen
                        stopped_count = 0
                        for hotfolder_id in list(self.file_watcher.observers.keys()):
                            try:
                                self.file_watcher.stop_watching(hotfolder_id)
                                stopped_count += 1
                                logger.info(f"Überwachung gestoppt für Hotfolder: {hotfolder_id}")
                            except Exception as e:
                                logger.error(f"Fehler beim Stoppen der Überwachung {hotfolder_id}: {e}")
                        
                        # Dann deaktiviere alle Hotfolder in der Konfiguration
                        deactivated_count = 0
                        for hotfolder in self.config_manager.hotfolders:
                            if hotfolder.enabled:
                                hotfolder.enabled = False
                                deactivated_count += 1
                                logger.info(f"Hotfolder '{hotfolder.name}' wurde wegen fehlender Lizenz deaktiviert")
                        
                        if deactivated_count > 0:
                            self.config_manager.save_config()
                            logger.info(f"{stopped_count} Überwachungen sofort gestoppt, "
                                      f"{deactivated_count} Hotfolder deaktiviert und gespeichert")
                        elif stopped_count > 0:
                            logger.info(f"{stopped_count} Überwachungen wurden sofort gestoppt")
                        
                        # Früher Ausstieg - keine weiteren Aktionen nötig
                        elapsed_time = time.time() - start_time
                        logger.info(f"Config-Reload abgeschlossen in {elapsed_time:.2f}s - "
                                  f"Alle Hotfolder wegen fehlender Lizenz deaktiviert")
                        return
                    
                    # Entferne gelöschte Hotfolder
                    for hotfolder_id in removed_ids:
                        try:
                            logger.info(f"Entferne Überwachung für gelöschten Hotfolder: {hotfolder_id}")
                            self.file_watcher.stop_watching(hotfolder_id)
                        except Exception as e:
                            logger.error(f"Fehler beim Entfernen der Überwachung {hotfolder_id}: {e}")
                    
                    # Aktualisiere bestehende Hotfolder
                    updated_count = 0
                    for hotfolder_id in existing_ids:
                        hotfolder = self.config_manager.get_hotfolder(hotfolder_id)
                        if hotfolder:
                            try:
                                # Stoppe alte Überwachung
                                self.file_watcher.stop_watching(hotfolder_id)
                                
                                # Starte neue Überwachung wenn aktiviert
                                if hotfolder.enabled:
                                    logger.info(f"Aktualisiere Überwachung für: {hotfolder.name}")
                                    self.file_watcher.start_watching(hotfolder)
                                    self.file_watcher.scan_existing_files(hotfolder)
                                    updated_count += 1
                                else:
                                    logger.info(f"Hotfolder deaktiviert: {hotfolder.name}")
                            except Exception as e:
                                logger.error(f"Fehler beim Aktualisieren von {hotfolder.name}: {e}")
                    
                    # Füge neue Hotfolder hinzu
                    added_count = 0
                    for hotfolder_id in added_ids:
                        hotfolder = self.config_manager.get_hotfolder(hotfolder_id)
                        if hotfolder and hotfolder.enabled:
                            try:
                                logger.info(f"Starte Überwachung für neuen Hotfolder: {hotfolder.name}")
                                self.file_watcher.start_watching(hotfolder)
                                self.file_watcher.scan_existing_files(hotfolder)
                                added_count += 1
                            except Exception as e:
                                logger.error(f"Fehler beim Hinzufügen von {hotfolder.name}: {e}")
                    
                    # Performance-Log
                    elapsed_time = time.time() - start_time
                    logger.info(f"Config-Reload abgeschlossen in {elapsed_time:.2f}s - "
                              f"Entfernt: {len(removed_ids)}, Hinzugefügt: {added_count}, "
                              f"Aktualisiert: {updated_count}")
                    
            except Exception as e:
                logger.error(f"Fehler beim Config-Reload: {e}", exc_info=True)

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
        """
        Erstellt einen neuen Hotfolder
        
        Returns:
            tuple[bool, str]: (Erfolg, Nachricht)
        """
        try:
            with self._lock:
                # Prüfe auf doppelten Input-Ordner nur bei AKTIVIERTEN Hotfoldern
                new_input_norm = os.path.normcase(os.path.abspath(input_path))
                for hf in self.config_manager.hotfolders:  # Direkt auf config_manager zugreifen
                    # Nur aktivierte Hotfolder prüfen
                    if hf.enabled:
                        existing_input_norm = os.path.normcase(os.path.abspath(hf.input_path))
                        if new_input_norm == existing_input_norm:
                            return False, f"Es existiert bereits ein aktivierter Hotfolder mit dem gleichen Input-Ordner: {hf.input_path}"

                # Lizenzprüfung
                is_licensed = get_license_manager().is_licensed()
                is_enabled = is_licensed  # Hotfolder ist nur aktiv, wenn eine Lizenz vorhanden ist

                # Generiere eindeutige ID
                hotfolder_id = str(uuid.uuid4())

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

                # Setze Standardwerte
                if xml_field_mappings is None:
                    xml_field_mappings = []
                if export_configs is None:
                    export_configs = []
                if stamp_configs is None:
                    stamp_configs = []
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

                # Starte Überwachung wenn aktiviert
                if hotfolder.enabled and self._running:
                    try:
                        self.file_watcher.start_watching(hotfolder)
                        self.file_watcher.scan_existing_files(hotfolder)
                        logger.info(f"Überwachung für neuen Hotfolder '{name}' gestartet")
                    except Exception as e:
                        logger.error(f"Fehler beim Starten der Überwachung für '{name}': {e}")

                # Erstelle Erfolgsmeldung
                success_message = f"Hotfolder '{name}' erfolgreich erstellt."
                if not is_enabled:
                    success_message += " (Deaktiviert, da keine gültige Lizenz vorhanden ist)"
                    
                logger.info(success_message)
                return True, success_message

        except Exception as e:
            logger.exception("Fehler beim Erstellen des Hotfolders")
            return False, f"Fehler beim Erstellen des Hotfolders: {str(e)}"

    def update_hotfolder(self, hotfolder_id: str, **kwargs) -> tuple[bool, str]:
        """
        Aktualisiert einen bestehenden Hotfolder
        
        Returns:
            tuple[bool, str]: (Erfolg, Nachricht)
        """
        try:
            with self._lock:
                hotfolder = self.config_manager.get_hotfolder(hotfolder_id)
                if not hotfolder:
                    return False, "Hotfolder nicht gefunden"
                
                # Prüfe auf doppelten Input-Ordner, falls input_path geändert wird
                if "input_path" in kwargs:
                    new_input_path = kwargs["input_path"]
                    new_input_norm = os.path.normcase(os.path.abspath(new_input_path))
                    for hf in self.config_manager.hotfolders:  # Direkt auf config_manager zugreifen
                        if hf.id == hotfolder_id:
                            continue  # Sich selbst überspringen
                        # Nur aktivierte Hotfolder prüfen
                        if hf.enabled:
                            existing_input_norm = os.path.normcase(os.path.abspath(hf.input_path))
                            if new_input_norm == existing_input_norm:
                                return False, f"Es existiert bereits ein aktivierter Hotfolder mit dem gleichen Input-Ordner: {hf.input_path}"
                
                # Stoppe aktuelle Überwachung
                if self._running:
                    try:
                        self.file_watcher.stop_watching(hotfolder_id)
                    except Exception as e:
                        logger.error(f"Fehler beim Stoppen der Überwachung: {e}")
                
                # Aktualisiere Felder
                for key, value in kwargs.items():
                    if hasattr(hotfolder, key):
                        if key == "actions" and isinstance(value, list):
                            # Konvertiere Action-Strings zu Enums
                            action_enums = []
                            for action in value:
                                try:
                                    action_enums.append(ProcessingAction(action))
                                except ValueError:
                                    return False, f"Unbekannte Aktion: {action}"
                            setattr(hotfolder, key, action_enums)
                        elif key == "ocr_zones" and isinstance(value, list):
                            # Konvertiere OCR-Zonen zu Objekten
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
                        else:
                            setattr(hotfolder, key, value)
                
                # Validiere Pfade wenn geändert
                if "input_path" in kwargs:
                    valid, message = self.config_manager.validate_paths(hotfolder)
                    if not valid:
                        return False, message
                
                # Speichere Änderungen (löst automatisch Config-Reload im Dienst aus)
                self.config_manager.update_hotfolder(hotfolder_id, hotfolder)
                
                logger.info(f"Hotfolder '{hotfolder.name}' aktualisiert")
                return True, "Hotfolder aktualisiert"
                
        except Exception as e:
            logger.exception("Fehler beim Aktualisieren")
            return False, f"Fehler beim Aktualisieren: {e}"
  
    def delete_hotfolder(self, hotfolder_id: str) -> tuple[bool, str]:
        """
        Löscht einen Hotfolder
        
        Returns:
            tuple[bool, str]: (Erfolg, Nachricht)
        """
        try:
            with self._lock:
                hotfolder = self.config_manager.get_hotfolder(hotfolder_id)
                if not hotfolder:
                    return False, "Hotfolder nicht gefunden"
                
                # Stoppe Überwachung
                if self._running:
                    try:
                        self.file_watcher.stop_watching(hotfolder_id)
                    except Exception as e:
                        logger.error(f"Fehler beim Stoppen der Überwachung: {e}")
                
                # Lösche Konfiguration
                self.config_manager.delete_hotfolder(hotfolder_id)
                
                logger.info(f"Hotfolder '{hotfolder.name}' gelöscht")
                return True, "Hotfolder gelöscht"
                
        except Exception as e:
            logger.exception("Fehler beim Löschen")
            return False, f"Fehler beim Löschen: {e}"
    
    def toggle_hotfolder(self, hotfolder_id: str, enabled: bool) -> tuple[bool, str]:
        """
        Aktiviert oder deaktiviert einen Hotfolder
        
        Returns:
            tuple[bool, str]: (Erfolg, Nachricht)
        """
        try:
            success, message = self.update_hotfolder(hotfolder_id, enabled=enabled)
            
            if success:
                status = "aktiviert" if enabled else "deaktiviert"
                logger.info(f"Hotfolder {hotfolder_id} {status}")
                return True, f"Hotfolder {status}"
            else:
                return False, message
                
        except Exception as e:
            logger.exception("Fehler beim Umschalten")
            return False, f"Fehler beim Umschalten: {e}"
    
    def get_hotfolders(self) -> List[HotfolderConfig]:
        """Gibt alle konfigurierten Hotfolder zurück"""
        # Thread-safe Kopie der Liste
        with self._lock:
            return list(self.config_manager.hotfolders)
    
    def get_hotfolder(self, hotfolder_id: str) -> Optional[HotfolderConfig]:
        """Gibt einen spezifischen Hotfolder zurück"""
        with self._lock:
            return self.config_manager.get_hotfolder(hotfolder_id)
    
    def _get_hotfolders_unlocked(self) -> List[HotfolderConfig]:
        """Interne Methode für Zugriff ohne Lock (nur innerhalb von Lock-Blöcken verwenden!)"""
        return list(self.config_manager.hotfolders)
    
    def set_rescan_interval(self, seconds: int):
        """Setzt das Rescan-Intervall in Sekunden"""
        if seconds >= self.MIN_RESCAN_INTERVAL:
            old_interval = self._rescan_interval
            self._rescan_interval = seconds
            logger.info(f"Rescan-Intervall geändert von {old_interval}s auf {seconds}s")
        else:
            logger.warning(f"Rescan-Intervall {seconds}s ist zu klein. Minimum: {self.MIN_RESCAN_INTERVAL}s")
