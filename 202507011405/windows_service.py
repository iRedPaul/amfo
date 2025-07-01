"""
Windows Service für Hotfolder PDF Processor
"""
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys
import os
import time
import logging
from pathlib import Path

# Füge das Hauptverzeichnis zum Python-Pfad hinzu
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from core.hotfolder_manager import HotfolderManager
from core.logging_config import setup_logging


class HotfolderService(win32serviceutil.ServiceFramework):
    """Windows Service für die Hotfolder-Verarbeitung"""

    _svc_name_ = "HotfolderPDFProcessor"
    _svc_display_name_ = "Hotfolder PDF Processor Service"
    _svc_description_ = "Überwacht Hotfolder und verarbeitet PDF-Dateien automatisch"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.manager = None
        self.logger = None
        self.running = False

    def _setup_service_logging(self):
        """Konfiguriert das Service-spezifische Logging"""
        # Verwende das zentrale Logging-System
        service_log_dir = Path(os.environ.get('PROGRAMDATA', 'C:\\ProgramData')) / 'HotfolderPDFProcessor' / 'logs'
        service_log_dir.mkdir(parents=True, exist_ok=True)
        
        # Initialisiere das Logging mit dem vorhandenen System
        self.logger = setup_logging(log_dir=service_log_dir, log_level=logging.INFO)
        
        # Zusätzlicher Handler für Windows Event Log
        try:
            if sys.platform == 'win32':
                import win32evtlogutil
                import win32evtlog
                
                # Event Log Handler für wichtige Meldungen
                class ServiceEventLogHandler(logging.Handler):
                    def __init__(self, service_name):
                        logging.Handler.__init__(self)
                        self.service_name = service_name
                        
                    def emit(self, record):
                        try:
                            msg = self.format(record)
                            if record.levelno >= logging.ERROR:
                                win32evtlogutil.ReportEvent(
                                    self.service_name,
                                    servicemanager.PYS_SERVICE_ERROR,
                                    0,
                                    servicemanager.EVENTLOG_ERROR_TYPE,
                                    (msg,)
                                )
                            elif record.levelno >= logging.WARNING:
                                win32evtlogutil.ReportEvent(
                                    self.service_name,
                                    servicemanager.PYS_SERVICE_WARNING,
                                    0,
                                    servicemanager.EVENTLOG_WARNING_TYPE,
                                    (msg,)
                                )
                        except:
                            pass
                
                # Füge Event Log Handler nur für Fehler und Warnungen hinzu
                event_handler = ServiceEventLogHandler(self._svc_name_)
                event_handler.setLevel(logging.WARNING)
                event_handler.setFormatter(logging.Formatter('%(levelname)s: %(message)s'))
                self.logger.addHandler(event_handler)
                
        except ImportError:
            # Win32-Module nicht verfügbar, nutze nur File-Logging
            pass
        
        return self.logger

    def SvcStop(self):
        """Stoppt den Service"""
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        
        if self.logger:
            self.logger.info("Service wird gestoppt...")
        
        self.running = False
        
        if self.manager:
            try:
                self.manager.stop()
                if self.logger:
                    self.logger.info("Hotfolder-Manager gestoppt")
            except Exception as e:
                if self.logger:
                    self.logger.error(f"Fehler beim Stoppen des Managers: {e}")

        win32event.SetEvent(self.hWaitStop)
        
        if self.logger:
            self.logger.info("Service gestoppt")

    def SvcDoRun(self):
        """Hauptschleife des Services"""
        try:
            # Initialisiere Logging
            self._setup_service_logging()
            
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, '')
            )

            self.logger.info(f"Service '{self._svc_name_}' wird gestartet...")
            self.running = True
            self.main()

        except Exception as e:
            if self.logger:
                self.logger.error(f"Kritischer Fehler im Service: {e}", exc_info=True)
            servicemanager.LogErrorMsg(f"Service Fehler: {e}")
            self.SvcStop()

    def main(self):
        """Hauptlogik des Services"""
        try:
            # Wechsle ins Arbeitsverzeichnis
            service_path = os.path.dirname(os.path.abspath(__file__))
            os.chdir(service_path)
            self.logger.info(f"Arbeitsverzeichnis: {os.getcwd()}")

            # Initialisiere und starte den Hotfolder-Manager
            self.logger.info("Initialisiere Hotfolder-Manager...")
            self.manager = HotfolderManager()
            self.manager.start()

            self.logger.info("Hotfolder-Manager erfolgreich gestartet")
            
            # Zeige aktive Hotfolder
            total_count = len(self.manager.config_manager.hotfolders)
            active_count = len([h for h in self.manager.config_manager.hotfolders if h.enabled])
            self.logger.info(f"Hotfolder geladen: {active_count} von {total_count} aktiv")
            
            # Liste der aktiven Hotfolder
            for hotfolder in self.manager.config_manager.hotfolders:
                if hotfolder.enabled:
                    self.logger.info(f"  - {hotfolder.name}: {hotfolder.input_path}")

            # Service-Hauptschleife
            last_status_time = 0
            
            while self.running:
                # Warte auf Stop-Event (alle 5 Sekunden prüfen)
                rc = win32event.WaitForSingleObject(self.hWaitStop, 5000)
                
                if rc == win32event.WAIT_OBJECT_0:
                    # Stop-Event empfangen
                    break

                # Status-Meldung alle 5 Minuten
                current_time = int(time.time())
                if current_time - last_status_time >= 300:  # 300 Sekunden = 5 Minuten
                    active_count = len([h for h in self.manager.config_manager.hotfolders if h.enabled])
                    self.logger.info(f"Service läuft - {active_count} aktive Hotfolder überwacht")
                    last_status_time = current_time

        except Exception as e:
            self.logger.error(f"Fehler in der Service-Hauptschleife: {e}", exc_info=True)
            raise
        finally:
            # Aufräumen
            if self.manager:
                try:
                    self.manager.stop()
                except:
                    pass
            
            self.logger.info("Service-Hauptschleife beendet")


def install_service():
    """Installiert den Windows Service"""
    if len(sys.argv) == 1:
        print("="*60)
        print("Hotfolder PDF Processor - Windows Service")
        print("="*60)
        print("\nVerwendung:")
        print("  python windows_service.py install    - Service installieren")
        print("  python windows_service.py start      - Service starten")
        print("  python windows_service.py stop       - Service stoppen")
        print("  python windows_service.py remove     - Service entfernen")
        print("  python windows_service.py status     - Service-Status anzeigen")
        print("\nHinweis: Administrator-Rechte erforderlich!")
        print("="*60)
        sys.exit(0)

    elif sys.argv[1].lower() == 'status':
        try:
            status = win32serviceutil.QueryServiceStatus(HotfolderService._svc_name_)
            state_map = {
                1: "STOPPED (Gestoppt)",
                2: "START_PENDING (Wird gestartet)",
                3: "STOP_PENDING (Wird gestoppt)",
                4: "RUNNING (Läuft)",
                5: "CONTINUE_PENDING (Wird fortgesetzt)",
                6: "PAUSE_PENDING (Wird pausiert)",
                7: "PAUSED (Pausiert)",
            }
            
            state = state_map.get(status[1], f'UNKNOWN ({status[1]})')
            print(f"\nService-Status: {state}")
            
            # Prüfe ob Service installiert ist
            try:
                config = win32serviceutil.GetServiceCustomOption(HotfolderService._svc_name_, 'config')
                print("Service ist installiert.")
            except:
                pass
                
            # Zeige Log-Verzeichnis
            log_dir = Path(os.environ.get('PROGRAMDATA', 'C:\\ProgramData')) / 'HotfolderPDFProcessor' / 'logs'
            print(f"\nLog-Dateien unter: {log_dir}")
            
            # Zeige neueste Log-Einträge wenn Service läuft
            if status[1] == 4:  # RUNNING
                log_files = sorted(log_dir.glob("hotfolder_*.log"), reverse=True)
                if log_files:
                    latest_log = log_files[0]
                    print(f"\nNeueste Log-Datei: {latest_log.name}")
                    
        except Exception as e:
            print(f"Fehler beim Abrufen des Service-Status: {e}")
            print("\nMögliche Ursachen:")
            print("- Service ist nicht installiert")
            print("- Keine Administrator-Rechte")
        sys.exit(0)

    # Alle anderen Kommandos an pywin32 weiterleiten
    try:
        win32serviceutil.HandleCommandLine(HotfolderService)
    except Exception as e:
        print(f"\nFehler: {e}")
        print("\nStellen Sie sicher, dass Sie Administrator-Rechte haben!")
        sys.exit(1)


if __name__ == '__main__':
    install_service()