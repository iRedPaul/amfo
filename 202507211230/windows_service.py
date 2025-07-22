"""
Windows Service für belegpilot
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
import traceback
from pathlib import Path
import threading

# Füge das Hauptverzeichnis zum Python-Pfad hinzu
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))


class BelegpilotService(win32serviceutil.ServiceFramework):
    """Windows Service für die belegpilot Hotfolder-Verarbeitung"""

    _svc_name_ = "belegpilot"
    _svc_display_name_ = "belegpilot Service"
    _svc_description_ = "Überwacht Hotfolder und verarbeitet PDF-Dateien automatisch mit belegpilot"

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.manager = None
        self.logger = None
        self.running = False
        self.main_thread = None

    def _setup_service_logging(self):
        """Konfiguriert das Service-spezifische Logging"""
        try:
            # Erstelle Log-Verzeichnis
            service_log_dir = Path(os.environ.get('PROGRAMDATA', 'C:\\ProgramData')) / 'belegpilot' / 'logs'
            service_log_dir.mkdir(parents=True, exist_ok=True)
            
            # Basis-Logger einrichten
            self.logger = logging.getLogger('BelegpilotService')
            self.logger.setLevel(logging.DEBUG)
            
            # Entferne alte Handler
            self.logger.handlers = []
            
            # File Handler
            log_file = service_log_dir / f'service_{time.strftime("%Y%m%d")}.log'
            file_handler = logging.FileHandler(log_file, encoding='utf-8')
            file_handler.setLevel(logging.DEBUG)
            file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(file_formatter)
            self.logger.addHandler(file_handler)
            
            self.logger.info("="*60)
            self.logger.info("Service-Logging initialisiert")
            self.logger.info(f"Log-Datei: {log_file}")
            self.logger.info(f"Python-Version: {sys.version}")
            self.logger.info(f"Executable: {sys.executable}")
            self.logger.info(f"Arbeitsverzeichnis: {os.getcwd()}")
            self.logger.info("="*60)
            
            return True
        except Exception as e:
            # Schreibe minimal in Event Log
            try:
                import win32evtlogutil
                win32evtlogutil.ReportEvent(
                    "belegpilot",
                    2,  # Warning
                    eventCategory=0,
                    eventType=win32evtlog.EVENTLOG_WARNING_TYPE,
                    strings=[f"Logging-Setup fehlgeschlagen: {str(e)}"]
                )
            except:
                pass
            return False

    def SvcStop(self):
        """Stoppt den Service"""
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        
        try:
            if self.logger:
                self.logger.info("Service Stop angefordert...")
            
            self.running = False
            win32event.SetEvent(self.hWaitStop)
            
            # Warte kurz auf Thread
            if self.main_thread and self.main_thread.is_alive():
                self.main_thread.join(timeout=5)
            
            # Stoppe Manager
            if self.manager:
                try:
                    self.manager.stop()
                    if self.logger:
                        self.logger.info("Hotfolder-Manager gestoppt")
                except Exception as e:
                    if self.logger:
                        self.logger.error(f"Fehler beim Stoppen: {e}")
            
            self.ReportServiceStatus(win32service.SERVICE_STOPPED)
            
            if self.logger:
                self.logger.info("Service vollständig gestoppt")
                
        except Exception as e:
            if self.logger:
                self.logger.error(f"Fehler in SvcStop: {e}", exc_info=True)

    def SvcDoRun(self):
        """Hauptschleife des Services - MUSS schnell antworten!"""
        # WICHTIG: Sofort als RUNNING melden!
        self.ReportServiceStatus(win32service.SERVICE_START_PENDING)
        
        try:
            # Minimale Initialisierung
            import win32evtlogutil
            
            # Event Log Meldung
            win32evtlogutil.ReportEvent(
                self._svc_name_,
                1,  # Information
                eventCategory=0,
                eventType=win32evtlog.EVENTLOG_INFORMATION_TYPE,
                strings=["belegpilot Service wird gestartet..."]
            )
            
            # Sofort als RUNNING melden - KRITISCH für Windows!
            self.ReportServiceStatus(win32service.SERVICE_RUNNING)
            
            # Starte Initialisierung asynchron
            init_thread = threading.Thread(target=self._async_init)
            init_thread.daemon = True
            init_thread.start()
            
            # Hauptschleife - warte auf Stop
            win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
            
        except Exception as e:
            # Kritischer Fehler - logge in Event Log
            try:
                import win32evtlogutil
                win32evtlogutil.ReportEvent(
                    self._svc_name_,
                    3,  # Error
                    eventCategory=0,
                    eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                    strings=[f"Service-Start fehlgeschlagen: {str(e)}"]
                )
            except:
                pass
            self.ReportServiceStatus(win32service.SERVICE_STOPPED)

    def _async_init(self):
        """Asynchrone Initialisierung - läuft NACH Service-Start-Bestätigung"""
        try:
            # Logging Setup
            self._setup_service_logging()
            
            if self.logger:
                self.logger.info("Beginne asynchrone Initialisierung...")
            
            self.running = True
            
            # Starte Hauptlogik
            self.main_thread = threading.Thread(target=self._run_main_logic)
            self.main_thread.daemon = True
            self.main_thread.start()
            
        except Exception as e:
            error_msg = f"Fehler bei Initialisierung: {str(e)}\n{traceback.format_exc()}"
            if self.logger:
                self.logger.error(error_msg)
            
            # Event Log
            try:
                import win32evtlogutil
                win32evtlogutil.ReportEvent(
                    self._svc_name_,
                    3,  # Error
                    eventCategory=0,
                    eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                    strings=[error_msg]
                )
            except:
                pass

    def _run_main_logic(self):
        """Führt die Hauptlogik aus"""
        try:
            if self.logger:
                self.logger.info("Starte Hauptlogik...")
            
            # Importiere Module
            try:
                from core.hotfolder_manager import HotfolderManager
                from core.license_manager import get_license_manager
                
                if self.logger:
                    self.logger.info("Module erfolgreich importiert")
                    
            except ImportError as e:
                error_msg = f"Import-Fehler: {str(e)}"
                if self.logger:
                    self.logger.error(error_msg, exc_info=True)
                
                # Event Log
                try:
                    import win32evtlogutil
                    win32evtlogutil.ReportEvent(
                        self._svc_name_,
                        3,  # Error
                        eventCategory=0,
                        eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                        strings=[error_msg]
                    )
                except:
                    pass
                return
            
            # Lizenzprüfung
            try:
                if self.logger:
                    self.logger.info("Prüfe Lizenz...")
                    
                license_manager = get_license_manager()
                valid, license_info, message = license_manager.validate_license()
                
                if not valid:
                    if self.logger:
                        self.logger.warning(f"Keine gültige Lizenz: {message}")
                else:
                    if self.logger:
                        self.logger.info(f"Lizenz OK: {license_info.get('type', 'unbekannt')}")
            except Exception as e:
                if self.logger:
                    self.logger.error(f"Lizenzprüfung fehlgeschlagen: {e}", exc_info=True)
            
            # Arbeitsverzeichnis
            service_path = os.path.dirname(os.path.abspath(__file__))
            os.chdir(service_path)
            
            if self.logger:
                self.logger.info(f"Arbeitsverzeichnis: {os.getcwd()}")
            
            # Starte Hotfolder-Manager
            try:
                if self.logger:
                    self.logger.info("Initialisiere Hotfolder-Manager...")
                
                self.manager = HotfolderManager()
                self.manager.start()
                
                if self.logger:
                    self.logger.info("Hotfolder-Manager erfolgreich gestartet")
                    
                    # Info über Hotfolder
                    total = len(self.manager.config_manager.hotfolders)
                    active = len([h for h in self.manager.config_manager.hotfolders if h.enabled])
                    self.logger.info(f"Hotfolder: {active} von {total} aktiv")
                    
                    for hf in self.manager.config_manager.hotfolders:
                        if hf.enabled:
                            self.logger.info(f"  - {hf.name}: {hf.input_path}")
                
                # Success Event Log
                try:
                    import win32evtlogutil
                    win32evtlogutil.ReportEvent(
                        self._svc_name_,
                        1,  # Information
                        eventCategory=0,
                        eventType=win32evtlog.EVENTLOG_INFORMATION_TYPE,
                        strings=[f"belegpilot Service erfolgreich gestartet. {active} Hotfolder aktiv."]
                    )
                except:
                    pass
                    
            except Exception as e:
                error_msg = f"Manager-Start fehlgeschlagen: {str(e)}"
                if self.logger:
                    self.logger.error(error_msg, exc_info=True)
                
                # Event Log
                try:
                    import win32evtlogutil
                    win32evtlogutil.ReportEvent(
                        self._svc_name_,
                        3,  # Error
                        eventCategory=0,
                        eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                        strings=[error_msg]
                    )
                except:
                    pass
                return
            
            # Hauptschleife
            last_log = 0
            
            while self.running:
                try:
                    time.sleep(5)
                    
                    # Periodisches Status-Log
                    current = int(time.time())
                    if current - last_log >= 300:  # alle 5 Min
                        if self.manager and self.logger:
                            active = len([h for h in self.manager.config_manager.hotfolders if h.enabled])
                            self.logger.info(f"Service läuft - {active} Hotfolder aktiv")
                        last_log = current
                        
                except Exception as e:
                    if self.logger:
                        self.logger.error(f"Fehler in Hauptschleife: {e}")
                    time.sleep(10)
            
        except Exception as e:
            error_msg = f"Kritischer Fehler: {str(e)}\n{traceback.format_exc()}"
            if self.logger:
                self.logger.error(error_msg)
            
            # Event Log
            try:
                import win32evtlogutil
                win32evtlogutil.ReportEvent(
                    self._svc_name_,
                    3,  # Error
                    eventCategory=0,
                    eventType=win32evtlog.EVENTLOG_ERROR_TYPE,
                    strings=[error_msg]
                )
            except:
                pass
                
        finally:
            if self.manager:
                try:
                    self.manager.stop()
                except:
                    pass
            
            if self.logger:
                self.logger.info("Hauptlogik beendet")


# DebugService und install_service bleiben gleich wie vorher...
class DebugService:
    """Einfache Klasse für Debug-Modus ohne Service-Framework"""
    
    def __init__(self):
        self.logger = None
        self.manager = None
        self.running = False
        
    def setup_logging(self):
        """Logging für Debug-Modus einrichten"""
        # Erstelle Log-Verzeichnis
        service_log_dir = Path(os.environ.get('PROGRAMDATA', 'C:\\ProgramData')) / 'belegpilot' / 'logs'
        service_log_dir.mkdir(parents=True, exist_ok=True)
        
        # Logger einrichten
        self.logger = logging.getLogger('BelegpilotServiceDebug')
        self.logger.setLevel(logging.DEBUG)
        
        # Konsolen-Handler
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.DEBUG)
        console_formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(console_formatter)
        self.logger.addHandler(console_handler)
        
        # File Handler
        log_file = service_log_dir / f'service_debug_{time.strftime("%Y%m%d_%H%M%S")}.log'
        file_handler = logging.FileHandler(log_file, encoding='utf-8')
        file_handler.setLevel(logging.DEBUG)
        file_formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        file_handler.setFormatter(file_formatter)
        self.logger.addHandler(file_handler)
        
        self.logger.info("="*60)
        self.logger.info("Debug-Modus gestartet")
        self.logger.info(f"Log-Datei: {log_file}")
        self.logger.info("="*60)
        
    def run(self):
        """Hauptlogik im Debug-Modus ausführen"""
        try:
            # Importiere Module
            from core.hotfolder_manager import HotfolderManager
            from core.license_manager import get_license_manager
            
            # Lizenzprüfung
            self.logger.info("Prüfe Lizenz...")
            license_manager = get_license_manager()
            valid, license_info, message = license_manager.validate_license()
            
            if not valid:
                self.logger.warning(f"Keine gültige Lizenz: {message}")
            else:
                self.logger.info(f"Lizenz OK: {license_info.get('type', 'unbekannt')}")
            
            # Starte Manager
            self.logger.info("Initialisiere Hotfolder-Manager...")
            self.manager = HotfolderManager()
            self.manager.start()
            
            self.logger.info("Hotfolder-Manager gestartet")
            
            # Zeige Hotfolder
            total_count = len(self.manager.config_manager.hotfolders)
            active_count = len([h for h in self.manager.config_manager.hotfolders if h.enabled])
            self.logger.info(f"Hotfolder: {active_count} von {total_count} aktiv")
            
            for hotfolder in self.manager.config_manager.hotfolders:
                if hotfolder.enabled:
                    self.logger.info(f"  - {hotfolder.name}: {hotfolder.input_path}")
            
            self.running = True
            self.logger.info("\nService läuft im Debug-Modus. Drücken Sie Strg+C zum Beenden.")
            
            # Hauptschleife
            while self.running:
                time.sleep(5)
                
        except KeyboardInterrupt:
            self.logger.info("\nBeenden durch Benutzer...")
        except Exception as e:
            self.logger.error(f"Fehler: {e}", exc_info=True)
        finally:
            if self.manager:
                self.logger.info("Stoppe Manager...")
                self.manager.stop()
            self.logger.info("Debug-Modus beendet")


def install_service():
    """Installiert den Windows Service"""
    # Füge fehlenden Import hinzu
    try:
        import win32evtlog
    except ImportError:
        print("FEHLER: win32evtlog konnte nicht importiert werden!")
        print("Bitte installieren Sie pywin32 neu.")
        sys.exit(1)
    
    if len(sys.argv) == 1:
        print("="*60)
        print("belegpilot - Windows Service")
        print("="*60)
        print("\nVerwendung:")
        print("  belegpilot_service.exe install    - Service installieren")
        print("  belegpilot_service.exe start      - Service starten")
        print("  belegpilot_service.exe stop       - Service stoppen")
        print("  belegpilot_service.exe remove     - Service entfernen")
        print("  belegpilot_service.exe status     - Service-Status anzeigen")
        print("  belegpilot_service.exe debug      - Service im Debug-Modus starten")
        print("\nHinweis: Administrator-Rechte erforderlich!")
        print("="*60)
        sys.exit(0)

    elif sys.argv[1].lower() == 'status':
        try:
            status = win32serviceutil.QueryServiceStatus(BelegpilotService._svc_name_)
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
            print(f"\nbelegpilot Service-Status: {state}")
            
            # Zeige Log-Verzeichnis
            log_dir = Path(os.environ.get('PROGRAMDATA', 'C:\\ProgramData')) / 'belegpilot' / 'logs'
            print(f"\nLog-Dateien unter: {log_dir}")
            
            # Prüfe Windows Event Log
            print("\nPrüfen Sie auch die Windows-Ereignisanzeige:")
            print("  Ereignisanzeige → Windows-Protokolle → Anwendung")
            print("  Filter: Quelle = 'belegpilot'")
            
            if log_dir.exists():
                log_files = sorted(log_dir.glob("service_*.log"), reverse=True)
                if log_files:
                    latest_log = log_files[0]
                    print(f"\nNeueste Log-Datei: {latest_log.name}")
                    
                    # Zeige letzte Zeilen
                    try:
                        with open(latest_log, 'r', encoding='utf-8') as f:
                            lines = f.readlines()
                            if lines:
                                print("\nLetzte Log-Einträge:")
                                print("-" * 60)
                                for line in lines[-10:]:
                                    print(line.rstrip())
                    except:
                        pass
                        
        except Exception as e:
            print(f"Fehler: {e}")
            print("\nService ist möglicherweise nicht installiert.")
        sys.exit(0)

    elif sys.argv[1].lower() == 'debug':
        print("Starte Service im Debug-Modus...")
        debug_service = DebugService()
        debug_service.setup_logging()
        debug_service.run()
        sys.exit(0)

    # Service-Kommandos
    try:
        win32serviceutil.HandleCommandLine(BelegpilotService)
    except Exception as e:
        print(f"\nFehler: {e}")
        print("\nStellen Sie sicher, dass Sie Administrator-Rechte haben!")
        sys.exit(1)


if __name__ == '__main__':
    # Globaler Import für win32evtlog
    try:
        import win32evtlog
    except ImportError:
        print("FEHLER: win32evtlog konnte nicht importiert werden!")
        sys.exit(1)
    
    install_service()
