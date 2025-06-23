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
        self._setup_logging()

    def _setup_logging(self):
        """Konfiguriert das Logging"""
        log_dir = Path(os.environ.get('PROGRAMDATA', 'C:\\ProgramData')) / 'HotfolderPDFProcessor'
        log_dir.mkdir(exist_ok=True)

        log_file = log_dir / 'service.log'

        logging.basicConfig(
            level=logging.INFO,
            format='%(asctime)s - %(levelname)s - %(message)s',
            handlers=[
                logging.FileHandler(log_file),
                logging.StreamHandler()
            ]
        )

        self.logger = logging.getLogger(__name__)

    def SvcStop(self):
        """Stoppt den Service"""
        self.logger.info("Service wird gestoppt...")
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)

        if self.manager:
            self.manager.stop()

        win32event.SetEvent(self.hWaitStop)
        self.logger.info("Service gestoppt")

    def SvcDoRun(self):
        """Hauptschleife des Services"""
        try:
            servicemanager.LogMsg(
                servicemanager.EVENTLOG_INFORMATION_TYPE,
                servicemanager.PYS_SERVICE_STARTED,
                (self._svc_name_, '')
            )

            self.logger.info("Service gestartet")
            self.main()

        except Exception as e:
            self.logger.error(f"Fehler im Service: {e}", exc_info=True)
            servicemanager.LogErrorMsg(f"Service Fehler: {e}")

    def main(self):
        """Hauptlogik des Services"""
        try:
            service_path = os.path.dirname(os.path.abspath(__file__))
            os.chdir(service_path)
            self.logger.info(f"Arbeitsverzeichnis: {os.getcwd()}")

            self.manager = HotfolderManager()
            self.manager.start()

            self.logger.info("Hotfolder-Manager gestartet")
            self.logger.info(f"Anzahl Hotfolder: {len(self.manager.config_manager.hotfolders)}")

            while True:
                rc = win32event.WaitForSingleObject(self.hWaitStop, 5000)
                if rc == win32event.WAIT_OBJECT_0:
                    break

                if int(time.time()) % 60 == 0:
                    active_count = len([h for h in self.manager.config_manager.hotfolders if h.enabled])
                    self.logger.info(f"Service läuft - {active_count} aktive Hotfolder")

        except Exception as e:
            self.logger.error(f"Fehler in main(): {e}", exc_info=True)
            raise


def install_service():
    """Installiert den Windows Service"""
    if len(sys.argv) == 1:
        print("Hotfolder PDF Processor - Windows Service")
        print("\nVerwendung:")
        print("  python windows_service.py install    - Service installieren")
        print("  python windows_service.py start      - Service starten")
        print("  python windows_service.py stop       - Service stoppen")
        print("  python windows_service.py remove     - Service entfernen")
        print("  python windows_service.py status     - Service-Status anzeigen")
        print("\nHinweis: Administrator-Rechte erforderlich!")
        sys.exit(0)

    elif sys.argv[1].lower() == 'status':
        try:
            status = win32serviceutil.QueryServiceStatus(HotfolderService._svc_name_)
            state_map = {
                1: "STOPPED",
                2: "START_PENDING",
                3: "STOP_PENDING",
                4: "RUNNING",
                5: "CONTINUE_PENDING",
                6: "PAUSE_PENDING",
                7: "PAUSED",
            }
            print(f"Status: {state_map.get(status[1], 'UNKNOWN')}")
        except Exception as e:
            print(f"Fehler beim Abrufen des Service-Status: {e}")
        sys.exit(0)

    # Alle anderen Kommandos an pywin32 weiterleiten
    win32serviceutil.HandleCommandLine(HotfolderService)


if __name__ == '__main__':
    install_service()
