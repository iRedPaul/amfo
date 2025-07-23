# belegpilot_service.py
import win32serviceutil
import win32service
import win32event
import servicemanager
import socket
import sys
import os
import threading
import logging

# Füge das Hauptverzeichnis zum Python-Pfad hinzu, damit die Module gefunden werden
# Wichtig, wenn es als Dienst läuft
if getattr(sys, 'frozen', False):
    # Wenn als PyInstaller-Bundle ausgeführt
    application_path = os.path.dirname(sys.executable)
else:
    # Wenn als normales Python-Skript ausgeführt
    application_path = os.path.dirname(os.path.abspath(__file__))

os.chdir(application_path)
sys.path.insert(0, application_path)

from core.hotfolder_manager import HotfolderManager
from core.logging_config import initialize_logging, cleanup_logging

class BelegpilotService(win32serviceutil.ServiceFramework):
    """
    Windows-Dienst für den belegpilot Hotfolder Manager.
    """
    # Name des Dienstes
    _svc_name_ = "BelegpilotService"
    # Angezeigter Name im Dienste-Manager
    _svc_display_name_ = "belegpilot Hotfolder Service"
    # Beschreibung des Dienstes
    _svc_description_ = "Überwacht Ordner auf neue Dokumente und verarbeitet sie automatisch."

    def __init__(self, args):
        """
        Konstruktor des Dienstes.
        """
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        # Ein Event-Objekt, um den Haupt-Thread zu signalisieren, dass er sich beenden soll
        self.is_running = threading.Event()

    def SvcStop(self):
        """
        Wird aufgerufen, wenn der Dienst gestoppt wird.
        """
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        # Signalisiert dem Haupt-Thread, dass er aufhören soll
        self.is_running.set()
        # Wartet auf das Haupt-Event, um sicherzustellen, dass alles sauber beendet wurde
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        """
        Die Hauptlogik des Dienstes.
        """
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        
        self.main()

    def main(self):
        """
        Haupt-Schleife des Dienstes. Hier wird die Kernfunktionalität ausgeführt.
        """
        # Logging initialisieren, sobald der Dienst startet
        initialize_logging()
        logger = logging.getLogger(__name__)
        logger.info("belegpilot Dienst startet...")

        # HotfolderManager initialisieren und starten
        self.manager = HotfolderManager()
        self.manager.start()
        
        logger.info("HotfolderManager gestartet. Dienst läuft.")

        # Der Dienst bleibt in dieser Schleife aktiv, bis er gestoppt wird
        while not self.is_running.is_set():
            # Warte 1 Sekunde oder bis das Stop-Event gesetzt wird
            # Dies ist eine effiziente Art zu warten, ohne die CPU zu belasten
            self.is_running.wait(1)
        
        # Aufräumarbeiten, wenn der Dienst gestoppt wird
        logger.info("belegpilot Dienst wird gestoppt...")
        self.manager.stop()
        cleanup_logging()
        logger.info("belegpilot Dienst erfolgreich gestoppt.")

if __name__ == '__main__':
    # Diese Zeile ermöglicht die Installation, das Starten, Stoppen und Entfernen
    # des Dienstes über die Kommandozeile.
    win32serviceutil.HandleCommandLine(BelegpilotService)
