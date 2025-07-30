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

# Füge das Hauptverzeichnis zum Python-Pfad hinzu
if getattr(sys, 'frozen', False):
    application_path = os.path.dirname(sys.executable)
else:
    application_path = os.path.dirname(os.path.abspath(__file__))

os.chdir(application_path)
sys.path.insert(0, application_path)

class BelegpilotService(win32serviceutil.ServiceFramework):
    _svc_name_ = "BelegpilotService"
    _svc_display_name_ = "belegpilot Service"
    _svc_description_ = "Überwacht Ordner auf neue Dokumente und verarbeitet sie automatisch."

    def __init__(self, args):
        win32serviceutil.ServiceFramework.__init__(self, args)
        self.hWaitStop = win32event.CreateEvent(None, 0, 0, None)
        socket.setdefaulttimeout(60)
        self.is_running = threading.Event()
        self.manager = None
        self.worker_thread = None

    def SvcStop(self):
        self.ReportServiceStatus(win32service.SERVICE_STOP_PENDING)
        self.is_running.set()
        win32event.SetEvent(self.hWaitStop)

    def SvcDoRun(self):
        servicemanager.LogMsg(servicemanager.EVENTLOG_INFORMATION_TYPE,
                              servicemanager.PYS_SERVICE_STARTED,
                              (self._svc_name_, ''))
        
        # Melde den Dienst als gestartet, bevor schwere Initialisierungen beginnen
        self.ReportServiceStatus(win32service.SERVICE_RUNNING)
        
        # Starte die eigentliche Arbeit in einem separaten Thread
        self.worker_thread = threading.Thread(target=self.main)
        self.worker_thread.start()
        
        # Warte auf Stop-Signal
        win32event.WaitForSingleObject(self.hWaitStop, win32event.INFINITE)
        
        # Stoppe den Worker-Thread
        if self.worker_thread and self.worker_thread.is_alive():
            self.worker_thread.join(timeout=10)

    def main(self):
        """Haupt-Schleife des Dienstes."""
        try:
            # Imports hier, um Startup-Zeit zu reduzieren
            from core.hotfolder_manager import HotfolderManager
            from core.logging_config import initialize_logging, cleanup_logging
            
            initialize_logging()
            logger = logging.getLogger(__name__)
            logger.info("belegpilot Dienst startet...")

            self.manager = HotfolderManager()
            self.manager.start()
            
            logger.info("HotfolderManager gestartet. Dienst läuft.")

            while not self.is_running.is_set():
                self.is_running.wait(1)
            
            logger.info("belegpilot Dienst wird gestoppt...")
            if self.manager:
                self.manager.stop()
            cleanup_logging()
            logger.info("belegpilot Dienst erfolgreich gestoppt.")
            
        except Exception as e:
            servicemanager.LogErrorMsg(f"Fehler im Dienst: {str(e)}")

if __name__ == '__main__':
    if len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(BelegpilotService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(BelegpilotService)