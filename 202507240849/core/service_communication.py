"""
Kommunikation zwischen GUI und Windows-Dienst
"""
import win32pipe
import win32file
import pywintypes
import threading
import logging
import json
import time

logger = logging.getLogger(__name__)

PIPE_NAME = r'\\.\pipe\belegpilot_service'
PIPE_BUFFER_SIZE = 512


class ServiceCommunicationServer:
    """Server-Seite für den Windows-Dienst"""
    
    def __init__(self, callback):
        self.callback = callback
        self._running = False
        self._thread = None
        
    def start(self):
        """Startet den Named Pipe Server"""
        if self._running:
            return
            
        self._running = True
        self._thread = threading.Thread(target=self._server_loop, daemon=True)
        self._thread.start()
        logger.info("Service-Kommunikation gestartet")
        
    def stop(self):
        """Stoppt den Named Pipe Server"""
        self._running = False
        if self._thread:
            self._thread.join(timeout=2)
        logger.info("Service-Kommunikation gestoppt")
        
    def _server_loop(self):
        """Hauptschleife des Named Pipe Servers"""
        while self._running:
            try:
                # Erstelle Named Pipe
                pipe = win32pipe.CreateNamedPipe(
                    PIPE_NAME,
                    win32pipe.PIPE_ACCESS_DUPLEX,
                    win32pipe.PIPE_TYPE_MESSAGE | win32pipe.PIPE_READMODE_MESSAGE | win32pipe.PIPE_WAIT,
                    1, PIPE_BUFFER_SIZE, PIPE_BUFFER_SIZE,
                    0,
                    None
                )
                
                # Warte auf Client-Verbindung
                win32pipe.ConnectNamedPipe(pipe, None)
                
                # Lese Nachricht
                result, data = win32file.ReadFile(pipe, PIPE_BUFFER_SIZE)
                if result == 0:  # Erfolg
                    message = data.decode('utf-8')
                    command = json.loads(message)
                    
                    logger.info(f"Kommando empfangen: {command}")
                    
                    # Verarbeite Kommando
                    response = self._process_command(command)
                    
                    # Sende Antwort
                    win32file.WriteFile(pipe, json.dumps(response).encode('utf-8'))
                
                # Schließe Pipe
                win32file.CloseHandle(pipe)
                
            except pywintypes.error as e:
                if e.args[0] == 232:  # Pipe wurde vom Client geschlossen
                    pass
                else:
                    logger.error(f"Pipe-Fehler: {e}")
                time.sleep(0.1)
            except Exception as e:
                logger.error(f"Fehler im Service-Kommunikations-Server: {e}")
                time.sleep(1)
    
    def _process_command(self, command):
        """Verarbeitet eingehende Kommandos"""
        cmd_type = command.get('type')
        
        if cmd_type == 'reload_config':
            if self.callback:
                self.callback()
            return {'status': 'success', 'message': 'Config-Reload ausgelöst'}
        
        elif cmd_type == 'ping':
            return {'status': 'success', 'message': 'pong'}
        
        else:
            return {'status': 'error', 'message': f'Unbekanntes Kommando: {cmd_type}'}


class ServiceCommunicationClient:
    """Client-Seite für die GUI"""
    
    @staticmethod
    def send_command(command, timeout=5):
        """Sendet ein Kommando an den Dienst"""
        try:
            # Verbinde zur Named Pipe
            pipe = win32file.CreateFile(
                PIPE_NAME,
                win32file.GENERIC_READ | win32file.GENERIC_WRITE,
                0,
                None,
                win32file.OPEN_EXISTING,
                0,
                None
            )
            
            # Sende Kommando
            message = json.dumps(command).encode('utf-8')
            win32file.WriteFile(pipe, message)
            
            # Empfange Antwort
            result, data = win32file.ReadFile(pipe, PIPE_BUFFER_SIZE)
            if result == 0:
                response = json.loads(data.decode('utf-8'))
                win32file.CloseHandle(pipe)
                return True, response
            
            win32file.CloseHandle(pipe)
            return False, "Keine Antwort vom Dienst"
            
        except pywintypes.error as e:
            if e.args[0] == 2:  # Datei nicht gefunden
                return False, "Dienst läuft nicht oder Named Pipe nicht verfügbar"
            else:
                return False, f"Pipe-Fehler: {e}"
        except Exception as e:
            return False, f"Fehler bei der Kommunikation: {e}"
    
    @staticmethod
    def reload_config():
        """Sendet Befehl zum Config-Reload an den Dienst"""
        return ServiceCommunicationClient.send_command({'type': 'reload_config'})
    
    @staticmethod
    def ping():
        """Prüft ob der Dienst erreichbar ist"""
        return ServiceCommunicationClient.send_command({'type': 'ping'}, timeout=2)