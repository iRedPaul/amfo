"""
Log-Handler für tägliche Log-Dateien
Leitet alle print-Ausgaben in Log-Dateien um
"""
import sys
import os
from datetime import datetime
from pathlib import Path
import threading


class DailyLogHandler:
    """Handler für tägliche Log-Dateien mit automatischer print-Umleitung"""
    
    def __init__(self, log_dir=None):
        """
        Initialisiert den Log-Handler
        
        Args:
            log_dir: Verzeichnis für Log-Dateien (None = Script-Verzeichnis)
        """
        # Bestimme Log-Verzeichnis
        if log_dir is None:
            # Verwende das Verzeichnis, in dem das Hauptskript liegt
            self.log_dir = Path(os.path.dirname(os.path.abspath(sys.argv[0])))
        else:
            self.log_dir = Path(log_dir)
        
        # Erstelle logs Unterordner
        self.log_dir = self.log_dir / "logs"
        self.log_dir.mkdir(exist_ok=True)
        
        # Speichere originale stdout
        self.original_stdout = sys.stdout
        
        # Thread-Lock für Thread-Sicherheit
        self.lock = threading.Lock()
        
        # Aktuelle Log-Datei
        self.current_date = None
        self.log_file = None
        
        # Buffer für unvollständige Zeilen
        self.line_buffer = ""
        
        # Starte Umleitung
        sys.stdout = self
    
    def _get_log_filename(self):
        """Generiert den Dateinamen für die aktuelle Log-Datei"""
        today = datetime.now().strftime("%Y-%m-%d")
        return self.log_dir / f"hotfolder_{today}.log"
    
    def _ensure_log_file(self):
        """Stellt sicher, dass die richtige Log-Datei geöffnet ist"""
        today = datetime.now().date()
        
        # Prüfe ob ein neuer Tag begonnen hat
        if self.current_date != today:
            # Schließe alte Log-Datei
            if self.log_file and not self.log_file.closed:
                self.log_file.close()
            
            # Öffne neue Log-Datei
            self.current_date = today
            log_path = self._get_log_filename()
            self.log_file = open(log_path, 'a', encoding='utf-8', buffering=1)
    
    def write(self, text):
        """
        Schreibt Text sowohl in die Konsole als auch in die Log-Datei
        
        Args:
            text: Zu schreibender Text
        """
        with self.lock:
            # Schreibe in Original-Stdout (Konsole)
            self.original_stdout.write(text)
            self.original_stdout.flush()
            
            # Verarbeite Text für Log-Datei
            if text:
                try:
                    self._ensure_log_file()
                    
                    # Füge Text zum Buffer hinzu
                    self.line_buffer += text
                    
                    # Verarbeite vollständige Zeilen
                    while '\n' in self.line_buffer:
                        line, self.line_buffer = self.line_buffer.split('\n', 1)
                        
                        # Ignoriere leere Zeilen
                        if line.strip():
                            timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            self.log_file.write(f"[{timestamp}] {line}\n")
                            self.log_file.flush()
                    
                except Exception as e:
                    # Bei Fehler nur in Konsole ausgeben
                    self.original_stdout.write(f"\n[LOG ERROR] Konnte nicht in Log-Datei schreiben: {e}\n")
    
    def flush(self):
        """Flush-Methode für Kompatibilität"""
        self.original_stdout.flush()
        
        # Schreibe verbleibenden Buffer-Inhalt
        with self.lock:
            if self.line_buffer.strip():
                try:
                    self._ensure_log_file()
                    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    self.log_file.write(f"[{timestamp}] {self.line_buffer}\n")
                    self.log_file.flush()
                    self.line_buffer = ""
                except Exception:
                    pass
            
            if self.log_file and not self.log_file.closed:
                self.log_file.flush()
    
    def cleanup(self):
        """Räumt auf und stellt originale stdout wieder her"""
        # Schreibe verbleibenden Buffer-Inhalt
        self.flush()
        
        sys.stdout = self.original_stdout
        if self.log_file and not self.log_file.closed:
            self.log_file.close()
    
    def cleanup_old_logs(self, days_to_keep=30):
        """
        Löscht alte Log-Dateien
        
        Args:
            days_to_keep: Anzahl Tage, die Log-Dateien behalten werden sollen
        """
        try:
            cutoff_date = datetime.now().date()
            
            for log_file in self.log_dir.glob("hotfolder_*.log"):
                # Extrahiere Datum aus Dateiname
                try:
                    date_str = log_file.stem.replace("hotfolder_", "")
                    file_date = datetime.strptime(date_str, "%Y-%m-%d").date()
                    
                    # Berechne Alter in Tagen
                    age_days = (cutoff_date - file_date).days
                    
                    if age_days > days_to_keep:
                        log_file.unlink()
                        print(f"Alte Log-Datei gelöscht: {log_file.name}")
                        
                except ValueError:
                    # Ignoriere Dateien mit ungültigem Datumformat
                    pass
                    
        except Exception as e:
            print(f"Fehler beim Aufräumen alter Logs: {e}")


# Globale Instanz
_log_handler = None

def initialize_logging():
    """Initialisiert das Logging-System"""
    global _log_handler
    if _log_handler is None:
        _log_handler = DailyLogHandler()
        print(f"Logging aktiviert. Log-Dateien werden gespeichert in: {_log_handler.log_dir}")
        print(f"Aktuelle Log-Datei: {_log_handler._get_log_filename().name}")
        
        # Räume alte Logs auf (älter als 30 Tage)
        _log_handler.cleanup_old_logs(30)
    return _log_handler

def cleanup_logging():
    """Beendet das Logging-System"""
    global _log_handler
    if _log_handler:
        _log_handler.cleanup()
        _log_handler = None