"""
Logging-Konfiguration für Hotfolder PDF Processor
"""
import logging
import logging.handlers
from pathlib import Path
import os
import sys
from datetime import datetime, timedelta

class HotfolderFileHandler(logging.handlers.TimedRotatingFileHandler):
    """Custom Handler für tägliche Logs mit speziellem Dateinamen-Format"""
    
    def __init__(self, log_dir, **kwargs):
        self.log_dir = Path(log_dir)
        self.log_dir.mkdir(exist_ok=True)
        
        # Basis-Dateiname für heute
        filename = self.log_dir / f"hotfolder_{datetime.now().strftime('%Y-%m-%d')}.log"
        
        super().__init__(
            filename=filename,
            when='midnight',
            interval=1,
            backupCount=0,  # Wir machen eigenes Cleanup
            encoding='utf-8',
            **kwargs
        )
        
        # Cleanup alter Logs beim Start
        self.cleanup_old_logs(30)
    
    def doRollover(self):
        """Überschreibe Rollover für custom Dateinamen"""
        super().doRollover()
        # Neuer Dateiname für den neuen Tag
        self.baseFilename = str(self.log_dir / f"hotfolder_{datetime.now().strftime('%Y-%m-%d')}.log")
        # Cleanup nach Rollover
        self.cleanup_old_logs(30)
    
    def cleanup_old_logs(self, days_to_keep=30):
        """Löscht alte Log-Dateien"""
        try:
            cutoff_date = datetime.now() - timedelta(days=days_to_keep)
            
            for log_file in self.log_dir.glob("hotfolder_*.log"):
                # Extrahiere Datum aus Dateiname
                try:
                    date_str = log_file.stem.replace("hotfolder_", "")
                    file_date = datetime.strptime(date_str, "%Y-%m-%d")
                    
                    if file_date < cutoff_date:
                        log_file.unlink()
                        logging.info(f"Alte Log-Datei gelöscht: {log_file.name}")
                        
                except ValueError:
                    # Ignoriere Dateien mit ungültigem Datumformat
                    pass
                    
        except Exception as e:
            logging.error(f"Fehler beim Aufräumen alter Logs: {e}")


def setup_logging(log_dir=None):
    """
    Initialisiert das Logging-System
    
    Args:
        log_dir: Verzeichnis für Log-Dateien (None = Script-Verzeichnis/logs)
    
    Returns:
        Logger-Instanz
    """
    if log_dir is None:
        log_dir = Path(os.path.dirname(os.path.abspath(sys.argv[0]))) / "logs"
    
    log_dir = Path(log_dir)
    
    # Root Logger konfigurieren
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)
    
    # Entferne existierende Handler
    root_logger.handlers.clear()
    
    # Formatter mit Log-Level-Anzeige
    formatter = logging.Formatter(
        '[%(asctime)s] %(levelname)s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Custom File Handler mit täglicher Rotation
    file_handler = HotfolderFileHandler(log_dir)
    file_handler.setLevel(logging.INFO)
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)
    
    # Console Handler mit Log-Level
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter('[%(levelname)s] %(message)s'))
    root_logger.addHandler(console_handler)
    
    # Log startup message
    logger = logging.getLogger(__name__)
    logger.info(f"Logging aktiviert. Log-Dateien werden gespeichert in: {log_dir}")
    logger.info(f"Aktuelle Log-Datei: hotfolder_{datetime.now().strftime('%Y-%m-%d')}.log")
    
    return logger


# Globale Funktionen für Kompatibilität
def initialize_logging():
    """Kompatibilitäts-Funktion für alte API"""
    return setup_logging()


def cleanup_logging():
    """Stub-Funktion für Kompatibilität - macht nichts mehr"""
    logging.shutdown()
