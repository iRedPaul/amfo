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
        """Löscht alte Log-Dateien mit robuster Fehlerbehandlung"""
        try:
            cutoff_date = datetime.now() - timedelta(days=days_to_keep)
            deleted_count = 0
            error_count = 0
            
            # Korrigiertes glob-Pattern
            for log_file in self.log_dir.glob("hotfolder_*.log"):
                # Extrahiere Datum aus Dateiname
                try:
                    date_str = log_file.stem.replace("hotfolder_", "")
                    file_date = datetime.strptime(date_str, "%Y-%m-%d")
                    
                    if file_date < cutoff_date:
                        try:
                            log_file.unlink()
                            deleted_count += 1
                            logging.info(f"Alte Log-Datei gelöscht: {log_file.name}")
                        except PermissionError:
                            error_count += 1
                            logging.warning(f"Keine Berechtigung zum Löschen von: {log_file.name}")
                        except Exception as e:
                            error_count += 1
                            logging.error(f"Fehler beim Löschen von {log_file.name}: {e}")
                            
                except ValueError:
                    # Ignoriere Dateien mit ungültigem Datumformat
                    logging.debug(f"Überspringe Datei mit ungültigem Datumformat: {log_file.name}")
                    
            # Log zusammenfassende Information
            if deleted_count > 0:
                logging.info(f"Log-Cleanup abgeschlossen: {deleted_count} Dateien gelöscht")
            if error_count > 0:
                logging.warning(f"Log-Cleanup: {error_count} Dateien konnten nicht gelöscht werden")
                    
        except Exception as e:
            logging.error(f"Kritischer Fehler beim Aufräumen alter Logs: {e}", exc_info=True)


def setup_logging(log_dir=None, log_level=logging.INFO):
    """
    Initialisiert das Logging-System
    
    Args:
        log_dir: Verzeichnis für Log-Dateien (None = Script-Verzeichnis/logs)
        log_level: Standard-Log-Level (default: INFO)
    
    Returns:
        Logger-Instanz
    """
    if log_dir is None:
        log_dir = Path(os.path.dirname(os.path.abspath(sys.argv[0]))) / "logs"
    
    log_dir = Path(log_dir)
    
    # Root Logger konfigurieren
    root_logger = logging.getLogger()
    root_logger.setLevel(logging.DEBUG)  # Root auf DEBUG für Flexibilität
    
    # Entferne existierende Handler
    root_logger.handlers.clear()
    
    # Formatter mit Log-Level-Anzeige
    formatter = logging.Formatter(
        '[%(asctime)s] %(levelname)-8s %(name)-25s - %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    
    # Custom File Handler mit täglicher Rotation
    file_handler = HotfolderFileHandler(log_dir)
    file_handler.setLevel(log_level)
    file_handler.setFormatter(formatter)
    root_logger.addHandler(file_handler)
    
    # Console Handler mit Log-Level
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_formatter = logging.Formatter('[%(levelname)-8s] %(message)s')
    console_handler.setFormatter(console_formatter)
    root_logger.addHandler(console_handler)
    
    # Spezifische Log-Level für verschiedene Module
    logging.getLogger('gui').setLevel(logging.INFO)  # GUI nur INFO und höher
    logging.getLogger('core.file_watcher').setLevel(logging.DEBUG)  # File Watcher detailliert
    logging.getLogger('core.export_processor').setLevel(logging.INFO)
    logging.getLogger('core.pdf_processor').setLevel(logging.INFO)
    logging.getLogger('core.ocr_processor').setLevel(logging.INFO)
    logging.getLogger('core.xml_field_processor').setLevel(logging.INFO)
    logging.getLogger('core.function_parser').setLevel(logging.WARNING)  # Weniger verbose
    logging.getLogger('core.hotfolder_manager').setLevel(logging.INFO)
    logging.getLogger('PIL').setLevel(logging.WARNING)  # Externe Bibliothek weniger verbose
    logging.getLogger('pytesseract').setLevel(logging.WARNING)
    
    # Log startup message
    logger = logging.getLogger(__name__)
    logger.info("="*70)
    logger.info(f"Logging aktiviert. Log-Dateien werden gespeichert in: {log_dir}")
    logger.info(f"Aktuelle Log-Datei: hotfolder_{datetime.now().strftime('%Y-%m-%d')}.log")
    logger.info(f"Standard Log-Level: {logging.getLevelName(log_level)}")
    logger.info("="*70)
    
    return logger


# Globale Funktionen für Kompatibilität
def initialize_logging(log_dir=None, log_level=logging.INFO):
    """Kompatibilitäts-Funktion für alte API"""
    return setup_logging(log_dir, log_level)


def cleanup_logging():
    """Stub-Funktion für Kompatibilität - macht nichts mehr"""
    logging.shutdown()
