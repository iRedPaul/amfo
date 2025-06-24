"""
Logging-Konfiguration für Hotfolder PDF Processor
Speichern Sie diese Datei als: hotfolder_pdf_processor/core/logger.py
"""
import logging
import logging.handlers
import os
from datetime import datetime
from pathlib import Path


class HotfolderLogger:
    """Zentrales Logging-System mit täglichen Log-Dateien"""
    
    _instance = None
    _logger = None
    
    def __new__(cls):
        if cls._instance is None:
            cls._instance = super().__new__(cls)
        return cls._instance
    
    def __init__(self):
        if self._logger is None:
            self._setup_logger()
    
    def _setup_logger(self):
        """Konfiguriert das Logging-System"""
        # Logger erstellen
        self._logger = logging.getLogger('HotfolderPDFProcessor')
        self._logger.setLevel(logging.DEBUG)
        
        # Verhindere doppelte Handler
        if self._logger.handlers:
            return
        
        # Log-Verzeichnis im Skript-Verzeichnis erstellen
        script_dir = Path(__file__).parent.parent
        log_dir = script_dir / 'logs'
        log_dir.mkdir(exist_ok=True)
        
        # Dateiname mit Datum
        log_filename = log_dir / f'hotfolder_{datetime.now().strftime("%Y-%m-%d")}.log'
        
        # File Handler mit täglicher Rotation
        file_handler = logging.handlers.TimedRotatingFileHandler(
            filename=log_filename,
            when='midnight',
            interval=1,
            backupCount=30,  # Behalte 30 Tage
            encoding='utf-8'
        )
        file_handler.suffix = "%Y-%m-%d"
        
        # Console Handler
        console_handler = logging.StreamHandler()
        
        # Formatter
        detailed_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(module)s.%(funcName)s:%(lineno)d - %(message)s',
            datefmt='%Y-%m-%d %H:%M:%S'
        )
        
        simple_formatter = logging.Formatter(
            '%(asctime)s - %(levelname)s - %(message)s',
            datefmt='%H:%M:%S'
        )
        
        # Formatter zuweisen
        file_handler.setFormatter(detailed_formatter)
        console_handler.setFormatter(simple_formatter)
        
        # Handler zum Logger hinzufügen
        self._logger.addHandler(file_handler)
        self._logger.addHandler(console_handler)
        
        # Log-Rotation bei Programmstart prüfen
        self._check_rotation()
        
        self._logger.info("="*60)
        self._logger.info("Hotfolder PDF Processor gestartet")
        self._logger.info(f"Log-Verzeichnis: {log_dir}")
        self._logger.info("="*60)
    
    def _check_rotation(self):
        """Prüft ob eine neue Log-Datei für heute erstellt werden muss"""
        for handler in self._logger.handlers:
            if isinstance(handler, logging.handlers.TimedRotatingFileHandler):
                # Trigger rotation wenn nötig
                if handler.shouldRollover(None):
                    handler.doRollover()
                break
    
    def get_logger(self, module_name: str = None) -> logging.Logger:
        """
        Gibt einen Logger für ein spezifisches Modul zurück
        
        Args:
            module_name: Name des Moduls (optional)
            
        Returns:
            Logger-Instanz
        """
        if module_name:
            return logging.getLogger(f'HotfolderPDFProcessor.{module_name}')
        return self._logger
    
    @classmethod
    def cleanup_old_logs(cls, days_to_keep: int = 30):
        """
        Löscht alte Log-Dateien
        
        Args:
            days_to_keep: Anzahl der Tage, die behalten werden sollen
        """
        script_dir = Path(__file__).parent.parent
        log_dir = script_dir / 'logs'
        
        if not log_dir.exists():
            return
        
        cutoff_date = datetime.now().timestamp() - (days_to_keep * 24 * 60 * 60)
        
        for log_file in log_dir.glob('hotfolder_*.log'):
            if log_file.stat().st_mtime < cutoff_date:
                try:
                    log_file.unlink()
                    if cls._instance and cls._instance._logger:
                        cls._instance._logger.info(f"Alte Log-Datei gelöscht: {log_file.name}")
                except Exception as e:
                    if cls._instance and cls._instance._logger:
                        cls._instance._logger.error(f"Fehler beim Löschen der Log-Datei {log_file.name}: {e}")


# Globale Logger-Instanz
def get_logger(module_name: str = None) -> logging.Logger:
    """
    Convenience-Funktion zum Abrufen eines Loggers
    
    Args:
        module_name: Name des Moduls (optional)
        
    Returns:
        Logger-Instanz
    """
    logger_instance = HotfolderLogger()
    return logger_instance.get_logger(module_name)


# Log-Level Funktionen für einfachen Zugriff
def log_debug(message: str, module_name: str = None):
    """Debug-Level Log"""
    get_logger(module_name).debug(message)

def log_info(message: str, module_name: str = None):
    """Info-Level Log"""
    get_logger(module_name).info(message)

def log_warning(message: str, module_name: str = None):
    """Warning-Level Log"""
    get_logger(module_name).warning(message)

def log_error(message: str, module_name: str = None, exc_info: bool = False):
    """Error-Level Log"""
    get_logger(module_name).error(message, exc_info=exc_info)

def log_critical(message: str, module_name: str = None, exc_info: bool = False):
    """Critical-Level Log"""
    get_logger(module_name).critical(message, exc_info=exc_info)