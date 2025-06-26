"""
Counter-Manager für persistente Auto-Inkrement Funktionalität
"""
import json
import os
import threading
from typing import Dict, Any
from pathlib import Path
import msvcrt
import errno
import time
import logging

# Logger für dieses Modul
logger = logging.getLogger(__name__)


class CounterManager:
    """Verwaltet persistente Counter für Auto-Inkrement Funktionen"""
    
    def __init__(self, counter_file: str = "counters.json"):
        self.counter_file = counter_file
        self.counters: Dict[str, int] = {}
        self._lock = threading.Lock()
        self.load_counters()
    
    def load_counters(self) -> None:
        """Lädt die Counter aus der Datei"""
        try:
            if os.path.exists(self.counter_file):
                with open(self.counter_file, 'r', encoding='utf-8') as f:
                    content = f.read().strip()
                    if content:
                        self.counters = json.loads(content)
                        logger.debug("Counter geladen", extra={
                            'counter_count': len(self.counters),
                            'file_path': self.counter_file
                        })
                    else:
                        self.counters = {}
                        logger.debug("Counter-Datei ist leer, initialisiere mit leeren Countern")
            else:
                self.counters = {}
                logger.debug("Counter-Datei existiert nicht, initialisiere mit leeren Countern")
        except (json.JSONDecodeError, FileNotFoundError) as e:
            logger.error(f"Fehler beim Laden der Counter-Datei: {e}")
            self.counters = {}
        except Exception as e:
            logger.exception("Unerwarteter Fehler beim Laden der Counter")
            self.counters = {}
    
    def save_counters(self) -> None:
        """Speichert die Counter in die Datei"""
        try:
            # Erstelle Backup der aktuellen Datei
            if os.path.exists(self.counter_file):
                backup_file = self.counter_file + ".backup"
                if os.path.exists(backup_file):
                    os.remove(backup_file)
                os.rename(self.counter_file, backup_file)
                logger.debug("Backup der Counter-Datei erstellt")
            
            # Speichere neue Counter
            with open(self.counter_file, 'w', encoding='utf-8') as f:
                json.dump(self.counters, f, indent=2, ensure_ascii=False)
            
            # Strukturiertes Logging
            logger.debug("Counter gespeichert", extra={
                'counter_count': len(self.counters),
                'file_path': self.counter_file
            })
            
            # Entferne Backup wenn erfolgreich
            backup_file = self.counter_file + ".backup"
            if os.path.exists(backup_file):
                os.remove(backup_file)
                
        except Exception as e:
            logger.exception("Fehler beim Speichern der Counter")
            # Stelle Backup wieder her wenn vorhanden
            backup_file = self.counter_file + ".backup"
            if os.path.exists(backup_file):
                try:
                    if os.path.exists(self.counter_file):
                        os.remove(self.counter_file)
                    os.rename(backup_file, self.counter_file)
                    logger.info("Counter-Backup wiederhergestellt")
                except Exception as restore_error:
                    logger.exception("Fehler beim Wiederherstellen des Counter-Backups")
    
    def get_and_increment(self, counter_name: str, start_value: int = 1, step: int = 1) -> int:
        """
        Gibt den aktuellen Counter-Wert zurück und erhöht ihn
        
        Args:
            counter_name: Name des Counters
            start_value: Startwert falls Counter nicht existiert
            step: Schrittweite für Erhöhung
            
        Returns:
            Der aktuelle Wert vor der Erhöhung
        """
        lock_file_path = self.counter_file + ".lock"
        
        # Prozess-übergreifendes Locking
        try:
            lock_file = open(lock_file_path, 'w')
            
            # Plattform-spezifisches File-Locking
            if os.name == 'nt':  # Windows
                while True:
                    try:
                        msvcrt.locking(lock_file.fileno(), msvcrt.LK_NBLCK, 1)
                        break
                    except IOError:
                        time.sleep(0.01)  # Kurz warten und erneut versuchen
            
            with self._lock:  # Thread-Lock für Thread-Sicherheit
                # Lade Counter neu vom Disk (für Prozess-Sicherheit)
                self.load_counters()
                
                # Initialisiere Counter falls nicht vorhanden
                if counter_name not in self.counters:
                    self.counters[counter_name] = start_value
                    logger.info(f"Neuer Counter '{counter_name}' mit Startwert {start_value} erstellt")
                
                # Hole aktuellen Wert
                current_value = self.counters[counter_name]
                
                # Erhöhe Counter
                self.counters[counter_name] += step
                
                # Strukturiertes Logging
                logger.debug("Counter erhöht", extra={
                    'counter_name': counter_name,
                    'old_value': current_value,
                    'new_value': self.counters[counter_name],
                    'step': step
                })
                
                # Speichere sofort
                self.save_counters()
                
                # Gebe Lock frei
                if os.name == 'nt':
                    msvcrt.locking(lock_file.fileno(), msvcrt.LK_UNLCK, 1)
                # fcntl gibt Lock automatisch beim Schließen frei
                
                lock_file.close()
                
                # Lösche Lock-Datei
                try:
                    os.remove(lock_file_path)
                except:
                    pass
                    
                return current_value
                
        except Exception as e:
            logger.error(f"Fehler beim Counter-Zugriff: {e}")
            # Fallback auf alte Implementierung
            return start_value
    
    def set_counter(self, counter_name: str, value: int) -> None:
        """
        Setzt einen Counter auf einen bestimmten Wert
        
        Args:
            counter_name: Name des Counters
            value: Neuer Wert
        """
        with self._lock:
            old_value = self.counters.get(counter_name, None)
            self.counters[counter_name] = value
            
            # Strukturiertes Logging
            if old_value is not None:
                logger.info("Counter manuell gesetzt", extra={
                    'counter_name': counter_name,
                    'old_value': old_value,
                    'new_value': value
                })
            else:
                logger.info("Counter initialisiert", extra={
                    'counter_name': counter_name,
                    'new_value': value
                })
            
            self.save_counters()
    
    def get_counter(self, counter_name: str, default: int = 0) -> int:
        """
        Gibt den aktuellen Counter-Wert zurück ohne ihn zu ändern
        
        Args:
            counter_name: Name des Counters
            default: Standardwert falls Counter nicht existiert
            
        Returns:
            Der aktuelle Counter-Wert
        """
        with self._lock:
            value = self.counters.get(counter_name, default)
            logger.debug(f"Counter '{counter_name}' abgefragt: {value}")
            return value
    
    def reset_counter(self, counter_name: str, new_value: int = 1) -> None:
        """
        Setzt einen Counter zurück
        
        Args:
            counter_name: Name des Counters
            new_value: Neuer Startwert
        """
        with self._lock:
            old_value = self.counters.get(counter_name, None)
            self.counters[counter_name] = new_value
            
            # Strukturiertes Logging
            if old_value is not None:
                logger.info("Counter zurückgesetzt", extra={
                    'counter_name': counter_name,
                    'old_value': old_value,
                    'new_value': new_value
                })
            else:
                logger.info("Counter initialisiert", extra={
                    'counter_name': counter_name,
                    'new_value': new_value
                })
            
            self.save_counters()
    
    def delete_counter(self, counter_name: str) -> bool:
        """
        Löscht einen Counter
        
        Args:
            counter_name: Name des Counters
            
        Returns:
            True wenn Counter existierte und gelöscht wurde
        """
        with self._lock:
            if counter_name in self.counters:
                value = self.counters[counter_name]
                del self.counters[counter_name]
                
                # Strukturiertes Logging
                logger.info("Counter gelöscht", extra={
                    'counter_name': counter_name,
                    'final_value': value
                })
                
                self.save_counters()
                return True
            else:
                logger.debug(f"Counter '{counter_name}' existiert nicht, kann nicht gelöscht werden")
                return False
    
    def list_counters(self) -> Dict[str, int]:
        """
        Gibt alle Counter zurück
        
        Returns:
            Dictionary mit allen Counter-Namen und ihren Werten
        """
        with self._lock:
            logger.debug(f"Alle Counter abgefragt: {len(self.counters)} Counter")
            return self.counters.copy()
    
    def clear_all_counters(self) -> None:
        """Löscht alle Counter"""
        with self._lock:
            count = len(self.counters)
            
            # Strukturiertes Logging vor dem Löschen
            logger.warning("Alle Counter werden gelöscht", extra={
                'counter_count': count,
                'counters': list(self.counters.keys())
            })
            
            self.counters.clear()
            self.save_counters()


# Globale Instanz für die Anwendung
_global_counter_manager = None

def get_counter_manager() -> CounterManager:
    """Gibt die globale CounterManager-Instanz zurück"""
    global _global_counter_manager
    if _global_counter_manager is None:
        _global_counter_manager = CounterManager()
        logger.debug("Globale CounterManager-Instanz erstellt")
    return _global_counter_manager
