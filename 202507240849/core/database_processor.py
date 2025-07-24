import json
import os
import logging
from typing import Dict, List, Optional, Any
from dataclasses import dataclass
import pyodbc
from datetime import datetime

logger = logging.getLogger(__name__)


@dataclass
class DatabaseConfig:
    """Konfiguration für eine Datenbank-Verbindung"""
    name: str
    connection_string: str
    driver: str
    server: str
    database: str
    username: str
    password: str
    trusted_connection: bool = False
    
    def to_dict(self) -> dict:
        return {
            'name': self.name,
            'connection_string': self.connection_string,
            'driver': self.driver,
            'server': self.server,
            'database': self.database,
            'username': self.username,
            'password': self.password,
            'trusted_connection': self.trusted_connection
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'DatabaseConfig':
        return cls(**data)


class DatabaseProcessor:
    """Verarbeitet Datenbank-Operationen"""
    
    CONFIG_FILE = "config/databases.json"
    
    def __init__(self):
        self._initialized = False
        self.configs: Dict[str, DatabaseConfig] = {}
        self.connections: Dict[str, pyodbc.Connection] = {}
        self.load_configs()
        self._initialized = True
    
    def load_configs(self):
        """Lädt Datenbank-Konfigurationen aus Datei"""
        if not os.path.exists(self.CONFIG_FILE):
            logger.info(f"Keine Datenbank-Konfigurationsdatei gefunden: {self.CONFIG_FILE}")
            self.create_default_config()
            return
        
        try:
            with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            self.configs.clear()  # Alte Configs löschen
            for config_data in data.get("databases", []):
                config = DatabaseConfig.from_dict(config_data)
                self.configs[config.name] = config
            
        except Exception as e:
            logger.error(f"Fehler beim Laden der Datenbank-Konfigurationen: {e}")
            # Bei Fehler erstelle neue Datei
            self.create_default_config()
    
    def create_default_config(self):
        """Erstellt eine Standard-Konfigurationsdatei"""
        default_data = {
            "databases": []
        }
        
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(default_data, f, indent=2, ensure_ascii=False)
            logger.info(f"Neue Datenbank-Konfigurationsdatei erstellt: {self.CONFIG_FILE}")
        except Exception as e:
            logger.error(f"Fehler beim Erstellen der Datenbank-Konfigurationsdatei: {e}")
    
    def save_configs(self):
        """Speichert Datenbank-Konfigurationen in Datei"""
        data = {
            "databases": [config.to_dict() for config in self.configs.values()]
        }
        
        try:
            with open(self.CONFIG_FILE, 'w', encoding='utf-8') as f:
                json.dump(data, f, indent=2, ensure_ascii=False)
            
            logger.info("Datenbank-Konfigurationen gespeichert")
            
        except Exception as e:
            logger.error(f"Fehler beim Speichern der Datenbank-Konfigurationen: {e}")
    
    def add_config(self, config: DatabaseConfig):
        """Fügt eine neue Datenbank-Konfiguration hinzu"""
        self.configs[config.name] = config
        self.save_configs()
        
    def update_config(self, config: DatabaseConfig):
        """Aktualisiert eine bestehende Datenbank-Konfiguration"""
        self.configs[config.name] = config
        self.save_configs()
        
    def delete_config(self, name: str):
        """Löscht eine Datenbank-Konfiguration"""
        if name in self.configs:
            del self.configs[name]
            # Schließe Verbindung falls vorhanden
            if name in self.connections:
                try:
                    self.connections[name].close()
                except:
                    pass
                del self.connections[name]
            self.save_configs()
    
    def get_config(self, name: str) -> Optional[DatabaseConfig]:
        """Gibt eine Datenbank-Konfiguration zurück"""
        return self.configs.get(name)
    
    def list_configs(self) -> List[DatabaseConfig]:
        """Listet alle Datenbank-Konfigurationen auf"""
        return list(self.configs.values())
    
    def connect(self, config_name: str) -> Optional[pyodbc.Connection]:
        """Stellt eine Verbindung zu einer Datenbank her"""
        if config_name not in self.configs:
            logger.error(f"Datenbank-Konfiguration '{config_name}' nicht gefunden")
            return None
        
        # Prüfe ob bereits verbunden
        if config_name in self.connections:
            try:
                # Teste ob Verbindung noch aktiv ist
                self.connections[config_name].execute("SELECT 1")
                return self.connections[config_name]
            except:
                # Verbindung ist tot, entfernen
                del self.connections[config_name]
        
        config = self.configs[config_name]
        
        try:
            # Baue Connection String
            if config.connection_string:
                conn_str = config.connection_string
            else:
                # Baue Connection String aus Einzelteilen
                conn_str = f"DRIVER={{{config.driver}}};SERVER={config.server};DATABASE={config.database};"
                
                if config.trusted_connection:
                    conn_str += "Trusted_Connection=yes;"
                else:
                    conn_str += f"UID={config.username};PWD={config.password};"
            
            # Verbinde
            connection = pyodbc.connect(conn_str, autocommit=False)
            self.connections[config_name] = connection
            
            logger.info(f"Datenbankverbindung '{config_name}' hergestellt")
            return connection
            
        except Exception as e:
            logger.error(f"Fehler beim Verbinden zu Datenbank '{config_name}': {e}")
            return None
    
    def disconnect(self, config_name: str):
        """Trennt eine Datenbankverbindung"""
        if config_name in self.connections:
            try:
                self.connections[config_name].close()
                del self.connections[config_name]
                logger.info(f"Datenbankverbindung '{config_name}' getrennt")
            except Exception as e:
                logger.error(f"Fehler beim Trennen der Verbindung '{config_name}': {e}")
    
    def disconnect_all(self):
        """Trennt alle Datenbankverbindungen"""
        for config_name in list(self.connections.keys()):
            self.disconnect(config_name)
    
    def execute_query(self, config_name: str, query: str, params: Optional[tuple] = None) -> Optional[List[Dict[str, Any]]]:
        """Führt eine SELECT-Abfrage aus und gibt die Ergebnisse zurück"""
        conn = self.connect(config_name)
        if not conn:
            return None
        
        try:
            cursor = conn.cursor()
            
            if params:
                cursor.execute(query, params)
            else:
                cursor.execute(query)
            
            # Hole Spaltennamen
            columns = [column[0] for column in cursor.description]
            
            # Hole alle Zeilen
            rows = cursor.fetchall()
            
            # Konvertiere zu Dictionary-Liste
            results = []
            for row in rows:
                results.append(dict(zip(columns, row)))
            
            cursor.close()
            return results
            
        except Exception as e:
            logger.error(f"Fehler beim Ausführen der Abfrage: {e}")
            return None
    
    def execute_command(self, config_name: str, command: str, params: Optional[tuple] = None) -> bool:
        """Führt einen SQL-Befehl aus (INSERT, UPDATE, DELETE)"""
        conn = self.connect(config_name)
        if not conn:
            return False
        
        try:
            cursor = conn.cursor()
            
            if params:
                cursor.execute(command, params)
            else:
                cursor.execute(command)
            
            conn.commit()
            cursor.close()
            
            logger.info(f"SQL-Befehl erfolgreich ausgeführt")
            return True
            
        except Exception as e:
            logger.error(f"Fehler beim Ausführen des SQL-Befehls: {e}")
            try:
                conn.rollback()
            except:
                pass
            return False
    
    def execute_many(self, config_name: str, command: str, params_list: List[tuple]) -> bool:
        """Führt einen SQL-Befehl mit mehreren Parameter-Sets aus"""
        conn = self.connect(config_name)
        if not conn:
            return False
        
        try:
            cursor = conn.cursor()
            cursor.executemany(command, params_list)
            conn.commit()
            cursor.close()
            
            logger.info(f"{len(params_list)} SQL-Befehle erfolgreich ausgeführt")
            return True
            
        except Exception as e:
            logger.error(f"Fehler beim Ausführen der SQL-Befehle: {e}")
            try:
                conn.rollback()
            except:
                pass
            return False
    
    def test_connection(self, config_name: str) -> tuple[bool, str]:
        """Testet eine Datenbankverbindung"""
        if config_name not in self.configs:
            return False, f"Konfiguration '{config_name}' nicht gefunden"
        
        try:
            conn = self.connect(config_name)
            if not conn:
                return False, "Verbindung konnte nicht hergestellt werden"
            
            # Teste mit einfacher Abfrage
            cursor = conn.cursor()
            cursor.execute("SELECT 1")
            cursor.close()
            
            return True, "Verbindung erfolgreich"
            
        except Exception as e:
            return False, f"Verbindungsfehler: {str(e)}"
    
    def get_drivers(self) -> List[str]:
        """Gibt eine Liste der verfügbaren ODBC-Treiber zurück"""
        try:
            return pyodbc.drivers()
        except:
            return []
    
    def __del__(self):
        """Cleanup beim Zerstören des Objekts"""
        if hasattr(self, '_initialized') and self._initialized:
            self.disconnect_all()


# Singleton-Instanz
_database_processor = None

def get_database_processor() -> DatabaseProcessor:
    """Gibt die globale DatabaseProcessor-Instanz zurück"""
    global _database_processor
    if _database_processor is None:
        _database_processor = DatabaseProcessor()
    return _database_processor