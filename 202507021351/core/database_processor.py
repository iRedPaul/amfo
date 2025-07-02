"""
Database Processor für ODBC-Verbindungen und SQL-Abfragen
"""
import pyodbc
import logging
import json
import os
from typing import Dict, Any, Optional, List, Tuple
from pathlib import Path

logger = logging.getLogger(__name__)


class DatabaseConfig:
    """Konfiguration für eine Datenbank-Verbindung"""
    
    def __init__(self, name: str, connection_string: str = "", 
                 driver: str = "", server: str = "", database: str = "",
                 username: str = "", password: str = "", 
                 trusted_connection: bool = False):
        self.name = name
        self.connection_string = connection_string
        self.driver = driver
        self.server = server
        self.database = database
        self.username = username
        self.password = password
        self.trusted_connection = trusted_connection
    
    def get_connection_string(self) -> str:
        """Gibt den Connection String zurück"""
        if self.connection_string:
            return self.connection_string
        
        # Baue Connection String aus Einzelteilen
        parts = []
        
        if self.driver:
            parts.append(f"DRIVER={{{self.driver}}}")
        if self.server:
            parts.append(f"SERVER={self.server}")
        if self.database:
            parts.append(f"DATABASE={self.database}")
        
        if self.trusted_connection:
            parts.append("Trusted_Connection=yes")
        else:
            if self.username:
                parts.append(f"UID={self.username}")
            if self.password:
                parts.append(f"PWD={self.password}")
        
        return ";".join(parts)
    
    def to_dict(self) -> Dict[str, Any]:
        """Konvertiert zu Dictionary"""
        return {
            "name": self.name,
            "connection_string": self.connection_string,
            "driver": self.driver,
            "server": self.server,
            "database": self.database,
            "username": self.username,
            "password": self.password,
            "trusted_connection": self.trusted_connection
        }
    
    @classmethod
    def from_dict(cls, data: Dict[str, Any]) -> 'DatabaseConfig':
        """Erstellt aus Dictionary"""
        return cls(
            name=data.get("name", ""),
            connection_string=data.get("connection_string", ""),
            driver=data.get("driver", ""),
            server=data.get("server", ""),
            database=data.get("database", ""),
            username=data.get("username", ""),
            password=data.get("password", ""),
            trusted_connection=data.get("trusted_connection", False)
        )


class DatabaseProcessor:
    """Verarbeitet Datenbank-Abfragen über ODBC"""
    
    CONFIG_FILE = "database_configs.json"
    _instance = None
    
    def __new__(cls):
        """Singleton-Pattern um sicherzustellen, dass nur eine Instanz existiert"""
        if cls._instance is None:
            cls._instance = super(DatabaseProcessor, cls).__new__(cls)
            cls._instance._initialized = False
        return cls._instance
    
    def __init__(self):
        """Initialisiert den Processor nur einmal"""
        if self._initialized:
            return
            
        self.configs: Dict[str, DatabaseConfig] = {}
        self.connections: Dict[str, pyodbc.Connection] = {}
        self.load_configs()
        self._initialized = True
    
    def load_configs(self):
        """Lädt Datenbank-Konfigurationen aus Datei"""
        if not os.path.exists(self.CONFIG_FILE):
            logger.info(f"Keine Datenbank-Konfigurationsdatei gefunden: {self.CONFIG_FILE}")
            return
        
        try:
            with open(self.CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            self.configs.clear()  # Alte Configs löschen
            for config_data in data.get("databases", []):
                config = DatabaseConfig.from_dict(config_data)
                self.configs[config.name] = config
            
            logger.info(f"{len(self.configs)} Datenbank-Konfigurationen geladen")
            
        except Exception as e:
            logger.error(f"Fehler beim Laden der Datenbank-Konfigurationen: {e}")
    
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
        # Schließe bestehende Verbindung falls vorhanden
        if config.name in self.connections:
            self.close_connection(config.name)
        
        self.configs[config.name] = config
        self.save_configs()
    
    def delete_config(self, name: str):
        """Löscht eine Datenbank-Konfiguration"""
        if name in self.configs:
            # Schließe Verbindung falls vorhanden
            if name in self.connections:
                self.close_connection(name)
            
            del self.configs[name]
            self.save_configs()
    
    def get_config(self, name: str) -> Optional[DatabaseConfig]:
        """Gibt eine Datenbank-Konfiguration zurück"""
        return self.configs.get(name)
    
    def list_configs(self) -> List[str]:
        """Gibt eine Liste aller konfigurierten Datenbanken zurück"""
        return list(self.configs.keys())
    
    def test_connection(self, config: DatabaseConfig) -> Tuple[bool, str]:
        """Testet eine Datenbank-Verbindung"""
        try:
            conn_str = config.get_connection_string()
            conn = pyodbc.connect(conn_str, timeout=5)
            
            # Teste mit einfacher Abfrage
            cursor = conn.cursor()
            cursor.execute("SELECT 1")
            cursor.close()
            conn.close()
            
            return True, "Verbindung erfolgreich!"
            
        except Exception as e:
            return False, str(e)
    
    def get_connection(self, database_name: str) -> Optional[pyodbc.Connection]:
        """Gibt eine Verbindung zur angegebenen Datenbank zurück"""
        # Lade Configs neu um sicherzustellen, dass wir die aktuellen haben
        self.load_configs()
        
        # Prüfe ob Verbindung bereits existiert
        if database_name in self.connections:
            try:
                # Teste ob Verbindung noch aktiv ist
                cursor = self.connections[database_name].cursor()
                cursor.execute("SELECT 1")
                cursor.close()
                return self.connections[database_name]
            except:
                # Verbindung ist tot, entferne sie
                del self.connections[database_name]
        
        # Erstelle neue Verbindung
        config = self.configs.get(database_name)
        if not config:
            logger.error(f"Datenbank-Konfiguration '{database_name}' nicht gefunden")
            logger.debug(f"Verfügbare Konfigurationen: {list(self.configs.keys())}")
            return None
        
        try:
            conn_str = config.get_connection_string()
            conn = pyodbc.connect(conn_str)
            self.connections[database_name] = conn
            logger.debug(f"Verbindung zu Datenbank '{database_name}' hergestellt")
            return conn
            
        except Exception as e:
            logger.error(f"Fehler beim Verbinden mit Datenbank '{database_name}': {e}")
            return None
    
    def close_connection(self, database_name: str):
        """Schließt eine Datenbank-Verbindung"""
        if database_name in self.connections:
            try:
                self.connections[database_name].close()
                del self.connections[database_name]
                logger.debug(f"Verbindung zu Datenbank '{database_name}' geschlossen")
            except Exception as e:
                logger.error(f"Fehler beim Schließen der Verbindung: {e}")
    
    def close_all_connections(self):
        """Schließt alle offenen Verbindungen"""
        for name in list(self.connections.keys()):
            self.close_connection(name)
    
    def execute_query(self, database_name: str, query: str, column_name: str) -> Optional[Any]:
        """
        Führt eine SQL-Abfrage aus und gibt den Wert der angegebenen Spalte zurück
        
        Args:
            database_name: Name der Datenbank-Konfiguration
            query: SQL-Abfrage
            column_name: Name der Spalte, deren Wert zurückgegeben werden soll
            
        Returns:
            Der Wert aus der ersten Zeile der angegebenen Spalte oder None
        """
        conn = self.get_connection(database_name)
        if not conn:
            return None
        
        try:
            cursor = conn.cursor()
            cursor.execute(query)
            
            # Hole erste Zeile
            row = cursor.fetchone()
            if not row:
                logger.debug(f"Keine Daten gefunden für Abfrage: {query}")
                return None
            
            # Finde Spalten-Index
            column_names = [column[0] for column in cursor.description]
            try:
                column_index = column_names.index(column_name)
            except ValueError:
                # Versuche case-insensitive Suche
                column_index = None
                for i, col in enumerate(column_names):
                    if col.lower() == column_name.lower():
                        column_index = i
                        break
                
                if column_index is None:
                    logger.error(f"Spalte '{column_name}' nicht gefunden. Verfügbare Spalten: {column_names}")
                    return None
            
            # Gib Wert zurück
            value = row[column_index]
            logger.debug(f"SQL-Ergebnis: {value}")
            
            cursor.close()
            return value
            
        except Exception as e:
            logger.error(f"Fehler bei SQL-Abfrage: {e}")
            logger.error(f"Query: {query}")
            return None
    
    def get_available_drivers(self) -> List[str]:
        """Gibt eine Liste der verfügbaren ODBC-Treiber zurück"""
        return pyodbc.drivers()
    
    def __del__(self):
        """Destruktor - schließt alle Verbindungen"""
        if hasattr(self, 'connections'):
            self.close_all_connections()