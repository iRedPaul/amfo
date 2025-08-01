"""
Microsoft Graph API Manager für E-Mail-Versand
"""
import os
import json
import base64
import secrets
import webbrowser
import threading
import time
from datetime import datetime, timedelta
from typing import Optional, Dict, Tuple
from urllib.parse import urlencode, parse_qs
from http.server import HTTPServer, BaseHTTPRequestHandler
from cryptography.fernet import Fernet
import keyring
import platform
import requests
import logging
import shutil

# Logger für dieses Modul
logger = logging.getLogger(__name__)


class MSGraphConfig:
    """Microsoft Graph API Konfiguration"""
    
    # Microsoft Graph API Endpoints
    AUTH_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/authorize"
    TOKEN_URL = "https://login.microsoftonline.com/common/oauth2/v2.0/token"
    GRAPH_API_URL = "https://graph.microsoft.com/v1.0"
    
    # Scopes für E-Mail-Versand
    SCOPES = "https://graph.microsoft.com/Mail.Send offline_access"
    
    # Redirect URI
    REDIRECT_URI = "http://localhost:8080/callback"


class MSGraphCallbackHandler(BaseHTTPRequestHandler):
    """HTTP Request Handler für Microsoft Graph Callback"""
    
    def do_GET(self):
        """Verarbeitet GET Request vom OAuth2 Callback"""
        if self.path.startswith('/callback'):
            # Parse Query String
            query_start = self.path.find('?')
            if query_start != -1:
                query_string = self.path[query_start + 1:]
                params = parse_qs(query_string)
                
                # Speichere Code oder Error und setze Event
                if 'code' in params:
                    self.server.auth_code = params['code'][0]
                    self.server.callback_received.set()  # Event setzen
                    self.send_response(200)
                    self.send_header('Content-type', 'text/html; charset=utf-8')
                    self.end_headers()
                    success_html = """
                        <html>
                        <head>
                            <meta charset="UTF-8">
                            <title>Authentifizierung erfolgreich</title>
                        </head>
                        <body style="font-family: Arial, sans-serif; text-align: center; padding-top: 50px;">
                            <h1 style="color: green;">Authentifizierung erfolgreich!</h1>
                            <p>Sie können dieses Fenster jetzt schließen.</p>
                            <script>
                                // Versuche das Fenster nach 3 Sekunden automatisch zu schließen
                                setTimeout(function() {
                                    window.close();
                                }, 3000);
                            </script>
                        </body>
                        </html>
                    """
                    self.wfile.write(success_html.encode('utf-8'))
                elif 'error' in params:
                    self.server.auth_error = params.get('error_description', ['Unbekannter Fehler'])[0]
                    self.server.callback_received.set()  # Event setzen
                    self.send_response(200)
                    self.send_header('Content-type', 'text/html; charset=utf-8')
                    self.end_headers()
                    error_desc = params.get('error_description', ['Unbekannter Fehler'])[0]
                    error_html = f"""
                        <html>
                        <head>
                            <meta charset="UTF-8">
                            <title>Authentifizierung fehlgeschlagen</title>
                        </head>
                        <body style="font-family: Arial, sans-serif; text-align: center; padding-top: 50px;">
                            <h1 style="color: red;">Authentifizierung fehlgeschlagen!</h1>
                            <p>{error_desc}</p>
                        </body>
                        </html>
                    """
                    self.wfile.write(error_html.encode('utf-8'))
            else:
                self.send_response(400)
                self.end_headers()
        else:
            self.send_response(404)
            self.end_headers()
    
    def log_message(self, format, *args):
        """Unterdrückt Log-Nachrichten"""
        pass


class MSGraphManager:
    """Verwaltet Microsoft Graph API Authentifizierung und E-Mail-Versand"""
    
    def __init__(self):
        self.config = MSGraphConfig()
        self.client_id = None
        self.client_secret = None
        self.server = None
        self.server_thread = None
        self.server_stop_event = None
    
    def set_client_credentials(self, client_id: str, client_secret: str):
        """Setzt Client-ID und Secret"""
        self.client_id = client_id
        self.client_secret = client_secret
    
    def start_auth_flow(self) -> Tuple[bool, str]:
        """
        Startet den Microsoft Graph Authentifizierungsflow
        
        Returns:
            Tuple[bool, str]: (Erfolg, Fehlermeldung oder Auth-Code)
        """
        if not self.client_id or not self.client_secret:
            return False, "Client-ID und Client-Secret müssen konfiguriert sein"
        
        # Generiere State für CSRF-Schutz
        state = secrets.token_urlsafe(16)
        
        # Starte lokalen Server für Callback
        try:
            self.server = HTTPServer(('localhost', 8080), MSGraphCallbackHandler)
            self.server.auth_code = None
            self.server.auth_error = None
            self.server.callback_received = threading.Event()  # Event für Callback
            self.server_stop_event = threading.Event()  # Event zum Stoppen des Servers
            
            # Starte Server in separatem Thread
            self.server_thread = threading.Thread(target=self._run_server)
            self.server_thread.daemon = True
            self.server_thread.start()
            
            # Kurz warten, damit Server startet
            time.sleep(0.5)
            
            # Baue Authorization URL
            auth_params = {
                'client_id': self.client_id,
                'response_type': 'code',
                'redirect_uri': self.config.REDIRECT_URI,
                'response_mode': 'query',
                'scope': self.config.SCOPES,
                'state': state
            }
            
            auth_url = f"{self.config.AUTH_URL}?{urlencode(auth_params)}"
            
            # Öffne Browser
            webbrowser.open(auth_url)
            logger.info("Browser für Microsoft Graph Authentifizierung geöffnet")
            
            # Warte auf Callback mit Event
            callback_received = self.server.callback_received.wait(timeout=120)  # 2 Minuten
            
            if callback_received:
                if self.server.auth_code:
                    auth_code = self.server.auth_code
                    logger.info("Microsoft Graph Autorisierungscode erhalten")
                    self._stop_server()
                    return True, auth_code
                elif self.server.auth_error:
                    error = self.server.auth_error
                    logger.error(f"Microsoft Graph Autorisierung fehlgeschlagen: {error}")
                    self._stop_server()
                    return False, error
            else:
                logger.error("Microsoft Graph Authentifizierung: Timeout erreicht")
                self._stop_server()
                return False, "Timeout - Keine Antwort vom Benutzer erhalten"
            
        except Exception as e:
            self._stop_server()
            logger.exception(f"Fehler beim Starten des Microsoft Graph Flows: {e}")
            return False, f"Fehler beim Starten des Microsoft Graph Flows: {str(e)}"
    
    def _run_server(self):
        """Führt den HTTP-Server aus"""
        try:
            logger.debug("Microsoft Graph Server gestartet")
            while not self.server.callback_received.is_set() and not self.server_stop_event.is_set():
                self.server.timeout = 0.5
                self.server.handle_request()
            logger.debug("Microsoft Graph Server beendet")
        except Exception as e:
            logger.error(f"Fehler im Microsoft Graph Server: {e}")
        finally:
            if hasattr(self.server, 'callback_received'):
                self.server.callback_received.set()
    
    def _stop_server(self):
        """Stoppt den HTTP-Server"""
        logger.debug("Stoppe Microsoft Graph Server...")
        
        if self.server_stop_event:
            self.server_stop_event.set()
        
        if self.server:
            try:
                if hasattr(self.server, 'callback_received'):
                    self.server.callback_received.set()
                
                time.sleep(0.1)
                self.server.server_close()
                logger.debug("Server Socket geschlossen")
            except Exception as e:
                logger.debug(f"Fehler beim Stoppen des Servers: {e}")
            finally:
                self.server = None
        
        if self.server_thread and self.server_thread.is_alive():
            logger.debug("Warte auf Server-Thread...")
            self.server_thread.join(timeout=2)
            if self.server_thread.is_alive():
                logger.warning("Server-Thread konnte nicht sauber beendet werden")
            else:
                logger.debug("Server-Thread beendet")
            self.server_thread = None
        
        self.server_stop_event = None
        logger.debug("Microsoft Graph Server vollständig gestoppt")
    
    def exchange_code_for_tokens(self, auth_code: str) -> Tuple[bool, Dict[str, str]]:
        """
        Tauscht Authorization Code gegen Access und Refresh Token
        
        Returns:
            Tuple[bool, Dict]: (Erfolg, Token-Dictionary oder Fehler-Dictionary)
        """
        token_data = {
            'client_id': self.client_id,
            'scope': self.config.SCOPES,
            'code': auth_code,
            'redirect_uri': self.config.REDIRECT_URI,
            'grant_type': 'authorization_code',
            'client_secret': self.client_secret
        }
        
        try:
            logger.debug("Tausche Autorisierungscode gegen Tokens")
            response = requests.post(self.config.TOKEN_URL, data=token_data)
            response.raise_for_status()
            
            tokens = response.json()
            
            # Berechne Ablaufzeit
            expires_in = tokens.get('expires_in', 3600)
            expiry = datetime.now() + timedelta(seconds=expires_in)
            
            logger.info("Microsoft Graph Tokens erfolgreich erhalten")
            return True, {
                'access_token': tokens.get('access_token', ''),
                'refresh_token': tokens.get('refresh_token', ''),
                'token_expiry': expiry.isoformat(),
                'token_type': tokens.get('token_type', 'Bearer')
            }
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Token-Austausch fehlgeschlagen: {e}")
            return False, {'error': f"Token-Austausch fehlgeschlagen: {str(e)}"}
    
    def refresh_access_token(self, refresh_token: str) -> Tuple[bool, Dict[str, str]]:
        """
        Erneuert den Access Token mit dem Refresh Token
        
        Returns:
            Tuple[bool, Dict]: (Erfolg, Token-Dictionary oder Fehler-Dictionary)
        """
        token_data = {
            'client_id': self.client_id,
            'scope': self.config.SCOPES,
            'refresh_token': refresh_token,
            'grant_type': 'refresh_token',
            'client_secret': self.client_secret
        }
        
        try:
            logger.debug("Erneuere Microsoft Graph Access-Token")
            response = requests.post(self.config.TOKEN_URL, data=token_data)
            response.raise_for_status()
            
            tokens = response.json()
            
            # Berechne neue Ablaufzeit
            expires_in = tokens.get('expires_in', 3600)
            expiry = datetime.now() + timedelta(seconds=expires_in)
            
            logger.info("Microsoft Graph Token erfolgreich erneuert")
            return True, {
                'access_token': tokens.get('access_token', ''),
                'refresh_token': tokens.get('refresh_token', refresh_token),
                'token_expiry': expiry.isoformat(),
                'token_type': tokens.get('token_type', 'Bearer')
            }
            
        except requests.exceptions.RequestException as e:
            logger.error(f"Token-Erneuerung fehlgeschlagen: {e}")
            return False, {'error': f"Token-Erneuerung fehlgeschlagen: {str(e)}"}
    
    def is_token_expired(self, token_expiry: str) -> bool:
        """Prüft ob ein Token abgelaufen ist"""
        try:
            expiry = datetime.fromisoformat(token_expiry)
            # 5 Minuten Puffer
            is_expired = datetime.now() >= expiry - timedelta(minutes=5)
            if is_expired:
                logger.debug("Microsoft Graph Token ist abgelaufen")
            return is_expired
        except Exception as e:
            logger.error(f"Fehler beim Prüfen der Token-Gültigkeit: {e}")
            return True
    
    def send_email(self, access_token: str, from_address: str, to_addresses: list, 
                   subject: str, body: str, attachments: list = None) -> Tuple[bool, str]:
        """
        Sendet eine E-Mail über Microsoft Graph API
        
        Args:
            access_token: Gültiger Access Token
            from_address: Absender-Adresse
            to_addresses: Liste von Empfänger-Adressen
            subject: Betreff
            body: E-Mail-Text
            attachments: Liste von Anhängen (optional)
            
        Returns:
            Tuple[bool, str]: (Erfolg, Fehlermeldung bei Misserfolg)
        """
        # Erstelle E-Mail-Objekt
        message = {
            "message": {
                "subject": subject,
                "body": {
                    "contentType": "Text",
                    "content": body
                },
                "toRecipients": [
                    {
                        "emailAddress": {
                            "address": addr
                        }
                    } for addr in to_addresses
                ]
            },
            "saveToSentItems": "true"
        }
        
        # Füge Anhänge hinzu wenn vorhanden
        if attachments:
            message["message"]["attachments"] = []
            for attachment in attachments:
                with open(attachment['path'], 'rb') as f:
                    content = base64.b64encode(f.read()).decode()
                
                message["message"]["attachments"].append({
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": attachment['name'],
                    "contentType": attachment.get('content_type', 'application/octet-stream'),
                    "contentBytes": content
                })
        
        # Sende E-Mail
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        try:
            response = requests.post(
                f"{self.config.GRAPH_API_URL}/users/{from_address}/sendMail",
                headers=headers,
                json=message
            )
            
            if response.status_code == 202:  # Accepted
                logger.info("E-Mail erfolgreich über Microsoft Graph gesendet")
                return True, ""
            else:
                error_msg = f"Microsoft Graph API Fehler: {response.status_code} - {response.text}"
                logger.error(error_msg)
                return False, error_msg
                
        except Exception as e:
            error_msg = f"Fehler beim Senden der E-Mail: {str(e)}"
            logger.error(error_msg)
            return False, error_msg


class MSGraphTokenStorage:
    """Speichert und lädt Microsoft Graph Tokens sicher verschlüsselt"""
    
    def __init__(self, storage_file: str = "msgraph_tokens.enc"):
        self.storage_file = storage_file
        self._tokens = {}
        self._cipher = self._get_cipher()
        self.load_tokens()
    
    def _get_cipher(self) -> Fernet:
        """Holt oder erstellt den Verschlüsselungsschlüssel"""
        try:
            # Versuche Schlüssel aus Keyring zu holen
            key = keyring.get_password("HotfolderPDFProcessor", "msgraph_key")
            
            if not key:
                # Generiere neuen Schlüssel
                key = Fernet.generate_key().decode()
                keyring.set_password("HotfolderPDFProcessor", "msgraph_key", key)
                logger.info("Neuer Verschlüsselungsschlüssel erstellt")
            
            return Fernet(key.encode())
            
        except Exception as e:
            logger.warning(f"Keyring nicht verfügbar: {e}. Verwende fallback.")
            # Fallback: Verwende maschinenspezifischen Schlüssel
            machine_id = platform.node() + platform.machine()
            import hashlib
            hash_obj = hashlib.sha256(machine_id.encode())
            key_base = base64.urlsafe_b64encode(hash_obj.digest())
            return Fernet(key_base)
    
    def load_tokens(self):
        """Lädt und entschlüsselt gespeicherte Tokens"""
        try:
            if os.path.exists(self.storage_file):
                with open(self.storage_file, 'rb') as f:
                    encrypted_data = f.read()
                
                if encrypted_data:
                    decrypted_data = self._cipher.decrypt(encrypted_data)
                    self._tokens = json.loads(decrypted_data.decode())
                    logger.debug(f"Microsoft Graph Tokens geladen: {len(self._tokens)} Konten")
                else:
                    self._tokens = {}
        except Exception as e:
            logger.error(f"Fehler beim Laden der Microsoft Graph Tokens: {e}")
            # Bei Entschlüsselungsfehler: Backup erstellen und neu starten
            if os.path.exists(self.storage_file):
                backup_file = f"{self.storage_file}.backup_{int(time.time())}"
                shutil.copy2(self.storage_file, backup_file)
                logger.info(f"Token-Backup erstellt: {backup_file}")
            self._tokens = {}
    
    def save_tokens(self):
        """Verschlüsselt und speichert Tokens"""
        try:
            # Erstelle Backup
            if os.path.exists(self.storage_file):
                backup_file = f"{self.storage_file}.backup"
                if os.path.exists(backup_file):
                    os.remove(backup_file)
                os.rename(self.storage_file, backup_file)
            
            # Verschlüssele und speichere Tokens
            json_data = json.dumps(self._tokens)
            encrypted_data = self._cipher.encrypt(json_data.encode())
            
            with open(self.storage_file, 'wb') as f:
                f.write(encrypted_data)
            
            # Entferne Backup
            backup_file = f"{self.storage_file}.backup"
            if os.path.exists(backup_file):
                os.remove(backup_file)
            
            logger.debug("Microsoft Graph Tokens sicher gespeichert")
                
        except Exception as e:
            logger.error(f"Fehler beim Speichern der Microsoft Graph Tokens: {e}")
            # Stelle Backup wieder her
            backup_file = f"{self.storage_file}.backup"
            if os.path.exists(backup_file):
                if os.path.exists(self.storage_file):
                    os.remove(self.storage_file)
                os.rename(backup_file, self.storage_file)
                logger.info("Microsoft Graph Token-Backup wiederhergestellt")
    
    def set_tokens(self, email: str, tokens: Dict[str, str]):
        """Speichert Tokens für ein Konto"""
        self._tokens[email] = tokens
        self.save_tokens()
        logger.debug(f"Tokens für {email} gespeichert")
    
    def get_tokens(self, email: str) -> Optional[Dict[str, str]]:
        """Lädt Tokens für ein Konto"""
        return self._tokens.get(email)
    
    def remove_tokens(self, email: str):
        """Entfernt Tokens für ein Konto"""
        if email in self._tokens:
            del self._tokens[email]
            self.save_tokens()
            logger.debug(f"Tokens für {email} entfernt")

# Globale Token-Storage Instanz
_token_storage = None

def get_token_storage() -> MSGraphTokenStorage:
    """Gibt die globale Token-Storage Instanz zurück"""
    global _token_storage
    if _token_storage is None:
        _token_storage = MSGraphTokenStorage()
        logger.debug("Globale Microsoft Graph Token-Storage Instanz erstellt")
    return _token_storage