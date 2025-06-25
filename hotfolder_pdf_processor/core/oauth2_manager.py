"""
OAuth2-Manager für E-Mail-Authentifizierung
"""
import os
import json
import base64
import hashlib
import secrets
import webbrowser
import threading
import time
from datetime import datetime, timedelta
from typing import Optional, Dict, Tuple
from urllib.parse import urlencode, parse_qs
from http.server import HTTPServer, BaseHTTPRequestHandler
import requests
import logging

# Logger für dieses Modul
logger = logging.getLogger(__name__)


class OAuth2Config:
    """OAuth2-Konfiguration für verschiedene Anbieter"""
    
    PROVIDERS = {
        "gmail": {
            "name": "Gmail",
            "auth_url": "https://accounts.google.com/o/oauth2/v2/auth",
            "token_url": "https://oauth2.googleapis.com/token",
            "scope": "https://www.googleapis.com/auth/gmail.send",
            "redirect_uri": "http://localhost:8080/callback",
            "client_id": "",  # Muss vom Benutzer bereitgestellt werden
            "client_secret": ""  # Muss vom Benutzer bereitgestellt werden
        },
        "outlook": {
            "name": "Outlook/Office 365",
            "auth_url": "https://login.microsoftonline.com/common/oauth2/v2.0/authorize",
            "token_url": "https://login.microsoftonline.com/common/oauth2/v2.0/token",
            "scope": "https://outlook.office.com/SMTP.Send offline_access",
            "redirect_uri": "http://localhost:8080/callback",
            "client_id": "",  # Muss vom Benutzer bereitgestellt werden
            "client_secret": ""  # Muss vom Benutzer bereitgestellt werden
        }
    }


class OAuth2CallbackHandler(BaseHTTPRequestHandler):
    """HTTP Request Handler für OAuth2 Callback"""
    
    def do_GET(self):
        """Verarbeitet GET Request vom OAuth2 Callback"""
        if self.path.startswith('/callback'):
            # Parse Query String
            query_start = self.path.find('?')
            if query_start != -1:
                query_string = self.path[query_start + 1:]
                params = parse_qs(query_string)
                
                # Speichere Code oder Error
                if 'code' in params:
                    self.server.auth_code = params['code'][0]
                    self.send_response(200)
                    self.send_header('Content-type', 'text/html')
                    self.end_headers()
                    # HTML als String erstellen und dann zu bytes konvertieren
                    success_html = """
                        <html>
                        <body style="font-family: Arial, sans-serif; text-align: center; padding-top: 50px;">
                            <h1 style="color: green;">Authentifizierung erfolgreich!</h1>
                            <p>Sie können dieses Fenster jetzt schließen.</p>
                        </body>
                        </html>
                    """
                    self.wfile.write(success_html.encode('utf-8'))
                elif 'error' in params:
                    self.server.auth_error = params.get('error_description', ['Unbekannter Fehler'])[0]
                    self.send_response(200)
                    self.send_header('Content-type', 'text/html')
                    self.end_headers()
                    error_desc = params.get('error_description', ['Unbekannter Fehler'])[0]
                    error_html = f"""
                        <html>
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


class OAuth2Manager:
    """Verwaltet OAuth2-Authentifizierung für E-Mail-Versand"""
    
    def __init__(self, provider: str):
        self.provider = provider.lower()
        
        # Office365 ist ein Alias für Outlook
        if self.provider == "office365":
            self.provider = "outlook"
        
        if self.provider not in OAuth2Config.PROVIDERS:
            raise ValueError(f"Unbekannter OAuth2-Provider: {provider}")
        
        self.config = OAuth2Config.PROVIDERS[self.provider].copy()
        self.server = None
        self.server_thread = None
    
    def set_client_credentials(self, client_id: str, client_secret: str):
        """Setzt Client-ID und Secret"""
        self.config['client_id'] = client_id
        self.config['client_secret'] = client_secret
    
    def start_auth_flow(self) -> Tuple[bool, str]:
        """
        Startet den OAuth2-Authentifizierungsflow
        
        Returns:
            Tuple[bool, str]: (Erfolg, Fehlermeldung oder Auth-Code)
        """
        if not self.config['client_id'] or not self.config['client_secret']:
            return False, "Client-ID und Client-Secret müssen konfiguriert sein"
        
        # Generiere PKCE Parameter (für zusätzliche Sicherheit)
        code_verifier = base64.urlsafe_b64encode(os.urandom(32)).decode('utf-8').rstrip('=')
        code_challenge = base64.urlsafe_b64encode(
            hashlib.sha256(code_verifier.encode('utf-8')).digest()
        ).decode('utf-8').rstrip('=')
        
        # Generiere State für CSRF-Schutz
        state = secrets.token_urlsafe(16)
        
        # Starte lokalen Server für Callback
        try:
            self.server = HTTPServer(('localhost', 8080), OAuth2CallbackHandler)
            self.server.auth_code = None
            self.server.auth_error = None
            self.server.timeout = 120  # 2 Minuten Timeout
            
            # Starte Server in separatem Thread
            self.server_thread = threading.Thread(target=self._run_server)
            self.server_thread.daemon = True
            self.server_thread.start()
            
            # Baue Authorization URL
            auth_params = {
                'client_id': self.config['client_id'],
                'redirect_uri': self.config['redirect_uri'],
                'response_type': 'code',
                'scope': self.config['scope'],
                'state': state,
                'access_type': 'offline',  # Für Refresh Token
                'prompt': 'consent'  # Zeige immer Consent Screen
            }
            
            if self.provider == 'gmail':
                auth_params['code_challenge'] = code_challenge
                auth_params['code_challenge_method'] = 'S256'
            
            auth_url = f"{self.config['auth_url']}?{urlencode(auth_params)}"
            
            # Öffne Browser
            webbrowser.open(auth_url)
            logger.info("Browser für OAuth2-Authentifizierung geöffnet")
            
            # Warte auf Callback
            timeout = time.time() + 120  # 2 Minuten
            while time.time() < timeout:
                if self.server.auth_code:
                    auth_code = self.server.auth_code
                    self._stop_server()
                    logger.info("OAuth2-Autorisierungscode erhalten")
                    return True, auth_code
                elif self.server.auth_error:
                    error = self.server.auth_error
                    self._stop_server()
                    logger.error(f"OAuth2-Autorisierung fehlgeschlagen: {error}")
                    return False, error
                time.sleep(0.1)
            
            self._stop_server()
            logger.error("OAuth2-Authentifizierung: Timeout erreicht")
            return False, "Timeout - Keine Antwort vom Benutzer erhalten"
            
        except Exception as e:
            self._stop_server()
            logger.exception(f"Fehler beim Starten des OAuth2-Flows: {e}")
            return False, f"Fehler beim Starten des OAuth2-Flows: {str(e)}"
    
    def _run_server(self):
        """Führt den HTTP-Server aus"""
        try:
            while self.server:
                self.server.handle_request()
        except Exception as e:
            logger.error(f"Fehler im OAuth2-Server: {e}")
    
    def _stop_server(self):
        """Stoppt den HTTP-Server"""
        if self.server:
            self.server.shutdown()
            self.server = None
        if self.server_thread:
            self.server_thread.join(timeout=1)
            self.server_thread = None
    
    def exchange_code_for_tokens(self, auth_code: str) -> Tuple[bool, Dict[str, str]]:
        """
        Tauscht Authorization Code gegen Access und Refresh Token
        
        Returns:
            Tuple[bool, Dict]: (Erfolg, Token-Dictionary oder Fehler-Dictionary)
        """
        token_data = {
            'client_id': self.config['client_id'],
            'client_secret': self.config['client_secret'],
            'code': auth_code,
            'redirect_uri': self.config['redirect_uri'],
            'grant_type': 'authorization_code'
        }
        
        try:
            logger.debug("Tausche Autorisierungscode gegen Tokens")
            response = requests.post(self.config['token_url'], data=token_data)
            response.raise_for_status()
            
            tokens = response.json()
            
            # Berechne Ablaufzeit
            expires_in = tokens.get('expires_in', 3600)
            expiry = datetime.now() + timedelta(seconds=expires_in)
            
            logger.info("OAuth2-Tokens erfolgreich erhalten")
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
            'client_id': self.config['client_id'],
            'client_secret': self.config['client_secret'],
            'refresh_token': refresh_token,
            'grant_type': 'refresh_token'
        }
        
        try:
            logger.debug("Erneuere OAuth2-Access-Token")
            response = requests.post(self.config['token_url'], data=token_data)
            response.raise_for_status()
            
            tokens = response.json()
            
            # Berechne neue Ablaufzeit
            expires_in = tokens.get('expires_in', 3600)
            expiry = datetime.now() + timedelta(seconds=expires_in)
            
            logger.info("OAuth2-Token erfolgreich erneuert")
            return True, {
                'access_token': tokens.get('access_token', ''),
                'refresh_token': tokens.get('refresh_token', refresh_token),  # Manche Provider geben neuen Refresh Token
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
                logger.debug("OAuth2-Token ist abgelaufen")
            return is_expired
        except Exception as e:
            logger.error(f"Fehler beim Prüfen der Token-Gültigkeit: {e}")
            return True
    
    def create_oauth2_sasl_string(self, username: str, access_token: str) -> str:
        """
        Erstellt den SASL XOAUTH2 String für SMTP-Authentifizierung
        
        Args:
            username: E-Mail-Adresse des Benutzers
            access_token: OAuth2 Access Token
            
        Returns:
            Base64-codierter SASL String
        """
        auth_string = f"user={username}\x01auth=Bearer {access_token}\x01\x01"
        return base64.b64encode(auth_string.encode()).decode()


class OAuth2TokenStorage:
    """Speichert und lädt OAuth2-Tokens sicher"""
    
    def __init__(self, storage_file: str = "oauth2_tokens.json"):
        self.storage_file = storage_file
        self._tokens = {}
        self.load_tokens()
    
    def load_tokens(self):
        """Lädt gespeicherte Tokens"""
        try:
            if os.path.exists(self.storage_file):
                with open(self.storage_file, 'r', encoding='utf-8') as f:
                    self._tokens = json.load(f)
                logger.debug(f"OAuth2-Tokens geladen: {len(self._tokens)} Konten")
        except Exception as e:
            logger.error(f"Fehler beim Laden der OAuth2-Tokens: {e}")
            self._tokens = {}
    
    def save_tokens(self):
        """Speichert Tokens"""
        try:
            # Erstelle Backup
            if os.path.exists(self.storage_file):
                backup_file = f"{self.storage_file}.backup"
                if os.path.exists(backup_file):
                    os.remove(backup_file)
                os.rename(self.storage_file, backup_file)
            
            # Speichere Tokens
            with open(self.storage_file, 'w', encoding='utf-8') as f:
                json.dump(self._tokens, f, indent=2)
            
            # Entferne Backup
            backup_file = f"{self.storage_file}.backup"
            if os.path.exists(backup_file):
                os.remove(backup_file)
            
            logger.debug("OAuth2-Tokens gespeichert")
                
        except Exception as e:
            logger.error(f"Fehler beim Speichern der OAuth2-Tokens: {e}")
            # Stelle Backup wieder her
            backup_file = f"{self.storage_file}.backup"
            if os.path.exists(backup_file):
                if os.path.exists(self.storage_file):
                    os.remove(self.storage_file)
                os.rename(backup_file, self.storage_file)
                logger.info("OAuth2-Token-Backup wiederhergestellt")
    
    def get_tokens(self, provider: str, username: str) -> Optional[Dict[str, str]]:
        """Holt gespeicherte Tokens für einen Benutzer"""
        key = f"{provider}:{username}"
        tokens = self._tokens.get(key)
        if tokens:
            logger.debug(f"OAuth2-Tokens abgerufen für {username}")
        return tokens
    
    def set_tokens(self, provider: str, username: str, tokens: Dict[str, str]):
        """Speichert Tokens für einen Benutzer"""
        key = f"{provider}:{username}"
        self._tokens[key] = tokens
        self.save_tokens()
        logger.info(f"OAuth2-Tokens gespeichert für {username}")
    
    def remove_tokens(self, provider: str, username: str):
        """Entfernt Tokens für einen Benutzer"""
        key = f"{provider}:{username}"
        if key in self._tokens:
            del self._tokens[key]
            self.save_tokens()
            logger.info(f"OAuth2-Tokens entfernt für {username}")


# Globale Token-Storage Instanz
_token_storage = None

def get_token_storage() -> OAuth2TokenStorage:
    """Gibt die globale Token-Storage Instanz zurück"""
    global _token_storage
    if _token_storage is None:
        _token_storage = OAuth2TokenStorage()
        logger.debug("Globale OAuth2-Token-Storage Instanz erstellt")
    return _token_storage
