"""
Lizenz-Manager für Hotfolder PDF Processor
Verwaltet Lizenzen mit Hardware-Fingerprint
"""
import os
import json
import logging
import platform
import hashlib
import subprocess
import socket
from datetime import datetime, timedelta
from typing import Dict, Optional, Tuple
from cryptography.fernet import Fernet
from cryptography.hazmat.primitives import hashes
from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
import base64
import uuid

logger = logging.getLogger(__name__)


class LicenseType:
    """Lizenztypen"""
    TRIAL = "trial"
    STANDARD = "standard"


class LicenseManager:
    """Verwaltet Lizenzen mit Hardware-Fingerprint"""
    MASTER_KEY = b"BbzykHng8SgKGyTkbYy7GXdpF8Z8MfuRZrrJ5DRDangvdpaz"
    
    def __init__(self, license_file: str = "license.dat"):
        self.license_file = license_file
        self.hardware_id = self._generate_hardware_id()
        self._cipher = self._create_cipher()
        self.current_license = None
        
    def _create_cipher(self) -> Fernet:
        """Erstellt Verschlüsselungsobjekt"""
        kdf = PBKDF2HMAC(
            algorithm=hashes.SHA256(),
            length=32,
            salt=b'HotfolderPDFSalt',
            iterations=100000,
        )
        key = base64.urlsafe_b64encode(kdf.derive(self.MASTER_KEY))
        return Fernet(key)
    
    def _generate_hardware_id(self) -> str:
        """Generiert einen eindeutigen Hardware-Fingerprint"""
        components = []
        
        # 1. MAC-Adresse
        try:
            # Hole alle Netzwerk-Interfaces
            mac_addresses = []
            for interface in socket.if_nameindex():
                try:
                    mac = ':'.join(['{:02x}'.format((uuid.getnode() >> ele) & 0xff) 
                                  for ele in range(0,8*6,8)][::-1])
                    if mac != '00:00:00:00:00:00':
                        mac_addresses.append(mac)
                except:
                    pass
            
            # Fallback wenn keine MAC gefunden
            if not mac_addresses:
                mac = ':'.join(['{:02x}'.format((uuid.getnode() >> ele) & 0xff) 
                              for ele in range(0,8*6,8)][::-1])
                mac_addresses = [mac]
            
            components.append(sorted(mac_addresses)[0])  # Erste MAC alphabetisch
        except Exception as e:
            logger.warning(f"Konnte MAC-Adresse nicht ermitteln: {e}")
            components.append("NO_MAC")
        
        # 2. Festplatten-Seriennummer (Windows)
        if platform.system() == "Windows":
            try:
                # Hole Seriennummer der System-Festplatte
                result = subprocess.run(
                    ['wmic', 'diskdrive', 'get', 'serialnumber'], 
                    capture_output=True, 
                    text=True, 
                    shell=True
                )
                lines = result.stdout.strip().split('\n')
                if len(lines) > 1:
                    serial = lines[1].strip()
                    if serial:
                        components.append(serial)
                    else:
                        components.append("NO_DISK_SERIAL")
                else:
                    components.append("NO_DISK_SERIAL")
            except Exception as e:
                logger.warning(f"Konnte Festplatten-Serial nicht ermitteln: {e}")
                components.append("NO_DISK_SERIAL")
        else:
            components.append("NOT_WINDOWS")
        
        # 3. CPU-Info
        try:
            if platform.system() == "Windows":
                result = subprocess.run(
                    ['wmic', 'cpu', 'get', 'processorid'], 
                    capture_output=True, 
                    text=True,
                    shell=True
                )
                lines = result.stdout.strip().split('\n')
                if len(lines) > 1:
                    cpu_id = lines[1].strip()
                    if cpu_id:
                        components.append(cpu_id)
                    else:
                        components.append(platform.processor())
                else:
                    components.append(platform.processor())
            else:
                components.append(platform.processor())
        except Exception as e:
            logger.warning(f"Konnte CPU-Info nicht ermitteln: {e}")
            components.append("NO_CPU")
        
        # 4. Hostname als zusätzliche Komponente
        try:
            components.append(socket.gethostname())
        except:
            components.append("NO_HOSTNAME")
        
        # 5. Windows Produkt-ID (zusätzliche Sicherheit)
        if platform.system() == "Windows":
            try:
                result = subprocess.run(
                    ['wmic', 'os', 'get', 'serialnumber'], 
                    capture_output=True, 
                    text=True,
                    shell=True
                )
                lines = result.stdout.strip().split('\n')
                if len(lines) > 1 and lines[1].strip():
                    components.append(lines[1].strip())
                else:
                    # Alternative: Registry-Schlüssel
                    import winreg
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                                       r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
                    product_id = winreg.QueryValueEx(key, "ProductId")[0]
                    components.append(product_id)
                    winreg.CloseKey(key)
            except:
                components.append("NO_PRODUCT_ID")
        
        # Erstelle Hash aus allen Komponenten
        hw_string = "|".join(components)
        
        # Verwende SHA256 für den Hash
        hw_hash = hashlib.sha256(hw_string.encode()).hexdigest()
        
        # Verwende mehr vom Hash für bessere Sicherheit (32 Zeichen = 128 Bit)
        # Format: XXXXXXXX-XXXXXXXX-XXXXXXXX-XXXXXXXX
        formatted_id = '-'.join([hw_hash[i:i+8].upper() for i in range(0, 32, 8)])
        
        logger.info(f"Hardware-ID generiert: {formatted_id}")
        return formatted_id
    
    def create_license_request(self, request_file: str = "license_request.txt") -> str:
        """Erstellt eine Lizenzanfrage-Datei (nur Hardware-ID)"""
        try:
            # Speichere nur die Hardware-ID
            with open(request_file, 'w', encoding='utf-8') as f:
                f.write(self.hardware_id)
            
            logger.info(f"Lizenzanfrage erstellt: {request_file}")
            return request_file
            
        except Exception as e:
            logger.error(f"Fehler beim Erstellen der Lizenzanfrage: {e}")
            raise
    
    def install_license(self, license_data: bytes) -> Tuple[bool, str]:
        """Installiert eine Lizenz aus verschlüsselten Daten"""
        try:
            # Entschlüssele Lizenz
            decrypted_data = self._cipher.decrypt(license_data)
            license_info = json.loads(decrypted_data.decode())
            
            # Validiere Hardware-ID
            if license_info.get("hardware_id") != self.hardware_id:
                return False, "Lizenz ist für andere Hardware ausgestellt"
            
            # Validiere Lizenztyp
            license_type = license_info.get("type")
            if license_type not in [LicenseType.TRIAL, LicenseType.STANDARD]:
                return False, "Ungültiger Lizenztyp"
            
            # Validiere Ablaufdatum (beide Lizenztypen haben jetzt ein Ablaufdatum)
            expiry_date = datetime.fromisoformat(license_info.get("expiry_date"))
            if datetime.now() > expiry_date:
                return False, "Lizenz ist abgelaufen"
            
            # Speichere Lizenz
            with open(self.license_file, 'wb') as f:
                f.write(license_data)
            
            self.current_license = license_info
            logger.info(f"Lizenz erfolgreich installiert: {license_type}")
            return True, "Lizenz erfolgreich installiert"
            
        except Exception as e:
            logger.error(f"Fehler beim Installieren der Lizenz: {e}")
            return False, f"Fehler: {str(e)}"
    
    def validate_license(self) -> Tuple[bool, Optional[Dict], str]:
        """Validiert die aktuelle Lizenz"""
        try:
            # Prüfe ob Lizenzdatei existiert
            if not os.path.exists(self.license_file):
                return False, None, "Keine Lizenz gefunden"
            
            # Lade und entschlüssele Lizenz
            with open(self.license_file, 'rb') as f:
                encrypted_data = f.read()
            
            try:
                decrypted_data = self._cipher.decrypt(encrypted_data)
                license_info = json.loads(decrypted_data.decode())
            except Exception as e:
                return False, None, "Lizenz ist beschädigt oder ungültig"
            
            # Validiere Hardware-ID
            if license_info.get("hardware_id") != self.hardware_id:
                return False, None, "Lizenz ist für andere Hardware ausgestellt"
            
            # Validiere Lizenztyp
            license_type = license_info.get("type")
            if license_type not in [LicenseType.TRIAL, LicenseType.STANDARD]:
                return False, None, "Ungültiger Lizenztyp"
            
            # Validiere Ablaufdatum (alle Lizenzen haben ein Ablaufdatum)
            expiry_date = datetime.fromisoformat(license_info.get("expiry_date"))
            if datetime.now() > expiry_date:
                days_expired = (datetime.now() - expiry_date).days
                return False, license_info, f"Lizenz ist seit {days_expired} Tagen abgelaufen"
            
            # Berechne verbleibende Tage
            days_remaining = (expiry_date - datetime.now()).days
            license_info["days_remaining"] = days_remaining
            
            self.current_license = license_info
            return True, license_info, "Lizenz ist gültig"
            
        except Exception as e:
            logger.error(f"Fehler bei der Lizenzvalidierung: {e}")
            return False, None, f"Fehler: {str(e)}"
    
    def get_license_info(self) -> Optional[Dict]:
        """Gibt Informationen zur aktuellen Lizenz zurück"""
        if self.current_license:
            return self.current_license.copy()
        
        # Versuche Lizenz zu laden
        valid, info, _ = self.validate_license()
        if valid and info:
            return info
        
        return None
    
    def remove_license(self) -> bool:
        """Entfernt die aktuelle Lizenz"""
        try:
            if os.path.exists(self.license_file):
                os.remove(self.license_file)
                self.current_license = None
                logger.info("Lizenz entfernt")
                return True
        except Exception as e:
            logger.error(f"Fehler beim Entfernen der Lizenz: {e}")
        return False
    
    def is_licensed(self) -> bool:
        """Prüft ob eine gültige Lizenz vorhanden ist"""
        valid, _, _ = self.validate_license()
        return valid


# Globale Instanz
_license_manager = None

def get_license_manager() -> LicenseManager:
    """Gibt die globale LicenseManager-Instanz zurück"""
    global _license_manager
    if _license_manager is None:
        _license_manager = LicenseManager()
    return _license_manager