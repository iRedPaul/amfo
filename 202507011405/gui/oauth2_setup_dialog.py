"""
OAuth2 Setup Dialog für E-Mail-Konfiguration
"""
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.oauth2_manager import OAuth2Manager, get_token_storage


class OAuth2SetupDialog:
    """Dialog zur OAuth2-Einrichtung"""
    
    def __init__(self, parent, provider: str, current_settings: Dict[str, str]):
        self.parent = parent
        self.provider = provider
        self.current_settings = current_settings
        self.result = None
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"OAuth2 Setup - {provider}")
        self.dialog.geometry("700x650")
        self.dialog.resizable(False, False)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # OAuth2 Manager
        self.oauth2_manager = OAuth2Manager(provider)
        self.token_storage = get_token_storage()
        
        self._create_widgets()
        self._layout_widgets()
        self._load_current_settings()
        
        # Bind Events
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())
    
    def _center_window(self):
        """Zentriert das Fenster"""
        self.dialog.update_idletasks()
        
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() - width) // 2
        y = (self.dialog.winfo_screenheight() - height) // 2
        
        self.dialog.geometry(f"+{x}+{y}")
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Hauptframe
        self.main_frame = ttk.Frame(self.dialog, padding="20")
        
        # Anleitung
        self.instruction_frame = ttk.LabelFrame(self.main_frame, text="Anleitung", padding="15")
        
        instructions = self._get_instructions()
        self.instruction_text = tk.Text(self.instruction_frame, height=10, width=70, wrap=tk.WORD)
        self.instruction_text.insert("1.0", instructions)
        self.instruction_text.config(state=tk.DISABLED)
        
        # Client-Credentials
        self.creds_frame = ttk.LabelFrame(self.main_frame, text="OAuth2 Client-Anmeldeinformationen", padding="15")
        
        self.client_id_label = ttk.Label(self.creds_frame, text="Client-ID:")
        self.client_id_var = tk.StringVar()
        self.client_id_entry = ttk.Entry(self.creds_frame, textvariable=self.client_id_var, width=60)
        
        self.client_secret_label = ttk.Label(self.creds_frame, text="Client-Secret:")
        self.client_secret_var = tk.StringVar()
        self.client_secret_entry = ttk.Entry(self.creds_frame, textvariable=self.client_secret_var, width=60, show="*")
        
        # E-Mail-Adresse
        self.email_frame = ttk.LabelFrame(self.main_frame, text="E-Mail-Konto", padding="15")
        
        self.email_label = ttk.Label(self.email_frame, text="E-Mail-Adresse:")
        self.email_var = tk.StringVar()
        self.email_entry = ttk.Entry(self.email_frame, textvariable=self.email_var, width=40)
        
        # Status
        self.status_frame = ttk.LabelFrame(self.main_frame, text="Authentifizierungsstatus", padding="15")
        
        self.status_label = ttk.Label(self.status_frame, text="Status: Nicht authentifiziert", foreground="red")
        self.auth_button = ttk.Button(self.status_frame, text="Jetzt authentifizieren", 
                                     command=self._start_authentication)
        self.remove_auth_button = ttk.Button(self.status_frame, text="Authentifizierung entfernen", 
                                            command=self._remove_authentication, state=tk.DISABLED)
        
        # Progress
        self.progress_label = ttk.Label(self.status_frame, text="")
        self.progress_bar = ttk.Progressbar(self.status_frame, mode='indeterminate', length=300)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save, state=tk.DISABLED)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Anleitung
        self.instruction_frame.pack(fill=tk.X, pady=(0, 15))
        self.instruction_text.pack(fill=tk.BOTH, expand=True)
        
        # Client-Credentials
        self.creds_frame.pack(fill=tk.X, pady=(0, 15))
        self.client_id_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.client_id_entry.grid(row=0, column=1, sticky="we", pady=(0, 5))
        self.client_secret_label.grid(row=1, column=0, sticky=tk.W)
        self.client_secret_entry.grid(row=1, column=1, sticky="we")
        self.creds_frame.columnconfigure(1, weight=1)
        
        # E-Mail
        self.email_frame.pack(fill=tk.X, pady=(0, 15))
        self.email_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.email_entry.grid(row=0, column=1, sticky="we")
        self.email_frame.columnconfigure(1, weight=1)
        
        # Status
        self.status_frame.pack(fill=tk.X, pady=(0, 15))
        self.status_label.pack(anchor=tk.W, pady=(0, 10))
        self.auth_button.pack(side=tk.LEFT, padx=(0, 10))
        self.remove_auth_button.pack(side=tk.LEFT)
        self.progress_label.pack(anchor=tk.W, pady=(10, 5))
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        self.progress_bar.pack_forget()  # Initial versteckt
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.save_button.pack(side=tk.RIGHT)
    
    def _get_instructions(self) -> str:
        """Gibt anbieterspezifische Anleitungen zurück"""
        if self.provider.lower() == "gmail":
            return """So richten Sie OAuth2 für Gmail ein:

1. Gehen Sie zur Google Cloud Console: https://console.cloud.google.com/
2. Erstellen Sie ein neues Projekt oder wählen Sie ein bestehendes
3. Aktivieren Sie die Gmail API für Ihr Projekt
4. Gehen Sie zu "Anmeldedaten" und erstellen Sie OAuth2-Client-ID
5. Wählen Sie "Desktop-App" als Anwendungstyp
6. Fügen Sie http://localhost:8080/callback als Redirect-URI hinzu
7. Kopieren Sie Client-ID und Client-Secret hierher

Hinweis: Ihr Google-Konto muss App-Zugriff erlauben."""
        
        elif self.provider.lower() in ["outlook", "office365"]:
            return """So richten Sie OAuth2 für Outlook/Office 365 ein:

1. Gehen Sie zum Azure Portal: https://portal.azure.com/
2. Navigieren Sie zu "App-Registrierungen" und erstellen Sie eine neue App
3. Wählen Sie "Konten in einem beliebigen Organisationsverzeichnis und persönliche Microsoft-Konten"
4. Fügen Sie http://localhost:8080/callback als Redirect-URI hinzu (Typ: Web)
5. Gehen Sie zu "Zertifikate & Geheimnisse" und erstellen Sie ein neues Client-Secret
6. Kopieren Sie die Application (client) ID und das Secret hierher
7. Fügen Sie unter "API-Berechtigungen" die Berechtigung "SMTP.Send" hinzu

Wichtig: Das Client-Secret läuft nach einer bestimmten Zeit ab und muss erneuert werden."""
        
        else:
            return "Konfigurationsanleitung für diesen Anbieter nicht verfügbar."
    
    def _load_current_settings(self):
        """Lädt die aktuellen Einstellungen"""
        # Client Credentials
        self.client_id_var.set(self.current_settings.get('oauth2_client_id', ''))
        self.client_secret_var.set(self.current_settings.get('oauth2_client_secret', ''))
        
        # E-Mail
        email = self.current_settings.get('smtp_from_address', '')
        if not email:
            email = self.current_settings.get('smtp_username', '')
        self.email_var.set(email)
        
        # Prüfe ob bereits authentifiziert
        if email and self.current_settings.get('oauth2_refresh_token'):
            self._update_status(True)
    
    def _update_status(self, authenticated: bool):
        """Aktualisiert den Authentifizierungsstatus"""
        if authenticated:
            self.status_label.config(text="Status: Authentifiziert ✓", foreground="green")
            self.remove_auth_button.config(state=tk.NORMAL)
            self.save_button.config(state=tk.NORMAL)
        else:
            self.status_label.config(text="Status: Nicht authentifiziert", foreground="red")
            self.remove_auth_button.config(state=tk.DISABLED)
            self.save_button.config(state=tk.DISABLED)
    
    def _start_authentication(self):
        """Startet den OAuth2-Authentifizierungsfluss"""
        # Validiere Eingaben
        if not self.client_id_var.get() or not self.client_secret_var.get():
            messagebox.showerror("Fehler", "Bitte geben Sie Client-ID und Client-Secret ein.")
            return
        
        if not self.email_var.get():
            messagebox.showerror("Fehler", "Bitte geben Sie Ihre E-Mail-Adresse ein.")
            return
        
        # UI-Update
        self.auth_button.config(state=tk.DISABLED)
        self.progress_label.config(text="Starte Authentifizierung...")
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))
        self.progress_bar.start()
        
        # Setze Client Credentials
        self.oauth2_manager.set_client_credentials(
            self.client_id_var.get(),
            self.client_secret_var.get()
        )
        
        # Starte OAuth2-Flow in Thread
        import threading
        thread = threading.Thread(target=self._auth_flow_thread)
        thread.daemon = True
        thread.start()
    
    def _auth_flow_thread(self):
        """Führt den OAuth2-Flow in einem separaten Thread aus"""
        try:
            # Phase 1: Authorization
            self.dialog.after(0, self._update_progress, "Öffne Browser für Autorisierung...")
            success, result = self.oauth2_manager.start_auth_flow()
            
            if not success:
                self.dialog.after(0, self._auth_failed, result)
                return
            
            auth_code = result
            
            # Phase 2: Token Exchange
            self.dialog.after(0, self._update_progress, "Tausche Autorisierungscode gegen Token...")
            success, tokens = self.oauth2_manager.exchange_code_for_tokens(auth_code)
            
            if not success:
                self.dialog.after(0, self._auth_failed, tokens.get('error', 'Unbekannter Fehler'))
                return
            
            # Speichere Tokens
            self.token_storage.set_tokens(
                self.provider.lower(),
                self.email_var.get(),
                tokens
            )
            
            # Erfolg
            self.dialog.after(0, self._auth_success, tokens)
            
        except Exception as e:
            self.dialog.after(0, self._auth_failed, str(e))
    
    def _update_progress(self, message: str):
        """Aktualisiert die Fortschrittsanzeige"""
        self.progress_label.config(text=message)
    
    def _auth_success(self, tokens: Dict[str, str]):
        """Wird bei erfolgreicher Authentifizierung aufgerufen"""
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.progress_label.config(text="Authentifizierung erfolgreich!")
        self.auth_button.config(state=tk.NORMAL)
        
        # Speichere Token-Infos
        self.current_settings['oauth2_access_token'] = tokens['access_token']
        self.current_settings['oauth2_refresh_token'] = tokens['refresh_token']
        self.current_settings['oauth2_token_expiry'] = tokens['token_expiry']
        
        self._update_status(True)
        
        messagebox.showinfo("Erfolg", "OAuth2-Authentifizierung erfolgreich abgeschlossen!")
    
    def _auth_failed(self, error: str):
        """Wird bei fehlgeschlagener Authentifizierung aufgerufen"""
        self.progress_bar.stop()
        self.progress_bar.pack_forget()
        self.progress_label.config(text="")
        self.auth_button.config(state=tk.NORMAL)
        
        messagebox.showerror("Authentifizierung fehlgeschlagen", f"Fehler: {error}")
    
    def _remove_authentication(self):
        """Entfernt die gespeicherte Authentifizierung"""
        if messagebox.askyesno("Authentifizierung entfernen", 
                              "Möchten Sie die gespeicherte Authentifizierung wirklich entfernen?"):
            # Entferne aus Token Storage
            self.token_storage.remove_tokens(self.provider.lower(), self.email_var.get())
            
            # Lösche Token aus Settings
            self.current_settings['oauth2_access_token'] = ''
            self.current_settings['oauth2_refresh_token'] = ''
            self.current_settings['oauth2_token_expiry'] = ''
            
            self._update_status(False)
            messagebox.showinfo("Entfernt", "Authentifizierung wurde entfernt.")
    
    def _on_save(self):
        """Speichert die Einstellungen"""
        self.result = {
            'oauth2_provider': self.provider.lower(),
            'oauth2_client_id': self.client_id_var.get(),
            'oauth2_client_secret': self.client_secret_var.get(),
            'smtp_from_address': self.email_var.get(),
            'oauth2_access_token': self.current_settings.get('oauth2_access_token', ''),
            'oauth2_refresh_token': self.current_settings.get('oauth2_refresh_token', ''),
            'oauth2_token_expiry': self.current_settings.get('oauth2_token_expiry', '')
        }
        
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab ohne zu speichern"""
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict[str, str]]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result