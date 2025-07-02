"""
Dialog für globale Anwendungseinstellungen
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional
import sys
import os
import json
import logging
import subprocess
import glob

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.export_config import ExportSettings, AuthMethod, ApplicationPaths
from gui.oauth2_setup_dialog import OAuth2SetupDialog

logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)  # GUI-Komponenten nur INFO und höher

class SettingsDialog:
    """Dialog für globale Einstellungen"""
    
    def __init__(self, parent):
        self.parent = parent
        self.settings_file = "settings.json"
        self.settings = self._load_settings()
        self.result = None
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Einstellungen")
        self.dialog.geometry("700x750")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Stelle sicher, dass ein Standard-Fehlerordner existiert
        self._ensure_default_error_path()
        
        self._create_widgets()
        self._layout_widgets()
        self._load_values()
        
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
    
    def _ensure_default_error_path(self):
        """Stellt sicher, dass ein Standard-Fehlerordner existiert"""
        if not self.settings.default_error_path:
            # Erstelle Standard-Fehlerordner in AppData
            appdata = os.getenv('APPDATA')
            if appdata:
                default_error_path = os.path.join(appdata, 'HotfolderPDFProcessor', 'errors')
                os.makedirs(default_error_path, exist_ok=True)
                self.settings.default_error_path = default_error_path
            else:
                # Fallback: Im Programmverzeichnis
                default_error_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'errors')
                os.makedirs(default_error_path, exist_ok=True)
                self.settings.default_error_path = default_error_path
    
    def _load_settings(self) -> ExportSettings:
        """Lädt die Einstellungen aus der Datei"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return ExportSettings.from_dict(data)
        except Exception as e:
            logger.error(f"Fehler beim Laden der Einstellungen: {e}")
        
        return ExportSettings()
    
    def _save_settings(self):
        """Speichert die Einstellungen in die Datei"""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings.to_dict(), f, indent=2, ensure_ascii=False)
        except Exception as e:
            logger.error(f"Fehler beim Speichern der Einstellungen: {e}")
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Hauptframe
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Notebook für verschiedene Einstellungskategorien
        self.notebook = ttk.Notebook(self.main_frame)
        
        # Allgemeine Einstellungen
        self.general_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.general_frame, text="Allgemein")
        
        # Fehlerbehandlung
        self.error_frame_content = ttk.LabelFrame(self.general_frame, 
                                                 text="Fehlerbehandlung", padding="10")
        
        self.error_path_label = ttk.Label(self.error_frame_content, 
            text="Standard-Fehlerpfad: *")
        self.error_path_desc = ttk.Label(self.error_frame_content, 
            text="Dieser Pfad wird verwendet, wenn im Hotfolder kein spezifischer Fehlerpfad definiert ist.\n"
                 "Dateien, die nicht verarbeitet werden können, werden hier abgelegt.",
            wraplength=500, foreground="gray")
        
        self.error_path_frame = ttk.Frame(self.error_frame_content)
        self.error_path_var = tk.StringVar()
        self.error_path_entry = ttk.Entry(self.error_path_frame, 
                                         textvariable=self.error_path_var, width=50)
        self.error_path_button = ttk.Button(self.error_path_frame, text="Durchsuchen...", 
                                           command=self._browse_error_path)
        
        # E-Mail-Einstellungen
        self.email_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.email_frame, text="E-Mail")
        
        # Authentifizierungsmethode
        self.auth_frame = ttk.LabelFrame(self.email_frame, text="Authentifizierung", padding="10")
        
        self.auth_method_var = tk.StringVar(value=self.settings.smtp_auth_method.value)
        self.auth_basic_radio = ttk.Radiobutton(self.auth_frame, text="Standard (Benutzername/Passwort)", 
                                               variable=self.auth_method_var, value=AuthMethod.BASIC.value,
                                               command=self._on_auth_method_changed)
        self.auth_oauth2_radio = ttk.Radiobutton(self.auth_frame, text="OAuth2 (Google, Microsoft)", 
                                                variable=self.auth_method_var, value=AuthMethod.OAUTH2.value,
                                                command=self._on_auth_method_changed)
        
        # SMTP-Server
        self.smtp_frame = ttk.LabelFrame(self.email_frame, text="SMTP-Server", padding="10")
        
        self.smtp_server_label = ttk.Label(self.smtp_frame, text="Server:")
        self.smtp_server_var = tk.StringVar()
        self.smtp_server_entry = ttk.Entry(self.smtp_frame, textvariable=self.smtp_server_var, 
                                          width=40)
        
        self.smtp_port_label = ttk.Label(self.smtp_frame, text="Port:")
        self.smtp_port_var = tk.IntVar()
        self.smtp_port_spinbox = ttk.Spinbox(self.smtp_frame, from_=1, to=65535, 
                                            textvariable=self.smtp_port_var, width=10)
        
        # SSL/TLS Optionen
        self.smtp_ssl_var = tk.BooleanVar()
        self.smtp_ssl_check = ttk.Checkbutton(self.smtp_frame, text="SSL verwenden (Port 465)", 
                                             variable=self.smtp_ssl_var,
                                             command=self._on_ssl_changed)
        
        self.smtp_tls_var = tk.BooleanVar()
        self.smtp_tls_check = ttk.Checkbutton(self.smtp_frame, text="TLS verwenden (STARTTLS, Port 587)", 
                                             variable=self.smtp_tls_var,
                                             command=self._on_tls_changed)
        
        # Standard-Anmeldung
        self.smtp_auth_frame = ttk.LabelFrame(self.email_frame, text="Standard-Anmeldung", padding="10")
        
        self.smtp_username_label = ttk.Label(self.smtp_auth_frame, text="Benutzername:")
        self.smtp_username_var = tk.StringVar()
        self.smtp_username_entry = ttk.Entry(self.smtp_auth_frame, 
                                            textvariable=self.smtp_username_var, width=40)
        
        self.smtp_password_label = ttk.Label(self.smtp_auth_frame, text="Passwort:")
        self.smtp_password_var = tk.StringVar()
        self.smtp_password_entry = ttk.Entry(self.smtp_auth_frame, 
                                            textvariable=self.smtp_password_var, width=40, 
                                            show="*")
        
        # OAuth2-Konfiguration
        self.oauth2_frame = ttk.LabelFrame(self.email_frame, text="OAuth2-Konfiguration", padding="10")
        
        self.oauth2_provider_label = ttk.Label(self.oauth2_frame, text="Anbieter:")
        self.oauth2_provider_var = tk.StringVar()
        self.oauth2_provider_combo = ttk.Combobox(self.oauth2_frame, 
                                                 textvariable=self.oauth2_provider_var,
                                                 values=["Gmail", "Outlook"],
                                                 state="readonly", width=20)
        
        self.oauth2_setup_button = ttk.Button(self.oauth2_frame, 
                                            text="OAuth2 einrichten...", 
                                            command=self._setup_oauth2)
        
        self.oauth2_status_label = ttk.Label(self.oauth2_frame, text="Status: Nicht konfiguriert", 
                                           foreground="red")
        
        # OAuth2-Info
        self.oauth2_info_frame = ttk.Frame(self.oauth2_frame)
        self.oauth2_email_label = ttk.Label(self.oauth2_info_frame, text="E-Mail: -")
        self.oauth2_client_label = ttk.Label(self.oauth2_info_frame, text="Client-ID: -")
        
        # Absender
        self.smtp_sender_frame = ttk.LabelFrame(self.email_frame, text="Absender", padding="10")
        
        self.smtp_from_label = ttk.Label(self.smtp_sender_frame, text="Absender-Adresse:")
        self.smtp_from_var = tk.StringVar()
        self.smtp_from_entry = ttk.Entry(self.smtp_sender_frame, 
                                        textvariable=self.smtp_from_var, width=40)
        
        # Test-Button
        self.email_test_button = ttk.Button(self.email_frame, 
                                           text="E-Mail-Einstellungen testen...", 
                                           command=self._test_email)
        
        # Anwendungspfade-Tab
        self.paths_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.paths_frame, text="Anwendungspfade")
        
        # Pfade-Einstellungen
        self.paths_info_label = ttk.Label(
            self.paths_frame,
            text="Konfigurieren Sie die Pfade zu den benötigten Anwendungen.\n" +
                 "Lassen Sie ein Feld leer, um die automatische Erkennung zu verwenden.",
            wraplength=550
        )
        
        # Tesseract
        self.tesseract_frame = ttk.LabelFrame(self.paths_frame, text="Tesseract OCR", padding="10")
        self.tesseract_path_var = tk.StringVar()
        self.tesseract_entry = ttk.Entry(self.tesseract_frame, textvariable=self.tesseract_path_var, width=50)
        self.tesseract_browse_btn = ttk.Button(self.tesseract_frame, text="Durchsuchen...", 
                                              command=lambda: self._browse_app_path('tesseract'))
        self.tesseract_auto_btn = ttk.Button(self.tesseract_frame, text="Auto-Erkennung", 
                                            command=lambda: self._auto_detect_path('tesseract'))
        self.tesseract_status_label = ttk.Label(self.tesseract_frame, text="")
        
        # Ghostscript
        self.ghostscript_frame = ttk.LabelFrame(self.paths_frame, text="Ghostscript", padding="10")
        self.ghostscript_path_var = tk.StringVar()
        self.ghostscript_entry = ttk.Entry(self.ghostscript_frame, textvariable=self.ghostscript_path_var, width=50)
        self.ghostscript_browse_btn = ttk.Button(self.ghostscript_frame, text="Durchsuchen...", 
                                                command=lambda: self._browse_app_path('ghostscript'))
        self.ghostscript_auto_btn = ttk.Button(self.ghostscript_frame, text="Auto-Erkennung", 
                                              command=lambda: self._auto_detect_path('ghostscript'))
        self.ghostscript_status_label = ttk.Label(self.ghostscript_frame, text="")
        
        # Poppler
        self.poppler_frame = ttk.LabelFrame(self.paths_frame, text="Poppler (für Bild-Export)", padding="10")
        self.poppler_path_var = tk.StringVar()
        self.poppler_entry = ttk.Entry(self.poppler_frame, textvariable=self.poppler_path_var, width=50)
        self.poppler_browse_btn = ttk.Button(self.poppler_frame, text="Durchsuchen...", 
                                            command=lambda: self._browse_app_path('poppler'))
        self.poppler_auto_btn = ttk.Button(self.poppler_frame, text="Auto-Erkennung", 
                                          command=lambda: self._auto_detect_path('poppler'))
        self.poppler_status_label = ttk.Label(self.poppler_frame, text="")
        
        # Test-Button für alle Pfade
        self.paths_test_button = ttk.Button(self.paths_frame, text="Alle Pfade testen", 
                                           command=self._test_all_paths)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save, default=tk.ACTIVE)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Notebook
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Allgemeine Einstellungen
        self.error_frame_content.pack(fill=tk.X, pady=(0, 10))
        self.error_path_label.pack(anchor=tk.W)
        self.error_path_desc.pack(anchor=tk.W, pady=(5, 10))
        self.error_path_frame.pack(fill=tk.X)
        self.error_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.error_path_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # E-Mail-Einstellungen
        # Authentifizierung
        self.auth_frame.pack(fill=tk.X, pady=(0, 10))
        self.auth_basic_radio.pack(anchor=tk.W)
        self.auth_oauth2_radio.pack(anchor=tk.W, pady=(5, 0))
        
        # SMTP-Server
        self.smtp_frame.pack(fill=tk.X, pady=(0, 10))
        self.smtp_server_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.smtp_server_entry.grid(row=0, column=1, sticky="we")
        self.smtp_port_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.smtp_port_spinbox.grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
        ttk.Separator(self.smtp_frame, orient='horizontal').grid(row=2, column=0, columnspan=2, 
                                                                sticky="ew", pady=10)
        self.smtp_ssl_check.grid(row=3, column=0, columnspan=2, sticky=tk.W)
        self.smtp_tls_check.grid(row=4, column=0, columnspan=2, sticky=tk.W, pady=(5, 0))
        self.smtp_frame.columnconfigure(1, weight=1)
        
        # Standard-Anmeldung (initial sichtbar)
        self.smtp_auth_frame.pack(fill=tk.X, pady=(0, 10))
        self.smtp_username_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.smtp_username_entry.grid(row=0, column=1, sticky="we")
        self.smtp_password_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.smtp_password_entry.grid(row=1, column=1, sticky="we", pady=(5, 0))
        self.smtp_auth_frame.columnconfigure(1, weight=1)
        
        # OAuth2 (initial versteckt)
        self.oauth2_frame.pack_forget()
        
        # Absender
        self.smtp_sender_frame.pack(fill=tk.X, pady=(0, 10))
        self.smtp_from_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.smtp_from_entry.grid(row=0, column=1, sticky="we")
        self.smtp_sender_frame.columnconfigure(1, weight=1)
        
        # Test-Button
        self.email_test_button.pack(anchor=tk.W)
        
        # Anwendungspfade Layout
        self.paths_info_label.pack(fill=tk.X, pady=(0, 20))
        
        # Tesseract
        self.tesseract_frame.pack(fill=tk.X, pady=(0, 10))
        self.tesseract_entry.pack(fill=tk.X)
        tesseract_btn_frame = ttk.Frame(self.tesseract_frame)
        tesseract_btn_frame.pack(fill=tk.X, pady=(5, 0))
        self.tesseract_browse_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.tesseract_auto_btn.pack(side=tk.LEFT)
        self.tesseract_status_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Ghostscript
        self.ghostscript_frame.pack(fill=tk.X, pady=(0, 10))
        self.ghostscript_entry.pack(fill=tk.X)
        ghostscript_btn_frame = ttk.Frame(self.ghostscript_frame)
        ghostscript_btn_frame.pack(fill=tk.X, pady=(5, 0))
        self.ghostscript_browse_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.ghostscript_auto_btn.pack(side=tk.LEFT)
        self.ghostscript_status_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Poppler
        self.poppler_frame.pack(fill=tk.X, pady=(0, 10))
        self.poppler_entry.pack(fill=tk.X)
        poppler_btn_frame = ttk.Frame(self.poppler_frame)
        poppler_btn_frame.pack(fill=tk.X, pady=(5, 0))
        self.poppler_browse_btn.pack(side=tk.LEFT, padx=(0, 5))
        self.poppler_auto_btn.pack(side=tk.LEFT)
        self.poppler_status_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Test-Button
        self.paths_test_button.pack(pady=(20, 0))
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.save_button.pack(side=tk.RIGHT)
    
    def _load_values(self):
        """Lädt die aktuellen Einstellungen in die UI"""
        self.error_path_var.set(self.settings.default_error_path)
        self.smtp_server_var.set(self.settings.smtp_server)
        self.smtp_port_var.set(self.settings.smtp_port)
        self.smtp_ssl_var.set(self.settings.smtp_use_ssl)
        self.smtp_tls_var.set(self.settings.smtp_use_tls)
        self.smtp_username_var.set(self.settings.smtp_username)
        self.smtp_password_var.set(self.settings.smtp_password)
        self.smtp_from_var.set(self.settings.smtp_from_address)
        
        # OAuth2
        self.auth_method_var.set(self.settings.smtp_auth_method.value)
        self.oauth2_provider_var.set(self.settings.oauth2_provider.capitalize() if self.settings.oauth2_provider else "Gmail")
        
        # Anwendungspfade laden
        if hasattr(self.settings, 'application_paths'):
            self.tesseract_path_var.set(self.settings.application_paths.tesseract or "")
            self.ghostscript_path_var.set(self.settings.application_paths.ghostscript or "")
            self.poppler_path_var.set(self.settings.application_paths.poppler or "")
        
        # Update OAuth2 Status
        self._update_oauth2_status()
        
        # Update UI basierend auf Auth-Methode
        self._on_auth_method_changed()
        
        # Teste Pfade beim Laden
        self.dialog.after(100, self._test_all_paths)
    
    def _update_oauth2_status(self):
        """Aktualisiert den OAuth2-Status in der UI"""
        if self.settings.oauth2_refresh_token:
            self.oauth2_status_label.config(text="Status: Konfiguriert ✓", foreground="green")
            
            # Zeige Info
            if self.settings.smtp_from_address:
                self.oauth2_email_label.config(text=f"E-Mail: {self.settings.smtp_from_address}")
            
            if self.settings.oauth2_client_id:
                # Zeige nur die ersten und letzten Zeichen der Client-ID
                client_id = self.settings.oauth2_client_id
                if len(client_id) > 20:
                    display_id = f"{client_id[:8]}...{client_id[-8:]}"
                else:
                    display_id = client_id
                self.oauth2_client_label.config(text=f"Client-ID: {display_id}")
        else:
            self.oauth2_status_label.config(text="Status: Nicht konfiguriert", foreground="red")
            self.oauth2_email_label.config(text="E-Mail: -")
            self.oauth2_client_label.config(text="Client-ID: -")
    
    def _browse_error_path(self):
        """Öffnet Dialog zur Auswahl des Fehler-Pfads"""
        folder = filedialog.askdirectory(
            title="Standard-Fehlerpfad auswählen",
            initialdir=self.error_path_var.get() or os.path.expanduser("~")
        )
        if folder:
            self.error_path_var.set(folder)
    
    def _browse_app_path(self, app_name: str):
        """Öffnet Dialog zur Auswahl eines Anwendungspfads"""
        if app_name == 'poppler':
            # Ordner-Auswahl für Poppler
            path = filedialog.askdirectory(
                title="Poppler bin-Ordner auswählen",
                initialdir=self.poppler_path_var.get() or os.path.expanduser("~")
            )
            if path:
                self.poppler_path_var.set(path)
        else:
            # Datei-Auswahl für ausführbare Dateien
            filetypes = [("Ausführbare Dateien", "*.exe"), ("Alle Dateien", "*.*")] if os.name == 'nt' else [("Alle Dateien", "*.*")]
            
            title = "Tesseract ausführbare Datei auswählen" if app_name == 'tesseract' else "Ghostscript ausführbare Datei auswählen"
            
            filename = filedialog.askopenfilename(
                title=title,
                filetypes=filetypes,
                initialdir=os.path.dirname(getattr(self, f"{app_name}_path_var").get()) or os.path.expanduser("~")
            )
            
            if filename:
                getattr(self, f"{app_name}_path_var").set(filename)
        
        # Teste den neuen Pfad
        self._test_single_path(app_name)
    
    def _auto_detect_path(self, app_name: str):
        """Versucht einen Anwendungspfad automatisch zu erkennen"""
        path = None
        
        if app_name == "tesseract":
            if os.name == 'nt':
                possible_paths = [
                    r"C:\Program Files\Tesseract-OCR\tesseract.exe",
                    r"C:\Program Files (x86)\Tesseract-OCR\tesseract.exe",
                    r"D:\Program Files\Tesseract-OCR\tesseract.exe",
                ]
                
                for p in possible_paths:
                    if os.path.exists(p):
                        path = p
                        break
            else:
                # Linux/Mac
                result = subprocess.run(['which', 'tesseract'], capture_output=True, text=True)
                if result.returncode == 0:
                    path = result.stdout.strip()
        
        elif app_name == "ghostscript":
            if os.name == 'nt':
                patterns = [
                    r"C:\Program Files\gs\gs*\bin\gswin64c.exe",
                    r"C:\Program Files (x86)\gs\gs*\bin\gswin32c.exe",
                ]
                
                for pattern in patterns:
                    for p in glob.glob(pattern):
                        if os.path.exists(p):
                            path = p
                            break
                    if path:
                        break
            else:
                result = subprocess.run(['which', 'gs'], capture_output=True, text=True)
                if result.returncode == 0:
                    path = result.stdout.strip()
        
        elif app_name == "poppler":
            # Poppler ist normalerweise im Programmverzeichnis
            app_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
            poppler_dir = os.path.join(app_dir, 'poppler', 'bin')
            if os.path.exists(poppler_dir):
                path = poppler_dir
        
        if path:
            getattr(self, f"{app_name}_path_var").set(path)
            messagebox.showinfo("Auto-Erkennung", f"{app_name.capitalize()} gefunden:\n{path}")
            self._test_single_path(app_name)
        else:
            messagebox.showwarning("Auto-Erkennung", f"{app_name.capitalize()} konnte nicht automatisch gefunden werden.")
    
    def _test_single_path(self, app_name: str):
        """Testet einen einzelnen Anwendungspfad"""
        path = getattr(self, f"{app_name}_path_var").get()
        status_label = getattr(self, f"{app_name}_status_label")
        
        if not path:
            status_label.config(text="ℹ Kein Pfad konfiguriert (Auto-Erkennung)", foreground="blue")
            return
        
        if app_name == "tesseract":
            if os.path.exists(path):
                try:
                    result = subprocess.run([path, "--version"], capture_output=True, text=True)
                    if result.returncode == 0:
                        version = result.stdout.split('\n')[0]
                        status_label.config(text=f"✓ {version}", foreground="green")
                    else:
                        status_label.config(text="✗ Pfad ungültig", foreground="red")
                except Exception as e:
                    status_label.config(text=f"✗ Fehler: {str(e)}", foreground="red")
            else:
                status_label.config(text="✗ Pfad existiert nicht", foreground="red")
        
        elif app_name == "ghostscript":
            if os.path.exists(path):
                try:
                    result = subprocess.run([path, "--version"], capture_output=True, text=True)
                    if result.returncode == 0:
                        version = result.stdout.strip()
                        status_label.config(text=f"✓ Version {version}", foreground="green")
                    else:
                        status_label.config(text="✗ Pfad ungültig", foreground="red")
                except Exception as e:
                    status_label.config(text=f"✗ Fehler: {str(e)}", foreground="red")
            else:
                status_label.config(text="✗ Pfad existiert nicht", foreground="red")
        
        elif app_name == "poppler":
            if os.path.exists(path):
                # Prüfe ob wichtige Dateien vorhanden sind
                pdftoppm = os.path.join(path, 'pdftoppm.exe' if os.name == 'nt' else 'pdftoppm')
                if os.path.exists(pdftoppm):
                    status_label.config(text="✓ Pfad gültig", foreground="green")
                else:
                    status_label.config(text="✗ Keine Poppler-Tools gefunden", foreground="red")
            else:
                status_label.config(text="✗ Pfad existiert nicht", foreground="red")
    
    def _test_all_paths(self):
        """Testet alle konfigurierten Pfade"""
        self._test_single_path('tesseract')
        self._test_single_path('ghostscript')
        self._test_single_path('poppler')
    
    def _on_auth_method_changed(self):
        """Wird aufgerufen wenn die Auth-Methode geändert wird"""
        if self.auth_method_var.get() == AuthMethod.BASIC.value:
            self.smtp_auth_frame.pack(fill=tk.X, pady=(0, 10), before=self.smtp_sender_frame)
            self.oauth2_frame.pack_forget()
        else:
            self.smtp_auth_frame.pack_forget()
            self.oauth2_frame.pack(fill=tk.X, pady=(0, 10), before=self.smtp_sender_frame)
            
            # Layout für OAuth2
            self.oauth2_provider_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
            self.oauth2_provider_combo.grid(row=0, column=1, sticky=tk.W)
            self.oauth2_setup_button.grid(row=1, column=0, columnspan=2, pady=(10, 5))
            self.oauth2_status_label.grid(row=2, column=0, columnspan=2, pady=(5, 10))
            
            self.oauth2_info_frame.grid(row=3, column=0, columnspan=2, sticky="w")
            self.oauth2_email_label.pack(anchor=tk.W)
            self.oauth2_client_label.pack(anchor=tk.W, pady=(2, 0))
            
            self.oauth2_frame.columnconfigure(1, weight=1)
    
    def _on_ssl_changed(self):
        """Wird aufgerufen wenn SSL-Option geändert wird"""
        if self.smtp_ssl_var.get():
            # SSL aktiviert, deaktiviere TLS
            self.smtp_tls_var.set(False)
            # Setze Port auf 465 wenn noch nicht gesetzt
            if self.smtp_port_var.get() != 465:
                self.smtp_port_var.set(465)
    
    def _on_tls_changed(self):
        """Wird aufgerufen wenn TLS-Option geändert wird"""
        if self.smtp_tls_var.get():
            # TLS aktiviert, deaktiviere SSL
            self.smtp_ssl_var.set(False)
            # Setze Port auf 587 wenn noch nicht gesetzt
            if self.smtp_port_var.get() != 587:
                self.smtp_port_var.set(587)
    
    def _setup_oauth2(self):
        """Öffnet OAuth2-Setup-Dialog"""
        provider = self.oauth2_provider_var.get()
        if not provider:
            messagebox.showerror("Fehler", "Bitte wählen Sie einen OAuth2-Anbieter aus.")
            return
        
        # Sammle aktuelle OAuth2-Settings
        current_oauth2_settings = {
            'oauth2_provider': provider.lower(),  # Konvertiere zu Kleinbuchstaben
            'oauth2_client_id': self.settings.oauth2_client_id,
            'oauth2_client_secret': self.settings.oauth2_client_secret,
            'oauth2_refresh_token': self.settings.oauth2_refresh_token,
            'oauth2_access_token': self.settings.oauth2_access_token,
            'oauth2_token_expiry': self.settings.oauth2_token_expiry,
            'smtp_from_address': self.smtp_from_var.get() or self.settings.smtp_from_address,
            'smtp_username': self.smtp_username_var.get() or self.settings.smtp_username
        }
        
        dialog = OAuth2SetupDialog(self.dialog, provider, current_oauth2_settings)
        result = dialog.show()
        
        if result:
            # Update Settings mit OAuth2-Config
            self.settings.oauth2_provider = result['oauth2_provider']
            self.settings.oauth2_client_id = result['oauth2_client_id']
            self.settings.oauth2_client_secret = result['oauth2_client_secret']
            self.settings.oauth2_refresh_token = result['oauth2_refresh_token']
            self.settings.oauth2_access_token = result['oauth2_access_token']
            self.settings.oauth2_token_expiry = result['oauth2_token_expiry']
            
            # Update E-Mail-Adresse
            if result['smtp_from_address']:
                self.smtp_from_var.set(result['smtp_from_address'])
            
            # Update Status
            self._update_oauth2_status()
    
    def _test_email(self):
        """Testet die E-Mail-Einstellungen"""
        # Prüfe ob alle erforderlichen Felder ausgefüllt sind
        if not all([self.smtp_server_var.get(), self.smtp_port_var.get(), 
                   self.smtp_from_var.get()]):
            messagebox.showerror("Fehler", 
                "Bitte füllen Sie mindestens Server, Port und Absender-Adresse aus.")
            return
        
        # Prüfe Auth-spezifische Felder
        if self.auth_method_var.get() == AuthMethod.BASIC.value:
            if not self.smtp_username_var.get() or not self.smtp_password_var.get():
                messagebox.showerror("Fehler", 
                    "Für Standard-Authentifizierung müssen Benutzername und Passwort ausgefüllt sein.")
                return
        else:  # OAuth2
            if not self.settings.oauth2_refresh_token:
                messagebox.showerror("Fehler", 
                    "Bitte richten Sie zuerst OAuth2 ein.")
                return
        
        # Test-Dialog
        config = {
            'server': self.smtp_server_var.get(),
            'port': self.smtp_port_var.get(),
            'use_ssl': self.smtp_ssl_var.get(),
            'use_tls': self.smtp_tls_var.get(),
            'username': self.smtp_username_var.get(),
            'password': self.smtp_password_var.get(),
            'from_address': self.smtp_from_var.get(),
            'auth_method': self.auth_method_var.get()
        }
        
        # Füge OAuth2-Config hinzu wenn OAuth2
        if self.auth_method_var.get() == AuthMethod.OAUTH2.value:
            config.update({
                'oauth2_provider': self.settings.oauth2_provider,
                'oauth2_access_token': self.settings.oauth2_access_token,
                'oauth2_refresh_token': self.settings.oauth2_refresh_token,
                'oauth2_token_expiry': self.settings.oauth2_token_expiry,
                'oauth2_client_id': self.settings.oauth2_client_id,
                'oauth2_client_secret': self.settings.oauth2_client_secret
            })
        
        dialog = EmailTestDialog(self.dialog, config)
        dialog.show()
    
    def _validate(self) -> bool:
        """Validiert die Eingaben"""
        # Fehlerordner ist Pflichtfeld
        if not self.error_path_var.get().strip():
            messagebox.showerror("Fehler", 
                "Der Standard-Fehlerpfad ist ein Pflichtfeld.\n"
                "Bitte wählen Sie einen Ordner aus.")
            return False
        
        # Prüfe ob Fehlerordner existiert oder erstellt werden kann
        error_path = self.error_path_var.get().strip()
        try:
            os.makedirs(error_path, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Fehler", 
                f"Der Fehlerpfad konnte nicht erstellt werden:\n{e}")
            return False
        
        return True
    
    def _on_save(self):
        """Speichert die Einstellungen"""
        if not self._validate():
            return
        
        # Aktualisiere Settings-Objekt
        self.settings.default_error_path = self.error_path_var.get().strip()
        self.settings.smtp_server = self.smtp_server_var.get()
        self.settings.smtp_port = self.smtp_port_var.get()
        self.settings.smtp_use_ssl = self.smtp_ssl_var.get()
        self.settings.smtp_use_tls = self.smtp_tls_var.get()
        self.settings.smtp_username = self.smtp_username_var.get()
        self.settings.smtp_password = self.smtp_password_var.get()
        self.settings.smtp_from_address = self.smtp_from_var.get()
        
        # Auth-Methode
        self.settings.smtp_auth_method = AuthMethod(self.auth_method_var.get())
        
        # OAuth2-Provider nur updaten wenn OAuth2 ausgewählt
        if self.auth_method_var.get() == AuthMethod.OAUTH2.value:
            self.settings.oauth2_provider = self.oauth2_provider_var.get().lower()
        
        # Anwendungspfade aktualisieren
        if not hasattr(self.settings, 'application_paths'):
            self.settings.application_paths = ApplicationPaths()
        
        self.settings.application_paths.tesseract = self.tesseract_path_var.get()
        self.settings.application_paths.ghostscript = self.ghostscript_path_var.get()
        self.settings.application_paths.poppler = self.poppler_path_var.get()
        
        # Speichere in Datei
        self._save_settings()
        
        self.result = True
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab ohne zu speichern"""
        self.dialog.destroy()
    
    def show(self) -> Optional[bool]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result


class EmailTestDialog:
    """Dialog zum Testen der E-Mail-Einstellungen"""
    
    def __init__(self, parent, smtp_config: dict):
        self.parent = parent
        self.smtp_config = smtp_config
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("E-Mail-Test")
        self.dialog.geometry("500x300")
        self.dialog.resizable(False, False)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._create_widgets()
        self._layout_widgets()
        
        # Fokus
        self.recipient_entry.focus()
    
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
        self.main_frame = ttk.Frame(self.dialog, padding="20")
        
        self.info_label = ttk.Label(self.main_frame, 
            text="Geben Sie eine E-Mail-Adresse ein, an die eine Test-Nachricht gesendet werden soll:")
        
        self.recipient_label = ttk.Label(self.main_frame, text="Empfänger:")
        self.recipient_var = tk.StringVar()
        self.recipient_entry = ttk.Entry(self.main_frame, textvariable=self.recipient_var, 
                                        width=40)
        
        self.progress_label = ttk.Label(self.main_frame, text="")
        self.progress_bar = ttk.Progressbar(self.main_frame, mode='indeterminate')
        
        self.button_frame = ttk.Frame(self.main_frame)
        self.send_button = ttk.Button(self.button_frame, text="Test senden", 
                                     command=self._send_test)
        self.close_button = ttk.Button(self.button_frame, text="Schließen", 
                                      command=self.dialog.destroy)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.info_label.pack(anchor=tk.W, pady=(0, 20))
        
        self.recipient_label.pack(anchor=tk.W)
        self.recipient_entry.pack(fill=tk.X, pady=(5, 20))
        
        self.progress_label.pack(anchor=tk.W, pady=(10, 5))
        self.progress_bar.pack(fill=tk.X, pady=(0, 20))
        
        self.button_frame.pack()
        self.send_button.pack(side=tk.LEFT, padx=(0, 10))
        self.close_button.pack(side=tk.LEFT)
    
    def _send_test(self):
        """Sendet die Test-E-Mail"""
        recipient = self.recipient_var.get().strip()
        if not recipient:
            messagebox.showerror("Fehler", "Bitte geben Sie eine Empfänger-Adresse ein.")
            return
        
        # UI-Update
        self.send_button.config(state=tk.DISABLED)
        self.progress_label.config(text="Sende Test-E-Mail...")
        self.progress_bar.start()
        
        # Sende E-Mail in Thread
        import threading
        thread = threading.Thread(target=self._send_email_thread, args=(recipient,))
        thread.daemon = True
        thread.start()
    
    def _send_email_thread(self, recipient: str):
        """Sendet die E-Mail in einem separaten Thread"""
        try:
            import smtplib
            import ssl
            from email.mime.text import MIMEText
            from email.mime.multipart import MIMEMultipart
            
            # OAuth2-Unterstützung
            if self.smtp_config.get('auth_method') == 'oauth2':
                # OAuth2-Authentifizierung
                from core.oauth2_manager import OAuth2Manager, get_token_storage
                
                provider = self.smtp_config.get('oauth2_provider', 'gmail')
                oauth2_manager = OAuth2Manager(provider)
                token_storage = get_token_storage()
                
                # Hole gespeicherte Tokens
                tokens = token_storage.get_tokens(provider, self.smtp_config['from_address'])
                if not tokens:
                    tokens = {
                        'access_token': self.smtp_config.get('oauth2_access_token', ''),
                        'refresh_token': self.smtp_config.get('oauth2_refresh_token', ''),
                        'token_expiry': self.smtp_config.get('oauth2_token_expiry', '')
                    }
                
                # Prüfe ob Token erneuert werden muss
                if oauth2_manager.is_token_expired(tokens.get('token_expiry', '')):
                    # Token erneuern
                    oauth2_manager.set_client_credentials(
                        self.smtp_config.get('oauth2_client_id', ''),
                        self.smtp_config.get('oauth2_client_secret', '')
                    )
                    
                    success, new_tokens = oauth2_manager.refresh_access_token(
                        tokens.get('refresh_token', '')
                    )
                    
                    if success:
                        tokens = new_tokens
                        # Speichere neue Tokens
                        token_storage.set_tokens(provider, self.smtp_config['from_address'], tokens)
                    else:
                        raise Exception(f"Token-Erneuerung fehlgeschlagen: {new_tokens.get('error', 'Unbekannter Fehler')}")
                
                access_token = tokens.get('access_token', '')
                if not access_token:
                    raise Exception("Kein gültiger Access Token vorhanden")
            
            # Erstelle Nachricht
            msg = MIMEMultipart()
            msg['From'] = self.smtp_config['from_address']
            msg['To'] = recipient
            msg['Subject'] = "Hotfolder PDF Processor - Test-E-Mail"
            
            auth_method = "OAuth2" if self.smtp_config.get('auth_method') == 'oauth2' else "Standard"
            
            body = f"""Dies ist eine Test-E-Mail vom Hotfolder PDF Processor.

Wenn Sie diese Nachricht erhalten, sind Ihre E-Mail-Einstellungen korrekt konfiguriert!

Server: {self.smtp_config['server']}
Port: {self.smtp_config['port']}
SSL: {"Ja" if self.smtp_config['use_ssl'] else "Nein"}
TLS: {"Ja" if self.smtp_config['use_tls'] else "Nein"}
Authentifizierung: {auth_method}
Absender: {self.smtp_config['from_address']}

Mit freundlichen Grüßen
Hotfolder PDF Processor"""
            
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # Verbinde zum Server
            if self.smtp_config['use_ssl'] and self.smtp_config['port'] == 465:
                # SSL direkt verwenden
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL(self.smtp_config['server'], 
                                         self.smtp_config['port'], 
                                         context=context)
            else:
                # Standard SMTP mit optionalem STARTTLS
                server = smtplib.SMTP(self.smtp_config['server'], self.smtp_config['port'])
                if self.smtp_config['use_tls']:
                    server.starttls()
            
            # Anmeldung
            if self.smtp_config.get('auth_method') == 'oauth2':
                # OAuth2-Anmeldung
                auth_string = oauth2_manager.create_oauth2_sasl_string(
                    self.smtp_config['from_address'],
                    access_token
                )
                server.docmd('AUTH', 'XOAUTH2 ' + auth_string)
            else:
                # Standard-Anmeldung
                if self.smtp_config['username'] and self.smtp_config['password']:
                    server.login(self.smtp_config['username'], self.smtp_config['password'])
            
            # Sende E-Mail
            server.send_message(msg)
            server.quit()
            
            # Erfolg
            self.dialog.after(0, self._on_success)
            
        except Exception as e:
            # Fehler
            self.dialog.after(0, self._on_error, str(e))
    
    def _on_success(self):
        """Wird bei erfolgreichem Versand aufgerufen"""
        self.progress_bar.stop()
        self.progress_label.config(text="Test-E-Mail erfolgreich gesendet!")
        self.send_button.config(state=tk.NORMAL)
        messagebox.showinfo("Erfolg", "Die Test-E-Mail wurde erfolgreich gesendet!")
    
    def _on_error(self, error_msg: str):
        """Wird bei Fehler aufgerufen"""
        self.progress_bar.stop()
        self.progress_label.config(text="Fehler beim Senden!")
        self.send_button.config(state=tk.NORMAL)
        messagebox.showerror("Fehler", f"Fehler beim Senden der Test-E-Mail:\n\n{error_msg}")
    
    def show(self):
        """Zeigt den Dialog"""
        self.dialog.wait_window()