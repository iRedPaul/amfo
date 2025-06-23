"""
Dialog für globale Anwendungseinstellungen
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional
import sys
import os
import json

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.export_config import ExportSettings


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
        self.dialog.geometry("700x600")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
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
    
    def _load_settings(self) -> ExportSettings:
        """Lädt die Einstellungen aus der Datei"""
        try:
            if os.path.exists(self.settings_file):
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    return ExportSettings.from_dict(data)
        except Exception as e:
            print(f"Fehler beim Laden der Einstellungen: {e}")
        
        return ExportSettings()
    
    def _save_settings(self):
        """Speichert die Einstellungen in die Datei"""
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(self.settings.to_dict(), f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Fehler beim Speichern der Einstellungen: {e}")
    
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
            text="Standard-Fehlerpfad:")
        self.error_path_desc = ttk.Label(self.error_frame_content, 
            text="Dieser Pfad wird verwendet, wenn im Hotfolder kein spezifischer Fehlerpfad definiert ist.",
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
        
        self.smtp_tls_var = tk.BooleanVar()
        self.smtp_tls_check = ttk.Checkbutton(self.smtp_frame, text="TLS verwenden", 
                                             variable=self.smtp_tls_var)
        
        # SMTP-Anmeldung
        self.smtp_auth_frame = ttk.LabelFrame(self.email_frame, text="Anmeldung", padding="10")
        
        self.smtp_username_label = ttk.Label(self.smtp_auth_frame, text="Benutzername:")
        self.smtp_username_var = tk.StringVar()
        self.smtp_username_entry = ttk.Entry(self.smtp_auth_frame, 
                                            textvariable=self.smtp_username_var, width=40)
        
        self.smtp_password_label = ttk.Label(self.smtp_auth_frame, text="Passwort:")
        self.smtp_password_var = tk.StringVar()
        self.smtp_password_entry = ttk.Entry(self.smtp_auth_frame, 
                                            textvariable=self.smtp_password_var, width=40, 
                                            show="*")
        
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
        
        # Vordefinierte Server
        self.email_presets_frame = ttk.LabelFrame(self.email_frame, 
                                                 text="Vordefinierte Server", padding="10")
        self.preset_label = ttk.Label(self.email_presets_frame, 
            text="Wählen Sie einen vordefinierten Server für automatische Konfiguration:")
        
        self.preset_var = tk.StringVar()
        self.preset_combo = ttk.Combobox(self.email_presets_frame, 
                                        textvariable=self.preset_var, 
                                        state="readonly", width=30)
        self.preset_combo['values'] = [
            "-- Auswählen --",
            "Gmail",
            "Outlook.com",
            "Yahoo Mail",
            "GMX",
            "Web.de",
            "1&1 / IONOS",
            "T-Online",
            "Office 365",
            "Eigener Server"
        ]
        self.preset_combo.set("-- Auswählen --")
        self.preset_combo.bind('<<ComboboxSelected>>', self._on_preset_selected)
        
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
        # Vordefinierte Server zuerst
        self.email_presets_frame.pack(fill=tk.X, pady=(0, 10))
        self.preset_label.pack(anchor=tk.W, pady=(0, 5))
        self.preset_combo.pack(anchor=tk.W)
        
        # SMTP-Server
        self.smtp_frame.pack(fill=tk.X, pady=(0, 10))
        self.smtp_server_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.smtp_server_entry.grid(row=0, column=1, sticky="we")
        self.smtp_port_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.smtp_port_spinbox.grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
        self.smtp_tls_check.grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        self.smtp_frame.columnconfigure(1, weight=1)
        
        # SMTP-Anmeldung
        self.smtp_auth_frame.pack(fill=tk.X, pady=(0, 10))
        self.smtp_username_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.smtp_username_entry.grid(row=0, column=1, sticky="we")
        self.smtp_password_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.smtp_password_entry.grid(row=1, column=1, sticky="we", pady=(5, 0))
        self.smtp_auth_frame.columnconfigure(1, weight=1)
        
        # Absender
        self.smtp_sender_frame.pack(fill=tk.X, pady=(0, 10))
        self.smtp_from_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.smtp_from_entry.grid(row=0, column=1, sticky="we")
        self.smtp_sender_frame.columnconfigure(1, weight=1)
        
        # Test-Button
        self.email_test_button.pack(anchor=tk.W)
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.save_button.pack(side=tk.RIGHT)
    
    def _load_values(self):
        """Lädt die aktuellen Einstellungen in die UI"""
        self.error_path_var.set(self.settings.default_error_path)
        self.smtp_server_var.set(self.settings.smtp_server)
        self.smtp_port_var.set(self.settings.smtp_port)
        self.smtp_tls_var.set(self.settings.smtp_use_tls)
        self.smtp_username_var.set(self.settings.smtp_username)
        self.smtp_password_var.set(self.settings.smtp_password)
        self.smtp_from_var.set(self.settings.smtp_from_address)
    
    def _browse_error_path(self):
        """Öffnet Dialog zur Auswahl des Fehler-Pfads"""
        folder = filedialog.askdirectory(
            title="Standard-Fehlerpfad auswählen",
            initialdir=self.error_path_var.get() or os.path.expanduser("~")
        )
        if folder:
            self.error_path_var.set(folder)
    
    def _on_preset_selected(self, event):
        """Wird aufgerufen wenn ein vordefinierter Server ausgewählt wird"""
        preset = self.preset_var.get()
        
        presets = {
            "Gmail": {
                "server": "smtp.gmail.com",
                "port": 587,
                "tls": True,
                "note": "Hinweis: Für Gmail benötigen Sie ein App-spezifisches Passwort!"
            },
            "Outlook.com": {
                "server": "smtp-mail.outlook.com",
                "port": 587,
                "tls": True
            },
            "Yahoo Mail": {
                "server": "smtp.mail.yahoo.com",
                "port": 587,
                "tls": True
            },
            "GMX": {
                "server": "mail.gmx.net",
                "port": 587,
                "tls": True
            },
            "Web.de": {
                "server": "smtp.web.de",
                "port": 587,
                "tls": True
            },
            "1&1 / IONOS": {
                "server": "smtp.1und1.de",
                "port": 587,
                "tls": True
            },
            "T-Online": {
                "server": "securesmtp.t-online.de",
                "port": 465,
                "tls": True
            },
            "Office 365": {
                "server": "smtp.office365.com",
                "port": 587,
                "tls": True
            }
        }
        
        if preset in presets:
            config = presets[preset]
            self.smtp_server_var.set(config["server"])
            self.smtp_port_var.set(config["port"])
            self.smtp_tls_var.set(config["tls"])
            
            if "note" in config:
                messagebox.showinfo("Hinweis", config["note"])
    
    def _test_email(self):
        """Testet die E-Mail-Einstellungen"""
        # Prüfe ob alle erforderlichen Felder ausgefüllt sind
        if not all([self.smtp_server_var.get(), self.smtp_port_var.get(), 
                   self.smtp_from_var.get()]):
            messagebox.showerror("Fehler", 
                "Bitte füllen Sie mindestens Server, Port und Absender-Adresse aus.")
            return
        
        # Test-Dialog
        dialog = EmailTestDialog(self.dialog, {
            'server': self.smtp_server_var.get(),
            'port': self.smtp_port_var.get(),
            'use_tls': self.smtp_tls_var.get(),
            'username': self.smtp_username_var.get(),
            'password': self.smtp_password_var.get(),
            'from_address': self.smtp_from_var.get()
        })
        dialog.show()
    
    def _on_save(self):
        """Speichert die Einstellungen"""
        # Aktualisiere Settings-Objekt
        self.settings.default_error_path = self.error_path_var.get()
        self.settings.smtp_server = self.smtp_server_var.get()
        self.settings.smtp_port = self.smtp_port_var.get()
        self.settings.smtp_use_tls = self.smtp_tls_var.get()
        self.settings.smtp_username = self.smtp_username_var.get()
        self.settings.smtp_password = self.smtp_password_var.get()
        self.settings.smtp_from_address = self.smtp_from_var.get()
        
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
            from email.mime.text import MIMEText
            from email.mime.multipart import MIMEMultipart
            
            # Erstelle Nachricht
            msg = MIMEMultipart()
            msg['From'] = self.smtp_config['from_address']
            msg['To'] = recipient
            msg['Subject'] = "Hotfolder PDF Processor - Test-E-Mail"
            
            body = """Dies ist eine Test-E-Mail vom Hotfolder PDF Processor.

Wenn Sie diese Nachricht erhalten, sind Ihre E-Mail-Einstellungen korrekt konfiguriert!

Server: {}
Port: {}
TLS: {}
Absender: {}

Mit freundlichen Grüßen
Hotfolder PDF Processor""".format(
                self.smtp_config['server'],
                self.smtp_config['port'],
                "Ja" if self.smtp_config['use_tls'] else "Nein",
                self.smtp_config['from_address']
            )
            
            msg.attach(MIMEText(body, 'plain', 'utf-8'))
            
            # Verbinde zum Server
            if self.smtp_config['use_tls']:
                server = smtplib.SMTP(self.smtp_config['server'], self.smtp_config['port'])
                server.starttls()
            else:
                server = smtplib.SMTP(self.smtp_config['server'], self.smtp_config['port'])
            
            # Anmelden wenn Credentials vorhanden
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