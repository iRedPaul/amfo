"""
Dialog für Datenbank-Konfigurationen
"""
import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.database_processor import DatabaseProcessor, DatabaseConfig


class DatabaseConfigDialog:
    """Dialog zum Verwalten von Datenbank-Verbindungen"""
    
    def __init__(self, parent):
        self.parent = parent
        self.db_processor = DatabaseProcessor()
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Datenbank-Verbindungen")
        self.dialog.geometry("800x600")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() - 800) // 2
        y = (self.dialog.winfo_screenheight() - 600) // 2
        self.dialog.geometry(f"800x600+{x}+{y}")
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._create_widgets()
        self._load_configs()
    
    def _create_widgets(self):
        """Erstellt die GUI-Elemente"""
        # Hauptframe
        main_frame = ttk.Frame(self.dialog, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Liste der Verbindungen
        list_frame = ttk.LabelFrame(main_frame, text="Konfigurierte Verbindungen", padding="10")
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Treeview für Verbindungen
        columns = ("driver", "server", "database")
        self.tree = ttk.Treeview(list_frame, columns=columns, height=10)
        self.tree.heading("#0", text="Name")
        self.tree.heading("driver", text="Treiber")
        self.tree.heading("server", text="Server")
        self.tree.heading("database", text="Datenbank")
        
        self.tree.column("#0", width=150)
        self.tree.column("driver", width=200)
        self.tree.column("server", width=150)
        self.tree.column("database", width=150)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(list_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Buttons für Verbindungen
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Button(button_frame, text="Neue Verbindung", 
                  command=self._new_connection).pack(side=tk.LEFT, padx=(0, 5))
        ttk.Button(button_frame, text="Bearbeiten", 
                  command=self._edit_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Löschen", 
                  command=self._delete_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Verbindung testen", 
                  command=self._test_connection).pack(side=tk.LEFT, padx=5)
        
        # Dialog-Buttons
        dialog_button_frame = ttk.Frame(main_frame)
        dialog_button_frame.pack(fill=tk.X, pady=(20, 0))
        
        ttk.Button(dialog_button_frame, text="Schließen", 
                  command=self.dialog.destroy).pack(side=tk.RIGHT)
    
    def _load_configs(self):
        """Lädt die Konfigurationen in die Liste"""
        # Lösche bestehende Einträge
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Lade Konfigurationen
        for name in self.db_processor.list_configs():
            config = self.db_processor.get_config(name)
            if config:
                self.tree.insert("", tk.END, text=config.name,
                               values=(config.driver, config.server, config.database))
    
    def _new_connection(self):
        """Erstellt eine neue Verbindung"""
        dialog = ConnectionEditDialog(self.dialog, self.db_processor)
        if dialog.show():
            self._load_configs()
    
    def _edit_connection(self):
        """Bearbeitet die ausgewählte Verbindung"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warnung", "Bitte wählen Sie eine Verbindung aus.")
            return
        
        item = self.tree.item(selection[0])
        config_name = item['text']
        config = self.db_processor.get_config(config_name)
        
        if config:
            dialog = ConnectionEditDialog(self.dialog, self.db_processor, config)
            if dialog.show():
                self._load_configs()
    
    def _delete_connection(self):
        """Löscht die ausgewählte Verbindung"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warnung", "Bitte wählen Sie eine Verbindung aus.")
            return
        
        item = self.tree.item(selection[0])
        config_name = item['text']
        
        if messagebox.askyesno("Löschen bestätigen", 
                              f"Möchten Sie die Verbindung '{config_name}' wirklich löschen?"):
            self.db_processor.delete_config(config_name)
            self._load_configs()
    
    def _test_connection(self):
        """Testet die ausgewählte Verbindung"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Warnung", "Bitte wählen Sie eine Verbindung aus.")
            return
        
        item = self.tree.item(selection[0])
        config_name = item['text']
        config = self.db_processor.get_config(config_name)
        
        if config:
            success, message = self.db_processor.test_connection(config)
            if success:
                messagebox.showinfo("Verbindungstest", message)
            else:
                messagebox.showerror("Verbindungstest", f"Verbindung fehlgeschlagen!\n\n{message}")


class ConnectionEditDialog:
    """Dialog zum Bearbeiten einer einzelnen Verbindung"""
    
    def __init__(self, parent, db_processor: DatabaseProcessor, config: DatabaseConfig = None):
        self.parent = parent
        self.db_processor = db_processor
        self.config = config
        self.result = False
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Verbindung bearbeiten" if config else "Neue Verbindung")
        self.dialog.geometry("500x400")
        self.dialog.resizable(False, False)
        
        # Zentriere Dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() - 500) // 2
        y = (self.dialog.winfo_screenheight() - 400) // 2
        self.dialog.geometry(f"500x400+{x}+{y}")
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._create_widgets()
        
        if config:
            self._load_config()
    
    def _create_widgets(self):
        """Erstellt die GUI-Elemente"""
        # Hauptframe
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Name
        ttk.Label(main_frame, text="Verbindungsname:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_var = tk.StringVar()
        self.name_entry = ttk.Entry(main_frame, textvariable=self.name_var, width=40)
        self.name_entry.grid(row=0, column=1, sticky=tk.W, pady=5)
        
        # Treiber
        ttk.Label(main_frame, text="ODBC-Treiber:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.driver_var = tk.StringVar()
        self.driver_combo = ttk.Combobox(main_frame, textvariable=self.driver_var, width=37)
        self.driver_combo['values'] = self.db_processor.get_available_drivers()
        self.driver_combo.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        # Server
        ttk.Label(main_frame, text="Server:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.server_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.server_var, width=40).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # Datenbank
        ttk.Label(main_frame, text="Datenbank:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.database_var = tk.StringVar()
        ttk.Entry(main_frame, textvariable=self.database_var, width=40).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # Authentifizierung
        auth_frame = ttk.LabelFrame(main_frame, text="Authentifizierung", padding="10")
        auth_frame.grid(row=4, column=0, columnspan=2, sticky=tk.EW, pady=10)
        
        # Windows-Authentifizierung
        self.trusted_var = tk.BooleanVar()
        self.trusted_check = ttk.Checkbutton(auth_frame, text="Windows-Authentifizierung verwenden", 
                                           variable=self.trusted_var, command=self._toggle_auth)
        self.trusted_check.pack(anchor=tk.W)
        
        # Benutzername und Passwort
        cred_frame = ttk.Frame(auth_frame)
        cred_frame.pack(fill=tk.X, pady=(10, 0))
        
        ttk.Label(cred_frame, text="Benutzername:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.username_var = tk.StringVar()
        self.username_entry = ttk.Entry(cred_frame, textvariable=self.username_var, width=30)
        self.username_entry.grid(row=0, column=1, sticky=tk.W, pady=5, padx=(10, 0))
        
        ttk.Label(cred_frame, text="Passwort:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.password_var = tk.StringVar()
        self.password_entry = ttk.Entry(cred_frame, textvariable=self.password_var, width=30, show="*")
        self.password_entry.grid(row=1, column=1, sticky=tk.W, pady=5, padx=(10, 0))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Speichern", command=self._save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Abbrechen", command=self.dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def _toggle_auth(self):
        """Aktiviert/Deaktiviert die Authentifizierungsfelder"""
        if self.trusted_var.get():
            self.username_entry.config(state='disabled')
            self.password_entry.config(state='disabled')
        else:
            self.username_entry.config(state='normal')
            self.password_entry.config(state='normal')
    
    def _load_config(self):
        """Lädt die Konfiguration in die Felder"""
        self.name_var.set(self.config.name)
        self.driver_var.set(self.config.driver)
        self.server_var.set(self.config.server)
        self.database_var.set(self.config.database)
        self.trusted_var.set(self.config.trusted_connection)
        self.username_var.set(self.config.username)
        self.password_var.set(self.config.password)
        self._toggle_auth()
    
    def _save(self):
        """Speichert die Konfiguration"""
        name = self.name_var.get().strip()
        if not name:
            messagebox.showerror("Fehler", "Bitte geben Sie einen Verbindungsnamen ein.")
            return
        
        # Prüfe ob Name bereits existiert (bei neuer Verbindung)
        if not self.config and name in self.db_processor.list_configs():
            messagebox.showerror("Fehler", f"Eine Verbindung mit dem Namen '{name}' existiert bereits.")
            return
        
        # Erstelle neue Konfiguration
        config = DatabaseConfig(
            name=name,
            driver=self.driver_var.get(),
            server=self.server_var.get(),
            database=self.database_var.get(),
            username=self.username_var.get() if not self.trusted_var.get() else "",
            password=self.password_var.get() if not self.trusted_var.get() else "",
            trusted_connection=self.trusted_var.get()
        )
        
        # Speichere Konfiguration
        if self.config:
            self.db_processor.update_config(config)
        else:
            self.db_processor.add_config(config)
        
        self.result = True
        self.dialog.destroy()
    
    def show(self) -> bool:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result


if __name__ == "__main__":
    # Test
    root = tk.Tk()
    root.withdraw()
    
    dialog = DatabaseConfigDialog(root)
    root.wait_window(dialog.dialog)