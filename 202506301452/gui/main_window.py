"""
Hauptfenster der Hotfolder-Anwendung
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import sys
import os
from typing import Optional

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gui.hotfolder_dialog import HotfolderDialog
from core.hotfolder_manager import HotfolderManager
from models.hotfolder_config import ProcessingAction


class MainWindow:
    """Hauptfenster der Anwendung"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Hotfolder PDF Processor")
        self.root.geometry("900x600")
        
        # Manager
        self.manager = HotfolderManager()
        
        # Style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Erstelle UI
        self._create_menu()
        self._create_widgets()
        self._layout_widgets()
        
        # Lade Hotfolder
        self._refresh_list()
        
        # Starte Manager
        self.manager.start()
        
        # Bind Events
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # Statusbar Update
        self._update_status()
    
    def _create_menu(self):
        """Erstellt die Menüleiste"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Datei-Menü
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Datei", menu=file_menu)
        file_menu.add_command(label="Neuer Hotfolder", command=self._new_hotfolder,
                             accelerator="Ctrl+N")
        file_menu.add_separator()
        file_menu.add_command(label="Hotfolder importieren...", command=self._import_hotfolder,
                             accelerator="Ctrl+I")
        file_menu.add_command(label="Hotfolder exportieren...", command=self._export_hotfolder,
                             accelerator="Ctrl+E")
        file_menu.add_separator()
        file_menu.add_command(label="Beenden", command=self._on_closing,
                             accelerator="Alt+F4")
        
        # Bearbeiten-Menü
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Bearbeiten", menu=edit_menu)
        edit_menu.add_command(label="Hotfolder bearbeiten", command=self._edit_hotfolder,
                             accelerator="F2")
        edit_menu.add_command(label="Hotfolder löschen", command=self._delete_hotfolder,
                             accelerator="Del")
        edit_menu.add_separator()
        edit_menu.add_command(label="Hotfolder Aktivieren/Deaktivieren", command=self._toggle_hotfolder,
                             accelerator="Leertaste")
        edit_menu.add_separator()
        edit_menu.add_command(label="Einstellungen", command=self._show_settings)
        
        # Extras-Menü
        extras_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Extras", menu=extras_menu)
        extras_menu.add_command(label="Counter-Management", command=self._manage_counters)
        
        # Hilfe-Menü
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Hilfe", menu=help_menu)
        help_menu.add_command(label="Über", command=self._show_about)
        
        # Keyboard Shortcuts
        self.root.bind("<Control-n>", lambda e: self._new_hotfolder())
        self.root.bind("<Control-i>", lambda e: self._import_hotfolder())
        self.root.bind("<Control-e>", lambda e: self._export_hotfolder())
        self.root.bind("<F2>", lambda e: self._edit_hotfolder())
        self.root.bind("<Delete>", lambda e: self._delete_hotfolder())
        self.root.bind("<space>", lambda e: self._toggle_hotfolder())
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Toolbar
        self.toolbar = ttk.Frame(self.root)
        self.new_button = ttk.Button(self.toolbar, text="➕ Neuer Hotfolder", 
                                    command=self._new_hotfolder)
        self.edit_button = ttk.Button(self.toolbar, text="✏️ Bearbeiten", 
                                     command=self._edit_hotfolder, state=tk.DISABLED)
        self.delete_button = ttk.Button(self.toolbar, text="🗑️ Löschen", 
                                       command=self._delete_hotfolder, state=tk.DISABLED)
        
        # Hauptbereich
        self.main_frame = ttk.Frame(self.root)
        
        # Treeview für Hotfolder-Liste - GEÄNDERT: Neue Spaltenstruktur
        self.tree_frame = ttk.Frame(self.main_frame)
        self.tree = ttk.Treeview(self.tree_frame, columns=("Aktiv", "Name", "Beschreibung"),
                                show="headings")
        
        # Spalten konfigurieren - GEÄNDERT: Anpassung an hotfolder_dialog Style
        self.tree.heading("Aktiv", text="Aktiv")
        self.tree.heading("Name", text="Name")
        self.tree.heading("Beschreibung", text="Beschreibung")
        
        self.tree.column("Aktiv", width=50, anchor=tk.CENTER)  # Gleiche Breite wie im Dialog
        self.tree.column("Name", width=200)
        self.tree.column("Beschreibung", width=500)
        
        # Scrollbars
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        
        # Kontextmenü
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Bearbeiten", command=self._edit_hotfolder)
        self.context_menu.add_command(label="Löschen", command=self._delete_hotfolder)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Aktivieren/Deaktivieren", 
                                     command=self._toggle_hotfolder)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Exportieren", command=self._export_hotfolder)
        
        # Statusbar
        self.statusbar = ttk.Frame(self.root)
        self.status_label = ttk.Label(self.statusbar, text="Bereit")
        
        # Events
        self.tree.bind("<Double-Button-1>", lambda e: self._edit_hotfolder())
        self.tree.bind("<Button-3>", self._show_context_menu)
        self.tree.bind("<<TreeviewSelect>>", self._on_selection_changed)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        # Toolbar
        self.toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        self.new_button.pack(side=tk.LEFT, padx=(0, 5))
        self.edit_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_button.pack(side=tk.LEFT)
        
        # Hauptbereich
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))
        
        # Treeview
        self.tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")
        
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)
        
        # Statusbar
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label.pack(side=tk.LEFT, padx=5)
    
    def _refresh_list(self):
        """Aktualisiert die Hotfolder-Liste"""
        # Lösche aktuelle Einträge
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Füge Hotfolder hinzu
        for hotfolder in self.manager.get_hotfolders():
            # GEÄNDERT: Gleiche Symbole wie im hotfolder_dialog
            aktiv_symbol = "✓" if hotfolder.enabled else "✗"
            
            # Beschreibung
            description = hotfolder.description if hasattr(hotfolder, 'description') else ""
            if not description:
                # Generiere eine Standardbeschreibung
                export_count = len(hotfolder.export_configs) if hasattr(hotfolder, 'export_configs') else 0
                if export_count > 0:
                    description = f"{export_count} Export(e) konfiguriert"
                else:
                    description = "Keine Exporte konfiguriert"
            
            # GEÄNDERT: Neue Spaltenreihenfolge mit Aktiv als erste Spalte
            item = self.tree.insert("", "end", values=(aktiv_symbol, hotfolder.name, description),
                                   tags=(hotfolder.id,))
            
            # Färbe inaktive Einträge grau
            if not hotfolder.enabled:
                self.tree.item(item, tags=(hotfolder.id, "disabled"))
        
        # Style für disabled
        self.tree.tag_configure("disabled", foreground="gray")
    
    def _on_selection_changed(self, event):
        """Wird aufgerufen wenn die Auswahl sich ändert"""
        selection = self.tree.selection()
        
        if selection:
            self.edit_button.config(state=tk.NORMAL)
            self.delete_button.config(state=tk.NORMAL)
        else:
            self.edit_button.config(state=tk.DISABLED)
            self.delete_button.config(state=tk.DISABLED)
    
    def _get_selected_hotfolder_id(self) -> Optional[str]:
        """Gibt die ID des ausgewählten Hotfolders zurück"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            tags = self.tree.item(item)["tags"]
            if tags:
                return tags[0]  # Erste Tag ist die ID
        return None
    
    def _new_hotfolder(self):
        """Erstellt einen neuen Hotfolder"""
        dialog = HotfolderDialog(self.root)
        result = dialog.show()
        
        if result:
            success, message = self.manager.create_hotfolder(
                name=str(result["name"] or ""),
                input_path=str(result["input_path"] or ""),
                description=str(result.get("description", "") or ""),
                actions=result["actions"],
                action_params=result["action_params"],
                xml_field_mappings=result.get("xml_field_mappings", []),
                output_filename_expression=str(result.get("output_filename_expression", "<FileName>") or "<FileName>"),
                ocr_zones=result.get("ocr_zones", []),
                export_configs=result.get("export_configs", []),
                error_path=str(result.get("error_path", "") or "")
            )
            
            if success:
                self._refresh_list()
                self._update_status("Hotfolder erstellt")
            else:
                messagebox.showerror("Fehler", message)
    
    def _edit_hotfolder(self):
        """Bearbeitet den ausgewählten Hotfolder"""
        hotfolder_id = self._get_selected_hotfolder_id()
        if not hotfolder_id:
            return
        
        hotfolder = self.manager.get_hotfolder(hotfolder_id)
        if not hotfolder:
            return
        
        dialog = HotfolderDialog(self.root, hotfolder)
        result = dialog.show()
        
        if result:
            # Fallback für hotfolder_id falls None
            if not hotfolder_id:
                hotfolder_id = ""
            success, message = self.manager.update_hotfolder(
                hotfolder_id,
                name=str(result["name"] or ""),
                input_path=str(result["input_path"] or ""),
                description=str(result.get("description", "") or ""),
                process_pairs=result["process_pairs"],
                actions=result["actions"],
                action_params=result["action_params"],
                xml_field_mappings=result.get("xml_field_mappings", []),
                output_filename_expression=str(result.get("output_filename_expression", "<FileName>") or "<FileName>"),
                ocr_zones=result.get("ocr_zones", []),
                export_configs=result.get("export_configs", []),
                error_path=str(result.get("error_path", "") or "")
            )
            
            if success:
                self._refresh_list()
                self._update_status("Hotfolder aktualisiert")
            else:
                messagebox.showerror("Fehler", message)
    
    def _delete_hotfolder(self):
        """Löscht den ausgewählten Hotfolder"""
        hotfolder_id = self._get_selected_hotfolder_id()
        if not hotfolder_id:
            return
        
        hotfolder = self.manager.get_hotfolder(hotfolder_id)
        if not hotfolder:
            return
        
        # Bestätigung
        result = messagebox.askyesno(
            "Hotfolder löschen",
            f"Möchten Sie den Hotfolder '{hotfolder.name}' wirklich löschen?"
        )
        
        if result:
            success, message = self.manager.delete_hotfolder(hotfolder_id)
            
            if success:
                self._refresh_list()
                self._update_status("Hotfolder gelöscht")
            else:
                messagebox.showerror("Fehler", message)
    
    def _toggle_hotfolder(self):
        """Aktiviert/Deaktiviert den ausgewählten Hotfolder"""
        hotfolder_id = self._get_selected_hotfolder_id()
        if not hotfolder_id:
            return
        
        hotfolder = self.manager.get_hotfolder(hotfolder_id)
        if not hotfolder:
            return
        
        new_state = not hotfolder.enabled
        
        # Bei Aktivierung: Prüfe auf doppelte Input-Pfade
        if new_state:  # Nur beim Aktivieren prüfen
            duplicate_name = self.manager.config_manager.check_duplicate_input_path(
                hotfolder.input_path, 
                exclude_id=hotfolder_id
            )
            if duplicate_name:
                messagebox.showerror(
                    "Doppelter Input-Ordner",
                    f"Der Input-Ordner wird bereits vom Hotfolder '{duplicate_name}' verwendet.\n\n"
                    f"Bitte ändern Sie zuerst den Input-Ordner dieses Hotfolders."
                )
                return
        
        success, message = self.manager.toggle_hotfolder(hotfolder_id, new_state)
        
        if success:
            self._refresh_list()
            self._update_status(message)
        else:
            messagebox.showerror("Fehler", message)
    
    def _import_hotfolder(self):
        """Importiert einen Hotfolder aus einer JSON-Datei"""
        filename = filedialog.askopenfilename(
            parent=self.root,
            title="Hotfolder importieren",
            filetypes=[
                ("JSON-Dateien", "*.json"),
                ("Alle Dateien", "*.*")
            ]
        )
        
        if filename:
            # Importiere mit neuer ID um Duplikate zu vermeiden
            success, message = self.manager.config_manager.import_hotfolder(filename, generate_new_id=True)
            
            if success:
                self._refresh_list()
                self._update_status(message)
                messagebox.showinfo("Import erfolgreich", 
                    message + "\n\nDer Hotfolder wurde deaktiviert importiert. "
                    "Sie können ihn bearbeiten und anschließend aktivieren.")
            else:
                messagebox.showerror("Fehler beim Import", message)
    
    def _export_hotfolder(self):
        """Exportiert den ausgewählten Hotfolder"""
        hotfolder_id = self._get_selected_hotfolder_id()
        if not hotfolder_id:
            messagebox.showwarning("Kein Hotfolder ausgewählt", 
                                  "Bitte wählen Sie einen Hotfolder aus, den Sie exportieren möchten.")
            return
        
        hotfolder = self.manager.get_hotfolder(hotfolder_id)
        if not hotfolder:
            return
        
        # Vorgeschlagener Dateiname
        default_filename = f"{hotfolder.name.replace(' ', '_')}_hotfolder.json"
        
        filename = filedialog.asksaveasfilename(
            parent=self.root,
            title="Hotfolder exportieren",
            defaultextension=".json",
            initialfile=default_filename,
            filetypes=[
                ("JSON-Dateien", "*.json"),
                ("Alle Dateien", "*.*")
            ]
        )
        
        if filename:
            success, message = self.manager.config_manager.export_hotfolder(hotfolder_id, filename)
            
            if success:
                self._update_status(message)
                messagebox.showinfo("Export erfolgreich", message)
            else:
                messagebox.showerror("Fehler beim Export", message)
    
    def _manage_counters(self):
        """Öffnet das Counter-Management-Tool"""
        try:
            from gui.counter_management_dialog import CounterManagementDialog
            dialog = CounterManagementDialog(self.root)
            dialog.show()
        except ImportError as e:
            messagebox.showerror("Fehler", f"Counter-Management nicht verfügbar: {e}")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Öffnen des Counter-Managements: {e}")
    
    def _show_settings(self):
        """Öffnet den Einstellungen-Dialog"""
        try:
            from gui.settings_dialog import SettingsDialog
            dialog = SettingsDialog(self.root)
            result = dialog.show()
            
            if result:
                self._update_status("Einstellungen gespeichert")
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Öffnen der Einstellungen: {e}")
    
    def _show_context_menu(self, event):
        """Zeigt das Kontextmenü"""
        # Wähle Item unter Mauszeiger
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def _show_about(self):
        """Zeigt den Über-Dialog"""
        messagebox.showinfo(
            "Über Hotfolder PDF Processor",
            "Hotfolder PDF Processor v1.0\n\n"
            "Ein skalierbares Tool zur automatischen PDF-Verarbeitung.\n\n"
            "Unterstützt:\n"
            "• Mehrere Hotfolder\n"
            "• Verschiedene PDF-Bearbeitungen\n"
            "• PDF-XML Dokumentenpaare\n"
            "• Automatische Ordnerüberwachung\n"
            "• Auto-Increment Counter\n"
            "• Erweiterte Funktions-Sprache\n"
            "• Multiple Export-Formate und -Methoden\n"
            "• Import/Export von Hotfolder-Konfigurationen"
        )
    
    def _update_status(self, message: str = None):
        """Aktualisiert die Statusleiste"""
        if message:
            self.status_label.config(text=message)
        else:
            active_count = len([h for h in self.manager.get_hotfolders() if h.enabled])
            total_count = len(self.manager.get_hotfolders())
            self.status_label.config(text=f"{active_count} von {total_count} Hotfoldern aktiv")
        
        # Reset nach 3 Sekunden
        if message:
            self.root.after(3000, lambda: self._update_status())
    
    def _on_closing(self):
        """Wird beim Schließen des Fensters aufgerufen"""
        # Stoppe Manager
        self.manager.stop()
        
        # Schließe Fenster
        self.root.destroy()
    
    def run(self):
        """Startet die Anwendung"""
        self.root.mainloop()