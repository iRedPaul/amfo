"""
Dialog zum Erstellen und Bearbeiten von Hotfoldern
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, List, Dict
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.hotfolder_config import HotfolderConfig, ProcessingAction, OCRZone
from gui.xml_field_dialog import XMLFieldDialog
from gui.zone_selector import ZoneSelector
from gui.expression_dialog import ExpressionDialog
from gui.export_list_dialog import ExportListDialog


class HotfolderDialog:
    """Dialog zum Erstellen/Bearbeiten eines Hotfolders"""
    
    def __init__(self, parent, hotfolder: Optional[HotfolderConfig] = None):
        self.parent = parent
        self.hotfolder = hotfolder
        self.result = None
        
        # Erstelle Dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Hotfolder bearbeiten" if hotfolder else "Neuer Hotfolder")
        self.dialog.geometry("700x900")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog relativ zum Parent
        self._center_window()
        
        # Dialog-Eigenschaften
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Variablen
        self.name_var = tk.StringVar(value=hotfolder.name if hotfolder else "")
        self.input_path_var = tk.StringVar(value=hotfolder.input_path if hotfolder else "")
        self.output_path_var = tk.StringVar(value=hotfolder.output_path if hotfolder else "")
        self.error_path_var = tk.StringVar(value=hotfolder.error_path if hotfolder else "")
        self.process_pairs_var = tk.BooleanVar(value=hotfolder.process_pairs if hotfolder else True)
        
        # Action Variablen
        self.action_vars = {}
        for action in ProcessingAction:
            is_selected = hotfolder and action in hotfolder.actions
            self.action_vars[action] = tk.BooleanVar(value=is_selected)
        
        # Action Parameter
        self.action_params = hotfolder.action_params.copy() if hotfolder else {}
        
        # XML-Feld-Mappings
        self.xml_field_mappings = hotfolder.xml_field_mappings.copy() if hotfolder else []
        
        # OCR-Zonen
        self.ocr_zones = []
        if hotfolder and hasattr(hotfolder, 'ocr_zones'):
            self.ocr_zones = [zone.to_dict() for zone in hotfolder.ocr_zones]
        
        # Export-Konfigurationen
        self.exports = []
        if hotfolder and hasattr(hotfolder, 'exports'):
            self.exports = hotfolder.exports.copy()
        
        self._create_widgets()
        self._layout_widgets()
        
        # Initialisiere Button-Texte
        self._update_button_texts()
        
        # Fokus auf erstes Eingabefeld
        self.name_entry.focus()
        
        # Bind Enter-Taste
        self.dialog.bind('<Return>', lambda e: self._on_save())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())
    
    def _center_window(self):
        """Zentriert das Fenster relativ zum Parent"""
        self.dialog.update_idletasks()
        
        # Parent-Geometrie
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        # Dialog-Größe
        dialog_width = self.dialog.winfo_width()
        dialog_height = self.dialog.winfo_height()
        
        # Berechne Position
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        # Stelle sicher, dass Dialog auf dem Bildschirm bleibt
        x = max(0, x)
        y = max(0, y)
        
        self.dialog.geometry(f"+{x}+{y}")
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Hauptframe
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Name
        self.name_label = ttk.Label(self.main_frame, text="Name:")
        self.name_entry = ttk.Entry(self.main_frame, textvariable=self.name_var, width=50)
        
        # Input-Pfad
        self.input_label = ttk.Label(self.main_frame, text="Input-Ordner:")
        self.input_frame = ttk.Frame(self.main_frame)
        self.input_entry = ttk.Entry(self.input_frame, textvariable=self.input_path_var, width=40)
        self.input_button = ttk.Button(self.input_frame, text="Durchsuchen...", 
                                      command=self._browse_input)
        
        # Output-Pfad
        self.output_label = ttk.Label(self.main_frame, text="Standard-Output:")
        self.output_frame = ttk.Frame(self.main_frame)
        self.output_entry = ttk.Entry(self.output_frame, textvariable=self.output_path_var, width=40)
        self.output_button = ttk.Button(self.output_frame, text="Durchsuchen...", 
                                       command=self._browse_output)
        
        # Fehler-Pfad
        self.error_label = ttk.Label(self.main_frame, text="Fehler-Ordner (optional):")
        self.error_frame = ttk.Frame(self.main_frame)
        self.error_entry = ttk.Entry(self.error_frame, textvariable=self.error_path_var, width=40)
        self.error_button = ttk.Button(self.error_frame, text="Durchsuchen...", 
                                      command=self._browse_error)
        
        # Export-Konfigurationen
        self.export_frame = ttk.LabelFrame(self.main_frame, text="Export-Konfigurationen", padding="10")
        self.export_info_label = ttk.Label(self.export_frame, 
                                          text="Konfigurieren Sie die Ausgabe-Optionen für verarbeitete Dokumente.")
        self.export_button = ttk.Button(self.export_frame, 
                                      text="Export-Optionen konfigurieren...",
                                      command=self._configure_exports)
        
        # Optionen
        self.options_frame = ttk.LabelFrame(self.main_frame, text="Optionen", padding="10")
        self.process_pairs_check = ttk.Checkbutton(
            self.options_frame,
            text="PDF-XML Paare verarbeiten",
            variable=self.process_pairs_var,
            command=self._on_process_pairs_toggle
        )
        
        # OCR-Zonen
        self.ocr_zones_frame = ttk.LabelFrame(self.main_frame, text="OCR-Zonen", padding="10")
        
        # OCR-Zonen Toolbar
        self.ocr_toolbar = ttk.Frame(self.ocr_zones_frame)
        self.add_zone_button = ttk.Button(self.ocr_toolbar, text="➕ Neue Zone", 
                                         command=self._add_ocr_zone)
        self.edit_zone_button = ttk.Button(self.ocr_toolbar, text="✏️ Bearbeiten", 
                                          command=self._edit_ocr_zone, state=tk.DISABLED)
        self.rename_zone_button = ttk.Button(self.ocr_toolbar, text="✏️ Umbenennen", 
                                            command=self._rename_ocr_zone, state=tk.DISABLED)
        self.delete_zone_button = ttk.Button(self.ocr_toolbar, text="🗑️ Löschen", 
                                            command=self._delete_ocr_zone, state=tk.DISABLED)
        
        # OCR-Zonen Liste
        self.zones_listbox = tk.Listbox(self.ocr_zones_frame, height=4)
        self.zones_listbox.bind('<<ListboxSelect>>', self._on_zone_selection)
        
        # XML-Felder Button
        self.xml_fields_button = ttk.Button(
            self.options_frame,
            text="XML-Felder konfigurieren...",
            command=self._configure_xml_fields,
            state=tk.NORMAL if self.process_pairs_var.get() else tk.DISABLED
        )
        
        # Aktionen
        self.actions_frame = ttk.LabelFrame(self.main_frame, text="PDF-Bearbeitungsaktionen", 
                                           padding="10")
        
        self.action_checks = {}
        action_descriptions = {
            ProcessingAction.COMPRESS: "PDF komprimieren",
            ProcessingAction.SPLIT: "In Einzelseiten aufteilen",
            ProcessingAction.OCR: "OCR durchführen (durchsuchbare PDF)",
            ProcessingAction.PDF_A: "In PDF/A konvertieren"
        }
        
        for action, description in action_descriptions.items():
            frame = ttk.Frame(self.actions_frame)
            check = ttk.Checkbutton(frame, text=description, 
                                   variable=self.action_vars[action])
            self.action_checks[action] = check
            check.pack(anchor=tk.W)
            frame.pack(anchor=tk.W, pady=2, fill=tk.X)
        
        # Info-Label
        self.info_label = ttk.Label(self.main_frame, 
                                   text="Hinweis: Die ausgewählten Aktionen werden in der "
                                        "angegebenen Reihenfolge ausgeführt.",
                                   wraplength=550)
        
        # Buttons - Reihenfolge getauscht!
        self.button_frame = ttk.Frame(self.main_frame)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save, default=tk.ACTIVE)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Name
        self.name_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.name_entry.grid(row=0, column=1, sticky="we", pady=(0, 5))
        
        # Input
        self.input_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        self.input_frame.grid(row=1, column=1, sticky="we", pady=(0, 5))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.input_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Output
        self.output_label.grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        self.output_frame.grid(row=2, column=1, sticky="we", pady=(0, 5))
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.output_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Fehler-Ordner
        self.error_label.grid(row=3, column=0, sticky=tk.W, pady=(0, 5))
        self.error_frame.grid(row=3, column=1, sticky="we", pady=(0, 5))
        self.error_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.error_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Export-Konfigurationen
        self.export_frame.grid(row=4, column=0, columnspan=2, sticky="we", pady=(10, 5))
        self.export_info_label.pack(anchor=tk.W, pady=(0, 5))
        self.export_button.pack(anchor=tk.W)
        
        # Optionen
        self.options_frame.grid(row=5, column=0, columnspan=2, sticky="we", pady=(10, 5))
        self.process_pairs_check.pack(anchor=tk.W)
        self.xml_fields_button.pack(anchor=tk.W, pady=(5, 0))
        
        # OCR-Zonen
        self.ocr_zones_frame.grid(row=6, column=0, columnspan=2, sticky="we", pady=(10, 5))
        self.ocr_toolbar.pack(fill=tk.X, pady=(0, 5))
        self.add_zone_button.pack(side=tk.LEFT, padx=(0, 5))
        self.edit_zone_button.pack(side=tk.LEFT, padx=(0, 5))
        self.rename_zone_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_zone_button.pack(side=tk.LEFT)
        self.zones_listbox.pack(fill=tk.BOTH, expand=True)
        
        # Aktionen
        self.actions_frame.grid(row=7, column=0, columnspan=2, sticky="we", pady=(10, 5))
        
        # Info
        self.info_label.grid(row=8, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        # Buttons - Reihenfolge beachten!
        self.button_frame.grid(row=9, column=0, columnspan=2, pady=(20, 0))
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.save_button.pack(side=tk.RIGHT)
        
        # Spalten-Konfiguration
        self.main_frame.columnconfigure(1, weight=1)
        
        # Lade OCR-Zonen in Liste
        self._refresh_zones_list()
    
    def _refresh_zones_list(self):
        """Aktualisiert die OCR-Zonen Liste"""
        self.zones_listbox.delete(0, tk.END)
        for zone in self.ocr_zones:
            zone_text = f"{zone['name']} - Seite {zone['page_num']}"
            self.zones_listbox.insert(tk.END, zone_text)
    
    def _update_button_texts(self):
        """Aktualisiert Button-Texte basierend auf Zustand"""
        if self.xml_field_mappings:
            count = len(self.xml_field_mappings)
            self.xml_fields_button.config(text=f"XML-Felder konfigurieren... ({count} Felder)")
        
        if self.ocr_zones:
            count = len(self.ocr_zones)
            # Aktualisiere OCR-Zonen Frame-Titel
            self.ocr_zones_frame.config(text=f"OCR-Zonen ({count} definiert)")
        
        if self.exports:
            count = len(self.exports)
            active_count = len([e for e in self.exports if e.get('enabled', True)])
            self.export_button.config(text=f"Export-Optionen konfigurieren... ({active_count}/{count} aktiv)")
        else:
            self.export_button.config(text="Export-Optionen konfigurieren... (Standard)")
    
    def _browse_input(self):
        """Öffnet Dialog zur Auswahl des Input-Ordners"""
        folder = filedialog.askdirectory(
            title="Input-Ordner auswählen",
            initialdir=self.input_path_var.get() or os.path.expanduser("~")
        )
        if folder:
            self.input_path_var.set(folder)
    
    def _browse_output(self):
        """Öffnet Dialog zur Auswahl des Output-Ordners"""
        folder = filedialog.askdirectory(
            title="Standard-Output-Ordner auswählen",
            initialdir=self.output_path_var.get() or os.path.expanduser("~")
        )
        if folder:
            self.output_path_var.set(folder)
    
    def _browse_error(self):
        """Öffnet Dialog zur Auswahl des Fehler-Ordners"""
        folder = filedialog.askdirectory(
            title="Fehler-Ordner auswählen (optional)",
            initialdir=self.error_path_var.get() or os.path.expanduser("~")
        )
        if folder:
            self.error_path_var.set(folder)
    
    def _configure_exports(self):
        """Öffnet Dialog zur Konfiguration der Export-Optionen"""
        # Wenn keine Exporte definiert, erstelle Standard-Export
        if not self.exports:
            self.exports = [{
                "id": "default",
                "name": "Standard-Ausgabe",
                "enabled": True,
                "export_type": "file",
                "output_path_expression": self.output_path_var.get() or "<OutputPath>",
                "filename_expression": "<FileName>",
                "create_path_if_not_exists": True,
                "append_to_existing_file": False,
                "error_output_path": self.error_path_var.get(),
                "metadata_config": {
                    "format": "none"
                }
            }]
        
        dialog = ExportListDialog(self.dialog, self.exports, 
                                 self.ocr_zones, self.xml_field_mappings)
        result = dialog.show()
        
        if result is not None:
            self.exports = result
            self._update_button_texts()
    
    def _on_zone_selection(self, event):
        """Wird aufgerufen wenn eine Zone ausgewählt wird"""
        selection = self.zones_listbox.curselection()
        if selection:
            self.edit_zone_button.config(state=tk.NORMAL)
            self.rename_zone_button.config(state=tk.NORMAL)
            self.delete_zone_button.config(state=tk.NORMAL)
        else:
            self.edit_zone_button.config(state=tk.DISABLED)
            self.rename_zone_button.config(state=tk.DISABLED)
            self.delete_zone_button.config(state=tk.DISABLED)
    
    def _add_ocr_zone(self):
        """Fügt eine neue OCR-Zone hinzu"""
        # PDF auswählen
        pdf_file = filedialog.askopenfilename(
            parent=self.dialog,
            title="PDF für Zone-Auswahl öffnen",
            filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
        )
        
        if pdf_file:
            selector = ZoneSelector(self.dialog, pdf_path=pdf_file)
            result = selector.show()
            
            if result:
                # Generiere eindeutigen Namen
                zone_index = len(self.ocr_zones) + 1
                zone_name = f"Zone_{zone_index}"
                
                # Frage nach benutzerdefiniertem Namen
                new_name = self._ask_zone_name(zone_name)
                if new_name:
                    zone_dict = {
                        'name': new_name,
                        'zone': result['zone'],
                        'page_num': result['page_num']
                    }
                    self.ocr_zones.append(zone_dict)
                    self._refresh_zones_list()
                    self._update_button_texts()
    
    def _edit_ocr_zone(self):
        """Bearbeitet die ausgewählte OCR-Zone"""
        selection = self.zones_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        zone = self.ocr_zones[index]
        
        # PDF für Bearbeitung auswählen
        pdf_file = filedialog.askopenfilename(
            parent=self.dialog,
            title="PDF für Zone-Bearbeitung öffnen",
            filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
        )
        
        if pdf_file:
            selector = ZoneSelector(
                self.dialog, 
                pdf_path=pdf_file,
                page_num=zone['page_num']
            )
            # Setze existierende Zone
            selector.zone = zone['zone']
            selector._draw_zone()  # Zeige existierende Zone
            
            result = selector.show()
            
            if result:
                # Aktualisiere Zone
                zone['zone'] = result['zone']
                zone['page_num'] = result['page_num']
                self._refresh_zones_list()
    
    def _rename_ocr_zone(self):
        """Benennt die ausgewählte OCR-Zone um"""
        selection = self.zones_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        zone = self.ocr_zones[index]
        
        new_name = self._ask_zone_name(zone['name'])
        if new_name and new_name != zone['name']:
            zone['name'] = new_name
            self._refresh_zones_list()
    
    def _delete_ocr_zone(self):
        """Löscht die ausgewählte OCR-Zone"""
        selection = self.zones_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        zone = self.ocr_zones[index]
        
        if messagebox.askyesno("Zone löschen", 
                              f"Möchten Sie die Zone '{zone['name']}' wirklich löschen?"):
            del self.ocr_zones[index]
            self._refresh_zones_list()
            self._update_button_texts()
    
    def _ask_zone_name(self, default_name: str = "") -> Optional[str]:
        """Fragt nach einem Namen für die OCR-Zone"""
        dialog = tk.Toplevel(self.dialog)
        dialog.title("Zone benennen")
        dialog.geometry("400x150")
        dialog.resizable(False, False)
        
        # Zentriere Dialog
        dialog.transient(self.dialog)
        dialog.grab_set()
        
        # Zentriere relativ zum Parent
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() - dialog.winfo_width()) // 2
        y = (dialog.winfo_screenheight() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")
        
        result = {'name': None}
        
        # Widgets
        frame = ttk.Frame(dialog, padding="20")
        frame.pack(fill=tk.BOTH, expand=True)
        
        ttk.Label(frame, text="Name für die OCR-Zone:").pack(anchor=tk.W)
        
        name_var = tk.StringVar(value=default_name)
        entry = ttk.Entry(frame, textvariable=name_var, width=40)
        entry.pack(fill=tk.X, pady=(5, 20))
        entry.focus()
        entry.select_range(0, tk.END)
        
        button_frame = ttk.Frame(frame)
        button_frame.pack(anchor=tk.E)
        
        def on_ok():
            name = name_var.get().strip()
            if not name:
                messagebox.showerror("Fehler", "Bitte geben Sie einen Namen ein.")
                return
            
            # Validiere Name (nur Buchstaben, Zahlen, Unterstrich)
            if not name.replace("_", "").isalnum():
                messagebox.showerror("Fehler", 
                    "Der Name darf nur Buchstaben, Zahlen und Unterstriche enthalten.")
                return
            
            # Prüfe ob Name bereits existiert
            existing_names = [z['name'] for z in self.ocr_zones if z['name'] != default_name]
            if name in existing_names:
                messagebox.showerror("Fehler", 
                    f"Eine Zone mit dem Namen '{name}' existiert bereits.")
                return
            
            result['name'] = name
            dialog.destroy()
        
        def on_cancel():
            dialog.destroy()
        
        ttk.Button(button_frame, text="Abbrechen", command=on_cancel).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="OK", command=on_ok, default=tk.ACTIVE).pack(side=tk.RIGHT)
        
        entry.bind('<Return>', lambda e: on_ok())
        entry.bind('<Escape>', lambda e: on_cancel())
        
        dialog.wait_window()
        return result['name']
    
    def _on_process_pairs_toggle(self):
        """Wird aufgerufen wenn Process Pairs Checkbox geändert wird"""
        if self.process_pairs_var.get():
            self.xml_fields_button.config(state=tk.NORMAL)
        else:
            self.xml_fields_button.config(state=tk.DISABLED)
    
    def _configure_xml_fields(self):
        """Öffnet Dialog zur Konfiguration der XML-Felder"""
        # Übergebe OCR-Zonen an den XML-Field-Dialog
        dialog = XMLFieldDialog(self.dialog, self.xml_field_mappings, self.ocr_zones)
        result = dialog.show()
        
        if result is not None:
            self.xml_field_mappings = result
            self._update_button_texts()
    
    def _validate(self) -> bool:
        """Validiert die Eingaben"""
        # Name prüfen
        if not self.name_var.get().strip():
            messagebox.showerror("Fehler", "Bitte geben Sie einen Namen ein.")
            self.name_entry.focus()
            return False
        
        # Input-Pfad prüfen
        if not self.input_path_var.get().strip():
            messagebox.showerror("Fehler", "Bitte wählen Sie einen Input-Ordner.")
            self.input_entry.focus()
            return False
        
        # Output-Pfad prüfen
        if not self.output_path_var.get().strip():
            messagebox.showerror("Fehler", "Bitte wählen Sie einen Standard-Output-Ordner.")
            self.output_entry.focus()
            return False
        
        # Pfade dürfen nicht identisch sein
        input_abs = os.path.abspath(self.input_path_var.get())
        output_abs = os.path.abspath(self.output_path_var.get())
        
        if input_abs == output_abs:
            messagebox.showerror("Fehler", "Input- und Output-Ordner dürfen nicht identisch sein.")
            return False
        
        # Fehler-Pfad prüfen wenn definiert
        if self.error_path_var.get().strip():
            error_abs = os.path.abspath(self.error_path_var.get())
            if error_abs == input_abs:
                messagebox.showerror("Fehler", "Fehler- und Input-Ordner dürfen nicht identisch sein.")
                return False
            if error_abs == output_abs:
                messagebox.showerror("Fehler", "Fehler- und Output-Ordner dürfen nicht identisch sein.")
                return False
        
        # Mindestens eine Aktion
        if not any(var.get() for var in self.action_vars.values()):
            messagebox.showerror("Fehler", "Bitte wählen Sie mindestens eine Bearbeitungsaktion.")
            return False
        
        # Mindestens ein aktiver Export
        if self.exports:
            active_exports = [e for e in self.exports if e.get('enabled', True)]
            if not active_exports:
                messagebox.showerror("Fehler", "Bitte aktivieren Sie mindestens einen Export.")
                return False
        
        return True
    
    def _on_save(self):
        """Speichert die Eingaben"""
        if not self._validate():
            return
        
        # Sammle ausgewählte Aktionen
        selected_actions = [action.value for action, var in self.action_vars.items() 
                           if var.get()]
        
        self.result = {
            "name": self.name_var.get().strip(),
            "input_path": self.input_path_var.get().strip(),
            "output_path": self.output_path_var.get().strip(),
            "error_path": self.error_path_var.get().strip(),
            "process_pairs": self.process_pairs_var.get(),
            "actions": selected_actions,
            "action_params": self.action_params,
            "xml_field_mappings": self.xml_field_mappings,
            "ocr_zones": self.ocr_zones,
            "exports": self.exports
        }
        
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht den Dialog ab"""
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result