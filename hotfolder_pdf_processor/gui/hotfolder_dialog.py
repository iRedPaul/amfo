"""
Dialog zum Erstellen und Bearbeiten von Hotfoldern - Überarbeitete Version
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
from gui.export_dialog import ExportDialog


class HotfolderDialog:
    """Dialog zum Erstellen/Bearbeiten eines Hotfolders"""
    
    def __init__(self, parent, hotfolder: Optional[HotfolderConfig] = None):
        self.parent = parent
        self.hotfolder = hotfolder
        self.result = None
        
        # Erstelle Dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Hotfolder bearbeiten" if hotfolder else "Neuer Hotfolder")
        self.dialog.geometry("800x700")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog relativ zum Parent
        self._center_window()
        
        # Dialog-Eigenschaften
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Variablen
        self.name_var = tk.StringVar(value=hotfolder.name if hotfolder else "")
        self.input_path_var = tk.StringVar(value=hotfolder.input_path if hotfolder else "")
        self.process_pairs_var = tk.BooleanVar(value=hotfolder.process_pairs if hotfolder else True)
        
        # Action Variablen - nur noch die Basis-Aktionen
        self.action_vars = {}
        # Entferne PDF_A aus den Basis-Aktionen
        basic_actions = [ProcessingAction.COMPRESS, ProcessingAction.SPLIT, ProcessingAction.OCR]
        for action in basic_actions:
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
        self.export_configs = hotfolder.export_configs.copy() if hotfolder and hasattr(hotfolder, 'export_configs') else []
        self.error_path = hotfolder.error_path if hotfolder and hasattr(hotfolder, 'error_path') else ""
        
        # Legacy: Output-Pfad für Abwärtskompatibilität
        self.output_path_var = tk.StringVar(value=hotfolder.output_path if hotfolder else "")
        
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
        
        # Notebook für bessere Organisation
        self.notebook = ttk.Notebook(self.main_frame)
        
        # Tab 1: Basis-Einstellungen
        self.basic_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.basic_frame, text="Grundeinstellungen")
        
        # Name und Pfade
        self.name_label = ttk.Label(self.basic_frame, text="Name:")
        self.name_entry = ttk.Entry(self.basic_frame, textvariable=self.name_var, width=50)
        
        self.input_label = ttk.Label(self.basic_frame, text="Überwachter Ordner (Input):")
        self.input_frame = ttk.Frame(self.basic_frame)
        self.input_entry = ttk.Entry(self.input_frame, textvariable=self.input_path_var, width=40)
        self.input_button = ttk.Button(self.input_frame, text="Durchsuchen...", 
                                      command=self._browse_input)
        
        # Verarbeitungsoptionen
        self.processing_frame = ttk.LabelFrame(self.basic_frame, text="Verarbeitungsoptionen", padding="10")
        
        self.process_pairs_check = ttk.Checkbutton(
            self.processing_frame,
            text="PDF-XML Paare verarbeiten (PDF und zugehörige XML-Datei gemeinsam verarbeiten)",
            variable=self.process_pairs_var,
            command=self._on_process_pairs_toggle
        )
        
        # Basis-Aktionen (ohne PDF/A, das ist jetzt nur noch bei Export)
        self.actions_frame = ttk.LabelFrame(self.basic_frame, text="Vorverarbeitungsschritte", padding="10")
        
        action_descriptions = {
            ProcessingAction.COMPRESS: {
                "text": "PDF komprimieren",
                "desc": "Reduziert die Dateigröße durch Komprimierung"
            },
            ProcessingAction.SPLIT: {
                "text": "In Einzelseiten aufteilen",
                "desc": "Teilt mehrseitige PDFs in einzelne Dateien auf"
            },
            ProcessingAction.OCR: {
                "text": "Texterkennung durchführen",
                "desc": "Macht PDFs durchsuchbar (OCR)"
            }
        }
        
        self.action_checks = {}
        for action, info in action_descriptions.items():
            if action in self.action_vars:
                frame = ttk.Frame(self.actions_frame)
                check = ttk.Checkbutton(frame, text=info["text"], 
                                       variable=self.action_vars[action])
                self.action_checks[action] = check
                check.pack(anchor=tk.W)
                desc_label = ttk.Label(frame, text=info["desc"], foreground="gray", 
                                      font=('TkDefaultFont', 9))
                desc_label.pack(anchor=tk.W, padx=(20, 0))
                frame.pack(anchor=tk.W, pady=5, fill=tk.X)
        
        # Tab 2: Datenextraktion
        self.extraction_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.extraction_frame, text="Datenextraktion")
        
        # OCR-Zonen
        self.ocr_zones_frame = ttk.LabelFrame(self.extraction_frame, text="OCR-Zonen", padding="10")
        self.ocr_zones_desc = ttk.Label(self.ocr_zones_frame, 
            text="Definieren Sie Bereiche im PDF, aus denen Text extrahiert werden soll.",
            wraplength=600, foreground="gray")
        
        # OCR-Zonen Toolbar
        self.ocr_toolbar = ttk.Frame(self.ocr_zones_frame)
        self.add_zone_button = ttk.Button(self.ocr_toolbar, text="➕ Neue Zone", 
                                         command=self._add_ocr_zone)
        self.edit_zone_button = ttk.Button(self.ocr_toolbar, text="✏️ Bearbeiten", 
                                          command=self._edit_ocr_zone, state=tk.DISABLED)
        self.rename_zone_button = ttk.Button(self.ocr_toolbar, text="📝 Umbenennen", 
                                            command=self._rename_ocr_zone, state=tk.DISABLED)
        self.delete_zone_button = ttk.Button(self.ocr_toolbar, text="🗑️ Löschen", 
                                            command=self._delete_ocr_zone, state=tk.DISABLED)
        
        # OCR-Zonen Liste
        self.zones_listbox = tk.Listbox(self.ocr_zones_frame, height=6)
        self.zones_listbox.bind('<<ListboxSelect>>', self._on_zone_selection)
        
        # XML-Felder
        self.xml_frame = ttk.LabelFrame(self.extraction_frame, text="XML-Feldverarbeitung", padding="10")
        self.xml_desc = ttk.Label(self.xml_frame, 
            text="Konfigurieren Sie, wie XML-Felder befüllt werden sollen.",
            wraplength=600, foreground="gray")
        
        self.xml_fields_button = ttk.Button(
            self.xml_frame,
            text="XML-Felder konfigurieren...",
            command=self._configure_xml_fields,
            state=tk.NORMAL if self.process_pairs_var.get() else tk.DISABLED
        )
        
        # Tab 3: Export
        self.export_frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(self.export_frame, text="Export")
        
        # Export-Beschreibung
        self.export_desc_frame = ttk.Frame(self.export_frame)
        self.export_desc = ttk.Label(self.export_desc_frame, 
            text="Konfigurieren Sie, wie und wohin die verarbeiteten Dokumente exportiert werden sollen. "
                 "Sie können mehrere Exporte definieren (z.B. PDF/A als Datei speichern UND per E-Mail versenden).",
            wraplength=650, justify=tk.LEFT)
        
        # Export-Button
        self.export_button_frame = ttk.Frame(self.export_frame)
        self.export_button = ttk.Button(
            self.export_button_frame,
            text="Exporte konfigurieren...",
            command=self._configure_exports,
            style="Accent.TButton"
        )
        
        # Info wenn keine Exporte
        self.no_export_info = ttk.Label(self.export_frame, 
            text="⚠️ Hinweis: Wenn keine Exporte konfiguriert sind, werden die Dateien nur in einen Output-Ordner verschoben.",
            foreground="orange", wraplength=650)
        
        # Legacy Output (versteckt, nur für Abwärtskompatibilität)
        # Wird intern auf einen Standardwert gesetzt
        
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
        
        # Tab 1: Basis-Einstellungen
        # Name
        self.name_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.name_entry.grid(row=0, column=1, sticky="we", pady=(0, 5))
        
        # Input
        self.input_label.grid(row=1, column=0, sticky=tk.W, pady=(10, 5))
        self.input_frame.grid(row=1, column=1, sticky="we", pady=(10, 5))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.input_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Verarbeitungsoptionen
        self.processing_frame.grid(row=2, column=0, columnspan=2, sticky="we", pady=(20, 10))
        self.process_pairs_check.pack(anchor=tk.W)
        
        # Aktionen
        self.actions_frame.grid(row=3, column=0, columnspan=2, sticky="we", pady=(10, 0))
        
        self.basic_frame.columnconfigure(1, weight=1)
        
        # Tab 2: Datenextraktion
        # OCR-Zonen
        self.ocr_zones_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.ocr_zones_desc.pack(anchor=tk.W, pady=(0, 10))
        self.ocr_toolbar.pack(fill=tk.X, pady=(0, 5))
        self.add_zone_button.pack(side=tk.LEFT, padx=(0, 5))
        self.edit_zone_button.pack(side=tk.LEFT, padx=(0, 5))
        self.rename_zone_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_zone_button.pack(side=tk.LEFT)
        self.zones_listbox.pack(fill=tk.BOTH, expand=True)
        
        # XML-Felder
        self.xml_frame.pack(fill=tk.X, pady=(10, 0))
        self.xml_desc.pack(anchor=tk.W, pady=(0, 10))
        self.xml_fields_button.pack(anchor=tk.W)
        
        # Tab 3: Export
        self.export_desc_frame.pack(fill=tk.X, pady=(0, 15))
        self.export_desc.pack()
        
        self.export_button_frame.pack(pady=(0, 20))
        self.export_button.pack()
        
        if not self.export_configs:
            self.no_export_info.pack(pady=(20, 0))
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.save_button.pack(side=tk.RIGHT)
        
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
            self.xml_fields_button.config(text=f"XML-Felder konfigurieren... ({count} Felder definiert)")
        else:
            self.xml_fields_button.config(text="XML-Felder konfigurieren...")
        
        if self.ocr_zones:
            count = len(self.ocr_zones)
            self.ocr_zones_frame.config(text=f"OCR-Zonen ({count} definiert)")
        else:
            self.ocr_zones_frame.config(text="OCR-Zonen")
        
        if self.export_configs:
            count = len(self.export_configs)
            active_count = sum(1 for ec in self.export_configs if ec.get('enabled', True))
            self.export_button.config(text=f"Exporte konfigurieren... ({active_count} von {count} aktiv)")
            # Verstecke Hinweis wenn Exporte konfiguriert
            if hasattr(self, 'no_export_info'):
                self.no_export_info.pack_forget()
        else:
            self.export_button.config(text="Exporte konfigurieren...")
            # Zeige Hinweis wenn keine Exporte
            if hasattr(self, 'no_export_info'):
                self.no_export_info.pack(pady=(20, 0))
    
    def _browse_input(self):
        """Öffnet Dialog zur Auswahl des Input-Ordners"""
        folder = filedialog.askdirectory(
            title="Überwachten Ordner auswählen",
            initialdir=self.input_path_var.get() or os.path.expanduser("~")
        )
        if folder:
            self.input_path_var.set(folder)
    
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
    
    def _configure_exports(self):
        """Öffnet Dialog zur Konfiguration der Exporte"""
        dialog = ExportDialog(
            self.dialog, 
            self.export_configs, 
            self.error_path,
            self.xml_field_mappings
        )
        result = dialog.show()
        
        if result is not None:
            self.export_configs = result['export_configs']
            self.error_path = result['error_path']
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
            messagebox.showerror("Fehler", "Bitte wählen Sie einen Ordner zur Überwachung.")
            self.input_entry.focus()
            return False
        
        return True
    
    def _on_save(self):
        """Speichert die Eingaben"""
        if not self._validate():
            return
        
        # Sammle ausgewählte Aktionen
        selected_actions = [action.value for action, var in self.action_vars.items() 
                           if var.get()]
        
        # Setze Output-Path auf Input-Path wenn keine Exporte definiert
        # Dies ist für Legacy-Kompatibilität
        if not self.export_configs:
            output_path = os.path.join(self.input_path_var.get().strip(), "output")
        else:
            # Wenn Exporte definiert sind, wird Output-Path nicht mehr verwendet
            output_path = self.input_path_var.get().strip()
        
        self.result = {
            "name": self.name_var.get().strip(),
            "input_path": self.input_path_var.get().strip(),
            "output_path": output_path,  # Legacy-Feld
            "process_pairs": self.process_pairs_var.get(),
            "actions": selected_actions,
            "action_params": self.action_params,
            "xml_field_mappings": self.xml_field_mappings,
            "output_filename_expression": "<FileName>",  # Default für Legacy
            "ocr_zones": self.ocr_zones,
            "export_configs": self.export_configs,
            "error_path": self.error_path
        }
        
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht den Dialog ab"""
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result