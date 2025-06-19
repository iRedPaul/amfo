"""
Dialog zur Konfiguration von XML-Feld-Mappings
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import List, Dict, Optional, Tuple
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.xml_field_processor import FieldMapping
from gui.zone_selector import ZoneSelector


class XMLFieldDialog:
    """Dialog zum Konfigurieren von XML-Feld-Mappings"""
    
    def __init__(self, parent, mappings: List[Dict] = None):
        self.parent = parent
        self.mappings = mappings or []
        self.result = None
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("XML-Felder konfigurieren")
        self.dialog.geometry("900x600")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._create_widgets()
        self._layout_widgets()
        self._load_mappings()
        
        # Bind Events
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Hauptframe
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Toolbar
        self.toolbar = ttk.Frame(self.main_frame)
        self.add_button = ttk.Button(self.toolbar, text="➕ Neues Feld", 
                                    command=self._add_field)
        self.edit_button = ttk.Button(self.toolbar, text="✏️ Bearbeiten", 
                                     command=self._edit_field, state=tk.DISABLED)
        self.delete_button = ttk.Button(self.toolbar, text="🗑️ Löschen", 
                                       command=self._delete_field, state=tk.DISABLED)
        self.test_button = ttk.Button(self.toolbar, text="🧪 Testen", 
                                     command=self._test_mapping, state=tk.DISABLED)
        
        # Liste der Mappings
        self.tree_frame = ttk.Frame(self.main_frame)
        self.tree = ttk.Treeview(self.tree_frame, 
                                columns=("Field", "Source", "Pattern"),
                                show="tree headings", height=15)
        
        # Spalten konfigurieren
        self.tree.heading("#0", text="")
        self.tree.heading("Field", text="XML-Feld")
        self.tree.heading("Source", text="Quelle")
        self.tree.heading("Pattern", text="Muster/RegEx")
        
        self.tree.column("#0", width=30)
        self.tree.column("Field", width=200)
        self.tree.column("Source", width=150)
        self.tree.column("Pattern", width=300)
        
        # Scrollbar
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", 
                                command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.vsb.set)
        
        # Info-Text
        self.info_frame = ttk.LabelFrame(self.main_frame, text="Hilfe", padding="10")
        self.info_text = tk.Text(self.info_frame, height=6, wrap=tk.WORD)
        self.info_text.insert("1.0", 
            "Quelltypen:\n"
            "• OCR-Text: Extrahiert Text aus der gesamten PDF\n"
            "• OCR-Zone: Extrahiert Text aus einem bestimmten Bereich\n"
            "• Fester Wert: Verwendet einen festen Wert\n"
            "• Datum: Fügt aktuelles Datum ein\n"
            "• Dateiname: Extrahiert aus dem PDF-Dateinamen")
        self.info_text.config(state=tk.DISABLED)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        
        # Events
        self.tree.bind("<<TreeviewSelect>>", self._on_selection_changed)
        self.tree.bind("<Double-Button-1>", lambda e: self._edit_field())
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Toolbar
        self.toolbar.pack(fill=tk.X, pady=(0, 10))
        self.add_button.pack(side=tk.LEFT, padx=(0, 5))
        self.edit_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_button.pack(side=tk.LEFT, padx=(0, 5))
        self.test_button.pack(side=tk.LEFT)
        
        # Tree
        self.tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)
        
        # Info
        self.info_frame.pack(fill=tk.X, pady=(0, 10))
        self.info_text.pack(fill=tk.BOTH)
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.save_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.cancel_button.pack(side=tk.RIGHT)
    
    def _load_mappings(self):
        """Lädt die Mappings in die Liste"""
        for mapping_dict in self.mappings:
            mapping = FieldMapping.from_dict(mapping_dict)
            self._add_to_tree(mapping)
    
    def _add_to_tree(self, mapping: FieldMapping):
        """Fügt ein Mapping zur Baumanzeige hinzu"""
        source_text = {
            "ocr_text": "OCR-Text",
            "ocr_zone": f"OCR-Zone (Seite {mapping.page_num})",
            "fixed": "Fester Wert",
            "date": "Datum",
            "filename": "Dateiname"
        }.get(mapping.source_type, mapping.source_type)
        
        pattern_text = mapping.pattern[:50] + "..." if len(mapping.pattern) > 50 else mapping.pattern
        
        self.tree.insert("", "end", 
                        values=(mapping.field_name, source_text, 
                               pattern_text))
    
    def _on_selection_changed(self, event):
        """Wird aufgerufen wenn die Auswahl sich ändert"""
        selection = self.tree.selection()
        if selection:
            self.edit_button.config(state=tk.NORMAL)
            self.delete_button.config(state=tk.NORMAL)
            self.test_button.config(state=tk.NORMAL)
        else:
            self.edit_button.config(state=tk.DISABLED)
            self.delete_button.config(state=tk.DISABLED)
            self.test_button.config(state=tk.DISABLED)
    
    def _add_field(self):
        """Fügt ein neues Feld hinzu"""
        dialog = FieldMappingEditDialog(self.dialog)
        result = dialog.show()
        
        if result:
            mapping = FieldMapping(**result)
            self._add_to_tree(mapping)
            self.mappings.append(mapping.to_dict())
    
    def _edit_field(self):
        """Bearbeitet das ausgewählte Feld"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if 0 <= index < len(self.mappings):
            mapping_dict = self.mappings[index]
            mapping = FieldMapping.from_dict(mapping_dict)
            
            dialog = FieldMappingEditDialog(self.dialog, mapping)
            result = dialog.show()
            
            if result:
                # Aktualisiere Mapping
                new_mapping = FieldMapping(**result)
                self.mappings[index] = new_mapping.to_dict()
                
                # Aktualisiere Tree
                self.tree.delete(item)
                self.tree.insert("", index, values=(
                    new_mapping.field_name,
                    {
                        "ocr_text": "OCR-Text",
                        "ocr_zone": f"OCR-Zone (Seite {new_mapping.page_num})",
                        "fixed": "Fester Wert",
                        "date": "Datum",
                        "filename": "Dateiname"
                    }.get(new_mapping.source_type, new_mapping.source_type),
                    new_mapping.pattern[:50] + "..." if len(new_mapping.pattern) > 50 else new_mapping.pattern
                ))
    
    def _delete_field(self):
        """Löscht das ausgewählte Feld"""
        selection = self.tree.selection()
        if not selection:
            return
        
        if messagebox.askyesno("Feld löschen", "Möchten Sie dieses Feld wirklich löschen?"):
            item = selection[0]
            index = self.tree.index(item)
            
            if 0 <= index < len(self.mappings):
                del self.mappings[index]
                self.tree.delete(item)
    
    def _test_mapping(self):
        """Testet das ausgewählte Mapping"""
        # TODO: Implementiere Test-Funktionalität
        messagebox.showinfo("Test", "Test-Funktionalität wird in einer späteren Version verfügbar sein.")
    
    def _on_save(self):
        """Speichert die Mappings"""
        self.result = self.mappings
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab ohne zu speichern"""
        self.dialog.destroy()
    
    def show(self) -> Optional[List[Dict]]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result


class FieldMappingEditDialog:
    """Dialog zum Bearbeiten eines einzelnen Field-Mappings"""
    
    def __init__(self, parent, mapping: Optional[FieldMapping] = None):
        self.parent = parent
        self.mapping = mapping
        self.result = None
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Feld-Mapping bearbeiten")
        self.dialog.geometry("700x600")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Variablen
        self.field_name_var = tk.StringVar(value=mapping.field_name if mapping else "")
        self.source_type_var = tk.StringVar(value=mapping.source_type if mapping else "ocr_text")
        self.pattern_var = tk.StringVar(value=mapping.pattern if mapping else "")
        self.page_num_var = tk.IntVar(value=mapping.page_num if mapping else 1)
        
        # Zone-Variablen
        if mapping and mapping.zone:
            self.zone_x_var = tk.IntVar(value=mapping.zone[0])
            self.zone_y_var = tk.IntVar(value=mapping.zone[1])
            self.zone_w_var = tk.IntVar(value=mapping.zone[2])
            self.zone_h_var = tk.IntVar(value=mapping.zone[3])
        else:
            self.zone_x_var = tk.IntVar(value=0)
            self.zone_y_var = tk.IntVar(value=0)
            self.zone_w_var = tk.IntVar(value=100)
            self.zone_h_var = tk.IntVar(value=50)
        
        self._create_widgets()
        self._layout_widgets()
        self._update_visibility()
        
        # Fokus
        self.field_name_entry.focus()
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Hauptframe
        self.main_frame = ttk.Frame(self.dialog, padding="20")
        
        # Feldname
        self.field_name_label = ttk.Label(self.main_frame, text="XML-Feldname:")
        self.field_name_entry = ttk.Entry(self.main_frame, textvariable=self.field_name_var, 
                                         width=40)
        
        # Quelltyp
        self.source_type_label = ttk.Label(self.main_frame, text="Quelltyp:")
        self.source_type_frame = ttk.Frame(self.main_frame)
        
        source_types = [
            ("ocr_text", "OCR-Text (gesamte PDF)"),
            ("ocr_zone", "OCR-Zone (bestimmter Bereich)"),
            ("fixed", "Fester Wert"),
            ("date", "Datum"),
            ("filename", "Dateiname")
        ]
        
        for value, text in source_types:
            ttk.Radiobutton(self.source_type_frame, text=text, 
                           variable=self.source_type_var, value=value,
                           command=self._update_visibility).pack(anchor=tk.W, pady=2)
        
        # Pattern/RegEx, Wert oder Datumsformat
        self.pattern_label = ttk.Label(self.main_frame, text="Muster/RegEx:")
        self.pattern_text = tk.Text(self.main_frame, height=4, width=50)
        if self.mapping and self.mapping.pattern:
            self.pattern_text.insert("1.0", self.mapping.pattern)
        # Hilfe-Button für Datumsformate
        self.date_help_button = ttk.Button(self.main_frame, text="Beispiele für Datumsformate", command=self._show_date_examples)
        
        # RegEx-Beispiele
        self.examples_button = ttk.Button(self.main_frame, text="Beispiele", 
                                         command=self._show_examples)
        
        # Seite (für OCR-Zone)
        self.page_label = ttk.Label(self.main_frame, text="Seite:")
        self.page_spinbox = ttk.Spinbox(self.main_frame, from_=1, to=100, 
                                       textvariable=self.page_num_var, width=10)
        
        # Zone (für OCR-Zone)
        self.zone_frame = ttk.LabelFrame(self.main_frame, text="Zone (Pixel)", 
                                        padding="10")
        
        # Zone-Auswahl Button
        self.zone_select_button = ttk.Button(self.zone_frame, 
                                           text="Zone grafisch auswählen...",
                                           command=self._select_zone_graphically)
        self.zone_select_button.grid(row=0, column=0, columnspan=4, pady=(0, 10))
        
        ttk.Label(self.zone_frame, text="X:").grid(row=1, column=0, padx=(0, 5))
        ttk.Spinbox(self.zone_frame, from_=0, to=9999, width=10,
                   textvariable=self.zone_x_var).grid(row=1, column=1, padx=(0, 15))
        
        ttk.Label(self.zone_frame, text="Y:").grid(row=1, column=2, padx=(0, 5))
        ttk.Spinbox(self.zone_frame, from_=0, to=9999, width=10,
                   textvariable=self.zone_y_var).grid(row=1, column=3)
        
        ttk.Label(self.zone_frame, text="Breite:").grid(row=2, column=0, padx=(0, 5), pady=(5, 0))
        ttk.Spinbox(self.zone_frame, from_=1, to=9999, width=10,
                   textvariable=self.zone_w_var).grid(row=2, column=1, padx=(0, 15), pady=(5, 0))
        
        ttk.Label(self.zone_frame, text="Höhe:").grid(row=2, column=2, padx=(0, 5), pady=(5, 0))
        ttk.Spinbox(self.zone_frame, from_=1, to=9999, width=10,
                   textvariable=self.zone_h_var).grid(row=2, column=3, pady=(5, 0))
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        row = 0
        self.field_name_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 10))
        self.field_name_entry.grid(row=row, column=1, columnspan=2, sticky='we', pady=(0, 10))
        
        row += 1
        self.source_type_label.grid(row=row, column=0, sticky=tk.NW, pady=(0, 10))
        self.source_type_frame.grid(row=row, column=1, columnspan=2, sticky='w', pady=(0, 10))
        
        row += 1
        self.pattern_label.grid(row=row, column=0, sticky=tk.NW, pady=(0, 5))
        self.pattern_text.grid(row=row, column=1, columnspan=2, sticky='we', pady=(0, 5))
        self.examples_button.grid(row=row+1, column=1, sticky=tk.W, pady=(0, 10))
        self.date_help_button.grid_remove()  # Standardmäßig nicht sichtbar
        
        row += 1
        self.page_label.grid(row=row, column=0, sticky=tk.W, pady=(0, 10))
        self.page_spinbox.grid(row=row, column=1, sticky=tk.W, pady=(0, 10))
        
        row += 1
        self.zone_frame.grid(row=row, column=0, columnspan=3, sticky='we', pady=(0, 10))
        
        row += 1
        self.button_frame.grid(row=row, column=0, columnspan=3, pady=(20, 0))
        self.save_button.pack(side=tk.LEFT, padx=(0, 5))
        self.cancel_button.pack(side=tk.LEFT)
        
        self.main_frame.columnconfigure(1, weight=1)
    
    def _update_visibility(self):
        """Aktualisiert die Sichtbarkeit von Widgets basierend auf Quelltyp"""
        source_type = self.source_type_var.get()
        
        if source_type == "fixed":
            self.pattern_label.config(text="Wert:")
            self.pattern_label.grid()
            self.pattern_text.grid()
            self.examples_button.grid_remove()
            self.date_help_button.grid_remove()
        elif source_type == "date":
            self.pattern_label.config(text="Datumsformat:")
            self.pattern_label.grid()
            self.pattern_text.grid()
            self.examples_button.grid_remove()
            self.date_help_button.grid(row=self.pattern_label.grid_info()['row']+1, column=1, sticky=tk.W, pady=(0, 10))
        elif source_type == "filename":
            self.pattern_label.grid_remove()
            self.pattern_text.grid_remove()
            self.examples_button.grid_remove()
            self.date_help_button.grid_remove()
        elif source_type in ["ocr_text", "ocr_zone"]:
            self.pattern_label.config(text="Muster/RegEx:")
            self.pattern_label.grid()
            self.pattern_text.grid()
            self.examples_button.grid()
            self.date_help_button.grid_remove()
        else:
            self.pattern_label.grid_remove()
            self.pattern_text.grid_remove()
            self.examples_button.grid_remove()
            self.date_help_button.grid_remove()
        
        # Seite und Zone
        if source_type == "ocr_zone":
            self.page_label.grid()
            self.page_spinbox.grid()
            self.zone_frame.grid()
        else:
            self.page_label.grid_remove()
            self.page_spinbox.grid_remove()
            self.zone_frame.grid_remove()
    
    def _show_examples(self):
        """Zeigt Beispiele für reguläre Ausdrücke"""
        examples = """Beispiele für reguläre Ausdrücke:

Rechnungsnummer:
Rechnungsnummer[:\\s]+([A-Z0-9\\-/]+)

Betrag:
([0-9.,]+)\\s*€

Datum:
(\\d{1,2}[.-/]\\d{1,2}[.-/]\\d{2,4})

E-Mail:
([a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\\.[a-zA-Z]{2,})

Erste Zahl nach einem Wort:
Betrag[:\\s]+([0-9.,]+)

Alles zwischen zwei Wörtern:
von\\s+(.+?)\\s+bis

Tipp: Verwenden Sie () um den zu extrahierenden Teil zu markieren!"""
        
        messagebox.showinfo("RegEx-Beispiele", examples)
    
    def _show_date_examples(self):
        examples = """Beispiele für Datumsformate:\n\n"""
        examples += "dd.mm.yyyy   →  31.12.2024\n"
        examples += "yyyy-mm-dd   →  2024-12-31\n"
        examples += "d.m.yy       →  1.2.24\n"
        examples += "d m yy       →  1 2 24\n"
        examples += "dd/mm/yyyy   →  31/12/2024\n"
        examples += "\nSie können beliebige Trennzeichen verwenden.\n"
        examples += "Weitere Infos: https://docs.python.org/3/library/datetime.html#strftime-and-strptime-format-codes"
        messagebox.showinfo("Datumsformat-Beispiele", examples)
    
    def _validate(self) -> bool:
        """Validiert die Eingaben"""
        if not self.field_name_var.get().strip():
            messagebox.showerror("Fehler", "Bitte geben Sie einen Feldnamen ein.")
            return False
        
        # Validiere Feldname (nur alphanumerisch und Unterstrich)
        field_name = self.field_name_var.get().strip()
        if not field_name.replace("_", "").isalnum():
            messagebox.showerror("Fehler", 
                "Der Feldname darf nur Buchstaben, Zahlen und Unterstriche enthalten.")
            return False
        
        return True
    
    def _on_save(self):
        """Speichert die Eingaben"""
        if not self._validate():
            return
        
        pattern = self.pattern_text.get("1.0", tk.END).strip()
        
        # Zone nur für OCR-Zone
        zone = None
        if self.source_type_var.get() == "ocr_zone":
            zone = (
                self.zone_x_var.get(),
                self.zone_y_var.get(),
                self.zone_w_var.get(),
                self.zone_h_var.get()
            )
        
        self.result = {
            "field_name": self.field_name_var.get().strip(),
            "source_type": self.source_type_var.get(),
            "pattern": pattern,
            "zone": zone,
            "page_num": self.page_num_var.get()
        }
        
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab"""
        self.dialog.destroy()
    
    def _select_zone_graphically(self):
        """Öffnet den grafischen Zone-Selector"""
        # Frage nach PDF-Datei wenn keine vorhanden
        pdf_path = None
        pdf_file = filedialog.askopenfilename(
            title="PDF für Zone-Auswahl öffnen",
            filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
        )
        
        if pdf_file:
            selector = ZoneSelector(
                self.dialog, 
                pdf_path=pdf_file,
                page_num=self.page_num_var.get()
            )
            result = selector.show()
            
            if result:
                # Übernehme Zone-Koordinaten
                zone = result['zone']
                self.zone_x_var.set(zone[0])
                self.zone_y_var.set(zone[1])
                self.zone_w_var.set(zone[2])
                self.zone_h_var.set(zone[3])
                
                # Übernehme Seitennummer
                self.page_num_var.set(result['page_num'])
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result