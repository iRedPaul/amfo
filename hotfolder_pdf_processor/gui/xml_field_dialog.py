"""
Dialog zur Konfiguration von XML-Feld-Mappings mit erweiterter Funktionsunterstützung
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import List, Dict, Optional, Tuple
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.xml_field_processor import FieldMapping, XMLFieldProcessor
from gui.zone_selector import ZoneSelector


class XMLFieldDialog:
    """Dialog zum Konfigurieren von XML-Feld-Mappings"""
    
    def __init__(self, parent, mappings: List[Dict] = None):
        self.parent = parent
        self.mappings = mappings or []
        self.result = None
        self.xml_processor = XMLFieldProcessor()
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("XML-Felder konfigurieren")
        self.dialog.geometry("1000x700")
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
                                columns=("Field", "Type", "Expression"),
                                show="tree headings", height=15)
        
        # Spalten konfigurieren
        self.tree.heading("#0", text="")
        self.tree.heading("Field", text="XML-Feld")
        self.tree.heading("Type", text="Typ")
        self.tree.heading("Expression", text="Ausdruck/Funktion")
        
        self.tree.column("#0", width=30)
        self.tree.column("Field", width=200)
        self.tree.column("Type", width=120)
        self.tree.column("Expression", width=450)
        
        # Scrollbar
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", 
                                command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.vsb.set)
        
        # Info-Text
        self.info_frame = ttk.LabelFrame(self.main_frame, text="Hilfe", padding="10")
        self.info_text = tk.Text(self.info_frame, height=8, wrap=tk.WORD)
        self.info_text.insert("1.0", 
            "XML-Felder mit dynamischen Werten befüllen:\n\n"
            "• Klicken Sie auf 'Neues Feld' um ein Mapping hinzuzufügen\n"
            "• Verwenden Sie Variablen in spitzen Klammern: <VariablenName>\n"
            "• Nutzen Sie Funktionen für erweiterte Verarbeitung\n"
            "• OCR-Zonen können grafisch definiert werden\n"
            "• Testen Sie Ihre Ausdrücke mit der Test-Funktion\n\n"
            "Doppelklick zum Bearbeiten eines Eintrags")
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
        if mapping.source_type == "ocr_zone" and mapping.zone:
            type_text = f"OCR-Zone ({mapping.zone[2]}x{mapping.zone[3]})"
        else:
            type_text = "Ausdruck"
        
        expression_text = mapping.expression[:60] + "..." if len(mapping.expression) > 60 else mapping.expression
        
        self.tree.insert("", "end", 
                        values=(mapping.field_name, type_text, expression_text))
    
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
        dialog = FieldMappingEditDialog(self.dialog, xml_processor=self.xml_processor)
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
            
            dialog = FieldMappingEditDialog(self.dialog, mapping, self.xml_processor)
            result = dialog.show()
            
            if result:
                # Aktualisiere Mapping
                new_mapping = FieldMapping(**result)
                self.mappings[index] = new_mapping.to_dict()
                
                # Aktualisiere Tree
                self.tree.delete(item)
                
                if new_mapping.source_type == "ocr_zone" and new_mapping.zone:
                    type_text = f"OCR-Zone ({new_mapping.zone[2]}x{new_mapping.zone[3]})"
                else:
                    type_text = "Ausdruck"
                    
                expression_text = new_mapping.expression[:60] + "..." if len(new_mapping.expression) > 60 else new_mapping.expression
                
                self.tree.insert("", index, values=(
                    new_mapping.field_name,
                    type_text,
                    expression_text
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
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if 0 <= index < len(self.mappings):
            mapping_dict = self.mappings[index]
            mapping = FieldMapping.from_dict(mapping_dict)
            
            # Test-Dialog
            dialog = TestMappingDialog(self.dialog, mapping, self.xml_processor)
            dialog.show()
    
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
    
    def __init__(self, parent, mapping: Optional[FieldMapping] = None, 
                 xml_processor: XMLFieldProcessor = None):
        self.parent = parent
        self.mapping = mapping
        self.result = None
        self.xml_processor = xml_processor or XMLFieldProcessor()
        self.ocr_zones = []  # Liste der definierten OCR-Zonen
        self.current_zone_index = 0
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Feld-Mapping bearbeiten")
        self.dialog.geometry("900x700")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Variablen
        self.field_name_var = tk.StringVar(value=mapping.field_name if mapping else "")
        self.source_type_var = tk.StringVar(value=mapping.source_type if mapping else "expression")
        self.expression_var = tk.StringVar(value=mapping.expression if mapping else "")
        
        # Lade existierende Zonen wenn vorhanden
        if mapping and mapping.zone:
            self.ocr_zones.append({
                'zone': mapping.zone,
                'page_num': mapping.page_num,
                'name': 'Zone_1'
            })
        
        self._create_widgets()
        self._layout_widgets()
        self._update_help("")
        
        # Fokus
        self.field_name_entry.focus()
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Hauptframe
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Feldname
        self.field_frame = ttk.LabelFrame(self.main_frame, text="Feld-Konfiguration", padding="10")
        self.field_name_label = ttk.Label(self.field_frame, text="XML-Feldname:")
        self.field_name_entry = ttk.Entry(self.field_frame, textvariable=self.field_name_var, width=40)
        
        # Notebook für verschiedene Bereiche
        self.notebook = ttk.Notebook(self.main_frame)
        
        # Tab 1: Expression Builder
        self.expr_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.expr_frame, text="Ausdruck erstellen")
        
        # Expression-Eingabe mit größerem Bereich
        self.expr_input_frame = ttk.LabelFrame(self.expr_frame, text="Ausdruck", padding="10")
        self.expr_text = tk.Text(self.expr_input_frame, height=6, width=70)
        if self.mapping and self.mapping.expression:
            self.expr_text.insert("1.0", self.mapping.expression)
        
        # Kombinierter Tree für Variablen und Funktionen
        self.tree_frame = ttk.LabelFrame(self.expr_frame, text="Verfügbare Variablen und Funktionen", padding="5")
        
        # Paned Window für Tree und Hilfe
        self.paned = ttk.PanedWindow(self.tree_frame, orient=tk.HORIZONTAL)
        
        # Tree
        self.var_func_tree = ttk.Treeview(self.paned, height=15)
        self.var_func_tree.heading("#0", text="Variablen und Funktionen")
        
        # Hilfe-Bereich
        self.help_frame = ttk.LabelFrame(self.paned, text="Beschreibung", padding="5")
        self.help_text = tk.Text(self.help_frame, width=40, height=15, wrap=tk.WORD)
        self.help_text.config(state=tk.DISABLED)
        
        self.paned.add(self.var_func_tree, weight=2)
        self.paned.add(self.help_frame, weight=1)
        
        # Tab 2: OCR-Zonen
        self.zone_tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.zone_tab_frame, text="OCR-Zonen")
        
        # Zone-Liste
        self.zone_list_frame = ttk.LabelFrame(self.zone_tab_frame, text="Definierte OCR-Zonen", padding="10")
        
        # Zone Toolbar
        self.zone_toolbar = ttk.Frame(self.zone_list_frame)
        self.add_zone_button = ttk.Button(self.zone_toolbar, text="➕ Neue Zone", 
                                         command=self._add_ocr_zone)
        self.edit_zone_button = ttk.Button(self.zone_toolbar, text="✏️ Bearbeiten", 
                                          command=self._edit_ocr_zone, state=tk.DISABLED)
        self.delete_zone_button = ttk.Button(self.zone_toolbar, text="🗑️ Löschen", 
                                            command=self._delete_ocr_zone, state=tk.DISABLED)
        self.test_zone_button = ttk.Button(self.zone_toolbar, text="🧪 Zone testen", 
                                          command=self._test_ocr_zone, state=tk.DISABLED)
        
        # Zone-Liste
        self.zone_listbox = tk.Listbox(self.zone_list_frame, height=6)
        for i, zone_info in enumerate(self.ocr_zones):
            self.zone_listbox.insert(tk.END, f"{zone_info['name']} - Seite {zone_info['page_num']}")
        
        # Zone-Expression
        self.zone_expr_frame = ttk.LabelFrame(self.zone_tab_frame, text="Zone-Verarbeitung", padding="10")
        self.zone_expr_label = ttk.Label(self.zone_expr_frame, 
                                        text="Optional: Ausdruck zur Verarbeitung des Zone-Texts")
        self.zone_expr_text = tk.Text(self.zone_expr_frame, height=3, width=60)
        self.zone_expr_help = ttk.Label(self.zone_expr_frame, 
                                       text="Verwenden Sie <ZONE_X> als Platzhalter für den jeweiligen Zone-Text",
                                       foreground="gray")
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        
        # Events
        self.var_func_tree.bind("<<TreeviewSelect>>", self._on_tree_selection)
        self.var_func_tree.bind("<Double-Button-1>", self._on_tree_double_click)
        self.zone_listbox.bind("<<ListboxSelect>>", self._on_zone_selection)
        
        # Lade Variablen und Funktionen
        self._load_variables_and_functions()
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Feld-Konfiguration
        self.field_frame.pack(fill=tk.X, pady=(0, 10))
        self.field_name_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.field_name_entry.grid(row=0, column=1, sticky="we")
        self.field_frame.columnconfigure(1, weight=1)
        
        # Notebook
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Expression Tab Layout
        self.expr_input_frame.pack(fill=tk.X, pady=(10, 10))
        self.expr_text.pack(fill=tk.BOTH, expand=True)
        
        self.tree_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 0))
        self.paned.pack(fill=tk.BOTH, expand=True)
        self.help_text.pack(fill=tk.BOTH, expand=True)
        
        # Zone Tab Layout
        self.zone_list_frame.pack(fill=tk.BOTH, expand=True, pady=(10, 10))
        self.zone_toolbar.pack(fill=tk.X, pady=(0, 5))
        self.add_zone_button.pack(side=tk.LEFT, padx=(0, 5))
        self.edit_zone_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_zone_button.pack(side=tk.LEFT, padx=(0, 5))
        self.test_zone_button.pack(side=tk.LEFT)
        self.zone_listbox.pack(fill=tk.BOTH, expand=True)
        
        self.zone_expr_frame.pack(fill=tk.X, pady=(10, 0))
        self.zone_expr_label.pack(anchor=tk.W)
        self.zone_expr_text.pack(fill=tk.X, pady=(5, 5))
        self.zone_expr_help.pack(anchor=tk.W)
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.save_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.cancel_button.pack(side=tk.RIGHT)
    
    def _load_variables_and_functions(self):
        """Lädt verfügbare Variablen und Funktionen in einen kombinierten Tree"""
        # Variablen-Hauptknoten
        var_root = self.var_func_tree.insert("", "end", text="Variablen", open=True)
        
        # Standard-Variablen
        std_node = self.var_func_tree.insert(var_root, "end", text="Standard", open=True)
        std_vars = [
            ("Date", "Aktuelles Datum"),
            ("Time", "Aktuelle Uhrzeit"),
            ("DateTime", "Datum und Uhrzeit"),
            ("Year", "Aktuelles Jahr"),
            ("Month", "Aktueller Monat"),
            ("Day", "Aktueller Tag"),
            ("Hour", "Aktuelle Stunde"),
            ("Minute", "Aktuelle Minute"),
            ("Second", "Aktuelle Sekunde"),
            ("Weekday", "Wochentag"),
            ("WeekNumber", "Kalenderwoche")
        ]
        for var, desc in std_vars:
            self.var_func_tree.insert(std_node, "end", text=var, tags=("variable", desc))
        
        # Datei-Variablen
        file_node = self.var_func_tree.insert(var_root, "end", text="Datei", open=True)
        file_vars = [
            ("FileName", "Dateiname ohne Erweiterung"),
            ("FileExtension", "Dateierweiterung"),
            ("FullFileName", "Vollständiger Dateiname"),
            ("FilePath", "Pfad zur Datei"),
            ("FullPath", "Vollständiger Pfad"),
            ("FileSize", "Dateigröße in Bytes")
        ]
        for var, desc in file_vars:
            self.var_func_tree.insert(file_node, "end", text=var, tags=("variable", desc))
        
        # OCR-Variablen
        ocr_node = self.var_func_tree.insert(var_root, "end", text="OCR", open=True)
        self.var_func_tree.insert(ocr_node, "end", text="OCR_FullText", 
                                 tags=("variable", "Kompletter OCR-Text der PDF"))
        
        # OCR-Zonen werden dynamisch hinzugefügt
        for i, zone_info in enumerate(self.ocr_zones):
            self.var_func_tree.insert(ocr_node, "end", text=f"ZONE_{i+1}", 
                                     tags=("variable", f"Text aus {zone_info['name']}"))
        
        # Funktionen-Hauptknoten
        func_root = self.var_func_tree.insert("", "end", text="Funktionen", open=True)
        
        # String-Funktionen
        string_node = self.var_func_tree.insert(func_root, "end", text="String", open=True)
        string_funcs = [
            ("FORMAT", 'FORMAT("<VAR>","<FORMATSTRING>")', "Formatiert eine Zeichenkette"),
            ("TRIM", 'TRIM("<VAR>")', "Entfernt Leerzeichen am Anfang und Ende"),
            ("LEFT", 'LEFT("<VAR>",<LENGTH>)', "Gibt die ersten n Zeichen zurück"),
            ("RIGHT", 'RIGHT("<VAR>",<LENGTH>)', "Gibt die letzten n Zeichen zurück"),
            ("MID", 'MID("<VAR>",<START>,<LENGTH>)', "Gibt Zeichen aus der Mitte zurück"),
            ("TOUPPER", 'TOUPPER("<VAR>")', "Konvertiert zu Großbuchstaben"),
            ("TOLOWER", 'TOLOWER("<VAR>")', "Konvertiert zu Kleinbuchstaben"),
            ("LEN", 'LEN("<VAR>")', "Gibt die Länge der Zeichenkette zurück"),
            ("INDEXOF", 'INDEXOF(<START>,"<STRING>","<SEARCH>",<CASE>)', "Findet Position eines Zeichens")
        ]
        for name, syntax, desc in string_funcs:
            self.var_func_tree.insert(string_node, "end", text=name, 
                                     tags=("function", syntax, desc))
        
        # Datumsfunktionen
        date_node = self.var_func_tree.insert(func_root, "end", text="Datum", open=True)
        self.var_func_tree.insert(date_node, "end", text="FORMATDATE", 
                                 tags=("function", 'FORMATDATE("<FORMAT>")', 
                                      "Formatiert das aktuelle Datum\n\nFormatzeichen:\n"
                                      "dd - Tag mit führender Null\n"
                                      "mm - Monat mit führender Null\n"
                                      "yyyy - Jahr vierstellig\n"
                                      "HH - Stunde (24h)\n"
                                      "MM - Minute\n"
                                      "ss - Sekunde"))
        
        # Bedingungen
        cond_node = self.var_func_tree.insert(func_root, "end", text="Bedingungen", open=True)
        self.var_func_tree.insert(cond_node, "end", text="IF", 
                                 tags=("function", 
                                      'IF("<VAR>","<OP>","<VALUE>","<TRUE>","<FALSE>")', 
                                      "Bedingte Auswertung\n\nOperatoren:\n"
                                      "= oder == - Gleich\n"
                                      "!= - Ungleich\n"
                                      "> - Größer als\n"
                                      "< - Kleiner als\n"
                                      ">= - Größer gleich\n"
                                      "<= - Kleiner gleich\n"
                                      "contains - Enthält\n"
                                      "startswith - Beginnt mit\n"
                                      "endswith - Endet mit"))
        
        # RegEx
        regex_node = self.var_func_tree.insert(func_root, "end", text="Reguläre Ausdrücke", open=True)
        regex_funcs = [
            ("REGEXP.MATCH", 'REGEXP.MATCH("<VAR>","<PATTERN>",<INDEX>)', 
             "Findet Muster mit regulären Ausdrücken\n\n"
             "INDEX = 0 gibt die erste Übereinstimmung zurück\n"
             "INDEX > 0 gibt die entsprechende Gruppe zurück"),
            ("REGEXP.REPLACE", 'REGEXP.REPLACE("<VAR>","<PATTERN>","<REPLACE>")', 
             "Ersetzt Muster mit regulären Ausdrücken")
        ]
        for name, syntax, desc in regex_funcs:
            self.var_func_tree.insert(regex_node, "end", text=name, 
                                     tags=("function", syntax, desc))
        
        # Numerisch
        num_node = self.var_func_tree.insert(func_root, "end", text="Numerisch", open=True)
        self.var_func_tree.insert(num_node, "end", text="AUTOINCREMENT", 
                                 tags=("function", 'AUTOINCREMENT("<VAR>",<STEP>)', 
                                      "Zählt einen Wert hoch\n\n"
                                      "VAR - Startwert\n"
                                      "STEP - Schrittweite"))
    
    def _on_tree_selection(self, event):
        """Wird aufgerufen wenn ein Element im Tree ausgewählt wird"""
        selection = self.var_func_tree.selection()
        if selection:
            item = self.var_func_tree.item(selection[0])
            tags = item.get('tags', [])
            
            if len(tags) >= 2:
                if tags[0] == "variable":
                    self._update_help(f"Variable: <{item['text']}>\n\n{tags[1]}")
                elif tags[0] == "function":
                    syntax = tags[1] if len(tags) > 1 else ""
                    desc = tags[2] if len(tags) > 2 else ""
                    self._update_help(f"Funktion: {item['text']}\n\nSyntax:\n{syntax}\n\n{desc}")
    
    def _on_tree_double_click(self, event):
        """Doppelklick fügt Variable oder Funktion ein"""
        selection = self.var_func_tree.selection()
        if selection:
            item = self.var_func_tree.item(selection[0])
            tags = item.get('tags', [])
            
            if tags and tags[0] == "variable":
                # Variable einfügen
                self.expr_text.insert(tk.INSERT, f"<{item['text']}>")
            elif tags and tags[0] == "function" and len(tags) > 1:
                # Funktion einfügen
                self.expr_text.insert(tk.INSERT, tags[1])
    
    def _update_help(self, text):
        """Aktualisiert den Hilfetext"""
        self.help_text.config(state=tk.NORMAL)
        self.help_text.delete("1.0", tk.END)
        self.help_text.insert("1.0", text)
        self.help_text.config(state=tk.DISABLED)
    
    def _add_ocr_zone(self):
        """Fügt eine neue OCR-Zone hinzu"""
        pdf_file = filedialog.askopenfilename(
            title="PDF für Zone-Auswahl öffnen",
            filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
        )
        
        if pdf_file:
            selector = ZoneSelector(self.dialog, pdf_path=pdf_file)
            result = selector.show()
            
            if result:
                zone_index = len(self.ocr_zones) + 1
                zone_info = {
                    'zone': result['zone'],
                    'page_num': result['page_num'],
                    'name': f'Zone_{zone_index}',
                    'pdf_path': result['pdf_path']
                }
                self.ocr_zones.append(zone_info)
                
                # Zur Liste hinzufügen
                self.zone_listbox.insert(tk.END, f"{zone_info['name']} - Seite {zone_info['page_num']}")
                
                # Variable im Tree aktualisieren
                self._refresh_ocr_variables()
    
    def _edit_ocr_zone(self):
        """Bearbeitet die ausgewählte OCR-Zone"""
        selection = self.zone_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        zone_info = self.ocr_zones[index]
        
        # PDF für Bearbeitung auswählen
        pdf_file = zone_info.get('pdf_path')
        if not pdf_file or not os.path.exists(pdf_file):
            pdf_file = filedialog.askopenfilename(
                title="PDF für Zone-Bearbeitung öffnen",
                filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
            )
        
        if pdf_file:
            selector = ZoneSelector(
                self.dialog, 
                pdf_path=pdf_file,
                page_num=zone_info['page_num']
            )
            # Setze existierende Zone
            selector.zone = zone_info['zone']
            result = selector.show()
            
            if result:
                # Aktualisiere Zone
                zone_info['zone'] = result['zone']
                zone_info['page_num'] = result['page_num']
                zone_info['pdf_path'] = result['pdf_path']
                
                # Liste aktualisieren
                self.zone_listbox.delete(index)
                self.zone_listbox.insert(index, f"{zone_info['name']} - Seite {zone_info['page_num']}")
    
    def _delete_ocr_zone(self):
        """Löscht die ausgewählte OCR-Zone"""
        selection = self.zone_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        if messagebox.askyesno("Zone löschen", "Möchten Sie diese OCR-Zone wirklich löschen?"):
            del self.ocr_zones[index]
            self.zone_listbox.delete(index)
            self._refresh_ocr_variables()
    
    def _test_ocr_zone(self):
        """Testet die ausgewählte OCR-Zone"""
        selection = self.zone_listbox.curselection()
        if not selection:
            return
        
        index = selection[0]
        zone_info = self.ocr_zones[index]
        
        # Test-Dialog
        dialog = TestOCRZoneDialog(self.dialog, zone_info, self.xml_processor.ocr_processor)
        dialog.show()
    
    def _refresh_ocr_variables(self):
        """Aktualisiert die OCR-Variablen im Tree"""
        # Finde OCR-Knoten
        for item in self.var_func_tree.get_children(""):
            if self.var_func_tree.item(item)['text'] == "Variablen":
                for child in self.var_func_tree.get_children(item):
                    if self.var_func_tree.item(child)['text'] == "OCR":
                        # Lösche alle Zone-Variablen
                        for zone_var in self.var_func_tree.get_children(child):
                            if self.var_func_tree.item(zone_var)['text'].startswith("ZONE_"):
                                self.var_func_tree.delete(zone_var)
                        
                        # Füge neue Zone-Variablen hinzu
                        for i, zone_info in enumerate(self.ocr_zones):
                            self.var_func_tree.insert(child, "end", text=f"ZONE_{i+1}", 
                                                     tags=("variable", f"Text aus {zone_info['name']}"))
                        break
                break
    
    def _on_zone_selection(self, event):
        """Wird aufgerufen wenn eine Zone ausgewählt wird"""
        selection = self.zone_listbox.curselection()
        if selection:
            self.edit_zone_button.config(state=tk.NORMAL)
            self.delete_zone_button.config(state=tk.NORMAL)
            self.test_zone_button.config(state=tk.NORMAL)
        else:
            self.edit_zone_button.config(state=tk.DISABLED)
            self.delete_zone_button.config(state=tk.DISABLED)
            self.test_zone_button.config(state=tk.DISABLED)
    
    def _validate(self) -> bool:
        """Validiert die Eingaben"""
        if not self.field_name_var.get().strip():
            messagebox.showerror("Fehler", "Bitte geben Sie einen Feldnamen ein.")
            return False
        
        # Validiere Feldname
        field_name = self.field_name_var.get().strip()
        if not field_name.replace("_", "").isalnum():
            messagebox.showerror("Fehler", 
                "Der Feldname darf nur Buchstaben, Zahlen und Unterstriche enthalten.")
            return False
        
        # Prüfe ob Ausdruck oder Zone definiert
        expression = self.expr_text.get("1.0", tk.END).strip()
        zone_expression = self.zone_expr_text.get("1.0", tk.END).strip()
        
        if not expression and not self.ocr_zones and not zone_expression:
            messagebox.showerror("Fehler", 
                "Bitte definieren Sie einen Ausdruck oder eine OCR-Zone.")
            return False
        
        return True
    
    def _on_save(self):
        """Speichert die Eingaben"""
        if not self._validate():
            return
        
        # Bestimme Source-Type und Expression
        current_tab = self.notebook.index(self.notebook.select())
        
        if current_tab == 1 and self.ocr_zones:  # OCR-Zone Tab
            # Verwende erste Zone (später kann erweitert werden)
            zone_info = self.ocr_zones[0]
            source_type = "ocr_zone"
            zone = zone_info['zone']
            page_num = zone_info['page_num']
            
            # Zone-Expression
            zone_expr = self.zone_expr_text.get("1.0", tk.END).strip()
            if zone_expr:
                # Ersetze ZONE_X Platzhalter
                expression = zone_expr
                for i in range(len(self.ocr_zones)):
                    expression = expression.replace(f"<ZONE_{i+1}>", f"<OCR_Zone_{i+1}>")
            else:
                expression = ""
        else:  # Expression Tab
            source_type = "expression"
            expression = self.expr_text.get("1.0", tk.END).strip()
            zone = None
            page_num = 1
        
        self.result = {
            "field_name": self.field_name_var.get().strip(),
            "source_type": source_type,
            "expression": expression,
            "zone": zone,
            "page_num": page_num
        }
        
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab"""
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result


class TestOCRZoneDialog:
    """Dialog zum Testen einer OCR-Zone"""
    
    def __init__(self, parent, zone_info: dict, ocr_processor):
        self.parent = parent
        self.zone_info = zone_info
        self.ocr_processor = ocr_processor
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("OCR-Zone testen")
        self.dialog.geometry("600x400")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._create_widgets()
        self._layout_widgets()
        
        # Automatisch testen wenn PDF vorhanden
        if zone_info.get('pdf_path') and os.path.exists(zone_info['pdf_path']):
            self._run_test()
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Info
        info_text = f"Zone: {self.zone_info['name']}\n"
        info_text += f"Seite: {self.zone_info['page_num']}\n"
        info_text += f"Position: X={self.zone_info['zone'][0]}, Y={self.zone_info['zone'][1]}\n"
        info_text += f"Größe: {self.zone_info['zone'][2]}x{self.zone_info['zone'][3]} Pixel"
        
        self.info_label = ttk.Label(self.main_frame, text=info_text)
        
        # PDF-Auswahl
        self.pdf_frame = ttk.Frame(self.main_frame)
        self.pdf_label = ttk.Label(self.pdf_frame, text="PDF-Datei:")
        self.pdf_path_var = tk.StringVar(value=self.zone_info.get('pdf_path', ''))
        self.pdf_entry = ttk.Entry(self.pdf_frame, textvariable=self.pdf_path_var, 
                                  width=40, state='readonly')
        self.pdf_button = ttk.Button(self.pdf_frame, text="Durchsuchen...", 
                                    command=self._select_pdf)
        
        # Ergebnis
        self.result_frame = ttk.LabelFrame(self.main_frame, text="OCR-Ergebnis", padding="10")
        self.result_text = tk.Text(self.result_frame, height=10, wrap=tk.WORD)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.test_button = ttk.Button(self.button_frame, text="Testen", 
                                     command=self._run_test)
        self.close_button = ttk.Button(self.button_frame, text="Schließen", 
                                      command=self.dialog.destroy)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.info_label.pack(anchor=tk.W, pady=(0, 10))
        
        self.pdf_frame.pack(fill=tk.X, pady=(0, 10))
        self.pdf_label.pack(side=tk.LEFT, padx=(0, 5))
        self.pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.pdf_button.pack(side=tk.LEFT)
        
        self.result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        self.button_frame.pack(fill=tk.X)
        self.test_button.pack(side=tk.LEFT, padx=(0, 5))
        self.close_button.pack(side=tk.LEFT)
    
    def _select_pdf(self):
        """Wählt eine PDF-Datei aus"""
        filename = filedialog.askopenfilename(
            title="PDF auswählen",
            filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")],
            initialfile=self.pdf_path_var.get()
        )
        
        if filename:
            self.pdf_path_var.set(filename)
            self.zone_info['pdf_path'] = filename
    
    def _run_test(self):
        """Führt den OCR-Test aus"""
        pdf_path = self.pdf_path_var.get()
        if not pdf_path or not os.path.exists(pdf_path):
            messagebox.showerror("Fehler", "Bitte wählen Sie eine gültige PDF-Datei.")
            return
        
        try:
            # Führe OCR aus
            text = self.ocr_processor.extract_text_from_zone(
                pdf_path,
                self.zone_info['page_num'],
                self.zone_info['zone'],
                language='deu'
            )
            
            # Zeige Ergebnis
            self.result_text.delete("1.0", tk.END)
            if text:
                self.result_text.insert("1.0", text)
            else:
                self.result_text.insert("1.0", "(Kein Text erkannt)")
                
        except Exception as e:
            self.result_text.delete("1.0", tk.END)
            self.result_text.insert("1.0", f"Fehler: {e}")
    
    def show(self):
        """Zeigt den Dialog"""
        self.dialog.wait_window()


class TestMappingDialog:
    """Dialog zum Testen eines Mappings"""
    
    def __init__(self, parent, mapping: FieldMapping, xml_processor: XMLFieldProcessor):
        self.parent = parent
        self.mapping = mapping
        self.xml_processor = xml_processor
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Mapping testen")
        self.dialog.geometry("700x600")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._create_widgets()
        self._layout_widgets()
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Info
        self.info_label = ttk.Label(self.main_frame, 
                                   text=f"Test für Feld: {self.mapping.field_name}")
        
        # Expression
        self.expr_frame = ttk.LabelFrame(self.main_frame, text="Ausdruck", padding="10")
        self.expr_text = tk.Text(self.expr_frame, height=3, wrap=tk.WORD)
        self.expr_text.insert("1.0", self.mapping.expression)
        self.expr_text.config(state=tk.DISABLED)
        
        # Test-Dateien
        self.files_frame = ttk.LabelFrame(self.main_frame, text="Test-Dateien", padding="10")
        
        # PDF
        self.pdf_frame = ttk.Frame(self.files_frame)
        self.pdf_label = ttk.Label(self.pdf_frame, text="PDF-Datei:")
        self.pdf_path_var = tk.StringVar()
        self.pdf_entry = ttk.Entry(self.pdf_frame, textvariable=self.pdf_path_var, 
                                  width=40, state='readonly')
        self.pdf_button = ttk.Button(self.pdf_frame, text="Durchsuchen...", 
                                    command=self._select_pdf)
        
        # XML
        self.xml_frame = ttk.Frame(self.files_frame)
        self.xml_label = ttk.Label(self.xml_frame, text="XML-Datei (optional):")
        self.xml_path_var = tk.StringVar()
        self.xml_entry = ttk.Entry(self.xml_frame, textvariable=self.xml_path_var, 
                                  width=40, state='readonly')
        self.xml_button = ttk.Button(self.xml_frame, text="Durchsuchen...", 
                                    command=self._select_xml)
        
        # Test-Variablen
        self.vars_frame = ttk.LabelFrame(self.main_frame, text="Test-Variablen überschreiben (optional)", padding="10")
        self.vars_text = tk.Text(self.vars_frame, height=4)
        self.vars_text.insert("1.0", 
            "# Format: VariablenName=Wert\n"
            "# Beispiel:\n"
            "# FileName=TestDokument\n"
            "# Date=01.01.2024\n")
        
        # Ergebnis
        self.result_frame = ttk.LabelFrame(self.main_frame, text="Ergebnis", padding="10")
        self.result_text = tk.Text(self.result_frame, height=6)
        self.result_text.config(state=tk.DISABLED)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.test_button = ttk.Button(self.button_frame, text="Testen", 
                                     command=self._run_test)
        self.close_button = ttk.Button(self.button_frame, text="Schließen", 
                                      command=self.dialog.destroy)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.info_label.pack(anchor=tk.W, pady=(0, 10))
        
        self.expr_frame.pack(fill=tk.X, pady=(0, 10))
        self.expr_text.pack(fill=tk.X)
        
        self.files_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.pdf_frame.pack(fill=tk.X, pady=(0, 5))
        self.pdf_label.pack(side=tk.LEFT, padx=(0, 5))
        self.pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.pdf_button.pack(side=tk.LEFT)
        
        self.xml_frame.pack(fill=tk.X)
        self.xml_label.pack(side=tk.LEFT, padx=(0, 5))
        self.xml_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.xml_button.pack(side=tk.LEFT)
        
        self.vars_frame.pack(fill=tk.X, pady=(0, 10))
        self.vars_text.pack(fill=tk.X)
        
        self.result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.result_text.pack(fill=tk.BOTH, expand=True)
        
        self.button_frame.pack(fill=tk.X)
        self.test_button.pack(side=tk.LEFT, padx=(0, 5))
        self.close_button.pack(side=tk.LEFT)
    
    def _select_pdf(self):
        """Wählt eine PDF-Datei aus"""
        filename = filedialog.askopenfilename(
            title="PDF für Test auswählen",
            filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
        )
        
        if filename:
            self.pdf_path_var.set(filename)
    
    def _select_xml(self):
        """Wählt eine XML-Datei aus"""
        filename = filedialog.askopenfilename(
            title="XML für Test auswählen",
            filetypes=[("XML Dateien", "*.xml"), ("Alle Dateien", "*.*")]
        )
        
        if filename:
            self.xml_path_var.set(filename)
    
    def _run_test(self):
        """Führt den Test aus"""
        pdf_path = self.pdf_path_var.get()
        if not pdf_path:
            messagebox.showerror("Fehler", "Bitte wählen Sie eine PDF-Datei für den Test.")
            return
        
        xml_path = self.xml_path_var.get() or None
        
        # Baue Kontext auf
        context = self.xml_processor._build_context(
            xml_path or "", 
            pdf_path, 
            [self.mapping]
        )
        
        # Überschreibe mit Test-Variablen
        for line in self.vars_text.get("1.0", tk.END).strip().split("\n"):
            if "=" in line and not line.startswith("#"):
                key, value = line.split("=", 1)
                context[key.strip()] = value.strip()
        
        # Evaluiere Expression
        try:
            result = self.xml_processor._evaluate_mapping(self.mapping, context)
            
            self.result_text.config(state=tk.NORMAL)
            self.result_text.delete("1.0", tk.END)
            
            if result is not None:
                self.result_text.insert("1.0", f"Ergebnis: {result}\n\n")
                
                # Zeige verwendete Variablen
                self.result_text.insert(tk.END, "Verwendete Variablen:\n")
                for key, value in sorted(context.items()):
                    if len(str(value)) > 50:
                        value_str = str(value)[:50] + "..."
                    else:
                        value_str = str(value)
                    self.result_text.insert(tk.END, f"{key} = {value_str}\n")
            else:
                self.result_text.insert("1.0", "Kein Ergebnis")
                
            self.result_text.config(state=tk.DISABLED)
            
        except Exception as e:
            self.result_text.config(state=tk.NORMAL)
            self.result_text.delete("1.0", tk.END)
            self.result_text.insert("1.0", f"Fehler: {e}")
            self.result_text.config(state=tk.DISABLED)
    
    def show(self):
        """Zeigt den Dialog"""
        self.dialog.wait_window()
