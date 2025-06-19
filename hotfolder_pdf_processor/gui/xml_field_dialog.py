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
            "Verwenden Sie Funktionen und Variablen um XML-Felder zu befüllen:\n\n"
            "Variablen: <VariablenName> z.B. <FileName>, <OCR_FullText>\n"
            "Funktionen: FUNKTIONSNAME(Parameter) z.B. FORMATDATE(\"dd.mm.yyyy\")\n"
            "Bedingungen: IF(Variable, Operator, Wert, WennWahr, WennFalsch)\n"
            "RegEx: REGEXP.MATCH(Variable, Pattern, Index)\n"
            "String: LEFT(Variable, Länge), TRIM(Variable), TOUPPER(Variable)\n"
            "Verkettung: Kombinieren Sie Text und Funktionen mit +")
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
        type_text = "OCR-Zone" if mapping.source_type == "ocr_zone" else "Ausdruck"
        
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
                type_text = "OCR-Zone" if new_mapping.source_type == "ocr_zone" else "Ausdruck"
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
    """Dialog zum Bearbeiten eines einzelnen Field-Mappings mit Funktionsunterstützung"""
    
    def __init__(self, parent, mapping: Optional[FieldMapping] = None, 
                 xml_processor: XMLFieldProcessor = None):
        self.parent = parent
        self.mapping = mapping
        self.result = None
        self.xml_processor = xml_processor or XMLFieldProcessor()
        
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
        # Hauptframe mit Notebook
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
        
        # Expression-Eingabe
        self.expr_label = ttk.Label(self.expr_frame, text="Ausdruck:")
        self.expr_text = tk.Text(self.expr_frame, height=4, width=70)
        if self.mapping and self.mapping.expression:
            self.expr_text.insert("1.0", self.mapping.expression)
        
        # Funktionen und Variablen
        self.builder_frame = ttk.Frame(self.expr_frame)
        
        # Variablen
        self.var_frame = ttk.LabelFrame(self.builder_frame, text="Variablen", padding="5")
        self.var_tree = ttk.Treeview(self.var_frame, height=10)
        self.var_tree.heading("#0", text="Verfügbare Variablen")
        
        # Funktionen
        self.func_frame = ttk.LabelFrame(self.builder_frame, text="Funktionen", padding="5")
        self.func_tree = ttk.Treeview(self.func_frame, columns=("syntax",), height=10)
        self.func_tree.heading("#0", text="Funktion")
        self.func_tree.heading("syntax", text="Syntax")
        self.func_tree.column("#0", width=150)
        self.func_tree.column("syntax", width=300)
        
        # Buttons für Expression Builder
        self.expr_button_frame = ttk.Frame(self.expr_frame)
        self.insert_var_button = ttk.Button(self.expr_button_frame, text="Variable einfügen",
                                           command=self._insert_variable)
        self.insert_func_button = ttk.Button(self.expr_button_frame, text="Funktion einfügen",
                                            command=self._insert_function)
        self.clear_expr_button = ttk.Button(self.expr_button_frame, text="Leeren",
                                           command=lambda: self.expr_text.delete("1.0", tk.END))
        
        # Tab 2: OCR-Zone
        self.zone_tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.zone_tab_frame, text="OCR-Zone")
        
        # Zone-Konfiguration
        self.zone_config_frame = ttk.LabelFrame(self.zone_tab_frame, text="Zone definieren", padding="10")
        
        # Zone-Auswahl Button
        self.zone_select_button = ttk.Button(self.zone_config_frame, 
                                           text="Zone grafisch auswählen...",
                                           command=self._select_zone_graphically)
        
        # Manuelle Zone-Eingabe
        self.zone_manual_frame = ttk.Frame(self.zone_config_frame)
        ttk.Label(self.zone_manual_frame, text="X:").grid(row=0, column=0, padx=(0, 5))
        ttk.Spinbox(self.zone_manual_frame, from_=0, to=9999, width=10,
                   textvariable=self.zone_x_var).grid(row=0, column=1, padx=(0, 15))
        
        ttk.Label(self.zone_manual_frame, text="Y:").grid(row=0, column=2, padx=(0, 5))
        ttk.Spinbox(self.zone_manual_frame, from_=0, to=9999, width=10,
                   textvariable=self.zone_y_var).grid(row=0, column=3)
        
        ttk.Label(self.zone_manual_frame, text="Breite:").grid(row=1, column=0, padx=(0, 5), pady=(5, 0))
        ttk.Spinbox(self.zone_manual_frame, from_=1, to=9999, width=10,
                   textvariable=self.zone_w_var).grid(row=1, column=1, padx=(0, 15), pady=(5, 0))
        
        ttk.Label(self.zone_manual_frame, text="Höhe:").grid(row=1, column=2, padx=(0, 5), pady=(5, 0))
        ttk.Spinbox(self.zone_manual_frame, from_=1, to=9999, width=10,
                   textvariable=self.zone_h_var).grid(row=1, column=3, pady=(5, 0))
        
        # Seite
        self.page_frame = ttk.Frame(self.zone_config_frame)
        ttk.Label(self.page_frame, text="Seite:").pack(side=tk.LEFT, padx=(0, 5))
        ttk.Spinbox(self.page_frame, from_=1, to=100, width=10,
                   textvariable=self.page_num_var).pack(side=tk.LEFT)
        
        # Zone-Expression
        self.zone_expr_frame = ttk.LabelFrame(self.zone_tab_frame, text="Zone-Verarbeitung", padding="10")
        self.zone_expr_label = ttk.Label(self.zone_expr_frame, 
                                        text="Optional: Ausdruck zur Verarbeitung des Zone-Texts (leer = nur Text)")
        self.zone_expr_text = tk.Text(self.zone_expr_frame, height=3, width=60)
        self.zone_expr_help = ttk.Label(self.zone_expr_frame, 
                                       text="Verwenden Sie <ZONE> als Platzhalter für den Zone-Text",
                                       foreground="gray")
        
        # Tab 3: Beispiele
        self.examples_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.examples_frame, text="Beispiele")
        
        # Beispiele Text
        self.examples_text = tk.Text(self.examples_frame, wrap=tk.WORD)
        self._fill_examples()
        self.examples_text.config(state=tk.DISABLED)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        
        # Events
        self.var_tree.bind("<Double-Button-1>", lambda e: self._insert_variable())
        self.func_tree.bind("<Double-Button-1>", lambda e: self._insert_function())
        
        # Lade verfügbare Variablen und Funktionen
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
        self.expr_label.pack(anchor=tk.W, pady=(10, 5))
        self.expr_text.pack(fill=tk.X, pady=(0, 10))
        
        self.builder_frame.pack(fill=tk.BOTH, expand=True)
        self.var_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        self.var_tree.pack(fill=tk.BOTH, expand=True)
        
        self.func_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.func_tree.pack(fill=tk.BOTH, expand=True)
        
        self.expr_button_frame.pack(fill=tk.X, pady=(10, 0))
        self.insert_var_button.pack(side=tk.LEFT, padx=(0, 5))
        self.insert_func_button.pack(side=tk.LEFT, padx=(0, 5))
        self.clear_expr_button.pack(side=tk.LEFT)
        
        # Zone Tab Layout
        self.zone_config_frame.pack(fill=tk.X, pady=(10, 10))
        self.zone_select_button.pack(pady=(0, 10))
        self.zone_manual_frame.pack(pady=(0, 10))
        self.page_frame.pack()
        
        self.zone_expr_frame.pack(fill=tk.X, pady=(10, 0))
        self.zone_expr_label.pack(anchor=tk.W)
        self.zone_expr_text.pack(fill=tk.X, pady=(5, 5))
        self.zone_expr_help.pack(anchor=tk.W)
        
        # Beispiele Tab Layout
        self.examples_text.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.save_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.cancel_button.pack(side=tk.RIGHT)
    
    def _load_variables_and_functions(self):
        """Lädt verfügbare Variablen und Funktionen"""
        # Variablen
        var_groups = self.xml_processor.get_available_variables()
        for group, variables in var_groups.items():
            group_node = self.var_tree.insert("", "end", text=group, open=True)
            for var in variables:
                self.var_tree.insert(group_node, "end", text=var)
        
        # Funktionen
        func_groups = self.xml_processor.get_available_functions()
        for group, functions in func_groups.items():
            group_node = self.func_tree.insert("", "end", text=group, open=True)
            for func in functions:
                self.func_tree.insert(group_node, "end", text=func["name"], 
                                     values=(func["syntax"],))
    
    def _fill_examples(self):
        """Füllt die Beispiele"""
        examples = """
BEISPIELE FÜR AUSDRÜCKE

1. EINFACHE VARIABLEN
   <FileName>                          - Dateiname ohne Erweiterung
   <Date>                             - Aktuelles Datum
   <OCR_FullText>                     - Kompletter OCR-Text

2. DATUMSFUNKTIONEN
   FORMATDATE("dd.mm.yyyy")           - 31.12.2024
   FORMATDATE("yyyy-mm-dd")           - 2024-12-31
   FORMATDATE("dd mmmm yyyy")         - 31 Dezember 2024

3. STRING-FUNKTIONEN
   LEFT("<FileName>", 5)              - Erste 5 Zeichen des Dateinamens
   RIGHT("<OCR_FullText>", 10)        - Letzte 10 Zeichen des OCR-Texts
   TRIM(" Text ")                     - Entfernt Leerzeichen
   TOUPPER("<FileName>")              - In Großbuchstaben
   MID("<FileName>", 3, 4)            - 4 Zeichen ab Position 3

4. BEDINGUNGEN
   IF("<FileName>", "contains", "Rechnung", "Invoice", "Document")
   IF(LEN("<FileName>"), ">", "10", "Langer Name", "Kurzer Name")

5. REGULÄRE AUSDRÜCKE
   REGEXP.MATCH("<OCR_FullText>", "Rechnungsnr[.:]\\s*(\\d+)", 1)
   REGEXP.MATCH("<OCR_FullText>", "(\\d{1,2}[.-/]\\d{1,2}[.-/]\\d{4})", 0)
   REGEXP.REPLACE("<FileName>", "[^a-zA-Z0-9]", "_")

6. KOMBINATIONEN
   "Rechnung_" + FORMATDATE("yyyymmdd")
   LEFT("<FileName>", 10) + "_" + FORMATDATE("HHMMss")
   IF(REGEXP.MATCH("<OCR_FullText>", "URGENT", 0), "!=", "", "URGENT_", "") + "<FileName>"

7. OCR-ZONEN
   <ZONE>                             - Zone-Text direkt verwenden
   TRIM("<ZONE>")                     - Zone-Text ohne Leerzeichen
   REGEXP.MATCH("<ZONE>", "(\\d+)", 1) - Erste Zahl aus Zone
"""
        self.examples_text.insert("1.0", examples)
    
    def _insert_variable(self):
        """Fügt die ausgewählte Variable ein"""
        selection = self.var_tree.selection()
        if selection:
            item = self.var_tree.item(selection[0])
            var_name = item['text']
            
            # Prüfe ob es eine Gruppe ist
            parent = self.var_tree.parent(selection[0])
            if parent:  # Ist eine Variable, keine Gruppe
                self.expr_text.insert(tk.INSERT, f"<{var_name}>")
    
    def _insert_function(self):
        """Fügt die ausgewählte Funktion ein"""
        selection = self.func_tree.selection()
        if selection:
            item = self.func_tree.item(selection[0])
            
            # Prüfe ob es eine Gruppe ist
            parent = self.func_tree.parent(selection[0])
            if parent:  # Ist eine Funktion, keine Gruppe
                syntax = item['values'][0]
                self.expr_text.insert(tk.INSERT, syntax)
    
    def _update_visibility(self):
        """Aktualisiert die Sichtbarkeit basierend auf dem Source-Typ"""
        # In dieser Version sind alle Tabs immer sichtbar
        pass
    
    def _select_zone_graphically(self):
        """Öffnet den grafischen Zone-Selector"""
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
        
        return True
    
    def _on_save(self):
        """Speichert die Eingaben"""
        if not self._validate():
            return
        
        # Bestimme Source-Type basierend auf aktuellem Tab
        current_tab = self.notebook.index(self.notebook.select())
        
        expression = ""
        source_type = "expression"
        zone = None
        
        if current_tab == 1:  # OCR-Zone Tab
            source_type = "ocr_zone"
            zone = (
                self.zone_x_var.get(),
                self.zone_y_var.get(),
                self.zone_w_var.get(),
                self.zone_h_var.get()
            )
            # Zone-Expression
            zone_expr = self.zone_expr_text.get("1.0", tk.END).strip()
            if zone_expr:
                expression = zone_expr
        else:  # Expression Tab
            expression = self.expr_text.get("1.0", tk.END).strip()
        
        self.result = {
            "field_name": self.field_name_var.get().strip(),
            "source_type": source_type,
            "expression": expression,
            "zone": zone,
            "page_num": self.page_num_var.get()
        }
        
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab"""
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result


class TestMappingDialog:
    """Dialog zum Testen eines Mappings"""
    
    def __init__(self, parent, mapping: FieldMapping, xml_processor: XMLFieldProcessor):
        self.parent = parent
        self.mapping = mapping
        self.xml_processor = xml_processor
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Mapping testen")
        self.dialog.geometry("600x400")
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
        
        # Test-Eingabe
        self.input_frame = ttk.LabelFrame(self.main_frame, text="Test-Variablen", padding="10")
        self.input_text = tk.Text(self.input_frame, height=6)
        self.input_text.insert("1.0", 
            "# Geben Sie Test-Werte für Variablen ein:\n"
            "FileName=TestDokument\n"
            "OCR_FullText=Dies ist ein Test-Text\n"
            "Date=31.12.2024\n")
        
        # Ergebnis
        self.result_frame = ttk.LabelFrame(self.main_frame, text="Ergebnis", padding="10")
        self.result_text = tk.Text(self.result_frame, height=4)
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
        
        self.input_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.input_text.pack(fill=tk.BOTH, expand=True)
        
        self.result_frame.pack(fill=tk.X, pady=(0, 10))
        self.result_text.pack(fill=tk.X)
        
        self.button_frame.pack(fill=tk.X)
        self.test_button.pack(side=tk.LEFT, padx=(0, 5))
        self.close_button.pack(side=tk.LEFT)
    
    def _run_test(self):
        """Führt den Test aus"""
        # Parse Test-Variablen
        context = {}
        for line in self.input_text.get("1.0", tk.END).strip().split("\n"):
            if "=" in line and not line.startswith("#"):
                key, value = line.split("=", 1)
                context[key.strip()] = value.strip()
        
        # Füge Standard-Variablen hinzu
        from core.function_parser import VariableExtractor
        context.update(VariableExtractor.get_standard_variables())
        
        # Evaluiere Expression
        try:
            from core.function_parser import FunctionParser
            parser = FunctionParser()
            result = parser.parse_and_evaluate(self.mapping.expression, context)
            
            self.result_text.config(state=tk.NORMAL)
            self.result_text.delete("1.0", tk.END)
            self.result_text.insert("1.0", result)
            self.result_text.config(state=tk.DISABLED)
            
        except Exception as e:
            self.result_text.config(state=tk.NORMAL)
            self.result_text.delete("1.0", tk.END)
            self.result_text.insert("1.0", f"Fehler: {e}")
            self.result_text.config(state=tk.DISABLED)
    
    def show(self):
        """Zeigt den Dialog"""
        self.dialog.wait_window()
