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
from gui.expression_editor_base import ExpressionEditorBase


class XMLFieldDialog:
    """Dialog zum Konfigurieren von XML-Feld-Mappings"""
    
    def __init__(self, parent, mappings: List[Dict] = [], ocr_zones: List[Dict] = []):
        self.parent = parent
        self.mappings = mappings or []
        self.ocr_zones = ocr_zones or []
        self.result = None
        self.xml_processor = XMLFieldProcessor()
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("XML-Felder konfigurieren")
        self.dialog.geometry("1000x700")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._create_widgets()
        self._layout_widgets()
        self._load_mappings()
        
        # Bind Events
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
        
        # Info über verfügbare OCR-Zonen
        if self.ocr_zones:
            self.info_frame = ttk.LabelFrame(self.main_frame, text="Verfügbare OCR-Zonen", padding="5")
            zone_names = [zone.get('name', f"Zone_{i+1}") for i, zone in enumerate(self.ocr_zones)]
            info_text = f"Definierte Zonen: {', '.join(zone_names)}"
            self.info_label = ttk.Label(self.info_frame, text=info_text, foreground="blue")
        
        # Toolbar
        self.toolbar = ttk.Frame(self.main_frame)
        self.add_button = ttk.Button(self.toolbar, text="➕ Neues Feld", 
                                    command=self._add_field)
        self.edit_button = ttk.Button(self.toolbar, text="✏️ Bearbeiten", 
                                     command=self._edit_field, state=tk.DISABLED)
        self.delete_button = ttk.Button(self.toolbar, text="🗑️ Löschen", 
                                       command=self._delete_field, state=tk.DISABLED)
        self.test_button = ttk.Button(self.toolbar, text="🧪 Testen", 
                                     command=self._test_all_mappings)
        
        # Separator
        self.separator = ttk.Separator(self.toolbar, orient=tk.VERTICAL)
        
        # Sortier-Buttons
        self.move_up_button = ttk.Button(self.toolbar, text="⬆", width=3,
                                        command=self._move_up, state=tk.DISABLED)
        self.move_down_button = ttk.Button(self.toolbar, text="⬇", width=3,
                                          command=self._move_down, state=tk.DISABLED)
        
        # Liste der Mappings
        self.tree_frame = ttk.Frame(self.main_frame)
        self.tree = ttk.Treeview(self.tree_frame, 
                                columns=("Field", "Description", "Expression"),
                                show="tree headings", height=20)
        
        # Spalten konfigurieren
        self.tree.heading("#0", text="")
        self.tree.heading("Field", text="XML-Feld")
        self.tree.heading("Description", text="Beschreibung")
        self.tree.heading("Expression", text="Ausdruck/Funktion")
        
        self.tree.column("#0", width=30)
        self.tree.column("Field", width=200)
        self.tree.column("Description", width=250)
        self.tree.column("Expression", width=400)
        
        # Scrollbar
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", 
                                command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.vsb.set)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save)
        
        # Events
        self.tree.bind("<<TreeviewSelect>>", self._on_selection_changed)
        self.tree.bind("<Double-Button-1>", lambda e: self._edit_field())
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Info über OCR-Zonen
        if hasattr(self, 'info_frame'):
            self.info_frame.pack(fill=tk.X, pady=(0, 10))
            self.info_label.pack(anchor=tk.W)
        
        # Toolbar
        self.toolbar.pack(fill=tk.X, pady=(0, 10))
        self.add_button.pack(side=tk.LEFT, padx=(0, 5))
        self.edit_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_button.pack(side=tk.LEFT, padx=(0, 5))
        self.test_button.pack(side=tk.LEFT)
        
        self.separator.pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        self.move_up_button.pack(side=tk.LEFT, padx=(0, 5))
        self.move_down_button.pack(side=tk.LEFT)
        
        # Tree
        self.tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.save_button.pack(side=tk.RIGHT)
    
    def _load_mappings(self):
        """Lädt die Mappings in die Liste"""
        for mapping_dict in self.mappings:
            mapping = FieldMapping.from_dict(mapping_dict)
            self._add_to_tree(mapping, mapping_dict.get('description', ''))
    
    def _add_to_tree(self, mapping: FieldMapping, description: str = ''):
        """Fügt ein Mapping zur Baumanzeige hinzu"""
        expression_text = mapping.expression[:60] + "..." if len(mapping.expression) > 60 else mapping.expression
        
        self.tree.insert("", "end", 
                        values=(mapping.field_name, description, expression_text))
    
    def _on_selection_changed(self, event):
        """Wird aufgerufen wenn die Auswahl sich ändert"""
        selection = self.tree.selection()
        if selection:
            self.edit_button.config(state=tk.NORMAL)
            self.delete_button.config(state=tk.NORMAL)
            self.test_button.config(state=tk.NORMAL)
            
            # Prüfe ob Bewegung möglich ist
            index = self.tree.index(selection[0])
            self.move_up_button.config(state=tk.NORMAL if index > 0 else tk.DISABLED)
            self.move_down_button.config(state=tk.NORMAL if index < len(self.mappings) - 1 else tk.DISABLED)
        else:
            self.edit_button.config(state=tk.DISABLED)
            self.delete_button.config(state=tk.DISABLED)
            self.test_button.config(state=tk.DISABLED)
            self.move_up_button.config(state=tk.DISABLED)
            self.move_down_button.config(state=tk.DISABLED)
    
    def _move_up(self):
        """Bewegt das ausgewählte Feld nach oben"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if index > 0:
            # Tausche in der Liste
            self.mappings[index], self.mappings[index-1] = self.mappings[index-1], self.mappings[index]
            
            # Aktualisiere Tree
            self._refresh_tree()
            
            # Behalte Auswahl
            new_items = self.tree.get_children()
            if index-1 < len(new_items):
                self.tree.selection_set(new_items[index-1])
                self.tree.focus(new_items[index-1])
    
    def _move_down(self):
        """Bewegt das ausgewählte Feld nach unten"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if index < len(self.mappings) - 1:
            # Tausche in der Liste
            self.mappings[index], self.mappings[index+1] = self.mappings[index+1], self.mappings[index]
            
            # Aktualisiere Tree
            self._refresh_tree()
            
            # Behalte Auswahl
            new_items = self.tree.get_children()
            if index+1 < len(new_items):
                self.tree.selection_set(new_items[index+1])
                self.tree.focus(new_items[index+1])
    
    def _refresh_tree(self):
        """Aktualisiert die komplette Tree-Anzeige"""
        # Lösche alle Einträge
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Füge alle wieder hinzu
        for mapping_dict in self.mappings:
            mapping = FieldMapping.from_dict(mapping_dict)
            self._add_to_tree(mapping, mapping_dict.get('description', ''))
    
    def _add_field(self):
        """Fügt ein neues Feld hinzu"""
        # Übergebe bereits definierte Mappings als xml_field_mappings
        dialog = FieldMappingEditDialog(
            self.dialog, 
            xml_processor=self.xml_processor, 
            ocr_zones=self.ocr_zones,
            xml_field_mappings=self.mappings  # WICHTIG: Übergebe bestehende Mappings
        )
        result = dialog.show()
        
        if result:
            # Füge description zum mapping_dict hinzu
            mapping_dict = result.copy()
            self.mappings.append(mapping_dict)
            
            # Für Anzeige
            mapping = FieldMapping(**{k: v for k, v in result.items() if k != 'description'})
            self._add_to_tree(mapping, result.get('description', ''))
    
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
            
            # Erstelle Liste der anderen Mappings (ohne das aktuell bearbeitete)
            other_mappings = [m for i, m in enumerate(self.mappings) if i != index]
            
            dialog = FieldMappingEditDialog(
                self.dialog, 
                mapping, 
                self.xml_processor,
                description=mapping_dict.get('description', ''),
                ocr_zones=self.ocr_zones,
                xml_field_mappings=other_mappings  # WICHTIG: Übergebe andere Mappings
            )
            result = dialog.show()
            
            if result:
                # Aktualisiere Mapping
                self.mappings[index] = result
                
                # Aktualisiere Tree
                self._refresh_tree()
    
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
    
    def _test_all_mappings(self):
        """Öffnet einen Dialog zum Testen aller Mappings mit Datei-Upload"""
        dialog = AllMappingsTestDialog(self.dialog, self.mappings, self.xml_processor, self.ocr_zones)
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


class FieldMappingEditDialog(ExpressionEditorBase):
    """Dialog zum Bearbeiten eines einzelnen Field-Mappings - erweitert die Basis-Klasse"""
    
    def __init__(self, parent, mapping: Optional[FieldMapping] = None, 
                 xml_processor: Optional[XMLFieldProcessor] = None, description: str = "",
                 ocr_zones: List[Dict] = [], xml_field_mappings: List[Dict] = []):
        self.mapping = mapping
        self.xml_processor = xml_processor or XMLFieldProcessor()
        
        # Variablen für zusätzliche Felder
        self.field_name_var = tk.StringVar(value=mapping.field_name if mapping else "")
        self.description_var = tk.StringVar(value=description)
        
        # Initialisiere mit der Basis-Klasse
        super().__init__(
            parent=parent,
            title="Feld-Mapping bearbeiten",
            expression=mapping.expression if mapping else "",
            description="",  # Keine Beschreibung hier, da wir eigene Felder haben
            ocr_zones=ocr_zones,
            geometry="1200x800",
            xml_field_mappings=xml_field_mappings  # WICHTIG: Übergebe XML-Feld-Mappings
        )
        
        # Fokus auf Feldname
        self.field_name_entry.focus()
    
    def _create_additional_widgets(self):
        """Erstellt zusätzliche Widgets für Feld-Konfiguration"""
        # Feld-Konfiguration
        self.field_frame = ttk.LabelFrame(self.main_frame, text="Feld-Konfiguration", padding="10")
        
        # Feldname
        self.field_name_label = ttk.Label(self.field_frame, text="XML-Feldname:")
        self.field_name_entry = ttk.Entry(self.field_frame, textvariable=self.field_name_var, width=40)
        
        # Beschreibung
        self.desc_label = ttk.Label(self.field_frame, text="Beschreibung:")
        self.desc_entry = ttk.Entry(self.field_frame, textvariable=self.description_var, width=40)
    
    def _layout_additional_widgets(self):
        """Layoutet die zusätzlichen Widgets"""
        # Feld-Konfiguration
        self.field_frame.pack(fill=tk.X, pady=(0, 10))
        self.field_name_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.field_name_entry.grid(row=0, column=1, sticky="we")
        self.desc_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.desc_entry.grid(row=1, column=1, sticky="we", pady=(5, 0))
        self.field_frame.columnconfigure(1, weight=1)
    
    def _create_buttons(self):
        """Erstellt spezielle Buttons für Feld-Mapping"""
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save)
    
    def _layout_buttons(self):
        """Layoutet die Buttons"""
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.save_button.pack(side=tk.RIGHT)
    
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
        
        # ÄNDERUNG: Leere Ausdrücke sind jetzt erlaubt
        # Keine Validierung des Ausdrucks mehr notwendig
        
        return True
    
    def _on_save(self):
        """Speichert die Eingaben - überschreibt die Basis-Methode"""
        if not self._validate():
            return
        
        expression = self.expr_text.get("1.0", tk.END).strip()
        
        self.result = {
            "field_name": self.field_name_var.get().strip(),
            "description": self.description_var.get().strip(),
            "source_type": "expression",
            "expression": expression,  # Kann jetzt auch leer sein
            "zone": None,
            "page_num": 1,
            "zones": []  # Keine Zonen hier, da sie vom Hotfolder kommen
        }
        
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis - überschreibt Return-Type"""
        self.dialog.wait_window()
        return self.result


class AllMappingsTestDialog:
    def __init__(self, parent, mappings: List[Dict], xml_processor: XMLFieldProcessor, ocr_zones: List[Dict] = []):
        self.parent = parent
        self.mappings = mappings
        self.xml_processor = xml_processor
        self.ocr_zones = ocr_zones
        self.result = None
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Alle Felder testen")
        self.dialog.geometry("800x600")
        self.dialog.resizable(True, True)
        self._center_window()
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self._create_widgets()
        self._layout_widgets()
    def _center_window(self):
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() - width) // 2
        y = (self.dialog.winfo_screenheight() - height) // 2
        self.dialog.geometry(f"+{x}+{y}")
    def _create_widgets(self):
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        # Test-Dateien
        self.files_frame = ttk.LabelFrame(self.main_frame, text="Test-Dateien", padding="10")
        # PDF
        self.pdf_frame = ttk.Frame(self.files_frame)
        self.pdf_label = ttk.Label(self.pdf_frame, text="PDF-Datei:")
        self.pdf_path_var = tk.StringVar()
        self.pdf_entry = ttk.Entry(self.pdf_frame, textvariable=self.pdf_path_var, width=40, state='readonly')
        self.pdf_button = ttk.Button(self.pdf_frame, text="Durchsuchen...", command=self._select_pdf)
        # XML
        self.xml_frame = ttk.Frame(self.files_frame)
        self.xml_label = ttk.Label(self.xml_frame, text="XML-Datei (optional):")
        self.xml_path_var = tk.StringVar()
        self.xml_entry = ttk.Entry(self.xml_frame, textvariable=self.xml_path_var, width=40, state='readonly')
        self.xml_button = ttk.Button(self.xml_frame, text="Durchsuchen...", command=self._select_xml)
        # Ergebnis
        self.result_frame = ttk.LabelFrame(self.main_frame, text="Ergebnisse aller Felder", padding="10")
        self.result_text = tk.Text(self.result_frame, height=20)
        self.result_text.config(state=tk.DISABLED)
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.test_button = ttk.Button(self.button_frame, text="Testen", command=self._run_test)
        self.close_button = ttk.Button(self.button_frame, text="Schließen", command=self.dialog.destroy)
    def _layout_widgets(self):
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.files_frame.pack(fill=tk.X, pady=(0, 10))
        self.pdf_frame.pack(fill=tk.X, pady=(0, 5))
        self.pdf_label.pack(side=tk.LEFT, padx=(0, 5))
        self.pdf_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.pdf_button.pack(side=tk.LEFT)
        self.xml_frame.pack(fill=tk.X)
        self.xml_label.pack(side=tk.LEFT, padx=(0, 5))
        self.xml_entry.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 5))
        self.xml_button.pack(side=tk.LEFT)
        self.result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.result_text.pack(fill=tk.BOTH, expand=True)
        self.button_frame.pack(fill=tk.X)
        self.test_button.pack(side=tk.LEFT, padx=(0, 5))
        self.close_button.pack(side=tk.LEFT)
    def _select_pdf(self):
        filename = filedialog.askopenfilename(
            title="PDF für Test auswählen",
            filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.pdf_path_var.set(filename)
    def _select_xml(self):
        filename = filedialog.askopenfilename(
            title="XML für Test auswählen",
            filetypes=[("XML Dateien", "*.xml"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.xml_path_var.set(filename)
    def _run_test(self):
        pdf_path = self.pdf_path_var.get()
        if not pdf_path:
            messagebox.showerror("Fehler", "Bitte wählen Sie eine PDF-Datei für den Test.")
            return
        xml_path = self.xml_path_var.get() or None
        # Baue Kontext auf
        mapping_objs = [FieldMapping.from_dict(m) for m in self.mappings]
        context = self.xml_processor._build_context(
            xml_path or "",
            pdf_path,
            mapping_objs,
            self.ocr_zones
        )
        # Ergebnisse für alle Felder berechnen
        results = {}
        for mapping in mapping_objs:
            try:
                value = self.xml_processor._evaluate_mapping(mapping, context)
            except Exception as e:
                value = f"Fehler: {e}"
            results[mapping.field_name] = value
        # Ergebnisse anzeigen
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete("1.0", tk.END)
        for field, value in results.items():
            if value is None:
                value = "(kein Ergebnis)"
            elif isinstance(value, str) and len(value) > 200:
                value = value[:200] + "..."
            self.result_text.insert(tk.END, f"{field}: {value}\n")
        self.result_text.config(state=tk.DISABLED)
    def show(self):
        self.dialog.wait_window()