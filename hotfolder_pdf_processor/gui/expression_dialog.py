"""
Dialog zur Bearbeitung von Ausdrücken mit Variablen und Funktionen
"""
import tkinter as tk
from tkinter import ttk
from typing import Optional
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.xml_field_processor import XMLFieldProcessor


class ExpressionDialog:
    """Dialog zur Bearbeitung von Ausdrücken"""
    
    def __init__(self, parent, title: str = "Ausdruck bearbeiten", 
                 expression: str = "", description: str = ""):
        self.parent = parent
        self.expression = expression
        self.result = None
        self.xml_processor = XMLFieldProcessor()
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry("800x600")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.description = description
        
        self._create_widgets()
        self._layout_widgets()
        
        # Setze initialen Ausdruck
        if expression:
            self.expr_text.insert("1.0", expression)
        
        # Fokus
        self.expr_text.focus()
        
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
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Hauptframe
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Beschreibung
        if self.description:
            self.desc_label = ttk.Label(self.main_frame, text=self.description)
        
        # Expression-Eingabe
        self.expr_frame = ttk.LabelFrame(self.main_frame, text="Ausdruck", padding="10")
        self.expr_text = tk.Text(self.expr_frame, height=5, width=70)
        
        # Variablen und Funktionen Tree
        self.tree_frame = ttk.LabelFrame(self.main_frame, 
                                        text="Verfügbare Variablen und Funktionen", 
                                        padding="10")
        
        # Tree
        self.var_func_tree = ttk.Treeview(self.tree_frame, height=15)
        self.var_func_tree.heading("#0", text="Variablen und Funktionen")
        
        # Scrollbar
        self.tree_scroll = ttk.Scrollbar(self.tree_frame, orient="vertical",
                                        command=self.var_func_tree.yview)
        self.var_func_tree.configure(yscrollcommand=self.tree_scroll.set)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        self.ok_button = ttk.Button(self.button_frame, text="OK", 
                                   command=self._on_ok, default=tk.ACTIVE)
        
        # Events
        self.var_func_tree.bind("<Double-Button-1>", self._on_tree_double_click)
        
        # Lade Variablen und Funktionen
        self._load_variables_and_functions()
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Beschreibung
        if hasattr(self, 'desc_label'):
            self.desc_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Expression
        self.expr_frame.pack(fill=tk.X, pady=(0, 10))
        self.expr_text.pack(fill=tk.BOTH, expand=True)
        
        # Tree
        self.tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.var_func_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.ok_button.pack(side=tk.RIGHT)
    
    def _load_variables_and_functions(self):
        """Lädt verfügbare Variablen und Funktionen"""
        # Variablen-Hauptknoten
        var_root = self.var_func_tree.insert("", "end", text="Variablen", open=True)
        
        # Standard-Variablen
        std_node = self.var_func_tree.insert(var_root, "end", text="Standard", open=True)
        std_vars = ["Date", "Time", "DateTime", "Year", "Month", "Day", 
                   "Hour", "Minute", "Second", "Weekday", "WeekNumber"]
        for var in std_vars:
            self.var_func_tree.insert(std_node, "end", text=var, tags=("variable",))
        
        # Datei-Variablen
        file_node = self.var_func_tree.insert(var_root, "end", text="Datei", open=True)
        file_vars = ["FileName", "FileExtension", "FullFileName", "FilePath", 
                    "FullPath", "FileSize"]
        for var in file_vars:
            self.var_func_tree.insert(file_node, "end", text=var, tags=("variable",))
        
        # OCR-Variablen
        ocr_node = self.var_func_tree.insert(var_root, "end", text="OCR", open=True)
        self.var_func_tree.insert(ocr_node, "end", text="OCR_FullText", tags=("variable",))
        
        # Funktionen-Hauptknoten
        func_root = self.var_func_tree.insert("", "end", text="Funktionen", open=True)
        
        # String-Funktionen
        string_node = self.var_func_tree.insert(func_root, "end", text="String", open=True)
        string_funcs = [
            ("FORMAT", 'FORMAT("<VAR>","<FORMATSTRING>")'),
            ("TRIM", 'TRIM("<VAR>")'),
            ("LEFT", 'LEFT("<VAR>",<LENGTH>)'),
            ("RIGHT", 'RIGHT("<VAR>",<LENGTH>)'),
            ("MID", 'MID("<VAR>",<START>,<LENGTH>)'),
            ("TOUPPER", 'TOUPPER("<VAR>")'),
            ("TOLOWER", 'TOLOWER("<VAR>")'),
            ("LEN", 'LEN("<VAR>")'),
        ]
        for name, syntax in string_funcs:
            self.var_func_tree.insert(string_node, "end", text=name, 
                                     tags=("function", syntax))
        
        # Datumsfunktionen
        date_node = self.var_func_tree.insert(func_root, "end", text="Datum", open=True)
        self.var_func_tree.insert(date_node, "end", text="FORMATDATE", 
                                 tags=("function", 'FORMATDATE("<FORMAT>")'))
        
        # Weitere Funktionskategorien nach Bedarf...
    
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
    
    def _on_ok(self):
        """Speichert den Ausdruck"""
        self.result = self.expr_text.get("1.0", tk.END).strip()
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab"""
        self.dialog.destroy()
    
    def show(self) -> Optional[str]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result