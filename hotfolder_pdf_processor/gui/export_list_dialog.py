"""
Dialog zur Verwaltung mehrerer Export-Konfigurationen
"""
import tkinter as tk
from tkinter import ttk, messagebox
from typing import List, Dict, Optional
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.export_action import ExportConfig
from gui.export_dialog import ExportDialog


class ExportListDialog:
    """Dialog zur Verwaltung der Export-Konfigurationen eines Hotfolders"""
    
    def __init__(self, parent, exports: List[Dict] = None,
                 ocr_zones: List[Dict] = None, xml_field_mappings: List[Dict] = None):
        self.parent = parent
        self.exports = exports or []
        self.ocr_zones = ocr_zones or []
        self.xml_field_mappings = xml_field_mappings or []
        self.result = None
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Export-Konfigurationen verwalten")
        self.dialog.geometry("800x600")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._create_widgets()
        self._layout_widgets()
        self._refresh_list()
        
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
        
        # Info
        info_text = "Konfigurieren Sie mehrere Export-Ausgaben für diesen Hotfolder.\nJeder Export kann unterschiedliche Ziele und Formate haben."
        self.info_label = ttk.Label(self.main_frame, text=info_text, wraplength=750)
        
        # Toolbar
        self.toolbar = ttk.Frame(self.main_frame)
        self.add_button = ttk.Button(self.toolbar, text="➕ Neuer Export", command=self._add_export)
        self.edit_button = ttk.Button(self.toolbar, text="✏️ Bearbeiten", command=self._edit_export, state=tk.DISABLED)
        self.delete_button = ttk.Button(self.toolbar, text="🗑️ Löschen", command=self._delete_export, state=tk.DISABLED)
        self.duplicate_button = ttk.Button(self.toolbar, text="📋 Duplizieren", command=self._duplicate_export, state=tk.DISABLED)
        
        ttk.Separator(self.toolbar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        self.up_button = ttk.Button(self.toolbar, text="⬆", width=3, command=self._move_up, state=tk.DISABLED)
        self.down_button = ttk.Button(self.toolbar, text="⬇", width=3, command=self._move_down, state=tk.DISABLED)
        
        # Liste
        self.tree_frame = ttk.Frame(self.main_frame)
        self.tree = ttk.Treeview(self.tree_frame, 
                                columns=("Name", "Type", "Target", "Status"),
                                show="tree headings", height=15)
        
        # Spalten konfigurieren
        self.tree.heading("#0", text="")
        self.tree.heading("Name", text="Export-Name")
        self.tree.heading("Type", text="Typ")
        self.tree.heading("Target", text="Ziel")
        self.tree.heading("Status", text="Status")
        
        self.tree.column("#0", width=30)
        self.tree.column("Name", width=200)
        self.tree.column("Type", width=100)
        self.tree.column("Target", width=300)
        self.tree.column("Status", width=100)
        
        # Scrollbar
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.vsb.set)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", command=self._on_cancel)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", command=self._on_save)
        
        # Events
        self.tree.bind("<<TreeviewSelect>>", self._on_selection_changed)
        self.tree.bind("<Double-Button-1>", lambda e: self._edit_export())
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.info_label.pack(anchor=tk.W, pady=(0, 10))
        
        # Toolbar
        self.toolbar.pack(fill=tk.X, pady=(0, 10))
        self.add_button.pack(side=tk.LEFT, padx=(0, 5))
        self.edit_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_button.pack(side=tk.LEFT, padx=(0, 5))
        self.duplicate_button.pack(side=tk.LEFT)
        
        ttk.Separator(self.toolbar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        self.up_button.pack(side=tk.LEFT, padx=(0, 5))
        self.down_button.pack(side=tk.LEFT)
        
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
    
    def _refresh_list(self):
        """Aktualisiert die Export-Liste"""
        # Lösche aktuelle Einträge
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Füge Exporte hinzu
        for export_dict in self.exports:
            export = ExportConfig.from_dict(export_dict)
            
            # Icon basierend auf Typ
            icon = "📄"  # Datei
            if export.export_type.value == "email":
                icon = "✉️"
            elif export.export_type.value == "script":
                icon = "⚙️"
            
            # Typ-Anzeige
            type_display = {
                "file": "Dateiausgabe",
                "email": "E-Mail",
                "script": "Programm"
            }.get(export.export_type.value, export.export_type.value)
            
            # Ziel bestimmen
            target = ""
            if export.export_type.value == "file":
                target = export.output_path_expression
            elif export.export_type.value == "email" and export.email_config:
                target = export.email_config.to_expression
            elif export.export_type.value == "script" and export.programs:
                target = f"{len(export.programs)} Programme"
            
            # Status
            status = "Aktiv" if export.enabled else "Inaktiv"
            
            # Füge zur Liste hinzu
            item = self.tree.insert("", "end", 
                                   text=icon,
                                   values=(export.name, type_display, target, status))
            
            # Färbe inaktive grau
            if not export.enabled:
                self.tree.item(item, tags=("disabled",))
        
        # Style für disabled
        self.tree.tag_configure("disabled", foreground="gray")
    
    def _on_selection_changed(self, event):
        """Wird aufgerufen wenn die Auswahl sich ändert"""
        selection = self.tree.selection()
        if selection:
            self.edit_button.config(state=tk.NORMAL)
            self.delete_button.config(state=tk.NORMAL)
            self.duplicate_button.config(state=tk.NORMAL)
            
            # Prüfe ob Bewegung möglich ist
            index = self.tree.index(selection[0])
            self.up_button.config(state=tk.NORMAL if index > 0 else tk.DISABLED)
            self.down_button.config(state=tk.NORMAL if index < len(self.exports) - 1 else tk.DISABLED)
        else:
            self.edit_button.config(state=tk.DISABLED)
            self.delete_button.config(state=tk.DISABLED)
            self.duplicate_button.config(state=tk.DISABLED)
            self.up_button.config(state=tk.DISABLED)
            self.down_button.config(state=tk.DISABLED)
    
    def _add_export(self):
        """Fügt einen neuen Export hinzu"""
        dialog = ExportDialog(self.dialog, 
                            ocr_zones=self.ocr_zones,
                            xml_field_mappings=self.xml_field_mappings)
        result = dialog.show()
        
        if result:
            self.exports.append(result)
            self._refresh_list()
    
    def _edit_export(self):
        """Bearbeitet den ausgewählten Export"""
        selection = self.tree.selection()
        if not selection:
            return
        
        index = self.tree.index(selection[0])
        if 0 <= index < len(self.exports):
            export_config = ExportConfig.from_dict(self.exports[index])
            
            dialog = ExportDialog(self.dialog, export_config,
                                ocr_zones=self.ocr_zones,
                                xml_field_mappings=self.xml_field_mappings)
            result = dialog.show()
            
            if result:
                self.exports[index] = result
                self._refresh_list()
    
    def _delete_export(self):
        """Löscht den ausgewählten Export"""
        selection = self.tree.selection()
        if not selection:
            return
        
        index = self.tree.index(selection[0])
        if 0 <= index < len(self.exports):
            export_name = self.exports[index].get('name', 'Export')
            
            if messagebox.askyesno("Export löschen", 
                                  f"Möchten Sie den Export '{export_name}' wirklich löschen?"):
                del self.exports[index]
                self._refresh_list()
    
    def _duplicate_export(self):
        """Dupliziert den ausgewählten Export"""
        selection = self.tree.selection()
        if not selection:
            return
        
        index = self.tree.index(selection[0])
        if 0 <= index < len(self.exports):
            # Erstelle Kopie
            import copy
            import uuid
            
            export_copy = copy.deepcopy(self.exports[index])
            export_copy['id'] = str(uuid.uuid4())
            export_copy['name'] = export_copy['name'] + " (Kopie)"
            
            self.exports.insert(index + 1, export_copy)
            self._refresh_list()
    
    def _move_up(self):
        """Bewegt Export nach oben"""
        selection = self.tree.selection()
        if not selection:
            return
        
        index = self.tree.index(selection[0])
        if index > 0:
            self.exports[index], self.exports[index-1] = self.exports[index-1], self.exports[index]
            self._refresh_list()
            
            # Behalte Auswahl
            new_items = self.tree.get_children()
            if index-1 < len(new_items):
                self.tree.selection_set(new_items[index-1])
                self.tree.focus(new_items[index-1])
    
    def _move_down(self):
        """Bewegt Export nach unten"""
        selection = self.tree.selection()
        if not selection:
            return
        
        index = self.tree.index(selection[0])
        if index < len(self.exports) - 1:
            self.exports[index], self.exports[index+1] = self.exports[index+1], self.exports[index]
            self._refresh_list()
            
            # Behalte Auswahl
            new_items = self.tree.get_children()
            if index+1 < len(new_items):
                self.tree.selection_set(new_items[index+1])
                self.tree.focus(new_items[index+1])
    
    def _on_save(self):
        """Speichert die Export-Liste"""
        self.result = self.exports
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab ohne zu speichern"""
        self.dialog.destroy()
    
    def show(self) -> Optional[List[Dict]]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result