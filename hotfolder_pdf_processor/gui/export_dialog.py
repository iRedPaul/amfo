"""
Dialog zur Konfiguration von Exporten - Verbesserte Version
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import List, Dict, Optional
import sys
import os
import uuid

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.export_config import ExportConfig, ExportFormat, ExportMethod, EmailConfig
from gui.expression_dialog import ExpressionDialog


class ExportDialog:
    """Dialog zum Konfigurieren von Exporten"""
    
    def __init__(self, parent, export_configs: List[Dict] = None, 
                 error_path: str = "", xml_field_mappings: List[Dict] = None):
        self.parent = parent
        self.export_configs = export_configs or []
        self.error_path = error_path
        self.xml_field_mappings = xml_field_mappings or []
        self.result = None
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Export-Einstellungen")
        self.dialog.geometry("1100x700")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self._create_widgets()
        self._layout_widgets()
        self._load_exports()
        
        # Bind Events
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())
    
    def _center_window(self):
        """Zentriert das Fenster relativ zum Parent"""
        self.dialog.update_idletasks()
        
        parent_x = self.parent.winfo_x()
        parent_y = self.parent.winfo_y()
        parent_width = self.parent.winfo_width()
        parent_height = self.parent.winfo_height()
        
        dialog_width = self.dialog.winfo_width()
        dialog_height = self.dialog.winfo_height()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        x = max(0, x)
        y = max(0, y)
        
        self.dialog.geometry(f"+{x}+{y}")
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Hauptframe
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Export-Liste
        self.export_frame = ttk.LabelFrame(self.main_frame, text="Export-Konfigurationen", padding="10")
        
        # Toolbar
        self.toolbar = ttk.Frame(self.export_frame)
        self.add_button = ttk.Button(self.toolbar, text="➕ Neuer Export", 
                                    command=self._add_export)
        self.edit_button = ttk.Button(self.toolbar, text="✏️ Bearbeiten", 
                                     command=self._edit_export, state=tk.DISABLED)
        self.duplicate_button = ttk.Button(self.toolbar, text="📋 Duplizieren", 
                                          command=self._duplicate_export, state=tk.DISABLED)
        self.delete_button = ttk.Button(self.toolbar, text="🗑️ Löschen", 
                                       command=self._delete_export, state=tk.DISABLED)
        
        ttk.Separator(self.toolbar, orient=tk.VERTICAL).pack(side=tk.LEFT, fill=tk.Y, padx=10)
        
        self.move_up_button = ttk.Button(self.toolbar, text="⬆", width=3,
                                        command=self._move_up, state=tk.DISABLED)
        self.move_down_button = ttk.Button(self.toolbar, text="⬇", width=3,
                                          command=self._move_down, state=tk.DISABLED)
        
        # Export-Liste
        self.tree_frame = ttk.Frame(self.export_frame)
        self.tree = ttk.Treeview(self.tree_frame, 
                                columns=("Name", "Aktiv", "Methode", "Format", "Pfad"),
                                show="tree headings", height=12)
        
        # Spalten konfigurieren
        self.tree.heading("#0", text="")
        self.tree.heading("Name", text="Name")
        self.tree.heading("Aktiv", text="Aktiv")
        self.tree.heading("Methode", text="Methode")
        self.tree.heading("Format", text="Format")
        self.tree.heading("Pfad", text="Pfad/Ziel")
        
        self.tree.column("#0", width=30)
        self.tree.column("Name", width=200)
        self.tree.column("Aktiv", width=60, anchor=tk.CENTER)
        self.tree.column("Methode", width=100)
        self.tree.column("Format", width=100)
        self.tree.column("Pfad", width=400)
        
        # Scrollbar
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", 
                                command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.vsb.set)
        
        # Fehler-Pfad
        self.error_frame = ttk.LabelFrame(self.main_frame, text="Fehlerbehandlung", padding="10")
        
        self.error_path_label = ttk.Label(self.error_frame, 
            text="Fehlerpfad (optional - leer = Standard aus Einstellungen):")
        self.error_path_desc = ttk.Label(self.error_frame,
            text="Dateien, die nicht verarbeitet werden können, werden in diesen Ordner verschoben.",
            foreground="gray", font=('TkDefaultFont', 9))
        
        self.error_path_frame = ttk.Frame(self.error_frame)
        self.error_path_var = tk.StringVar(value=self.error_path)
        self.error_path_entry = ttk.Entry(self.error_path_frame, 
                                         textvariable=self.error_path_var, width=50)
        self.error_path_button = ttk.Button(self.error_path_frame, text="Durchsuchen...", 
                                           command=self._browse_error_path)
        
        # Info
        self.info_label = ttk.Label(self.main_frame, 
            text="Hinweis: Exporte werden in der angegebenen Reihenfolge ausgeführt. "
                 "Verwenden Sie Ausdrücke für dynamische Pfade und Dateinamen.",
            wraplength=1000, foreground="gray")
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save)
        
        # Events
        self.tree.bind("<<TreeviewSelect>>", self._on_selection_changed)
        self.tree.bind("<Double-Button-1>", lambda e: self._edit_export())
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Export-Liste
        self.export_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        self.toolbar.pack(fill=tk.X, pady=(0, 10))
        self.add_button.pack(side=tk.LEFT, padx=(0, 5))
        self.edit_button.pack(side=tk.LEFT, padx=(0, 5))
        self.duplicate_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_button.pack(side=tk.LEFT)
        
        self.move_up_button.pack(side=tk.LEFT, padx=(0, 5))
        self.move_down_button.pack(side=tk.LEFT)
        
        self.tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)
        
        # Fehler-Pfad
        self.error_frame.pack(fill=tk.X, pady=(0, 10))
        self.error_path_label.pack(anchor=tk.W)
        self.error_path_desc.pack(anchor=tk.W, pady=(2, 8))
        self.error_path_frame.pack(fill=tk.X)
        self.error_path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.error_path_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Info
        self.info_label.pack(fill=tk.X, pady=(0, 10))
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.save_button.pack(side=tk.RIGHT)
    
    def _load_exports(self):
        """Lädt die Export-Konfigurationen in die Liste"""
        for export_dict in self.export_configs:
            export = ExportConfig.from_dict(export_dict)
            self._add_to_tree(export)
    
    def _add_to_tree(self, export: ExportConfig):
        """Fügt einen Export zur Baumanzeige hinzu"""
        aktiv = "✓" if export.enabled else "✗"
        
        methode = {
            ExportMethod.FILE: "Datei",
            ExportMethod.EMAIL: "E-Mail",
            ExportMethod.FTP: "FTP",
            ExportMethod.WEBDAV: "WebDAV"
        }.get(export.export_method, export.export_method.value)
        
        format_name = {
            ExportFormat.PDF: "PDF",
            ExportFormat.PDF_A: "PDF/A",
            ExportFormat.SEARCHABLE_PDF_A: "Durchsuchbares PDF/A",
            ExportFormat.PNG: "PNG",
            ExportFormat.JPG: "JPG",
            ExportFormat.TIFF: "TIFF",
            ExportFormat.XML: "XML",
            ExportFormat.JSON: "JSON",
            ExportFormat.TXT: "Text",
            ExportFormat.CSV: "CSV"
        }.get(export.export_format, export.export_format.value)
        
        # Ziel je nach Methode
        if export.export_method == ExportMethod.EMAIL and export.email_config:
            ziel = export.email_config.recipient
        else:
            ziel = export.export_path_expression[:50] + "..." if len(export.export_path_expression) > 50 else export.export_path_expression
        
        self.tree.insert("", "end", values=(export.name, aktiv, methode, format_name, ziel))
    
    def _on_selection_changed(self, event):
        """Wird aufgerufen wenn die Auswahl sich ändert"""
        selection = self.tree.selection()
        if selection:
            self.edit_button.config(state=tk.NORMAL)
            self.duplicate_button.config(state=tk.NORMAL)
            self.delete_button.config(state=tk.NORMAL)
            
            # Prüfe ob Bewegung möglich ist
            index = self.tree.index(selection[0])
            self.move_up_button.config(state=tk.NORMAL if index > 0 else tk.DISABLED)
            self.move_down_button.config(state=tk.NORMAL if index < len(self.export_configs) - 1 else tk.DISABLED)
        else:
            self.edit_button.config(state=tk.DISABLED)
            self.duplicate_button.config(state=tk.DISABLED)
            self.delete_button.config(state=tk.DISABLED)
            self.move_up_button.config(state=tk.DISABLED)
            self.move_down_button.config(state=tk.DISABLED)
    
    def _add_export(self):
        """Fügt einen neuen Export hinzu"""
        dialog = ExportEditDialog(self.dialog, xml_field_mappings=self.xml_field_mappings)
        result = dialog.show()
        
        if result:
            export = ExportConfig(**result)
            self.export_configs.append(export.to_dict())
            self._add_to_tree(export)
    
    def _edit_export(self):
        """Bearbeitet den ausgewählten Export"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if 0 <= index < len(self.export_configs):
            export_dict = self.export_configs[index]
            export = ExportConfig.from_dict(export_dict)
            
            dialog = ExportEditDialog(self.dialog, export, self.xml_field_mappings)
            result = dialog.show()
            
            if result:
                # Behalte die ID bei
                result['id'] = export.id
                self.export_configs[index] = result
                
                # Aktualisiere Tree
                self._refresh_tree()
    
    def _duplicate_export(self):
        """Dupliziert den ausgewählten Export"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if 0 <= index < len(self.export_configs):
            export_dict = self.export_configs[index].copy()
            export = ExportConfig.from_dict(export_dict)
            
            # Neue ID und Name
            export.id = str(uuid.uuid4())
            export.name = f"{export.name} (Kopie)"
            
            self.export_configs.append(export.to_dict())
            self._add_to_tree(export)
    
    def _delete_export(self):
        """Löscht den ausgewählten Export"""
        selection = self.tree.selection()
        if not selection:
            return
        
        if messagebox.askyesno("Export löschen", "Möchten Sie diesen Export wirklich löschen?"):
            item = selection[0]
            index = self.tree.index(item)
            
            if 0 <= index < len(self.export_configs):
                del self.export_configs[index]
                self.tree.delete(item)
    
    def _move_up(self):
        """Bewegt den Export nach oben"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if index > 0:
            # Tausche in der Liste
            self.export_configs[index], self.export_configs[index-1] = \
                self.export_configs[index-1], self.export_configs[index]
            
            # Aktualisiere Tree
            self._refresh_tree()
            
            # Behalte Auswahl
            new_items = self.tree.get_children()
            if index-1 < len(new_items):
                self.tree.selection_set(new_items[index-1])
                self.tree.focus(new_items[index-1])
    
    def _move_down(self):
        """Bewegt den Export nach unten"""
        selection = self.tree.selection()
        if not selection:
            return
        
        item = selection[0]
        index = self.tree.index(item)
        
        if index < len(self.export_configs) - 1:
            # Tausche in der Liste
            self.export_configs[index], self.export_configs[index+1] = \
                self.export_configs[index+1], self.export_configs[index]
            
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
        for export_dict in self.export_configs:
            export = ExportConfig.from_dict(export_dict)
            self._add_to_tree(export)
    
    def _browse_error_path(self):
        """Öffnet Dialog zur Auswahl des Fehler-Pfads"""
        folder = filedialog.askdirectory(
            title="Fehlerpfad auswählen",
            initialdir=self.error_path_var.get() or os.path.expanduser("~")
        )
        if folder:
            self.error_path_var.set(folder)
    
    def _on_save(self):
        """Speichert die Export-Konfigurationen"""
        self.result = {
            'export_configs': self.export_configs,
            'error_path': self.error_path_var.get().strip()
        }
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab ohne zu speichern"""
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result


class ExportEditDialog:
    """Dialog zum Bearbeiten eines einzelnen Exports"""
    
    def __init__(self, parent, export: Optional[ExportConfig] = None,
                 xml_field_mappings: List[Dict] = None):
        self.parent = parent
        self.export = export
        self.xml_field_mappings = xml_field_mappings or []
        self.result = None
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Export bearbeiten" if export else "Neuer Export")
        self.dialog.geometry("800x750")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Variablen
        self.name_var = tk.StringVar(value=export.name if export else "")
        self.enabled_var = tk.BooleanVar(value=export.enabled if export else True)
        self.method_var = tk.StringVar(value=export.export_method.value if export else ExportMethod.FILE.value)
        self.format_var = tk.StringVar(value=export.export_format.value if export else ExportFormat.PDF.value)
        self.path_var = tk.StringVar(value=export.export_path_expression if export else "<OutputPath>")
        self.filename_var = tk.StringVar(value=export.export_filename_expression if export else "<FileName>")
        
        # E-Mail Variablen
        if export and export.email_config:
            self.email_recipient_var = tk.StringVar(value=export.email_config.recipient)
            self.email_subject_var = tk.StringVar(value=export.email_config.subject_expression)
            self.email_body_var = tk.StringVar(value=export.email_config.body_expression)
            self.email_cc_var = tk.StringVar(value=export.email_config.cc)
            self.email_bcc_var = tk.StringVar(value=export.email_config.bcc)
        else:
            self.email_recipient_var = tk.StringVar()
            self.email_subject_var = tk.StringVar(value="Dokument: <FileName>")
            self.email_body_var = tk.StringVar(value="Anbei finden Sie das verarbeitete Dokument.")
            self.email_cc_var = tk.StringVar()
            self.email_bcc_var = tk.StringVar()
        
        # Format-Parameter
        self.format_params = export.format_params.copy() if export else {}
        
        self._create_widgets()
        self._layout_widgets()
        
        # Initialisiere UI-Zustand
        self._on_method_changed()
        self._on_format_changed()
        
        # Fokus
        self.name_entry.focus()
        
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
        
        # Basis-Einstellungen
        self.base_frame = ttk.LabelFrame(self.main_frame, text="Basis-Einstellungen", padding="10")
        
        self.name_label = ttk.Label(self.base_frame, text="Name:")
        self.name_entry = ttk.Entry(self.base_frame, textvariable=self.name_var, width=40)
        
        self.enabled_check = ttk.Checkbutton(self.base_frame, text="Export aktiviert", 
                                            variable=self.enabled_var)
        
        # Export-Methode
        self.method_frame = ttk.LabelFrame(self.main_frame, text="Export-Methode", padding="10")
        
        self.method_file = ttk.Radiobutton(self.method_frame, text="Als Datei speichern", 
                                          variable=self.method_var, value=ExportMethod.FILE.value,
                                          command=self._on_method_changed)
        self.method_email = ttk.Radiobutton(self.method_frame, text="Per E-Mail versenden", 
                                           variable=self.method_var, value=ExportMethod.EMAIL.value,
                                           command=self._on_method_changed)
        
        # Export-Format
        self.format_frame = ttk.LabelFrame(self.main_frame, text="Export-Format", padding="10")
        
        # Format-Auswahl
        self.format_label = ttk.Label(self.format_frame, text="Format:")
        self.format_combo = ttk.Combobox(self.format_frame, textvariable=self.format_var, 
                                        state="readonly", width=25)
        self.format_combo['values'] = [
            (ExportFormat.PDF.value, "PDF"),
            (ExportFormat.PDF_A.value, "PDF/A"),
            (ExportFormat.SEARCHABLE_PDF_A.value, "Durchsuchbares PDF/A"),
            (ExportFormat.PNG.value, "PNG (Bilder)"),
            (ExportFormat.JPG.value, "JPG (Bilder)"),
            (ExportFormat.TIFF.value, "TIFF (Bilder)"),
            (ExportFormat.XML.value, "XML"),
            (ExportFormat.JSON.value, "JSON"),
            (ExportFormat.TXT.value, "Text (nur OCR-Text)"),
            (ExportFormat.CSV.value, "CSV (strukturierte Daten)")
        ]
        # Setze Display-Werte
        format_display = {
            ExportFormat.PDF.value: "PDF",
            ExportFormat.PDF_A.value: "PDF/A",
            ExportFormat.SEARCHABLE_PDF_A.value: "Durchsuchbares PDF/A",
            ExportFormat.PNG.value: "PNG (Bilder)",
            ExportFormat.JPG.value: "JPG (Bilder)",
            ExportFormat.TIFF.value: "TIFF (Bilder)",
            ExportFormat.XML.value: "XML",
            ExportFormat.JSON.value: "JSON",
            ExportFormat.TXT.value: "Text (nur OCR-Text)",
            ExportFormat.CSV.value: "CSV (strukturierte Daten)"
        }
        self.format_combo['values'] = list(format_display.values())
        # Setze aktuellen Wert
        current_format = self.format_var.get()
        self.format_combo.set(format_display.get(current_format, "PDF"))
        
        self.format_combo.bind('<<ComboboxSelected>>', lambda e: self._on_format_changed())
        
        # Format-spezifische Optionen (wird dynamisch gefüllt)
        self.format_options_frame = ttk.Frame(self.format_frame)
        
        # Datei-Export Einstellungen
        self.file_frame = ttk.LabelFrame(self.main_frame, text="Datei-Export", padding="10")
        
        self.path_label = ttk.Label(self.file_frame, text="Export-Pfad:")
        self.path_frame = ttk.Frame(self.file_frame)
        self.path_entry = ttk.Entry(self.path_frame, textvariable=self.path_var, width=40)
        self.path_expr_button = ttk.Button(self.path_frame, text="📝", width=3,
                                          command=self._edit_path_expression)
        
        self.filename_label = ttk.Label(self.file_frame, text="Dateiname:")
        self.filename_frame = ttk.Frame(self.file_frame)
        self.filename_entry = ttk.Entry(self.filename_frame, textvariable=self.filename_var, width=40)
        self.filename_expr_button = ttk.Button(self.filename_frame, text="📝", width=3,
                                              command=self._edit_filename_expression)
        
        # E-Mail-Export Einstellungen
        self.email_frame = ttk.LabelFrame(self.main_frame, text="E-Mail-Export", padding="10")
        
        self.email_recipient_label = ttk.Label(self.email_frame, text="Empfänger:")
        self.email_recipient_frame = ttk.Frame(self.email_frame)
        self.email_recipient_entry = ttk.Entry(self.email_recipient_frame, textvariable=self.email_recipient_var, width=40)
        self.email_recipient_expr_button = ttk.Button(self.email_recipient_frame, text="📝", width=3,
                                                     command=self._edit_recipient_expression)
        
        self.email_cc_label = ttk.Label(self.email_frame, text="CC (optional):")
        self.email_cc_frame = ttk.Frame(self.email_frame)
        self.email_cc_entry = ttk.Entry(self.email_cc_frame, textvariable=self.email_cc_var, width=40)
        self.email_cc_expr_button = ttk.Button(self.email_cc_frame, text="📝", width=3,
                                              command=self._edit_cc_expression)
        
        self.email_bcc_label = ttk.Label(self.email_frame, text="BCC (optional):")
        self.email_bcc_frame = ttk.Frame(self.email_frame)
        self.email_bcc_entry = ttk.Entry(self.email_bcc_frame, textvariable=self.email_bcc_var, width=40)
        self.email_bcc_expr_button = ttk.Button(self.email_bcc_frame, text="📝", width=3,
                                               command=self._edit_bcc_expression)
        
        self.email_subject_label = ttk.Label(self.email_frame, text="Betreff:")
        self.email_subject_frame = ttk.Frame(self.email_frame)
        self.email_subject_entry = ttk.Entry(self.email_subject_frame, textvariable=self.email_subject_var, width=40)
        self.email_subject_expr_button = ttk.Button(self.email_subject_frame, text="📝", width=3,
                                                   command=self._edit_subject_expression)
        
        # Dateiname für E-Mail Anhang
        self.email_filename_label = ttk.Label(self.email_frame, text="Anhang-Dateiname:")
        self.email_filename_frame = ttk.Frame(self.email_frame)
        self.email_filename_entry = ttk.Entry(self.email_filename_frame, textvariable=self.filename_var, width=40)
        self.email_filename_expr_button = ttk.Button(self.email_filename_frame, text="📝", width=3,
                                                    command=self._edit_filename_expression)
        
        self.email_body_label = ttk.Label(self.email_frame, text="Nachricht:")
        self.email_body_frame = ttk.Frame(self.email_frame)
        self.email_body_text = tk.Text(self.email_body_frame, height=4, width=50)
        self.email_body_text.insert("1.0", self.email_body_var.get())
        self.email_body_expr_button = ttk.Button(self.email_body_frame, text="📝", width=3,
                                                command=self._edit_body_expression)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save, default=tk.ACTIVE)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Basis-Einstellungen
        self.base_frame.pack(fill=tk.X, pady=(0, 10))
        self.name_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.name_entry.grid(row=0, column=1, sticky="we")
        self.enabled_check.grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=(10, 0))
        self.base_frame.columnconfigure(1, weight=1)
        
        # Export-Methode
        self.method_frame.pack(fill=tk.X, pady=(0, 10))
        self.method_file.pack(anchor=tk.W)
        self.method_email.pack(anchor=tk.W, pady=(5, 0))
        
        # Export-Format
        self.format_frame.pack(fill=tk.X, pady=(0, 10))
        self.format_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.format_combo.grid(row=0, column=1, sticky="w")
        self.format_options_frame.grid(row=1, column=0, columnspan=2, sticky="we", pady=(10, 0))
        
        # Datei-Export
        self.file_frame.pack(fill=tk.X, pady=(0, 10))
        self.path_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.path_frame.grid(row=0, column=1, sticky="we", pady=(0, 5))
        self.path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.path_expr_button.pack(side=tk.LEFT, padx=(5, 0))
        
        self.filename_label.grid(row=1, column=0, sticky=tk.W)
        self.filename_frame.grid(row=1, column=1, sticky="we")
        self.filename_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.filename_expr_button.pack(side=tk.LEFT, padx=(5, 0))
        self.file_frame.columnconfigure(1, weight=1)
        
        # E-Mail-Export
        self.email_frame.pack(fill=tk.X, pady=(0, 10))
        
        # Empfänger
        self.email_recipient_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.email_recipient_frame.grid(row=0, column=1, sticky="we", pady=(0, 5))
        self.email_recipient_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.email_recipient_expr_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # CC
        self.email_cc_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10))
        self.email_cc_frame.grid(row=1, column=1, sticky="we", pady=(0, 5))
        self.email_cc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.email_cc_expr_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # BCC
        self.email_bcc_label.grid(row=2, column=0, sticky=tk.W, padx=(0, 10))
        self.email_bcc_frame.grid(row=2, column=1, sticky="we", pady=(0, 5))
        self.email_bcc_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.email_bcc_expr_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Betreff
        self.email_subject_label.grid(row=3, column=0, sticky=tk.W, padx=(0, 10))
        self.email_subject_frame.grid(row=3, column=1, sticky="we", pady=(0, 5))
        self.email_subject_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.email_subject_expr_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Dateiname
        self.email_filename_label.grid(row=4, column=0, sticky=tk.W, padx=(0, 10))
        self.email_filename_frame.grid(row=4, column=1, sticky="we", pady=(0, 5))
        self.email_filename_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.email_filename_expr_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Nachricht
        self.email_body_label.grid(row=5, column=0, sticky="nw", padx=(0, 10))
        self.email_body_frame.grid(row=5, column=1, sticky="we")
        self.email_body_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.email_body_expr_button.pack(side=tk.LEFT, anchor=tk.N, padx=(5, 0))
        
        self.email_frame.columnconfigure(1, weight=1)
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.save_button.pack(side=tk.RIGHT)
    
    def _on_method_changed(self):
        """Wird aufgerufen wenn die Export-Methode geändert wird"""
        method = self.method_var.get()
        
        if method == ExportMethod.FILE.value:
            self.file_frame.pack(fill=tk.X, pady=(0, 10), before=self.button_frame)
            self.email_frame.pack_forget()
        elif method == ExportMethod.EMAIL.value:
            self.file_frame.pack_forget()
            self.email_frame.pack(fill=tk.X, pady=(0, 10), before=self.button_frame)
    
    def _on_format_changed(self):
        """Wird aufgerufen wenn das Export-Format geändert wird"""
        # Lösche alte Format-Optionen
        for widget in self.format_options_frame.winfo_children():
            widget.destroy()
        
        # Map Display-Wert zurück zum Enum-Wert
        format_display_map = {
            "PDF": ExportFormat.PDF.value,
            "PDF/A": ExportFormat.PDF_A.value,
            "Durchsuchbares PDF/A": ExportFormat.SEARCHABLE_PDF_A.value,
            "PNG (Bilder)": ExportFormat.PNG.value,
            "JPG (Bilder)": ExportFormat.JPG.value,
            "TIFF (Bilder)": ExportFormat.TIFF.value,
            "XML": ExportFormat.XML.value,
            "JSON": ExportFormat.JSON.value,
            "Text (nur OCR-Text)": ExportFormat.TXT.value,
            "CSV (strukturierte Daten)": ExportFormat.CSV.value
        }
        
        display_value = self.format_combo.get()
        format_value = format_display_map.get(display_value, ExportFormat.PDF.value)
        self.format_var.set(format_value)
        
        # Format-spezifische Optionen
        if format_value in [ExportFormat.PNG.value, ExportFormat.JPG.value]:
            # Bild-Optionen
            ttk.Label(self.format_options_frame, text="Bild-Qualität:").grid(row=0, column=0, sticky=tk.W)
            quality_var = tk.IntVar(value=self.format_params.get('quality', 95))
            quality_scale = ttk.Scale(self.format_options_frame, from_=50, to=100, 
                                     variable=quality_var, orient=tk.HORIZONTAL, length=200)
            quality_scale.grid(row=0, column=1, sticky="we")
            quality_label = ttk.Label(self.format_options_frame, text=f"{quality_var.get()}%")
            quality_label.grid(row=0, column=2, padx=(5, 0))
            
            def update_quality_label(value):
                quality_label.config(text=f"{int(float(value))}%")
                self.format_params['quality'] = int(float(value))
            
            quality_scale.config(command=update_quality_label)
            
            ttk.Label(self.format_options_frame, text="DPI:").grid(row=1, column=0, sticky=tk.W, pady=(5, 0))
            dpi_var = tk.IntVar(value=self.format_params.get('dpi', 300))
            dpi_combo = ttk.Combobox(self.format_options_frame, textvariable=dpi_var, 
                                    values=[72, 150, 300, 600], width=10)
            dpi_combo.grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
            dpi_combo.bind('<<ComboboxSelected>>', lambda e: self.format_params.update({'dpi': dpi_var.get()}))
            
        elif format_value == ExportFormat.TIFF.value:
            # TIFF-Optionen
            multipage_var = tk.BooleanVar(value=self.format_params.get('multipage', True))
            ttk.Checkbutton(self.format_options_frame, text="Mehrseitiges TIFF", 
                           variable=multipage_var,
                           command=lambda: self.format_params.update({'multipage': multipage_var.get()})).pack(anchor=tk.W)
    
    def _edit_path_expression(self):
        """Öffnet Dialog zur Bearbeitung des Pfad-Ausdrucks"""
        current_expr = self.path_var.get()
        dialog = ExpressionDialog(
            self.dialog, 
            title="Export-Pfad bearbeiten",
            expression=current_expr,
            description="Definieren Sie den Export-Pfad.\n"
                       "Verwenden Sie Variablen für dynamische Pfade.",
            xml_field_mappings=self.xml_field_mappings
        )
        result = dialog.show()
        
        if result is not None:
            self.path_var.set(result)
    
    def _edit_filename_expression(self):
        """Öffnet Dialog zur Bearbeitung des Dateiname-Ausdrucks"""
        current_expr = self.filename_var.get()
        dialog = ExpressionDialog(
            self.dialog, 
            title="Export-Dateiname bearbeiten",
            expression=current_expr,
            description="Definieren Sie den Dateinamen für den Export.\n"
                       "Die Dateierweiterung wird automatisch hinzugefügt.",
            xml_field_mappings=self.xml_field_mappings
        )
        result = dialog.show()
        
        if result is not None:
            self.filename_var.set(result)
    
    def _edit_recipient_expression(self):
        """Öffnet Dialog zur Bearbeitung des Empfänger-Ausdrucks"""
        current_expr = self.email_recipient_var.get()
        dialog = ExpressionDialog(
            self.dialog, 
            title="E-Mail Empfänger bearbeiten",
            expression=current_expr,
            description="Definieren Sie die E-Mail-Adresse des Empfängers.\n"
                       "Sie können auch Variablen verwenden (z.B. aus XML-Feldern).",
            xml_field_mappings=self.xml_field_mappings
        )
        result = dialog.show()
        
        if result is not None:
            self.email_recipient_var.set(result)
    
    def _edit_cc_expression(self):
        """Öffnet Dialog zur Bearbeitung des CC-Ausdrucks"""
        current_expr = self.email_cc_var.get()
        dialog = ExpressionDialog(
            self.dialog, 
            title="CC-Empfänger bearbeiten",
            expression=current_expr,
            description="Definieren Sie CC-Empfänger (optional).\n"
                       "Mehrere Adressen mit Komma trennen.",
            xml_field_mappings=self.xml_field_mappings
        )
        result = dialog.show()
        
        if result is not None:
            self.email_cc_var.set(result)
    
    def _edit_bcc_expression(self):
        """Öffnet Dialog zur Bearbeitung des BCC-Ausdrucks"""
        current_expr = self.email_bcc_var.get()
        dialog = ExpressionDialog(
            self.dialog, 
            title="BCC-Empfänger bearbeiten",
            expression=current_expr,
            description="Definieren Sie BCC-Empfänger (optional).\n"
                       "Mehrere Adressen mit Komma trennen.",
            xml_field_mappings=self.xml_field_mappings
        )
        result = dialog.show()
        
        if result is not None:
            self.email_bcc_var.set(result)
    
    def _edit_subject_expression(self):
        """Öffnet Dialog zur Bearbeitung des Betreff-Ausdrucks"""
        current_expr = self.email_subject_var.get()
        dialog = ExpressionDialog(
            self.dialog, 
            title="E-Mail-Betreff bearbeiten",
            expression=current_expr,
            description="Definieren Sie den Betreff für die E-Mail.\n"
                       "Verwenden Sie Variablen für dynamische Inhalte.",
            xml_field_mappings=self.xml_field_mappings
        )
        result = dialog.show()
        
        if result is not None:
            self.email_subject_var.set(result)
    
    def _edit_body_expression(self):
        """Öffnet Dialog zur Bearbeitung des Nachrichtentext-Ausdrucks"""
        current_expr = self.email_body_text.get("1.0", tk.END).strip()
        dialog = ExpressionDialog(
            self.dialog, 
            title="E-Mail-Nachricht bearbeiten",
            expression=current_expr,
            description="Definieren Sie den Nachrichtentext für die E-Mail.\n"
                       "Verwenden Sie Variablen für dynamische Inhalte.",
            xml_field_mappings=self.xml_field_mappings
        )
        result = dialog.show()
        
        if result is not None:
            self.email_body_text.delete("1.0", tk.END)
            self.email_body_text.insert("1.0", result)
    
    def _validate(self) -> bool:
        """Validiert die Eingaben"""
        # Name prüfen
        if not self.name_var.get().strip():
            messagebox.showerror("Fehler", "Bitte geben Sie einen Namen ein.")
            self.name_entry.focus()
            return False
        
        # Je nach Methode validieren
        method = self.method_var.get()
        
        if method == ExportMethod.FILE.value:
            # Pfad prüfen
            if not self.path_var.get().strip():
                messagebox.showerror("Fehler", "Bitte geben Sie einen Export-Pfad ein.")
                self.path_entry.focus()
                return False
            
            # Dateiname prüfen
            if not self.filename_var.get().strip():
                messagebox.showerror("Fehler", "Bitte geben Sie einen Dateinamen ein.")
                self.filename_entry.focus()
                return False
                
        elif method == ExportMethod.EMAIL.value:
            # E-Mail-Empfänger prüfen
            if not self.email_recipient_var.get().strip():
                messagebox.showerror("Fehler", "Bitte geben Sie einen E-Mail-Empfänger ein.")
                self.email_recipient_entry.focus()
                return False
            
            # Dateiname für Anhang prüfen
            if not self.filename_var.get().strip():
                messagebox.showerror("Fehler", "Bitte geben Sie einen Dateinamen für den E-Mail-Anhang ein.")
                self.email_filename_entry.focus()
                return False
        
        return True
    
    def _on_save(self):
        """Speichert die Export-Konfiguration"""
        if not self._validate():
            return
        
        # Erstelle Result-Dictionary
        self.result = {
            'id': self.export.id if self.export else str(uuid.uuid4()),
            'name': self.name_var.get().strip(),
            'enabled': self.enabled_var.get(),
            'export_method': self.method_var.get(),
            'export_format': self.format_var.get(),
            'export_path_expression': self.path_var.get().strip(),
            'export_filename_expression': self.filename_var.get().strip(),
            'format_params': self.format_params
        }
        
        # E-Mail-Konfiguration hinzufügen wenn E-Mail-Export
        if self.method_var.get() == ExportMethod.EMAIL.value:
            email_config = EmailConfig(
                recipient=self.email_recipient_var.get().strip(),
                subject_expression=self.email_subject_var.get().strip(),
                body_expression=self.email_body_text.get("1.0", tk.END).strip(),
                cc=self.email_cc_var.get().strip(),
                bcc=self.email_bcc_var.get().strip()
            )
            self.result['email_config'] = email_config.to_dict()
        
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab ohne zu speichern"""
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result