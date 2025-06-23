"""
Dialog zur Konfiguration von Export-Aktionen
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Optional, List, Dict
import sys
import os
import uuid

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.export_action import ExportConfig, ExportType, MetaDataFormat, MetaDataConfig, EmailConfig, ExportProgram
from gui.expression_dialog import ExpressionDialog


class ExportDialog:
    """Dialog zur Konfiguration eines Exports"""
    
    def __init__(self, parent, export_config: Optional[ExportConfig] = None,
                 ocr_zones: List[Dict] = None, xml_field_mappings: List[Dict] = None):
        self.parent = parent
        self.export_config = export_config
        self.ocr_zones = ocr_zones or []
        self.xml_field_mappings = xml_field_mappings or []
        self.result = None
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Export konfigurieren" if export_config else "Neuer Export")
        self.dialog.geometry("900x800")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Variablen
        self._init_variables()
        
        self._create_widgets()
        self._layout_widgets()
        self._load_config()
        
        # Initial-Zustand
        self._on_export_type_changed()
        
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
    
    def _init_variables(self):
        """Initialisiert alle Variablen"""
        config = self.export_config
        
        # Basis
        self.name_var = tk.StringVar(value=config.name if config else "")
        self.enabled_var = tk.BooleanVar(value=config.enabled if config else True)
        self.export_type_var = tk.StringVar(value=config.export_type.value if config else "file")
        
        # Dateiausgabe
        self.output_path_var = tk.StringVar(value=config.output_path_expression if config else "<OutputPath>")
        self.filename_var = tk.StringVar(value=config.filename_expression if config else "<FileName>")
        self.create_path_var = tk.BooleanVar(value=config.create_path_if_not_exists if config else True)
        self.append_file_var = tk.BooleanVar(value=config.append_to_existing_file if config else False)
        self.max_size_var = tk.StringVar(value=str(config.max_file_size_mb) if config and config.max_file_size_mb else "")
        self.max_pages_var = tk.StringVar(value=str(config.max_pages_per_file) if config and config.max_pages_per_file else "")
        self.append_position_var = tk.StringVar(value=config.append_position if config else "end")
        self.ocr_text_file_var = tk.BooleanVar(value=config.include_ocr_text_file if config else False)
        self.error_path_var = tk.StringVar(value=config.error_output_path if config else "")
        
        # Metadaten
        if config and config.metadata_config:
            meta = config.metadata_config
            self.meta_format_var = tk.StringVar(value=meta.format.value)
            self.meta_path_var = tk.StringVar(value=meta.output_path_expression)
            self.meta_filename_var = tk.StringVar(value=meta.filename_expression)
            self.meta_use_doc_name_var = tk.BooleanVar(value=meta.use_document_filename)
            self.meta_append_var = tk.BooleanVar(value=meta.append_to_existing)
            self.meta_ocr_text_var = tk.BooleanVar(value=meta.include_ocr_text)
            self.meta_ocr_zones_var = tk.BooleanVar(value=meta.include_ocr_zones)
            self.meta_xml_fields_var = tk.BooleanVar(value=meta.include_xml_fields)
            self.csv_delimiter_var = tk.StringVar(value=meta.csv_delimiter)
            self.csv_extension_var = tk.StringVar(value=meta.csv_extension)
            self.xml_extension_var = tk.StringVar(value=meta.xml_extension)
        else:
            self.meta_format_var = tk.StringVar(value="none")
            self.meta_path_var = tk.StringVar(value="<OutputPath>")
            self.meta_filename_var = tk.StringVar(value="<FileName>")
            self.meta_use_doc_name_var = tk.BooleanVar(value=True)
            self.meta_append_var = tk.BooleanVar(value=False)
            self.meta_ocr_text_var = tk.BooleanVar(value=False)
            self.meta_ocr_zones_var = tk.BooleanVar(value=False)
            self.meta_xml_fields_var = tk.BooleanVar(value=True)
            self.csv_delimiter_var = tk.StringVar(value=";")
            self.csv_extension_var = tk.StringVar(value="csv")
            self.xml_extension_var = tk.StringVar(value="xml")
        
        # E-Mail
        if config and config.email_config:
            email = config.email_config
            self.smtp_server_var = tk.StringVar(value=email.smtp_server)
            self.smtp_port_var = tk.StringVar(value=str(email.smtp_port))
            self.smtp_ssl_var = tk.BooleanVar(value=email.use_ssl)
            self.smtp_user_var = tk.StringVar(value=email.username)
            self.smtp_pass_var = tk.StringVar(value=email.password)
            self.from_addr_var = tk.StringVar(value=email.from_address)
            self.to_expr_var = tk.StringVar(value=email.to_expression)
            self.subject_expr_var = tk.StringVar(value=email.subject_expression)
            self.body_expr_var = tk.StringVar(value=email.body_expression)
            self.attach_pdf_var = tk.BooleanVar(value=email.attach_pdf)
            self.attach_xml_var = tk.BooleanVar(value=email.attach_xml)
        else:
            self.smtp_server_var = tk.StringVar()
            self.smtp_port_var = tk.StringVar(value="587")
            self.smtp_ssl_var = tk.BooleanVar(value=True)
            self.smtp_user_var = tk.StringVar()
            self.smtp_pass_var = tk.StringVar()
            self.from_addr_var = tk.StringVar()
            self.to_expr_var = tk.StringVar()
            self.subject_expr_var = tk.StringVar(value="Dokument: <FileName>")
            self.body_expr_var = tk.StringVar(value="Anbei erhalten Sie das angeforderte Dokument.")
            self.attach_pdf_var = tk.BooleanVar(value=True)
            self.attach_xml_var = tk.BooleanVar(value=True)
        
        # Bedingung
        self.condition_var = tk.StringVar(value=config.condition_expression if config else "")
        
        # Programme
        self.programs = []
        if config and config.programs:
            self.programs = [p.to_dict() for p in config.programs]
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Hauptframe
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Notebook für Tabs
        self.notebook = ttk.Notebook(self.main_frame)
        
        # Tab: Allgemein
        self.general_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.general_tab, text="Allgemein")
        self._create_general_widgets()
        
        # Tab: Dateiausgabe
        self.file_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.file_tab, text="Dateiausgabe")
        self._create_file_widgets()
        
        # Tab: Metadaten
        self.meta_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.meta_tab, text="Meta-Datei")
        self._create_metadata_widgets()
        
        # Tab: E-Mail
        self.email_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.email_tab, text="E-Mail")
        self._create_email_widgets()
        
        # Tab: Programme
        self.program_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.program_tab, text="Programme")
        self._create_program_widgets()
        
        # Tab: Bedingungen
        self.condition_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.condition_tab, text="Bedingungen")
        self._create_condition_widgets()
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save)
    
    def _create_general_widgets(self):
        """Erstellt Widgets für Allgemein-Tab"""
        frame = ttk.Frame(self.general_tab, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Name
        ttk.Label(frame, text="Export-Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(frame, textvariable=self.name_var, width=40).grid(row=0, column=1, sticky="we", pady=5)
        
        # Aktiviert
        ttk.Checkbutton(frame, text="Export aktiviert", 
                       variable=self.enabled_var).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Export-Typ
        ttk.Label(frame, text="Export-Typ:").grid(row=2, column=0, sticky=tk.W, pady=5)
        export_combo = ttk.Combobox(frame, textvariable=self.export_type_var, width=30, state="readonly")
        export_combo['values'] = [
            ("file", "Dateiausgabe"),
            ("email", "E-Mail"),
            ("script", "Externes Programm"),
            # Weitere Typen können hier ergänzt werden
        ]
        export_combo.set_values = lambda: [export_combo.set(v[0]) for v in export_combo['values'] if v[0] == self.export_type_var.get()]
        export_combo.grid(row=2, column=1, sticky="we", pady=5)
        export_combo.bind('<<ComboboxSelected>>', lambda e: self._on_export_type_changed())
        
        # Konfiguriere Export-Typ Dropdown richtig
        export_values = []
        export_display = {}
        for value, display in [
            ("file", "Dateiausgabe"),
            ("email", "E-Mail"),
            ("script", "Externes Programm")
        ]:
            export_values.append(display)
            export_display[display] = value
        
        export_combo['values'] = export_values
        # Setze den angezeigten Wert
        for display, value in export_display.items():
            if value == self.export_type_var.get():
                export_combo.set(display)
                break
        
        # Update export_type_var wenn Auswahl sich ändert
        def on_export_type_select(event):
            selected_display = export_combo.get()
            if selected_display in export_display:
                self.export_type_var.set(export_display[selected_display])
            self._on_export_type_changed()
        
        export_combo.bind('<<ComboboxSelected>>', on_export_type_select)
        
        frame.columnconfigure(1, weight=1)
    
    def _create_file_widgets(self):
        """Erstellt Widgets für Dateiausgabe-Tab"""
        frame = ttk.Frame(self.file_tab, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Ausgabepfad
        ttk.Label(frame, text="Ausgabepfad:").grid(row=0, column=0, sticky=tk.W, pady=5)
        path_frame = ttk.Frame(frame)
        path_frame.grid(row=0, column=1, sticky="we", pady=5)
        ttk.Entry(path_frame, textvariable=self.output_path_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(path_frame, text="...", width=3,
                  command=lambda: self._browse_path(self.output_path_var)).pack(side=tk.LEFT, padx=(5,0))
        ttk.Button(path_frame, text="fx", width=3,
                  command=lambda: self._edit_expression("Ausgabepfad", self.output_path_var)).pack(side=tk.LEFT, padx=(5,0))
        
        # Dateiname
        ttk.Label(frame, text="Dateiname:").grid(row=1, column=0, sticky=tk.W, pady=5)
        name_frame = ttk.Frame(frame)
        name_frame.grid(row=1, column=1, sticky="we", pady=5)
        ttk.Entry(name_frame, textvariable=self.filename_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(name_frame, text="fx", width=3,
                  command=lambda: self._edit_expression("Dateiname", self.filename_var)).pack(side=tk.LEFT, padx=(5,0))
        
        # Optionen
        ttk.Checkbutton(frame, text="Ordner erstellen, wenn nicht vorhanden",
                       variable=self.create_path_var).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        ttk.Checkbutton(frame, text="An vorhandene Datei anhängen",
                       variable=self.append_file_var,
                       command=self._on_append_toggled).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Erweiterte Optionen für Anhängen
        self.append_options_frame = ttk.LabelFrame(frame, text="Anhänge-Optionen", padding="10")
        self.append_options_frame.grid(row=4, column=0, columnspan=2, sticky="we", pady=10)
        
        ttk.Label(self.append_options_frame, text="Position:").grid(row=0, column=0, sticky=tk.W)
        position_frame = ttk.Frame(self.append_options_frame)
        position_frame.grid(row=0, column=1, sticky=tk.W)
        ttk.Radiobutton(position_frame, text="Am Ende", variable=self.append_position_var, 
                       value="end").pack(side=tk.LEFT)
        ttk.Radiobutton(position_frame, text="Am Anfang", variable=self.append_position_var, 
                       value="start").pack(side=tk.LEFT, padx=(20,0))
        
        ttk.Label(self.append_options_frame, text="Max. Dateigröße (MB):").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.append_options_frame, textvariable=self.max_size_var, width=10).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(self.append_options_frame, text="Max. Seitenzahl:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.append_options_frame, textvariable=self.max_pages_var, width=10).grid(row=2, column=1, sticky=tk.W, pady=5)
        
        # OCR-Textdatei
        ttk.Checkbutton(frame, text="OCR-Textdatei erstellen",
                       variable=self.ocr_text_file_var).grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Fehlerausgabe
        ttk.Label(frame, text="Ausgabeordner im Fehlerfall:").grid(row=6, column=0, sticky=tk.W, pady=5)
        error_frame = ttk.Frame(frame)
        error_frame.grid(row=6, column=1, sticky="we", pady=5)
        ttk.Entry(error_frame, textvariable=self.error_path_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(error_frame, text="...", width=3,
                  command=lambda: self._browse_path(self.error_path_var)).pack(side=tk.LEFT, padx=(5,0))
        
        frame.columnconfigure(1, weight=1)
    
    def _create_metadata_widgets(self):
        """Erstellt Widgets für Metadaten-Tab"""
        frame = ttk.Frame(self.meta_tab, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Format
        ttk.Label(frame, text="Meta-Datei Format:").grid(row=0, column=0, sticky=tk.W, pady=5)
        format_combo = ttk.Combobox(frame, textvariable=self.meta_format_var, width=20, state="readonly")
        format_combo['values'] = ["none", "csv", "xml", "json"]
        format_combo.grid(row=0, column=1, sticky=tk.W, pady=5)
        format_combo.bind('<<ComboboxSelected>>', lambda e: self._on_meta_format_changed())
        
        # Meta-Datei Optionen
        self.meta_options_frame = ttk.LabelFrame(frame, text="Meta-Datei Einstellungen", padding="10")
        self.meta_options_frame.grid(row=1, column=0, columnspan=2, sticky="we", pady=10)
        
        # Verzeichnis
        ttk.Label(self.meta_options_frame, text="Verzeichnis:").grid(row=0, column=0, sticky=tk.W, pady=5)
        meta_path_frame = ttk.Frame(self.meta_options_frame)
        meta_path_frame.grid(row=0, column=1, sticky="we", pady=5)
        ttk.Entry(meta_path_frame, textvariable=self.meta_path_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(meta_path_frame, text="fx", width=3,
                  command=lambda: self._edit_expression("Meta-Datei Verzeichnis", self.meta_path_var)).pack(side=tk.LEFT, padx=(5,0))
        
        # Dateiname
        ttk.Label(self.meta_options_frame, text="Dateiname:").grid(row=1, column=0, sticky=tk.W, pady=5)
        meta_name_frame = ttk.Frame(self.meta_options_frame)
        meta_name_frame.grid(row=1, column=1, sticky="we", pady=5)
        ttk.Entry(meta_name_frame, textvariable=self.meta_filename_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(meta_name_frame, text="fx", width=3,
                  command=lambda: self._edit_expression("Meta-Dateiname", self.meta_filename_var)).pack(side=tk.LEFT, padx=(5,0))
        
        ttk.Checkbutton(self.meta_options_frame, text="Dateiname des Dokumentes übernehmen",
                       variable=self.meta_use_doc_name_var).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        ttk.Checkbutton(self.meta_options_frame, text="An vorhandene Datei anhängen",
                       variable=self.meta_append_var).grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        # Inhalt
        content_frame = ttk.LabelFrame(self.meta_options_frame, text="Inhalt", padding="5")
        content_frame.grid(row=4, column=0, columnspan=2, sticky="we", pady=10)
        
        ttk.Checkbutton(content_frame, text="OCR-Text hinzufügen",
                       variable=self.meta_ocr_text_var).pack(anchor=tk.W)
        ttk.Checkbutton(content_frame, text="OCR-Zonen-Text hinzufügen",
                       variable=self.meta_ocr_zones_var).pack(anchor=tk.W)
        ttk.Checkbutton(content_frame, text="XML-Felder hinzufügen",
                       variable=self.meta_xml_fields_var).pack(anchor=tk.W)
        
        # Format-spezifische Optionen
        self.csv_options_frame = ttk.LabelFrame(self.meta_options_frame, text="CSV-Optionen", padding="5")
        self.csv_options_frame.grid(row=5, column=0, columnspan=2, sticky="we", pady=10)
        
        ttk.Label(self.csv_options_frame, text="Trennzeichen:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.csv_options_frame, textvariable=self.csv_delimiter_var, width=5).grid(row=0, column=1, sticky=tk.W)
        
        ttk.Label(self.csv_options_frame, text="Dateiendung:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self.csv_options_frame, textvariable=self.csv_extension_var, width=10).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        self.xml_options_frame = ttk.LabelFrame(self.meta_options_frame, text="XML-Optionen", padding="5")
        self.xml_options_frame.grid(row=6, column=0, columnspan=2, sticky="we", pady=10)
        
        ttk.Label(self.xml_options_frame, text="Dateiendung:").grid(row=0, column=0, sticky=tk.W)
        ttk.Entry(self.xml_options_frame, textvariable=self.xml_extension_var, width=10).grid(row=0, column=1, sticky=tk.W)
        
        self.meta_options_frame.columnconfigure(1, weight=1)
        frame.columnconfigure(1, weight=1)
    
    def _create_email_widgets(self):
        """Erstellt Widgets für E-Mail-Tab"""
        frame = ttk.Frame(self.email_tab, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Server-Einstellungen
        server_frame = ttk.LabelFrame(frame, text="Server-Einstellungen", padding="10")
        server_frame.pack(fill=tk.X, pady=(0,10))
        
        ttk.Label(server_frame, text="SMTP-Server:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(server_frame, textvariable=self.smtp_server_var, width=40).grid(row=0, column=1, sticky="we", pady=5)
        
        ttk.Label(server_frame, text="Port:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(server_frame, textvariable=self.smtp_port_var, width=10).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Checkbutton(server_frame, text="SSL/TLS verwenden",
                       variable=self.smtp_ssl_var).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=5)
        
        ttk.Label(server_frame, text="Benutzername:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(server_frame, textvariable=self.smtp_user_var, width=40).grid(row=3, column=1, sticky="we", pady=5)
        
        ttk.Label(server_frame, text="Passwort:").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Entry(server_frame, textvariable=self.smtp_pass_var, width=40, show="*").grid(row=4, column=1, sticky="we", pady=5)
        
        server_frame.columnconfigure(1, weight=1)
        
        # E-Mail-Einstellungen
        email_frame = ttk.LabelFrame(frame, text="E-Mail-Einstellungen", padding="10")
        email_frame.pack(fill=tk.X, pady=(0,10))
        
        ttk.Label(email_frame, text="Betreff:").grid(row=2, column=0, sticky=tk.W, pady=5)
        subject_frame = ttk.Frame(email_frame)
        subject_frame.grid(row=2, column=1, sticky="we", pady=5)
        ttk.Entry(subject_frame, textvariable=self.subject_expr_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(subject_frame, text="fx", width=3,
                  command=lambda: self._edit_expression("E-Mail Betreff", self.subject_expr_var)).pack(side=tk.LEFT, padx=(5,0))
        
        ttk.Label(email_frame, text="Text:").grid(row=3, column=0, sticky="nw", pady=5)
        body_frame = ttk.Frame(email_frame)
        body_frame.grid(row=3, column=1, sticky="we", pady=5)
        self.body_text = tk.Text(body_frame, height=5, width=40)
        self.body_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.body_text.insert("1.0", self.body_expr_var.get())
        ttk.Button(body_frame, text="fx", width=3,
                  command=lambda: self._edit_body_expression()).pack(side=tk.LEFT, anchor="n", padx=(5,0))
        
        email_frame.columnconfigure(1, weight=1)
        
        # Anhänge
        attach_frame = ttk.LabelFrame(frame, text="Anhänge", padding="10")
        attach_frame.pack(fill=tk.X)
        
        ttk.Checkbutton(attach_frame, text="PDF-Dokument anhängen",
                       variable=self.attach_pdf_var).pack(anchor=tk.W)
        ttk.Checkbutton(attach_frame, text="XML-Datei anhängen (wenn vorhanden)",
                       variable=self.attach_xml_var).pack(anchor=tk.W)
    
    def _create_program_widgets(self):
        """Erstellt Widgets für Programme-Tab"""
        frame = ttk.Frame(self.program_tab, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Info
        info_label = ttk.Label(frame, text="Externe Programme, die nach dem Export ausgeführt werden:")
        info_label.pack(anchor=tk.W, pady=(0,10))
        
        # Toolbar
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill=tk.X, pady=(0,5))
        
        ttk.Button(toolbar, text="➕ Hinzufügen", command=self._add_program).pack(side=tk.LEFT, padx=(0,5))
        ttk.Button(toolbar, text="✏️ Bearbeiten", command=self._edit_program).pack(side=tk.LEFT, padx=(0,5))
        ttk.Button(toolbar, text="🗑️ Löschen", command=self._delete_program).pack(side=tk.LEFT, padx=(0,5))
        ttk.Button(toolbar, text="⬆", width=3, command=self._move_program_up).pack(side=tk.LEFT, padx=(0,5))
        ttk.Button(toolbar, text="⬇", width=3, command=self._move_program_down).pack(side=tk.LEFT)
        
        # Liste
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        self.program_tree = ttk.Treeview(list_frame, columns=("Path", "Parameters", "64bit"), height=10)
        self.program_tree.heading("#0", text="")
        self.program_tree.heading("Path", text="Programmpfad")
        self.program_tree.heading("Parameters", text="Aufrufparameter")
        self.program_tree.heading("64bit", text="64-bit")
        
        self.program_tree.column("#0", width=30)
        self.program_tree.column("Path", width=300)
        self.program_tree.column("Parameters", width=300)
        self.program_tree.column("64bit", width=60)
        
        vsb = ttk.Scrollbar(list_frame, orient="vertical", command=self.program_tree.yview)
        self.program_tree.configure(yscrollcommand=vsb.set)
        
        self.program_tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        
        list_frame.grid_columnconfigure(0, weight=1)
        list_frame.grid_rowconfigure(0, weight=1)
        
        # Checkbutton
        ttk.Checkbutton(frame, text="Jedes Dokument einzeln ausführen",
                       variable=tk.BooleanVar(value=True)).pack(anchor=tk.W, pady=(10,0))
    
    def _create_condition_widgets(self):
        """Erstellt Widgets für Bedingungs-Tab"""
        frame = ttk.Frame(self.condition_tab, padding="10")
        frame.pack(fill=tk.BOTH, expand=True)
        
        # Info
        info_text = ("Definieren Sie eine Bedingung, wann dieser Export ausgeführt werden soll.\n"
                    "Lassen Sie das Feld leer, um den Export immer auszuführen.")
        ttk.Label(frame, text=info_text, wraplength=600).pack(anchor=tk.W, pady=(0,20))
        
        # Bedingung
        ttk.Label(frame, text="Bedingungsausdruck:").pack(anchor=tk.W)
        
        cond_frame = ttk.Frame(frame)
        cond_frame.pack(fill=tk.X, pady=(5,0))
        
        self.condition_entry = ttk.Entry(cond_frame, textvariable=self.condition_var, width=60)
        self.condition_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(cond_frame, text="fx", width=3,
                  command=lambda: self._edit_expression("Bedingung", self.condition_var)).pack(side=tk.LEFT, padx=(5,0))
        
        # Beispiele
        example_frame = ttk.LabelFrame(frame, text="Beispiele", padding="10")
        example_frame.pack(fill=tk.X, pady=(20,0))
        
        examples = [
            'IF("<FileSize>", ">", "1000000", "true", "false")  # Nur Dateien größer als 1MB',
            'IF("<FileName>", "contains", "Rechnung", "true", "false")  # Nur Rechnungen',
            'IF("<Kundenname>", "!=", "", "true", "false")  # Nur wenn Kundenname vorhanden'
        ]
        
        for example in examples:
            ttk.Label(example_frame, text=example, font=("Courier", 9)).pack(anchor=tk.W, pady=2)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.notebook.pack(fill=tk.BOTH, expand=True, pady=(0,10))
        
        self.button_frame.pack(fill=tk.X)
        self.cancel_button.pack(side=tk.RIGHT, padx=(5,0))
        self.save_button.pack(side=tk.RIGHT)
    
    def _load_config(self):
        """Lädt die Konfiguration"""
        if not self.export_config:
            return
        
        # Programme laden
        self._refresh_program_list()
        
        # Initial-Zustand für abhängige Widgets
        self._on_append_toggled()
        self._on_meta_format_changed()
    
    def _on_export_type_changed(self):
        """Wird aufgerufen wenn der Export-Typ geändert wird"""
        export_type = self.export_type_var.get()
        
        # Aktiviere/Deaktiviere Tabs basierend auf Export-Typ
        if export_type == "file":
            self._enable_tab(1)  # Dateiausgabe
            self._enable_tab(2)  # Metadaten
            self._disable_tab(3)  # E-Mail
        elif export_type == "email":
            self._disable_tab(1)  # Dateiausgabe
            self._disable_tab(2)  # Metadaten
            self._enable_tab(3)  # E-Mail
        elif export_type == "script":
            self._disable_tab(1)  # Dateiausgabe
            self._disable_tab(2)  # Metadaten
            self._disable_tab(3)  # E-Mail
    
    def _enable_tab(self, index: int):
        """Aktiviert einen Tab"""
        self.notebook.tab(index, state="normal")
    
    def _disable_tab(self, index: int):
        """Deaktiviert einen Tab"""
        self.notebook.tab(index, state="disabled")
    
    def _on_append_toggled(self):
        """Wird aufgerufen wenn Anhängen-Option geändert wird"""
        if self.append_file_var.get():
            self.append_options_frame.configure(state="normal")
            for child in self.append_options_frame.winfo_children():
                child.configure(state="normal")
        else:
            self.append_options_frame.configure(state="disabled")
            for child in self.append_options_frame.winfo_children():
                child.configure(state="disabled")
    
    def _on_meta_format_changed(self):
        """Wird aufgerufen wenn Meta-Format geändert wird"""
        format_type = self.meta_format_var.get()
        
        if format_type == "none":
            self.meta_options_frame.configure(state="disabled")
            for child in self.meta_options_frame.winfo_children():
                child.configure(state="disabled")
        else:
            self.meta_options_frame.configure(state="normal")
            for child in self.meta_options_frame.winfo_children():
                child.configure(state="normal")
            
            # Format-spezifische Optionen
            if format_type == "csv":
                self.csv_options_frame.configure(state="normal")
                for child in self.csv_options_frame.winfo_children():
                    child.configure(state="normal")
                self.xml_options_frame.configure(state="disabled")
                for child in self.xml_options_frame.winfo_children():
                    child.configure(state="disabled")
            elif format_type == "xml":
                self.csv_options_frame.configure(state="disabled")
                for child in self.csv_options_frame.winfo_children():
                    child.configure(state="disabled")
                self.xml_options_frame.configure(state="normal")
                for child in self.xml_options_frame.winfo_children():
                    child.configure(state="normal")
            else:
                self.csv_options_frame.configure(state="disabled")
                for child in self.csv_options_frame.winfo_children():
                    child.configure(state="disabled")
                self.xml_options_frame.configure(state="disabled")
                for child in self.xml_options_frame.winfo_children():
                    child.configure(state="disabled")
    
    def _browse_path(self, var: tk.StringVar):
        """Öffnet Dialog zur Pfadauswahl"""
        path = filedialog.askdirectory(
            title="Ordner auswählen",
            initialdir=var.get() or os.path.expanduser("~")
        )
        if path:
            var.set(path)
    
    def _edit_expression(self, title: str, var: tk.StringVar):
        """Öffnet Expression-Editor"""
        dialog = ExpressionDialog(
            self.dialog,
            title=f"{title} - Ausdruck bearbeiten",
            expression=var.get(),
            xml_field_mappings=self.xml_field_mappings
        )
        result = dialog.show()
        if result is not None:
            var.set(result)
    
    def _edit_body_expression(self):
        """Öffnet Expression-Editor für E-Mail Body"""
        dialog = ExpressionDialog(
            self.dialog,
            title="E-Mail Text - Ausdruck bearbeiten",
            expression=self.body_text.get("1.0", tk.END).strip(),
            xml_field_mappings=self.xml_field_mappings
        )
        result = dialog.show()
        if result is not None:
            self.body_text.delete("1.0", tk.END)
            self.body_text.insert("1.0", result)
            self.body_expr_var.set(result)
    
    def _add_program(self):
        """Fügt ein neues Programm hinzu"""
        dialog = ProgramEditDialog(self.dialog, xml_field_mappings=self.xml_field_mappings)
        result = dialog.show()
        if result:
            self.programs.append(result)
            self._refresh_program_list()
    
    def _edit_program(self):
        """Bearbeitet das ausgewählte Programm"""
        selection = self.program_tree.selection()
        if not selection:
            return
        
        index = self.program_tree.index(selection[0])
        if 0 <= index < len(self.programs):
            dialog = ProgramEditDialog(self.dialog, self.programs[index], 
                                     xml_field_mappings=self.xml_field_mappings)
            result = dialog.show()
            if result:
                self.programs[index] = result
                self._refresh_program_list()
    
    def _delete_program(self):
        """Löscht das ausgewählte Programm"""
        selection = self.program_tree.selection()
        if not selection:
            return
        
        index = self.program_tree.index(selection[0])
        if 0 <= index < len(self.programs):
            if messagebox.askyesno("Programm löschen", "Möchten Sie dieses Programm wirklich löschen?"):
                del self.programs[index]
                self._refresh_program_list()
    
    def _move_program_up(self):
        """Verschiebt Programm nach oben"""
        selection = self.program_tree.selection()
        if not selection:
            return
        
        index = self.program_tree.index(selection[0])
        if index > 0:
            self.programs[index], self.programs[index-1] = self.programs[index-1], self.programs[index]
            self._refresh_program_list()
            # Behalte Auswahl
            self.program_tree.selection_set(self.program_tree.get_children()[index-1])
    
    def _move_program_down(self):
        """Verschiebt Programm nach unten"""
        selection = self.program_tree.selection()
        if not selection:
            return
        
        index = self.program_tree.index(selection[0])
        if index < len(self.programs) - 1:
            self.programs[index], self.programs[index+1] = self.programs[index+1], self.programs[index]
            self._refresh_program_list()
            # Behalte Auswahl
            self.program_tree.selection_set(self.program_tree.get_children()[index+1])
    
    def _refresh_program_list(self):
        """Aktualisiert die Programmliste"""
        for item in self.program_tree.get_children():
            self.program_tree.delete(item)
        
        for program in self.programs:
            self.program_tree.insert("", "end", 
                values=(program['path'], program['parameters'], 
                       "Ja" if program.get('run_as_64bit', True) else "Nein"))
    
    def _validate(self) -> bool:
        """Validiert die Eingaben"""
        if not self.name_var.get().strip():
            messagebox.showerror("Fehler", "Bitte geben Sie einen Export-Namen ein.")
            return False
        
        export_type = self.export_type_var.get()
        
        if export_type == "email":
            if not self.smtp_server_var.get():
                messagebox.showerror("Fehler", "Bitte geben Sie einen SMTP-Server ein.")
                return False
            if not self.from_addr_var.get():
                messagebox.showerror("Fehler", "Bitte geben Sie eine Absender-Adresse ein.")
                return False
            if not self.to_expr_var.get():
                messagebox.showerror("Fehler", "Bitte geben Sie einen Empfänger-Ausdruck ein.")
                return False
        
        return True
    
    def _on_save(self):
        """Speichert die Konfiguration"""
        if not self._validate():
            return
        
        # Erstelle ExportConfig
        config_id = self.export_config.id if self.export_config else str(uuid.uuid4())
        
        # Basis-Konfiguration
        result = {
            "id": config_id,
            "name": self.name_var.get().strip(),
            "enabled": self.enabled_var.get(),
            "export_type": self.export_type_var.get(),
            "output_path_expression": self.output_path_var.get(),
            "filename_expression": self.filename_var.get(),
            "create_path_if_not_exists": self.create_path_var.get(),
            "append_to_existing_file": self.append_file_var.get(),
            "max_file_size_mb": int(self.max_size_var.get()) if self.max_size_var.get() else None,
            "max_pages_per_file": int(self.max_pages_var.get()) if self.max_pages_var.get() else None,
            "append_position": self.append_position_var.get(),
            "include_ocr_text_file": self.ocr_text_file_var.get(),
            "error_output_path": self.error_path_var.get(),
            "condition_expression": self.condition_var.get(),
            "programs": self.programs
        }
        
        # Metadaten-Konfiguration
        result["metadata_config"] = {
            "format": self.meta_format_var.get(),
            "output_path_expression": self.meta_path_var.get(),
            "filename_expression": self.meta_filename_var.get(),
            "use_document_filename": self.meta_use_doc_name_var.get(),
            "append_to_existing": self.meta_append_var.get(),
            "include_ocr_text": self.meta_ocr_text_var.get(),
            "include_ocr_zones": self.meta_ocr_zones_var.get(),
            "include_xml_fields": self.meta_xml_fields_var.get(),
            "csv_delimiter": self.csv_delimiter_var.get(),
            "csv_extension": self.csv_extension_var.get(),
            "xml_extension": self.xml_extension_var.get()
        }
        
        # E-Mail-Konfiguration
        if self.export_type_var.get() == "email":
            result["email_config"] = {
                "smtp_server": self.smtp_server_var.get(),
                "smtp_port": int(self.smtp_port_var.get()) if self.smtp_port_var.get() else 587,
                "use_ssl": self.smtp_ssl_var.get(),
                "username": self.smtp_user_var.get(),
                "password": self.smtp_pass_var.get(),
                "from_address": self.from_addr_var.get(),
                "to_expression": self.to_expr_var.get(),
                "subject_expression": self.subject_expr_var.get(),
                "body_expression": self.body_text.get("1.0", tk.END).strip(),
                "attach_pdf": self.attach_pdf_var.get(),
                "attach_xml": self.attach_xml_var.get()
            }
        
        self.result = result
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab ohne zu speichern"""
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result


class ProgramEditDialog:
    """Dialog zum Bearbeiten eines Programms"""
    
    def __init__(self, parent, program_data: Dict = None, xml_field_mappings: List[Dict] = None):
        self.parent = parent
        self.program_data = program_data
        self.xml_field_mappings = xml_field_mappings or []
        self.result = None
        
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Programm bearbeiten")
        self.dialog.geometry("600x400")
        self.dialog.resizable(True, True)
        
        self._center_window()
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Variablen
        self.path_var = tk.StringVar(value=program_data['path'] if program_data else "")
        self.params_var = tk.StringVar(value=program_data['parameters'] if program_data else "")
        self.bit64_var = tk.BooleanVar(value=program_data.get('run_as_64bit', True) if program_data else True)
        self.each_doc_var = tk.BooleanVar(value=program_data.get('run_for_each_document', True) if program_data else True)
        
        self._create_widgets()
        self._layout_widgets()
    
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
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Programmpfad
        ttk.Label(self.main_frame, text="Programmpfad:").grid(row=0, column=0, sticky=tk.W, pady=5)
        path_frame = ttk.Frame(self.main_frame)
        path_frame.grid(row=0, column=1, sticky="we", pady=5)
        
        ttk.Entry(path_frame, textvariable=self.path_var, width=50).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(path_frame, text="...", width=3, command=self._browse_program).pack(side=tk.LEFT, padx=(5,0))
        
        # Aufrufparameter
        ttk.Label(self.main_frame, text="Aufrufparameter:").grid(row=1, column=0, sticky="nw", pady=5)
        params_frame = ttk.Frame(self.main_frame)
        params_frame.grid(row=1, column=1, sticky="we", pady=5)
        
        self.params_text = tk.Text(params_frame, height=8, width=50)
        self.params_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.params_text.insert("1.0", self.params_var.get())
        
        ttk.Button(params_frame, text="fx", width=3,
                  command=self._edit_parameters).pack(side=tk.LEFT, anchor="n", padx=(5,0))
        
        # Info
        info_frame = ttk.LabelFrame(self.main_frame, text="Verfügbare Variablen", padding="10")
        info_frame.grid(row=2, column=0, columnspan=2, sticky="we", pady=10)
        
        info_text = ("Spezielle Variablen für Programme:\n"
                    "<ProcessedFile> - Vollständiger Pfad der verarbeiteten Datei\n"
                    "<ProcessedFileName> - Dateiname der verarbeiteten Datei\n"
                    "<ProcessedFilePath> - Verzeichnis der verarbeiteten Datei\n"
                    "\nAlle anderen Variablen und Funktionen sind ebenfalls verfügbar.")
        ttk.Label(info_frame, text=info_text, font=("Courier", 9)).pack(anchor=tk.W)
        
        # Optionen
        options_frame = ttk.Frame(self.main_frame)
        options_frame.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        ttk.Checkbutton(options_frame, text="Als 64-Bit-Programm ausführen",
                       variable=self.bit64_var).pack(anchor=tk.W)
        ttk.Checkbutton(options_frame, text="Für jedes Dokument einzeln ausführen",
                       variable=self.each_doc_var).pack(anchor=tk.W)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", command=self._on_cancel)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", command=self._on_save)
        
        self.main_frame.columnconfigure(1, weight=1)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        self.button_frame.grid(row=4, column=0, columnspan=2, pady=(20,0))
        self.cancel_button.pack(side=tk.RIGHT, padx=(5,0))
        self.save_button.pack(side=tk.RIGHT)
    
    def _browse_program(self):
        """Öffnet Dialog zur Programmauswahl"""
        filename = filedialog.askopenfilename(
            title="Programm auswählen",
            filetypes=[("Ausführbare Dateien", "*.exe;*.bat;*.cmd"), ("Alle Dateien", "*.*")]
        )
        if filename:
            self.path_var.set(filename)
    
    def _edit_parameters(self):
        """Öffnet Expression-Editor für Parameter"""
        dialog = ExpressionDialog(
            self.dialog,
            title="Aufrufparameter - Ausdruck bearbeiten",
            expression=self.params_text.get("1.0", tk.END).strip(),
            description="Definieren Sie die Aufrufparameter für das Programm.\n"
                       "Verwenden Sie Variablen wie <ProcessedFile> für die verarbeitete Datei.",
            xml_field_mappings=self.xml_field_mappings
        )
        result = dialog.show()
        if result is not None:
            self.params_text.delete("1.0", tk.END)
            self.params_text.insert("1.0", result)
    
    def _on_save(self):
        """Speichert die Eingaben"""
        if not self.path_var.get().strip():
            messagebox.showerror("Fehler", "Bitte geben Sie einen Programmpfad ein.")
            return
        
        self.result = {
            "path": self.path_var.get().strip(),
            "parameters": self.params_text.get("1.0", tk.END).strip(),
            "run_as_64bit": self.bit64_var.get(),
            "run_for_each_document": self.each_doc_var.get()
        }
        
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab ohne zu speichern"""
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result_frame, text="Von:").grid(row=0, column=0, sticky=tk.W, pady=5)
        ttk.Entry(email_frame, textvariable=self.from_addr_var, width=40).grid(row=0, column=1, sticky="we", pady=5)
        
        ttk.Label(email_frame, text="An:").grid(row=1, column=0, sticky=tk.W, pady=5)
        to_frame = ttk.Frame(email_frame)
        to_frame.grid(row=1, column=1, sticky="we", pady=5)
        ttk.Entry(to_frame, textvariable=self.to_expr_var, width=40).pack(side=tk.LEFT, fill=tk.X, expand=True)
        ttk.Button(to_frame, text="fx", width=3,
                  command=lambda: self._edit_expression("E-Mail Empfänger", self.to_expr_var)).pack(side=tk.LEFT, padx=(5,0))
        
        ttk.Label(email