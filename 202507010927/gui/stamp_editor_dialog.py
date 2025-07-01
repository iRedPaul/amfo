"""
Dialog zum Erstellen und Bearbeiten von Stempeln
"""
import tkinter as tk
from tkinter import ttk, colorchooser, messagebox, font
from typing import Optional, List, Dict, Any
import uuid
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.stamp_config import StampConfig, StampPosition, StampOrientation
from gui.expression_dialog import ExpressionDialog
from gui.stamp_preview import StampPreview


class StampEditorDialog:
    """Dialog zum Bearbeiten eines Stempels"""
    
    def __init__(self, parent, stamp_config: Optional[StampConfig] = None, 
                 ocr_zones: List[Dict] = None, xml_fields: List[Dict] = None):
        self.parent = parent
        self.stamp_config = stamp_config
        self.ocr_zones = ocr_zones or []
        self.xml_fields = xml_fields or []
        self.result = None
        
        # Erstelle Dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Stempel bearbeiten" if stamp_config else "Neuer Stempel")
        self.dialog.geometry("1200x800")
        self.dialog.resizable(True, True)
        
        # Dialog-Eigenschaften
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Hauptcontainer mit zwei Spalten
        main_frame = ttk.Frame(self.dialog)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Linke Spalte: Einstellungen
        left_frame = ttk.Frame(main_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        
        # Rechte Spalte: Vorschau
        right_frame = ttk.Frame(main_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        
        # Erstelle UI
        self._create_settings_ui(left_frame)
        self._create_preview_ui(right_frame)
        
        # Buttons
        button_frame = ttk.Frame(self.dialog)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))
        
        ttk.Button(button_frame, text="Abbrechen", 
                  command=self._on_cancel).pack(side=tk.RIGHT, padx=(5, 0))
        ttk.Button(button_frame, text="Speichern", 
                  command=self._on_save).pack(side=tk.RIGHT)
        
        # Lade vorhandene Konfiguration
        if stamp_config:
            self._load_config(stamp_config)
        else:
            # Standard-Stempel mit Beispielzeilen
            self._add_default_lines()
        
        # Initial Preview
        self._update_preview()
    
    def _create_settings_ui(self, parent):
        """Erstellt die Einstellungs-UI"""
        # Notebook für Tabs
        self.notebook = ttk.Notebook(parent)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Tab 1: Grundeinstellungen
        self._create_basic_tab()
        
        # Tab 2: Position
        self._create_position_tab()
        
        # Tab 3: Aussehen
        self._create_appearance_tab()
        
        # Tab 4: Inhalt
        self._create_content_tab()
        
        # Tab 5: Seiten
        self._create_pages_tab()
    
    def _create_basic_tab(self):
        """Erstellt den Tab für Grundeinstellungen"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Grundeinstellungen")
        
        # Name
        ttk.Label(frame, text="Name:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.name_var = tk.StringVar(value=self.stamp_config.name if self.stamp_config else "Eingangsstempel")
        ttk.Entry(frame, textvariable=self.name_var, width=40).grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        
        # Aktiviert
        self.enabled_var = tk.BooleanVar(value=self.stamp_config.enabled if self.stamp_config else True)
        ttk.Checkbutton(frame, text="Stempel aktiviert", 
                       variable=self.enabled_var).grid(row=1, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        # Orientierung
        ttk.Label(frame, text="Orientierung:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.orientation_var = tk.StringVar(value=self.stamp_config.orientation.value if self.stamp_config else StampOrientation.HORIZONTAL.value)
        orientation_frame = ttk.Frame(frame)
        orientation_frame.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        for i, (text, value) in enumerate([
            ("Horizontal", StampOrientation.HORIZONTAL.value),
            ("Vertikal", StampOrientation.VERTICAL.value),
            ("Diagonal", StampOrientation.DIAGONAL.value)
        ]):
            ttk.Radiobutton(orientation_frame, text=text, value=value,
                           variable=self.orientation_var,
                           command=self._update_preview).pack(side=tk.LEFT, padx=5)
        
        # Rotation
        ttk.Label(frame, text="Rotation (Grad):").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.rotation_var = tk.DoubleVar(value=self.stamp_config.rotation if self.stamp_config else 0)
        rotation_frame = ttk.Frame(frame)
        rotation_frame.grid(row=3, column=1, sticky=(tk.W, tk.E), pady=5)
        
        self.rotation_scale = ttk.Scale(rotation_frame, from_=0, to=360, 
                                       variable=self.rotation_var,
                                       command=lambda x: self._update_preview())
        self.rotation_scale.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        self.rotation_label = ttk.Label(rotation_frame, text="0°")
        self.rotation_label.pack(side=tk.LEFT, padx=5)
        
        # Update label when scale changes
        self.rotation_var.trace('w', lambda *args: self.rotation_label.config(text=f"{int(self.rotation_var.get())}°"))
        
        frame.columnconfigure(1, weight=1)
    
    def _create_position_tab(self):
        """Erstellt den Tab für Position"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Position")
        
        # Position
        ttk.Label(frame, text="Position:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.position_var = tk.StringVar(value=self.stamp_config.position.value if self.stamp_config else StampPosition.TOP_RIGHT.value)
        
        position_frame = ttk.LabelFrame(frame, text="Vordefinierte Positionen", padding="10")
        position_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        # Position Grid
        positions = [
            [StampPosition.TOP_LEFT, StampPosition.TOP_CENTER, StampPosition.TOP_RIGHT],
            [StampPosition.MIDDLE_LEFT, StampPosition.MIDDLE_CENTER, StampPosition.MIDDLE_RIGHT],
            [StampPosition.BOTTOM_LEFT, StampPosition.BOTTOM_CENTER, StampPosition.BOTTOM_RIGHT]
        ]
        
        for r, row in enumerate(positions):
            for c, pos in enumerate(row):
                btn = ttk.Radiobutton(position_frame, text=pos.value.replace('_', ' ').title(),
                                     value=pos.value, variable=self.position_var,
                                     command=self._on_position_change)
                btn.grid(row=r, column=c, padx=10, pady=5)
        
        # Custom Position
        ttk.Radiobutton(frame, text="Benutzerdefinierte Position",
                       value=StampPosition.CUSTOM.value, variable=self.position_var,
                       command=self._on_position_change).grid(row=2, column=0, columnspan=2, sticky=tk.W, pady=10)
        
        custom_frame = ttk.Frame(frame)
        custom_frame.grid(row=3, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Label(custom_frame, text="X:").pack(side=tk.LEFT)
        self.custom_x_var = tk.DoubleVar(value=self.stamp_config.custom_x if self.stamp_config else 100)
        self.custom_x_spin = ttk.Spinbox(custom_frame, from_=0, to=1000, increment=10,
                                        textvariable=self.custom_x_var, width=10,
                                        command=self._update_preview)
        self.custom_x_spin.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(custom_frame, text="Y:").pack(side=tk.LEFT, padx=(20, 0))
        self.custom_y_var = tk.DoubleVar(value=self.stamp_config.custom_y if self.stamp_config else 100)
        self.custom_y_spin = ttk.Spinbox(custom_frame, from_=0, to=1000, increment=10,
                                        textvariable=self.custom_y_var, width=10,
                                        command=self._update_preview)
        self.custom_y_spin.pack(side=tk.LEFT, padx=5)
        
        # Margins
        margin_frame = ttk.LabelFrame(frame, text="Abstände vom Rand", padding="10")
        margin_frame.grid(row=4, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=10)
        
        ttk.Label(margin_frame, text="Horizontal:").grid(row=0, column=0, sticky=tk.W)
        self.margin_x_var = tk.DoubleVar(value=self.stamp_config.margin_x if self.stamp_config else 20)
        ttk.Spinbox(margin_frame, from_=0, to=100, increment=5,
                   textvariable=self.margin_x_var, width=10,
                   command=self._update_preview).grid(row=0, column=1, padx=5)
        
        ttk.Label(margin_frame, text="Vertikal:").grid(row=0, column=2, sticky=tk.W, padx=(20, 0))
        self.margin_y_var = tk.DoubleVar(value=self.stamp_config.margin_y if self.stamp_config else 20)
        ttk.Spinbox(margin_frame, from_=0, to=100, increment=5,
                   textvariable=self.margin_y_var, width=10,
                   command=self._update_preview).grid(row=0, column=3, padx=5)
        
        self._on_position_change()
    
    def _create_appearance_tab(self):
        """Erstellt den Tab für Aussehen"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Aussehen")
        
        # Rahmen
        border_frame = ttk.LabelFrame(frame, text="Rahmen", padding="10")
        border_frame.pack(fill=tk.X, pady=5)
        
        self.show_border_var = tk.BooleanVar(value=self.stamp_config.show_border if self.stamp_config else True)
        ttk.Checkbutton(border_frame, text="Rahmen anzeigen",
                       variable=self.show_border_var,
                       command=self._update_preview).grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        ttk.Label(border_frame, text="Breite:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.border_width_var = tk.DoubleVar(value=self.stamp_config.border_width if self.stamp_config else 2.0)
        ttk.Spinbox(border_frame, from_=0.5, to=10, increment=0.5,
                   textvariable=self.border_width_var, width=10,
                   command=self._update_preview).grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(border_frame, text="Farbe:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.border_color_var = tk.StringVar(value=self.stamp_config.border_color if self.stamp_config else "#000000")
        self.border_color_btn = tk.Button(border_frame, text="   ", bg=self.border_color_var.get(),
                                         command=lambda: self._choose_color('border'))
        self.border_color_btn.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(border_frame, text="Eckenradius:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.border_radius_var = tk.DoubleVar(value=self.stamp_config.border_radius if self.stamp_config else 5.0)
        ttk.Spinbox(border_frame, from_=0, to=20, increment=1,
                   textvariable=self.border_radius_var, width=10,
                   command=self._update_preview).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        # Hintergrund
        bg_frame = ttk.LabelFrame(frame, text="Hintergrund", padding="10")
        bg_frame.pack(fill=tk.X, pady=5)
        
        self.show_background_var = tk.BooleanVar(value=self.stamp_config.show_background if self.stamp_config else True)
        ttk.Checkbutton(bg_frame, text="Hintergrund anzeigen",
                       variable=self.show_background_var,
                       command=self._update_preview).grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        ttk.Label(bg_frame, text="Farbe:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.background_color_var = tk.StringVar(value=self.stamp_config.background_color if self.stamp_config else "#FFFFFF")
        self.background_color_btn = tk.Button(bg_frame, text="   ", bg=self.background_color_var.get(),
                                            command=lambda: self._choose_color('background'))
        self.background_color_btn.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(bg_frame, text="Transparenz:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.background_opacity_var = tk.DoubleVar(value=self.stamp_config.background_opacity if self.stamp_config else 0.8)
        opacity_frame = ttk.Frame(bg_frame)
        opacity_frame.grid(row=2, column=1, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Scale(opacity_frame, from_=0, to=1, variable=self.background_opacity_var,
                 command=lambda x: self._update_preview()).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.opacity_label = ttk.Label(opacity_frame, text=f"{int(self.background_opacity_var.get()*100)}%")
        self.opacity_label.pack(side=tk.LEFT, padx=5)
        
        self.background_opacity_var.trace('w', lambda *args: self.opacity_label.config(
            text=f"{int(self.background_opacity_var.get()*100)}%"))
        
        # Schatten
        shadow_frame = ttk.LabelFrame(frame, text="Schatten", padding="10")
        shadow_frame.pack(fill=tk.X, pady=5)
        
        self.show_shadow_var = tk.BooleanVar(value=self.stamp_config.show_shadow if self.stamp_config else True)
        ttk.Checkbutton(shadow_frame, text="Schatten anzeigen",
                       variable=self.show_shadow_var,
                       command=self._update_preview).grid(row=0, column=0, columnspan=4, sticky=tk.W)
        
        ttk.Label(shadow_frame, text="X-Offset:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.shadow_offset_x_var = tk.DoubleVar(value=self.stamp_config.shadow_offset_x if self.stamp_config else 2)
        ttk.Spinbox(shadow_frame, from_=-10, to=10, increment=1,
                   textvariable=self.shadow_offset_x_var, width=8,
                   command=self._update_preview).grid(row=1, column=1, pady=5)
        
        ttk.Label(shadow_frame, text="Y-Offset:").grid(row=1, column=2, sticky=tk.W, pady=5, padx=(10, 0))
        self.shadow_offset_y_var = tk.DoubleVar(value=self.stamp_config.shadow_offset_y if self.stamp_config else 2)
        ttk.Spinbox(shadow_frame, from_=-10, to=10, increment=1,
                   textvariable=self.shadow_offset_y_var, width=8,
                   command=self._update_preview).grid(row=1, column=3, pady=5)
        
        ttk.Label(shadow_frame, text="Farbe:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.shadow_color_var = tk.StringVar(value=self.stamp_config.shadow_color if self.stamp_config else "#808080")
        self.shadow_color_btn = tk.Button(shadow_frame, text="   ", bg=self.shadow_color_var.get(),
                                         command=lambda: self._choose_color('shadow'))
        self.shadow_color_btn.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(shadow_frame, text="Transparenz:").grid(row=2, column=2, sticky=tk.W, pady=5, padx=(10, 0))
        self.shadow_opacity_var = tk.DoubleVar(value=self.stamp_config.shadow_opacity if self.stamp_config else 0.3)
        shadow_opacity_frame = ttk.Frame(shadow_frame)
        shadow_opacity_frame.grid(row=2, column=3, sticky=(tk.W, tk.E), pady=5)
        
        ttk.Scale(shadow_opacity_frame, from_=0, to=1, variable=self.shadow_opacity_var,
                 command=lambda x: self._update_preview()).pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.shadow_opacity_label = ttk.Label(shadow_opacity_frame, text=f"{int(self.shadow_opacity_var.get()*100)}%")
        self.shadow_opacity_label.pack(side=tk.LEFT, padx=5)
        
        self.shadow_opacity_var.trace('w', lambda *args: self.shadow_opacity_label.config(
            text=f"{int(self.shadow_opacity_var.get()*100)}%"))
        
        # Größe
        size_frame = ttk.LabelFrame(frame, text="Größe", padding="10")
        size_frame.pack(fill=tk.X, pady=5)
        
        self.auto_size_var = tk.BooleanVar(value=self.stamp_config.auto_size if self.stamp_config else True)
        ttk.Checkbutton(size_frame, text="Automatische Größenanpassung",
                       variable=self.auto_size_var,
                       command=self._on_auto_size_change).grid(row=0, column=0, columnspan=2, sticky=tk.W)
        
        ttk.Label(size_frame, text="Breite:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.fixed_width_var = tk.DoubleVar(value=self.stamp_config.fixed_width if self.stamp_config else 200)
        self.width_spin = ttk.Spinbox(size_frame, from_=50, to=500, increment=10,
                                     textvariable=self.fixed_width_var, width=10,
                                     command=self._update_preview)
        self.width_spin.grid(row=1, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(size_frame, text="Höhe:").grid(row=2, column=0, sticky=tk.W, pady=5)
        self.fixed_height_var = tk.DoubleVar(value=self.stamp_config.fixed_height if self.stamp_config else 100)
        self.height_spin = ttk.Spinbox(size_frame, from_=20, to=300, increment=10,
                                      textvariable=self.fixed_height_var, width=10,
                                      command=self._update_preview)
        self.height_spin.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        ttk.Label(size_frame, text="Innenabstand:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.padding_var = tk.DoubleVar(value=self.stamp_config.padding if self.stamp_config else 10)
        ttk.Spinbox(size_frame, from_=0, to=50, increment=5,
                   textvariable=self.padding_var, width=10,
                   command=self._update_preview).grid(row=3, column=1, sticky=tk.W, pady=5)
        
        self._on_auto_size_change()
    
    def _create_content_tab(self):
        """Erstellt den Tab für Inhalt"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Inhalt")
        
        # Toolbar
        toolbar = ttk.Frame(frame)
        toolbar.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Button(toolbar, text="+ Zeile hinzufügen",
                  command=self._add_line).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="- Zeile entfernen",
                  command=self._remove_line).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="↑ Nach oben",
                  command=self._move_line_up).pack(side=tk.LEFT, padx=2)
        ttk.Button(toolbar, text="↓ Nach unten",
                  command=self._move_line_down).pack(side=tk.LEFT, padx=2)
        
        # Zeilen-Liste
        list_frame = ttk.Frame(frame)
        list_frame.pack(fill=tk.BOTH, expand=True)
        
        # Listbox mit Scrollbar
        scrollbar = ttk.Scrollbar(list_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.lines_listbox = tk.Listbox(list_frame, height=8,
                                       yscrollcommand=scrollbar.set,
                                       selectmode=tk.SINGLE)
        self.lines_listbox.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.config(command=self.lines_listbox.yview)
        
        self.lines_listbox.bind('<<ListboxSelect>>', self._on_line_select)
        
        # Zeilen-Editor
        editor_frame = ttk.LabelFrame(frame, text="Zeile bearbeiten", padding="10")
        editor_frame.pack(fill=tk.X, pady=10)
        
        # Text
        ttk.Label(editor_frame, text="Text:").grid(row=0, column=0, sticky=tk.W, pady=5)
        text_frame = ttk.Frame(editor_frame)
        text_frame.grid(row=0, column=1, sticky=(tk.W, tk.E), pady=5)
        
        self.line_text_var = tk.StringVar()
        self.line_text_entry = ttk.Entry(text_frame, textvariable=self.line_text_var)
        self.line_text_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        ttk.Button(text_frame, text="Expression...", width=12,
                  command=self._edit_expression).pack(side=tk.LEFT, padx=(5, 0))
        
        # Font
        ttk.Label(editor_frame, text="Schriftart:").grid(row=1, column=0, sticky=tk.W, pady=5)
        font_frame = ttk.Frame(editor_frame)
        font_frame.grid(row=1, column=1, sticky=(tk.W, tk.E), pady=5)
        
        self.line_font_var = tk.StringVar(value="Helvetica")
        self.font_combo = ttk.Combobox(font_frame, textvariable=self.line_font_var,
                                      values=["Helvetica", "Times-Roman", "Courier"],
                                      width=15)
        self.font_combo.pack(side=tk.LEFT)
        
        ttk.Label(font_frame, text="Größe:").pack(side=tk.LEFT, padx=(10, 5))
        self.line_size_var = tk.IntVar(value=12)
        ttk.Spinbox(font_frame, from_=6, to=72, increment=1,
                   textvariable=self.line_size_var, width=8).pack(side=tk.LEFT)
        
        ttk.Label(font_frame, text="Farbe:").pack(side=tk.LEFT, padx=(10, 5))
        self.line_color_var = tk.StringVar(value="#000000")
        self.line_color_btn = tk.Button(font_frame, text="   ", bg=self.line_color_var.get(),
                                       command=lambda: self._choose_color('line'))
        self.line_color_btn.pack(side=tk.LEFT)
        
        # Style
        ttk.Label(editor_frame, text="Stil:").grid(row=2, column=0, sticky=tk.W, pady=5)
        style_frame = ttk.Frame(editor_frame)
        style_frame.grid(row=2, column=1, sticky=tk.W, pady=5)
        
        self.line_bold_var = tk.BooleanVar()
        ttk.Checkbutton(style_frame, text="Fett",
                       variable=self.line_bold_var).pack(side=tk.LEFT, padx=5)
        
        self.line_italic_var = tk.BooleanVar()
        ttk.Checkbutton(style_frame, text="Kursiv",
                       variable=self.line_italic_var).pack(side=tk.LEFT, padx=5)
        
        # Alignment
        ttk.Label(editor_frame, text="Ausrichtung:").grid(row=3, column=0, sticky=tk.W, pady=5)
        self.line_align_var = tk.StringVar(value="left")
        align_frame = ttk.Frame(editor_frame)
        align_frame.grid(row=3, column=1, sticky=tk.W, pady=5)
        
        for text, value in [("Links", "left"), ("Zentriert", "center"), ("Rechts", "right")]:
            ttk.Radiobutton(align_frame, text=text, value=value,
                           variable=self.line_align_var).pack(side=tk.LEFT, padx=5)
        
        # Update button
        ttk.Button(editor_frame, text="Zeile aktualisieren",
                  command=self._update_current_line).grid(row=4, column=1, sticky=tk.E, pady=10)
        
        editor_frame.columnconfigure(1, weight=1)
        
        # Zeilen-Daten
        self.stamp_lines = []
    
    def _create_pages_tab(self):
        """Erstellt den Tab für Seiten-Einstellungen"""
        frame = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(frame, text="Seiten")
        
        ttk.Label(frame, text="Stempel anwenden auf:").pack(anchor=tk.W, pady=(0, 10))
        
        self.pages_var = tk.StringVar(value=self.stamp_config.apply_to_pages if self.stamp_config else "first")
        
        ttk.Radiobutton(frame, text="Nur erste Seite", value="first",
                       variable=self.pages_var,
                       command=self._on_pages_change).pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(frame, text="Nur letzte Seite", value="last",
                       variable=self.pages_var,
                       command=self._on_pages_change).pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(frame, text="Alle Seiten", value="all",
                       variable=self.pages_var,
                       command=self._on_pages_change).pack(anchor=tk.W, pady=2)
        ttk.Radiobutton(frame, text="Bestimmte Seiten:", value="custom",
                       variable=self.pages_var,
                       command=self._on_pages_change).pack(anchor=tk.W, pady=2)
        
        # Custom pages
        self.custom_pages_frame = ttk.Frame(frame)
        self.custom_pages_frame.pack(anchor=tk.W, padx=(20, 0), pady=5)
        
        ttk.Label(self.custom_pages_frame, text="Seitenzahlen (kommagetrennt):").pack(side=tk.LEFT)
        self.custom_pages_var = tk.StringVar()
        if self.stamp_config and self.stamp_config.custom_pages:
            self.custom_pages_var.set(','.join(map(str, self.stamp_config.custom_pages)))
        
        self.custom_pages_entry = ttk.Entry(self.custom_pages_frame, 
                                           textvariable=self.custom_pages_var,
                                           width=30)
        self.custom_pages_entry.pack(side=tk.LEFT, padx=5)
        
        ttk.Label(self.custom_pages_frame, text="(z.B. 1,3,5-8)", 
                 foreground="gray").pack(side=tk.LEFT, padx=5)
        
        self._on_pages_change()
    
    def _create_preview_ui(self, parent):
        """Erstellt die Vorschau-UI"""
        preview_frame = ttk.LabelFrame(parent, text="Vorschau", padding="10")
        preview_frame.pack(fill=tk.BOTH, expand=True)
        
        # Canvas für Vorschau
        self.preview_canvas = tk.Canvas(preview_frame, bg='white', highlightthickness=1)
        self.preview_canvas.pack(fill=tk.BOTH, expand=True)
        
        # Info
        info_frame = ttk.Frame(preview_frame)
        info_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.preview_info = ttk.Label(info_frame, text="", foreground="gray")
        self.preview_info.pack()
        
        # Preview object
        self.preview = StampPreview(self.preview_canvas)
    
    def _add_line(self):
        """Fügt eine neue Zeile hinzu"""
        new_line = {
            "text": "Neue Zeile",
            "font_name": "Helvetica",
            "font_size": 12,
            "font_color": "#000000",
            "bold": False,
            "italic": False,
            "alignment": "left"
        }
        self.stamp_lines.append(new_line)
        self.lines_listbox.insert(tk.END, self._format_line_display(new_line))
        self.lines_listbox.selection_clear(0, tk.END)
        self.lines_listbox.selection_set(tk.END)
        self._on_line_select()
        self._update_preview()
    
    def _add_default_lines(self):
        """Fügt Standard-Zeilen hinzu"""
        default_lines = [
            {
                "text": "EINGANG",
                "font_name": "Helvetica",
                "font_size": 16,
                "font_color": "#FF0000",
                "bold": True,
                "italic": False,
                "alignment": "center"
            },
            {
                "text": "<Date>",
                "font_name": "Helvetica",
                "font_size": 12,
                "font_color": "#000000",
                "bold": False,
                "italic": False,
                "alignment": "center"
            },
            {
                "text": "<Time>",
                "font_name": "Helvetica",
                "font_size": 10,
                "font_color": "#000000",
                "bold": False,
                "italic": False,
                "alignment": "center"
            }
        ]
        
        for line in default_lines:
            self.stamp_lines.append(line)
            self.lines_listbox.insert(tk.END, self._format_line_display(line))
    
    def _remove_line(self):
        """Entfernt die ausgewählte Zeile"""
        selection = self.lines_listbox.curselection()
        if selection:
            index = selection[0]
            self.lines_listbox.delete(index)
            del self.stamp_lines[index]
            self._update_preview()
    
    def _move_line_up(self):
        """Verschiebt die ausgewählte Zeile nach oben"""
        selection = self.lines_listbox.curselection()
        if selection and selection[0] > 0:
            index = selection[0]
            # Swap in list
            self.stamp_lines[index], self.stamp_lines[index-1] = \
                self.stamp_lines[index-1], self.stamp_lines[index]
            # Update listbox
            self._refresh_listbox()
            self.lines_listbox.selection_set(index-1)
            self._update_preview()
    
    def _move_line_down(self):
        """Verschiebt die ausgewählte Zeile nach unten"""
        selection = self.lines_listbox.curselection()
        if selection and selection[0] < len(self.stamp_lines)-1:
            index = selection[0]
            # Swap in list
            self.stamp_lines[index], self.stamp_lines[index+1] = \
                self.stamp_lines[index+1], self.stamp_lines[index]
            # Update listbox
            self._refresh_listbox()
            self.lines_listbox.selection_set(index+1)
            self._update_preview()
    
    def _refresh_listbox(self):
        """Aktualisiert die Listbox"""
        self.lines_listbox.delete(0, tk.END)
        for line in self.stamp_lines:
            self.lines_listbox.insert(tk.END, self._format_line_display(line))
    
    def _format_line_display(self, line):
        """Formatiert eine Zeile für die Anzeige"""
        style = []
        if line.get('bold', False):
            style.append('B')
        if line.get('italic', False):
            style.append('I')
        style_str = f" [{','.join(style)}]" if style else ""
        
        text = line['text'][:50]
        if len(line['text']) > 50:
            text += "..."
            
        return f"{text} ({line['font_size']}pt{style_str})"
    
    def _on_line_select(self, event=None):
        """Wird aufgerufen wenn eine Zeile ausgewählt wird"""
        selection = self.lines_listbox.curselection()
        if selection:
            index = selection[0]
            line = self.stamp_lines[index]
            
            # Lade Zeilen-Einstellungen
            self.line_text_var.set(line['text'])
            self.line_font_var.set(line['font_name'])
            self.line_size_var.set(line['font_size'])
            self.line_color_var.set(line['font_color'])
            self.line_bold_var.set(line.get('bold', False))
            self.line_italic_var.set(line.get('italic', False))
            self.line_align_var.set(line.get('alignment', 'left'))
            
            # Update color button
            self.line_color_btn.config(bg=line['font_color'])
    
    def _update_current_line(self):
        """Aktualisiert die aktuelle Zeile"""
        selection = self.lines_listbox.curselection()
        if selection:
            index = selection[0]
            
            # Update line data
            self.stamp_lines[index] = {
                "text": self.line_text_var.get(),
                "font_name": self.line_font_var.get(),
                "font_size": self.line_size_var.get(),
                "font_color": self.line_color_var.get(),
                "bold": self.line_bold_var.get(),
                "italic": self.line_italic_var.get(),
                "alignment": self.line_align_var.get()
            }
            
            # Update listbox display
            self.lines_listbox.delete(index)
            self.lines_listbox.insert(index, self._format_line_display(self.stamp_lines[index]))
            self.lines_listbox.selection_set(index)
            
            self._update_preview()
    
    def _edit_expression(self):
        """Öffnet den Expression Editor"""
        dialog = ExpressionDialog(self.dialog, 
                                 title="Stempel-Text bearbeiten",
                                 expression=self.line_text_var.get(),
                                 description="Verwenden Sie Variablen für dynamische Inhalte",
                                 xml_field_mappings=self.xml_fields)
        
        # Erstelle kombinierte Zonen und Felder für den Dialog
        dialog.ocr_zones = self.ocr_zones
        
        result = dialog.show()
        if result:
            self.line_text_var.set(result)
    
    def _choose_color(self, target):
        """Öffnet den Farbauswahl-Dialog"""
        if target == 'border':
            color = colorchooser.askcolor(initialcolor=self.border_color_var.get())
            if color[1]:
                self.border_color_var.set(color[1])
                self.border_color_btn.config(bg=color[1])
        elif target == 'background':
            color = colorchooser.askcolor(initialcolor=self.background_color_var.get())
            if color[1]:
                self.background_color_var.set(color[1])
                self.background_color_btn.config(bg=color[1])
        elif target == 'shadow':
            color = colorchooser.askcolor(initialcolor=self.shadow_color_var.get())
            if color[1]:
                self.shadow_color_var.set(color[1])
                self.shadow_color_btn.config(bg=color[1])
        elif target == 'line':
            color = colorchooser.askcolor(initialcolor=self.line_color_var.get())
            if color[1]:
                self.line_color_var.set(color[1])
                self.line_color_btn.config(bg=color[1])
        
        self._update_preview()
    
    def _on_position_change(self):
        """Wird aufgerufen wenn die Position geändert wird"""
        is_custom = self.position_var.get() == StampPosition.CUSTOM.value
        state = tk.NORMAL if is_custom else tk.DISABLED
        self.custom_x_spin.config(state=state)
        self.custom_y_spin.config(state=state)
        self._update_preview()
    
    def _on_auto_size_change(self):
        """Wird aufgerufen wenn Auto-Size geändert wird"""
        state = tk.DISABLED if self.auto_size_var.get() else tk.NORMAL
        self.width_spin.config(state=state)
        self.height_spin.config(state=state)
        self._update_preview()
    
    def _on_pages_change(self):
        """Wird aufgerufen wenn die Seiten-Einstellung geändert wird"""
        is_custom = self.pages_var.get() == "custom"
        state = tk.NORMAL if is_custom else tk.DISABLED
        self.custom_pages_entry.config(state=state)
    
    def _update_preview(self):
        """Aktualisiert die Vorschau"""
        if hasattr(self, 'preview'):
            config = self._build_config()
            self.preview.update_stamp(config)
            
            # Update info
            if config.auto_size:
                self.preview_info.config(text="Größe: Automatisch angepasst")
            else:
                self.preview_info.config(text=f"Größe: {config.fixed_width} x {config.fixed_height}")
    
    def _build_config(self):
        """Erstellt eine StampConfig aus den aktuellen Einstellungen"""
        return StampConfig(
            id=self.stamp_config.id if self.stamp_config else str(uuid.uuid4()),
            name=self.name_var.get(),
            enabled=self.enabled_var.get(),
            position=StampPosition(self.position_var.get()),
            custom_x=self.custom_x_var.get(),
            custom_y=self.custom_y_var.get(),
            margin_x=self.margin_x_var.get(),
            margin_y=self.margin_y_var.get(),
            orientation=StampOrientation(self.orientation_var.get()),
            rotation=self.rotation_var.get(),
            show_border=self.show_border_var.get(),
            border_width=self.border_width_var.get(),
            border_color=self.border_color_var.get(),
            border_radius=self.border_radius_var.get(),
            show_background=self.show_background_var.get(),
            background_color=self.background_color_var.get(),
            background_opacity=self.background_opacity_var.get(),
            show_shadow=self.show_shadow_var.get(),
            shadow_offset_x=self.shadow_offset_x_var.get(),
            shadow_offset_y=self.shadow_offset_y_var.get(),
            shadow_color=self.shadow_color_var.get(),
            shadow_opacity=self.shadow_opacity_var.get(),
            stamp_lines=self.stamp_lines,
            apply_to_pages=self.pages_var.get(),
            custom_pages=self._parse_custom_pages(),
            auto_size=self.auto_size_var.get(),
            fixed_width=self.fixed_width_var.get(),
            fixed_height=self.fixed_height_var.get(),
            padding=self.padding_var.get()
        )
    
    def _parse_custom_pages(self):
        """Parst die benutzerdefinierten Seitenzahlen"""
        pages = []
        if self.custom_pages_var.get():
            parts = self.custom_pages_var.get().split(',')
            for part in parts:
                part = part.strip()
                if '-' in part:
                    # Range
                    try:
                        start, end = part.split('-')
                        pages.extend(range(int(start), int(end)+1))
                    except:
                        pass
                else:
                    # Single page
                    try:
                        pages.append(int(part))
                    except:
                        pass
        return sorted(set(pages))
    
    def _load_config(self, config: StampConfig):
        """Lädt eine vorhandene Konfiguration"""
        # Lade Zeilen
        self.stamp_lines = config.stamp_lines.copy()
        for line in self.stamp_lines:
            self.lines_listbox.insert(tk.END, self._format_line_display(line))
        
        # Setze custom pages
        if config.custom_pages:
            self.custom_pages_var.set(','.join(map(str, config.custom_pages)))
    
    def _validate(self):
        """Validiert die Eingaben"""
        if not self.name_var.get().strip():
            messagebox.showerror("Fehler", "Bitte geben Sie einen Namen ein.")
            return False
        
        if not self.stamp_lines:
            messagebox.showerror("Fehler", "Der Stempel muss mindestens eine Zeile enthalten.")
            return False
        
        # Validiere custom pages
        if self.pages_var.get() == "custom":
            pages = self._parse_custom_pages()
            if not pages:
                messagebox.showerror("Fehler", "Bitte geben Sie gültige Seitenzahlen ein.")
                return False
        
        return True
    
    def _on_save(self):
        """Speichert die Konfiguration"""
        if not self._validate():
            return
        
        self.result = self._build_config()
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab"""
        self.dialog.destroy()
    
    def show(self):
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result