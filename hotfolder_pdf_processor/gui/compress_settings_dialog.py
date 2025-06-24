"""
Dialog für PDF-Komprimierungseinstellungen
"""
import tkinter as tk
from tkinter import ttk, messagebox
from typing import Dict, Optional
import sys
import os
import subprocess

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


class CompressSettingsDialog:
    """Dialog zur Konfiguration der PDF-Komprimierung"""
    
    def __init__(self, parent, initial_params: Optional[Dict] = None):
        self.parent = parent
        self.result = None
        
        # Standard-Parameter
        self.default_params = {
            'compression_level': 'custom',
            'color_dpi': 150,
            'gray_dpi': 150,
            'mono_dpi': 150,
            'jpeg_quality': 85,
            'color_compression': 'jpeg',
            'gray_compression': 'jpeg',
            'mono_compression': 'ccitt',
            'downsample_images': True,
            'subset_fonts': True,
            'remove_duplicates': True,
            'optimize': True
        }
        
        # Initiale Parameter überschreiben Standard
        self.params = self.default_params.copy()
        if initial_params:
            self.params.update(initial_params)
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("PDF-Komprimierungseinstellungen")
        self.dialog.geometry("700x750")
        self.dialog.resizable(False, False)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Variablen
        self.preset_var = tk.StringVar(value=self.params.get('compression_level', 'custom'))
        self.color_dpi_var = tk.IntVar(value=self.params.get('color_dpi', 150))
        self.gray_dpi_var = tk.IntVar(value=self.params.get('gray_dpi', 150))
        self.mono_dpi_var = tk.IntVar(value=self.params.get('mono_dpi', 150))
        self.jpeg_quality_var = tk.IntVar(value=self.params.get('jpeg_quality', 85))
        self.color_compression_var = tk.StringVar(value=self.params.get('color_compression', 'jpeg'))
        self.gray_compression_var = tk.StringVar(value=self.params.get('gray_compression', 'jpeg'))
        self.mono_compression_var = tk.StringVar(value=self.params.get('mono_compression', 'ccitt'))
        self.downsample_var = tk.BooleanVar(value=self.params.get('downsample_images', True))
        self.subset_fonts_var = tk.BooleanVar(value=self.params.get('subset_fonts', True))
        self.remove_duplicates_var = tk.BooleanVar(value=self.params.get('remove_duplicates', True))
        self.optimize_var = tk.BooleanVar(value=self.params.get('optimize', True))
        
        self._create_widgets()
        self._layout_widgets()
        
        # Prüfe Ghostscript-Verfügbarkeit
        if not self._check_ghostscript():
            messagebox.showerror(
                "Ghostscript fehlt", 
                "Ghostscript wurde nicht gefunden!\n\n"
                "Die PDF-Komprimierung benötigt Ghostscript.\n"
                "Bitte installieren Sie Ghostscript von:\n"
                "https://www.ghostscript.com/download/gsdnld.html"
            )
            self.save_button.config(state=tk.DISABLED)
        
        # Bind Events
        self.dialog.bind('<Return>', lambda e: self._on_save())
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
        self.main_frame = ttk.Frame(self.dialog, padding="20")
        
        # Voreinstellungen
        self.preset_frame = ttk.LabelFrame(self.main_frame, text="Voreinstellungen", padding="10")
        
        self.preset_label = ttk.Label(self.preset_frame,
            text="Wählen Sie eine Voreinstellung oder passen Sie die Werte manuell an:")
        
        # Dropdown für Voreinstellungen
        self.preset_combo = ttk.Combobox(self.preset_frame, 
            textvariable=self.preset_var,
            values=["Niedrig (beste Qualität)", "Mittel (ausgewogen)", 
                    "Hoch (kleine Dateigröße)", "Maximum (kleinste Dateigröße)", 
                    "Benutzerdefiniert"],
            state="readonly", width=30)
        self.preset_combo.bind('<<ComboboxSelected>>', self._on_preset_changed)
        
        # Setze initiale Anzeige
        self._update_preset_display()
        
        # Bildauflösung
        self.resolution_frame = ttk.LabelFrame(self.main_frame, text="Bildauflösung (DPI)", padding="10")
        
        # Farbbilder
        self.color_dpi_label = ttk.Label(self.resolution_frame, text="Farbbilder:")
        self.color_dpi_spinbox = ttk.Spinbox(self.resolution_frame, 
            from_=72, to=600, increment=50,
            textvariable=self.color_dpi_var, width=10,
            command=self._on_manual_change)
        
        # Graustufenbilder
        self.gray_dpi_label = ttk.Label(self.resolution_frame, text="Graustufenbilder:")
        self.gray_dpi_spinbox = ttk.Spinbox(self.resolution_frame, 
            from_=72, to=600, increment=50,
            textvariable=self.gray_dpi_var, width=10,
            command=self._on_manual_change)
        
        # Schwarz-Weiß-Bilder
        self.mono_dpi_label = ttk.Label(self.resolution_frame, text="Schwarz-Weiß-Bilder:")
        self.mono_dpi_spinbox = ttk.Spinbox(self.resolution_frame, 
            from_=72, to=600, increment=50,
            textvariable=self.mono_dpi_var, width=10,
            command=self._on_manual_change)
        
        # Komprimierungsmethoden
        self.compression_frame = ttk.LabelFrame(self.main_frame, text="Komprimierungsmethoden", padding="10")
        
        # Farbbilder
        self.color_comp_label = ttk.Label(self.compression_frame, text="Farbbilder:")
        self.color_comp_combo = ttk.Combobox(self.compression_frame, 
            textvariable=self.color_compression_var,
            values=["jpeg", "jpeg2000", "zip", "none"],
            state="readonly", width=15)
        self.color_comp_combo.bind('<<ComboboxSelected>>', lambda e: self._on_manual_change())
        
        # Graustufenbilder
        self.gray_comp_label = ttk.Label(self.compression_frame, text="Graustufenbilder:")
        self.gray_comp_combo = ttk.Combobox(self.compression_frame, 
            textvariable=self.gray_compression_var,
            values=["jpeg", "jpeg2000", "zip", "none"],
            state="readonly", width=15)
        self.gray_comp_combo.bind('<<ComboboxSelected>>', lambda e: self._on_manual_change())
        
        # Schwarz-Weiß-Bilder
        self.mono_comp_label = ttk.Label(self.compression_frame, text="Schwarz-Weiß-Bilder:")
        self.mono_comp_combo = ttk.Combobox(self.compression_frame, 
            textvariable=self.mono_compression_var,
            values=["ccitt", "zip", "none"],  # JBIG2 entfernt wegen Kompatibilität
            state="readonly", width=15)
        self.mono_comp_combo.bind('<<ComboboxSelected>>', lambda e: self._on_manual_change())
        
        # JPEG-Qualität
        self.quality_frame = ttk.LabelFrame(self.main_frame, text="JPEG-Qualität", padding="10")
        
        self.quality_label = ttk.Label(self.quality_frame, text="Qualität (gilt für JPEG und JPEG2000):")
        self.quality_scale = ttk.Scale(self.quality_frame, 
            from_=10, to=100,
            variable=self.jpeg_quality_var,
            orient=tk.HORIZONTAL,
            length=500,
            command=self._update_quality_label)
        
        self.quality_value_label = ttk.Label(self.quality_frame, text="85%")
        
        # Erweiterte Optionen
        self.advanced_frame = ttk.LabelFrame(self.main_frame, text="Erweiterte Optionen", padding="10")
        
        self.downsample_check = ttk.Checkbutton(self.advanced_frame,
            text="Bilder herunterskalieren (Downsampling)",
            variable=self.downsample_var,
            command=self._on_manual_change)
        
        self.subset_fonts_check = ttk.Checkbutton(self.advanced_frame,
            text="Nur verwendete Zeichen in Schriften einbetten (Font Subsetting)",
            variable=self.subset_fonts_var,
            command=self._on_manual_change)
        
        self.remove_duplicates_check = ttk.Checkbutton(self.advanced_frame,
            text="Doppelte Bilder und Objekte entfernen",
            variable=self.remove_duplicates_var,
            command=self._on_manual_change)
        
        self.optimize_check = ttk.Checkbutton(self.advanced_frame,
            text="PDF-Struktur optimieren",
            variable=self.optimize_var,
            command=self._on_manual_change)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save, default=tk.ACTIVE)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Voreinstellungen
        self.preset_frame.pack(fill=tk.X, pady=(0, 15))
        self.preset_label.pack(anchor=tk.W, pady=(0, 10))
        self.preset_combo.pack(fill=tk.X)
        
        # Bildauflösung
        self.resolution_frame.pack(fill=tk.X, pady=(0, 15))
        self.color_dpi_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.color_dpi_spinbox.grid(row=0, column=1, sticky=tk.W)
        self.gray_dpi_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.gray_dpi_spinbox.grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
        self.mono_dpi_label.grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.mono_dpi_spinbox.grid(row=2, column=1, sticky=tk.W, pady=(5, 0))
        
        # Komprimierungsmethoden
        self.compression_frame.pack(fill=tk.X, pady=(0, 15))
        self.color_comp_label.grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        self.color_comp_combo.grid(row=0, column=1, sticky=tk.W)
        self.gray_comp_label.grid(row=1, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.gray_comp_combo.grid(row=1, column=1, sticky=tk.W, pady=(5, 0))
        self.mono_comp_label.grid(row=2, column=0, sticky=tk.W, padx=(0, 10), pady=(5, 0))
        self.mono_comp_combo.grid(row=2, column=1, sticky=tk.W, pady=(5, 0))
        
        # JPEG-Qualität
        self.quality_frame.pack(fill=tk.X, pady=(0, 15))
        self.quality_label.grid(row=0, column=0, columnspan=2, sticky=tk.W)
        self.quality_scale.grid(row=1, column=0, sticky="we", pady=(5, 0))
        self.quality_value_label.grid(row=1, column=1, padx=(10, 0))
        self.quality_frame.columnconfigure(0, weight=1)
        
        # Erweiterte Optionen
        self.advanced_frame.pack(fill=tk.X, pady=(0, 15))
        self.downsample_check.pack(anchor=tk.W, pady=2)
        self.subset_fonts_check.pack(anchor=tk.W, pady=2)
        self.remove_duplicates_check.pack(anchor=tk.W, pady=2)
        self.optimize_check.pack(anchor=tk.W, pady=2)
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.save_button.pack(side=tk.RIGHT)
        
        # Update initial quality label
        self._update_quality_label(self.jpeg_quality_var.get())
        
        # Wende initiale Voreinstellung an
        if self.preset_var.get() != 'custom':
            self._apply_preset()
    
    def _check_ghostscript(self) -> bool:
        """Prüft ob Ghostscript verfügbar ist"""
        try:
            # Windows
            if os.name == 'nt':
                try:
                    result = subprocess.run(['gswin64c', '--version'], 
                                          capture_output=True, text=True)
                    if result.returncode == 0:
                        return True
                except:
                    pass
                
                try:
                    result = subprocess.run(['gswin32c', '--version'], 
                                          capture_output=True, text=True)
                    if result.returncode == 0:
                        return True
                except:
                    pass
            else:
                # Unix/Linux
                result = subprocess.run(['gs', '--version'], 
                                      capture_output=True, text=True)
                if result.returncode == 0:
                    return True
            
            return False
            
        except Exception:
            return False
    
    def _update_preset_display(self):
        """Aktualisiert die Anzeige des Preset-Dropdowns"""
        preset_map = {
            'low': 'Niedrig (beste Qualität)',
            'medium': 'Mittel (ausgewogen)',
            'high': 'Hoch (kleine Dateigröße)',
            'maximum': 'Maximum (kleinste Dateigröße)',
            'custom': 'Benutzerdefiniert'
        }
        display_value = preset_map.get(self.preset_var.get(), 'Benutzerdefiniert')
        self.preset_combo.set(display_value)
    
    def _on_preset_changed(self, event=None):
        """Wird aufgerufen wenn eine Voreinstellung ausgewählt wird"""
        selected = self.preset_combo.get()
        
        # Map Display-Wert zum internen Wert
        preset_map = {
            'Niedrig (beste Qualität)': 'low',
            'Mittel (ausgewogen)': 'medium',
            'Hoch (kleine Dateigröße)': 'high',
            'Maximum (kleinste Dateigröße)': 'maximum',
            'Benutzerdefiniert': 'custom'
        }
        
        preset = preset_map.get(selected, 'custom')
        self.preset_var.set(preset)
        
        if preset != 'custom':
            self._apply_preset()
    
    def _apply_preset(self):
        """Wendet die gewählte Voreinstellung an"""
        preset = self.preset_var.get()
        
        if preset == 'low':
            # Beste Qualität
            self.color_dpi_var.set(300)
            self.gray_dpi_var.set(300)
            self.mono_dpi_var.set(300)
            self.jpeg_quality_var.set(95)
            self.color_compression_var.set('jpeg')
            self.gray_compression_var.set('jpeg')
            self.mono_compression_var.set('ccitt')
            self.downsample_var.set(False)
            self.subset_fonts_var.set(True)
            self.remove_duplicates_var.set(True)
            self.optimize_var.set(True)
            
        elif preset == 'medium':
            # Ausgewogen
            self.color_dpi_var.set(150)
            self.gray_dpi_var.set(150)
            self.mono_dpi_var.set(150)
            self.jpeg_quality_var.set(85)
            self.color_compression_var.set('jpeg')
            self.gray_compression_var.set('jpeg')
            self.mono_compression_var.set('ccitt')
            self.downsample_var.set(True)
            self.subset_fonts_var.set(True)
            self.remove_duplicates_var.set(True)
            self.optimize_var.set(True)
            
        elif preset == 'high':
            # Kleine Dateigröße
            self.color_dpi_var.set(72)
            self.gray_dpi_var.set(72)
            self.mono_dpi_var.set(72)
            self.jpeg_quality_var.set(70)
            self.color_compression_var.set('jpeg')
            self.gray_compression_var.set('jpeg')
            self.mono_compression_var.set('ccitt')
            self.downsample_var.set(True)
            self.subset_fonts_var.set(True)
            self.remove_duplicates_var.set(True)
            self.optimize_var.set(True)
            
        elif preset == 'maximum':
            # Kleinste Dateigröße
            self.color_dpi_var.set(72)
            self.gray_dpi_var.set(72)
            self.mono_dpi_var.set(72)
            self.jpeg_quality_var.set(50)
            self.color_compression_var.set('jpeg2000')
            self.gray_compression_var.set('jpeg2000')
            self.mono_compression_var.set('ccitt')  # CCITT statt JBIG2
            self.downsample_var.set(True)
            self.subset_fonts_var.set(True)
            self.remove_duplicates_var.set(True)
            self.optimize_var.set(True)
        
        self._update_quality_label(self.jpeg_quality_var.get())
    
    def _on_manual_change(self):
        """Wird aufgerufen wenn ein Wert manuell geändert wird"""
        # Setze auf "Benutzerdefiniert"
        self.preset_var.set('custom')
        self._update_preset_display()
    
    def _update_quality_label(self, value):
        """Aktualisiert das Qualitätslabel"""
        try:
            quality = int(float(value))
            self.quality_value_label.config(text=f"{quality}%")
            
            # Farbcodierung
            if quality >= 90:
                color = "green"
            elif quality >= 70:
                color = "orange"
            else:
                color = "red"
            
            self.quality_value_label.config(foreground=color)
            
            # Bei manueller Änderung
            self._on_manual_change()
        except:
            pass
    
    def _on_save(self):
        """Speichert die Einstellungen"""
        self.result = {
            'compression_level': self.preset_var.get(),
            'color_dpi': self.color_dpi_var.get(),
            'gray_dpi': self.gray_dpi_var.get(),
            'mono_dpi': self.mono_dpi_var.get(),
            'jpeg_quality': self.jpeg_quality_var.get(),
            'color_compression': self.color_compression_var.get(),
            'gray_compression': self.gray_compression_var.get(),
            'mono_compression': self.mono_compression_var.get(),
            'downsample_images': self.downsample_var.get(),
            'subset_fonts': self.subset_fonts_var.get(),
            'remove_duplicates': self.remove_duplicates_var.get(),
            'optimize': self.optimize_var.get()
        }
        
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht den Dialog ab"""
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result