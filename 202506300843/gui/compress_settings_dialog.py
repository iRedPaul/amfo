"""
Dialog für PDF-Komprimierungseinstellungen
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, Optional, Any
import sys
import os
import subprocess
import threading
import shutil
import tempfile
import logging

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

logger = logging.getLogger(__name__)


class CompressSettingsDialog:
    """Dialog zur Konfiguration der PDF-Komprimierung"""
    
    def __init__(self, parent, initial_params: Optional[Dict] = None):
        self.parent = parent
        self.result = None
        self.test_running = False
        self.compressed_pdf_path = None
        self.temp_dir = None
        
        # Standard-Parameter
        self.default_params = {
            'compression_level': 'Niedrig - Beste Qualität',
            'color_dpi': 200,
            'gray_dpi': 200,
            'mono_dpi': 200,
            'jpeg_quality': 90,
            'color_compression': 'jpeg',
            'gray_compression': 'jpeg',
            'mono_compression': 'ccitt',
            'downsample_images': False,
            'subset_fonts': True,
            'remove_duplicates': True,
            'optimize': True
        }
        
        # Übernehme initiale Parameter
        self.params = self.default_params.copy()
        if initial_params:
            self.params.update(initial_params)
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("PDF-Komprimierungseinstellungen")
        self.dialog.geometry("600x560")
        self.dialog.resizable(False, False)
        
        # Dialog zentrieren
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() - 600) // 2
        y = (self.dialog.winfo_screenheight() - 560) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Flag für Preset-Änderungen
        self.changing_preset = False
        
        # GUI erstellen
        self._create_gui()
        
        # Events binden
        self.dialog.bind('<Return>', lambda e: self._on_ok())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)
        
        # Ghostscript prüfen
        self._check_ghostscript()
    
    def _create_gui(self):
        """Erstellt die komplette GUI"""
        # Hauptframe
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Preset-Auswahl
        preset_frame = ttk.LabelFrame(main_frame, text="Komprimierungsstärke", padding="10")
        preset_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(preset_frame, text="Voreinstellung:").pack(anchor=tk.W, pady=(0, 5))
        
        self.preset_var = tk.StringVar()
        self.preset_combo = ttk.Combobox(
            preset_frame,
            textvariable=self.preset_var,
            values=[
                "Niedrig - Beste Qualität",
                "Mittel - Ausgewogen", 
                "Hoch - Kleine Dateigröße",
                "Benutzerdefiniert"
            ],
            state="readonly",
            width=35
        )
        self.preset_combo.pack(fill=tk.X)
        self.preset_combo.bind('<<ComboboxSelected>>', self._on_preset_changed)
        
        # Einstellungen
        settings_frame = ttk.LabelFrame(main_frame, text="Einstellungen", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 15))
        
        # DPI
        dpi_frame = ttk.Frame(settings_frame)
        dpi_frame.pack(fill=tk.X, pady=(0, 10))
        
        ttk.Label(dpi_frame, text="Bildauflösung (DPI):").pack(anchor=tk.W)
        
        dpi_control_frame = ttk.Frame(dpi_frame)
        dpi_control_frame.pack(fill=tk.X)
        
        self.dpi_var = tk.IntVar(value=self.params.get('color_dpi', 150))
        self.dpi_scale = ttk.Scale(
            dpi_control_frame,
            from_=72, to=300,
            variable=self.dpi_var,
            orient=tk.HORIZONTAL,
            length=350,
            command=lambda v: self._on_value_changed()
        )
        self.dpi_scale.pack(side=tk.LEFT, padx=(0, 10))
        
        self.dpi_label = ttk.Label(dpi_control_frame, text="150 DPI", width=10)
        self.dpi_label.pack(side=tk.LEFT)
        
        # JPEG-Qualität
        quality_frame = ttk.Frame(settings_frame)
        quality_frame.pack(fill=tk.X)
        
        ttk.Label(quality_frame, text="JPEG-Qualität:").pack(anchor=tk.W)
        
        quality_control_frame = ttk.Frame(quality_frame)
        quality_control_frame.pack(fill=tk.X)
        
        self.quality_var = tk.IntVar(value=self.params.get('jpeg_quality', 85))
        self.quality_scale = ttk.Scale(
            quality_control_frame,
            from_=10, to=100,
            variable=self.quality_var,
            orient=tk.HORIZONTAL,
            length=350,
            command=lambda v: self._on_value_changed()
        )
        self.quality_scale.pack(side=tk.LEFT, padx=(0, 10))
        
        self.quality_label = ttk.Label(quality_control_frame, text="85%", width=10)
        self.quality_label.pack(side=tk.LEFT)
        
        # Erweiterte Optionen
        advanced_frame = ttk.LabelFrame(main_frame, text="Erweiterte Optionen", padding="10")
        advanced_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Optionen in 2 Spalten
        left_frame = ttk.Frame(advanced_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        right_frame = ttk.Frame(advanced_frame)
        right_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(20, 0))
        
        self.downsample_var = tk.BooleanVar(value=self.params.get('downsample_images', True))
        self.downsample_check = ttk.Checkbutton(
            left_frame,
            text="Bilder herunterskalieren",
            variable=self.downsample_var,
            command=self._on_value_changed
        )
        self.downsample_check.pack(anchor=tk.W, pady=2)
        
        self.subset_fonts_var = tk.BooleanVar(value=self.params.get('subset_fonts', True))
        self.subset_fonts_check = ttk.Checkbutton(
            left_frame,
            text="Schriften optimieren",
            variable=self.subset_fonts_var,
            command=self._on_value_changed
        )
        self.subset_fonts_check.pack(anchor=tk.W, pady=2)
        
        self.remove_duplicates_var = tk.BooleanVar(value=self.params.get('remove_duplicates', True))
        self.remove_duplicates_check = ttk.Checkbutton(
            right_frame,
            text="Duplikate entfernen",
            variable=self.remove_duplicates_var,
            command=self._on_value_changed
        )
        self.remove_duplicates_check.pack(anchor=tk.W, pady=2)
        
        self.optimize_var = tk.BooleanVar(value=self.params.get('optimize', True))
        self.optimize_check = ttk.Checkbutton(
            right_frame,
            text="PDF optimieren",
            variable=self.optimize_var,
            command=self._on_value_changed
        )
        self.optimize_check.pack(anchor=tk.W, pady=2)
        
        # Test-Bereich
        test_frame = ttk.LabelFrame(main_frame, text="Test", padding="10")
        test_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.test_button = ttk.Button(
            test_frame,
            text="Mit PDF-Datei testen...",
            command=self._test_compression
        )
        self.test_button.pack(anchor=tk.W)
        
        self.preview_button = ttk.Button(
            test_frame,
            text="Komprimierte PDF anzeigen",
            command=self._show_preview,
            state=tk.DISABLED
        )
        self.preview_button.pack(anchor=tk.W, pady=(5, 0))
        
        self.test_progress = ttk.Progressbar(
            test_frame,
            mode='indeterminate',
            length=400
        )
        
        self.test_result_label = ttk.Label(test_frame, text="")
        self.test_result_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.cancel_button = ttk.Button(
            button_frame,
            text="Abbrechen",
            command=self._on_cancel
        )
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        
        self.ok_button = ttk.Button(
            button_frame,
            text="OK",
            command=self._on_ok,
            default=tk.ACTIVE
        )
        self.ok_button.pack(side=tk.RIGHT)
        
        # Initiale Werte setzen
        self._set_initial_values()
        self._update_labels()
    
    def _set_initial_values(self):
        """Setzt die initialen Werte basierend auf den Parametern"""
        # Preset bestimmen basierend auf compression_level oder DPI/Quality
        preset = self.params.get('compression_level', '')
        
        # Wenn preset direkt gesetzt ist, verwende es
        if preset in ["Niedrig - Beste Qualität", "Mittel - Ausgewogen", 
                      "Hoch - Kleine Dateigröße", "Benutzerdefiniert"]:
            self.preset_var.set(preset)
        else:
            # Ansonsten bestimme Preset basierend auf DPI/Quality
            dpi = self.params.get('color_dpi', 200)
            quality = self.params.get('jpeg_quality', 90)
            
            if dpi == 200 and quality == 90:
                preset = "Niedrig - Beste Qualität"
            elif dpi == 150 and quality == 85:
                preset = "Mittel - Ausgewogen"
            elif dpi == 72 and quality == 65:
                preset = "Hoch - Kleine Dateigröße"
            else:
                preset = "Benutzerdefiniert"
            
            self.preset_var.set(preset)
    
    def _check_ghostscript(self):
        """Prüft ob Ghostscript verfügbar ist"""
        gs_available = False
        
        if os.name == 'nt':
            # Windows
            for cmd in ['gswin64c', 'gswin32c']:
                try:
                    result = subprocess.run([cmd, '--version'], capture_output=True)
                    if result.returncode == 0:
                        gs_available = True
                        break
                except:
                    continue
        else:
            # Linux/Mac
            try:
                result = subprocess.run(['gs', '--version'], capture_output=True)
                gs_available = result.returncode == 0
            except:
                pass
        
        if not gs_available:
            messagebox.showwarning(
                "Ghostscript nicht gefunden",
                "Ghostscript wurde nicht gefunden!\n\n"
                "Die PDF-Komprimierung benötigt Ghostscript.\n"
                "Bitte installieren Sie es von:\n"
                "https://www.ghostscript.com/download/gsdnld.html"
            )
            self.ok_button.config(state=tk.DISABLED)
            self.test_button.config(state=tk.DISABLED)
    
    def _on_preset_changed(self, event=None):
        """Wird aufgerufen wenn ein Preset ausgewählt wird"""
        if self.changing_preset:
            return
        
        preset = self.preset_var.get()
        self.changing_preset = True
        
        if preset == "Niedrig - Beste Qualität":
            self.dpi_var.set(200)
            self.quality_var.set(90)
            self.downsample_var.set(False)
            self.subset_fonts_var.set(True)
            self.remove_duplicates_var.set(True)
            self.optimize_var.set(True)
        
        elif preset == "Mittel - Ausgewogen":
            self.dpi_var.set(150)
            self.quality_var.set(85)
            self.downsample_var.set(True)
            self.subset_fonts_var.set(True)
            self.remove_duplicates_var.set(True)
            self.optimize_var.set(True)
        
        elif preset == "Hoch - Kleine Dateigröße":
            self.dpi_var.set(72)
            self.quality_var.set(65)
            self.downsample_var.set(True)
            self.subset_fonts_var.set(True)
            self.remove_duplicates_var.set(True)
            self.optimize_var.set(True)
        
        self.changing_preset = False
        self._update_labels()
    
    def _on_value_changed(self):
        """Wird aufgerufen wenn ein Wert geändert wird"""
        if not self.changing_preset:
            self.preset_var.set("Benutzerdefiniert")
        self._update_labels()
    
    def _update_labels(self):
        """Aktualisiert die Wertanzeigen"""
        # DPI
        dpi = int(self.dpi_var.get())
        self.dpi_label.config(text=f"{dpi} DPI")
        
        # Qualität mit Farbcodierung
        quality = int(self.quality_var.get())
        self.quality_label.config(text=f"{quality}%")
        
        if quality >= 85:
            color = "darkgreen"
        elif quality >= 70:
            color = "orange"
        else:
            color = "red"
        
        self.quality_label.config(foreground=color)
    
    def _test_compression(self):
        """Testet die Komprimierung mit einer ausgewählten PDF"""
        if self.test_running:
            return
        
        # PDF auswählen
        filename = filedialog.askopenfilename(
            parent=self.dialog,
            title="PDF-Datei auswählen",
            filetypes=[("PDF-Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
        )
        
        if not filename:
            return
        
        # Test vorbereiten
        self.test_running = True
        self.test_button.config(state=tk.DISABLED)
        self.test_result_label.config(text="Teste Komprimierung...", foreground="black")
        self.test_progress.pack(fill=tk.X, pady=(5, 0))
        self.test_progress.start()
        
        # Test in Thread ausführen
        thread = threading.Thread(target=self._run_test, args=(filename,), daemon=True)
        thread.start()
    
    def _run_test(self, filename):
        """Führt den Test aus (in separatem Thread)"""
        try:
            # Altes temp_dir aufräumen wenn vorhanden
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir, ignore_errors=True)
            
            # Temporäre Dateien erstellen
            self.temp_dir = tempfile.mkdtemp()
            temp_input = os.path.join(self.temp_dir, "input.pdf")
            temp_output = os.path.join(self.temp_dir, "compressed.pdf")
            
            # Kopiere Original
            shutil.copy2(filename, temp_input)
            original_size = os.path.getsize(temp_input)
            
            # Ghostscript-Befehl erstellen
            gs_cmd = self._get_gs_command()
            if not gs_cmd:
                raise Exception("Ghostscript nicht gefunden")
            
            # Parameter holen
            dpi = self.dpi_var.get()
            quality = self.quality_var.get()
            
            # Befehl zusammenbauen
            cmd = [
                gs_cmd,
                '-sDEVICE=pdfwrite',
                '-dCompatibilityLevel=1.4',
                '-dNOPAUSE',
                '-dBATCH',
                '-dQUIET',
                f'-dColorImageResolution={dpi}',
                f'-dGrayImageResolution={dpi}',
                f'-dMonoImageResolution={dpi}',
                f'-dJPEGQ={quality/100.0:.2f}',
                '-dAutoFilterColorImages=false',
                '-dColorImageFilter=/DCTEncode',
                '-dAutoFilterGrayImages=false', 
                '-dGrayImageFilter=/DCTEncode',
                '-dMonoImageFilter=/CCITTFaxEncode'
            ]
            
            # Optionale Parameter
            if self.downsample_var.get():
                cmd.extend([
                    '-dDownsampleColorImages=true',
                    '-dDownsampleGrayImages=true',
                    '-dDownsampleMonoImages=true'
                ])
            
            if self.subset_fonts_var.get():
                cmd.append('-dSubsetFonts=true')
            
            if self.remove_duplicates_var.get():
                cmd.append('-dDetectDuplicateImages=true')
            
            if self.optimize_var.get():
                cmd.append('-dOptimize=true')
            
            # Output und Input
            output_arg = '-sOutputFile=' + temp_output  # String-Konkatenation statt f-string
            cmd.append(output_arg)
            cmd.append(temp_input)
            
            # Komprimierung ausführen
            result = subprocess.run(cmd, capture_output=True, text=True)
            
            if result.returncode == 0 and os.path.exists(temp_output):
                compressed_size = os.path.getsize(temp_output)
                reduction = (1 - compressed_size/original_size) * 100
                
                # Speichere komprimierte PDF für Vorschau
                self.compressed_pdf_path = temp_output
                
                # Größenanzeige in KB oder MB je nach Größe
                if original_size < 1024 * 1024:  # Kleiner als 1 MB
                    orig_str = f"{original_size/1024:.0f} KB"
                else:
                    orig_str = f"{original_size/1024/1024:.1f} MB"
                
                if compressed_size < 1024 * 1024:  # Kleiner als 1 MB
                    comp_str = f"{compressed_size/1024:.0f} KB"
                else:
                    comp_str = f"{compressed_size/1024/1024:.1f} MB"
                
                msg = f"Original: {orig_str} → Komprimiert: {comp_str} (−{reduction:.1f}%)"
                
                color = "darkgreen" if reduction > 0 else "orange"
                self.dialog.after(0, self._test_complete, msg, color, True)
            else:
                error = result.stderr if result.stderr else "Unbekannter Fehler"
                self.dialog.after(0, self._test_complete, f"Fehler: {error}", "red", False)
            
        except Exception as e:
            self.dialog.after(0, self._test_complete, f"Fehler: {str(e)}", "red", False)
    
    def _test_complete(self, message, color, success):
        """Wird aufgerufen wenn der Test abgeschlossen ist"""
        self.test_running = False
        self.test_button.config(state=tk.NORMAL)
        self.test_progress.stop()
        self.test_progress.pack_forget()
        self.test_result_label.config(text=message, foreground=color)
        
        # Preview-Button aktivieren wenn erfolgreich
        if success and self.compressed_pdf_path:
            self.preview_button.config(state=tk.NORMAL)
        else:
            self.preview_button.config(state=tk.DISABLED)
    
    def _show_preview(self):
        """Öffnet die komprimierte PDF zur Vorschau"""
        if self.compressed_pdf_path and os.path.exists(self.compressed_pdf_path):
            try:
                if os.name == 'nt':  # Windows
                    os.startfile(self.compressed_pdf_path)
                elif os.name == 'posix':  # macOS and Linux
                    subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', 
                                  self.compressed_pdf_path])
            except Exception as e:
                messagebox.showerror("Fehler", f"Konnte PDF nicht öffnen: {str(e)}")
    
    def _get_gs_command(self):
        """Ermittelt den Ghostscript-Befehl"""
        if os.name == 'nt':
            for cmd in ['gswin64c', 'gswin32c']:
                try:
                    subprocess.run([cmd, '--version'], capture_output=True, check=True)
                    return cmd
                except:
                    continue
        else:
            try:
                subprocess.run(['gs', '--version'], capture_output=True, check=True)
                return 'gs'
            except:
                pass
        return None
    
    def _cleanup(self):
        """Räumt temporäre Dateien auf"""
        if self.temp_dir and os.path.exists(self.temp_dir):
            try:
                shutil.rmtree(self.temp_dir, ignore_errors=True)
            except:
                pass
    
    def _on_ok(self):
        """Speichert die Einstellungen und schließt den Dialog"""
        dpi = self.dpi_var.get()
        
        self.result = {
            'compression_level': self.preset_var.get(),
            'color_dpi': dpi,
            'gray_dpi': dpi,
            'mono_dpi': dpi,
            'jpeg_quality': self.quality_var.get(),
            'color_compression': 'jpeg',
            'gray_compression': 'jpeg',
            'mono_compression': 'ccitt',
            'downsample_images': self.downsample_var.get(),
            'subset_fonts': self.subset_fonts_var.get(),
            'remove_duplicates': self.remove_duplicates_var.get(),
            'optimize': self.optimize_var.get()
        }
        
        self._cleanup()
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Schließt den Dialog ohne zu speichern"""
        self._cleanup()
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result