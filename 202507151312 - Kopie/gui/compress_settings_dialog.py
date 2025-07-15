"""
Dialog für PDF-Komprimierungseinstellungen - Professionelle Version
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
import json
from datetime import datetime
import fitz  # PyMuPDF für PDF-Analyse

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

logger = logging.getLogger(__name__)


class CompressSettingsDialog:
    """Professioneller Dialog für PDF-Komprimierung"""
    
    # Vordefinierte Profile für Geschäftsdokumente
    COMPRESSION_PROFILES = {
        "Rechnung/Geschäftsdokument": {
            "description": "Optimiert für Lesbarkeit von Text und Zahlen, erhält Stempel und Unterschriften",
            "color_dpi": 300,
            "gray_dpi": 300,
            "mono_dpi": 600,
            "jpeg_quality": 85,
            "downsample_images": True,
            "subset_fonts": True,
            "remove_duplicates": True,
            "optimize": True
        },
        "Langzeitarchiv": {
            "description": "Ausgewogene Komprimierung für dauerhafte Archivierung",
            "color_dpi": 200,
            "gray_dpi": 200,
            "mono_dpi": 400,
            "jpeg_quality": 80,
            "downsample_images": True,
            "subset_fonts": True,
            "remove_duplicates": True,
            "optimize": True
        },
        "Gescanntes Dokument": {
            "description": "Stärkere Komprimierung für bereits gescannte Dokumente",
            "color_dpi": 150,
            "gray_dpi": 150,
            "mono_dpi": 300,
            "jpeg_quality": 75,
            "downsample_images": True,
            "subset_fonts": True,
            "remove_duplicates": True,
            "optimize": True
        },
        "E-Mail-Versand": {
            "description": "Maximale Komprimierung für kleine Dateigrößen",
            "color_dpi": 100,
            "gray_dpi": 100,
            "mono_dpi": 200,
            "jpeg_quality": 65,
            "downsample_images": True,
            "subset_fonts": True,
            "remove_duplicates": True,
            "optimize": True
        },
        "Benutzerdefiniert": {
            "description": "Eigene Einstellungen verwenden",
            "color_dpi": 150,
            "gray_dpi": 150,
            "mono_dpi": 300,
            "jpeg_quality": 80,
            "downsample_images": True,
            "subset_fonts": True,
            "remove_duplicates": True,
            "optimize": True
        }
    }
    
    def __init__(self, parent, initial_params: Optional[Dict] = None):
        self.parent = parent
        self.result = None
        self.test_running = False
        self.compressed_pdf_path = None
        self.temp_dir = None
        self.test_pdf_info = None
        
        # Lade gespeicherte Einstellungen
        self.saved_profiles = self._load_saved_profiles()
        
        # Standard-Parameter aus Profil
        default_profile = "Rechnung/Geschäftsdokument"
        self.default_params = self.COMPRESSION_PROFILES[default_profile].copy()
        self.default_params['compression_profile'] = default_profile
        
        # Übernehme initiale Parameter
        self.params = self.default_params.copy()
        if initial_params:
            self.params.update(initial_params)
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("PDF-Komprimierungseinstellungen")
        self.dialog.geometry("700x680")
        self.dialog.resizable(False, False)
        
        # Dialog zentrieren
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() - 700) // 2
        y = (self.dialog.winfo_screenheight() - 680) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Flags
        self.changing_preset = False
        
        # GUI erstellen
        self._create_gui()
        
        # Events
        self.dialog.bind('<Return>', lambda e: self._on_ok())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)
        
        # Prüfe Abhängigkeiten
        self._check_dependencies()
    
    def _create_gui(self):
        """Erstellt die GUI mit professionellem Layout"""
        # Hauptframe mit Scrollbar
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Profil-Auswahl
        profile_frame = ttk.LabelFrame(main_frame, text="Komprimierungsprofil", padding="10")
        profile_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(profile_frame, text="Dokumenttyp:").pack(anchor=tk.W, pady=(0, 5))
        
        self.profile_var = tk.StringVar()
        self.profile_combo = ttk.Combobox(
            profile_frame,
            textvariable=self.profile_var,
            values=list(self.COMPRESSION_PROFILES.keys()),
            state="readonly",
            width=40
        )
        self.profile_combo.pack(fill=tk.X)
        self.profile_combo.bind('<<ComboboxSelected>>', self._on_profile_changed)
        
        # Profilbeschreibung
        self.profile_desc_label = ttk.Label(
            profile_frame, 
            text="",
            wraplength=650,
            font=('TkDefaultFont', 9, 'italic')
        )
        self.profile_desc_label.pack(anchor=tk.W, pady=(5, 0))
        
        # Detaillierte Einstellungen
        settings_frame = ttk.LabelFrame(main_frame, text="Detaileinstellungen", padding="10")
        settings_frame.pack(fill=tk.X, pady=(0, 15))
        
        # Notebook für Kategorien
        self.notebook = ttk.Notebook(settings_frame)
        self.notebook.pack(fill=tk.BOTH, expand=True)
        
        # Tab 1: Bildeinstellungen
        image_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(image_tab, text="Bilder")
        
        # Farbbilder
        color_frame = ttk.Frame(image_tab)
        color_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(color_frame, text="Farbbilder (DPI):").pack(anchor=tk.W)
        
        color_control = ttk.Frame(color_frame)
        color_control.pack(fill=tk.X)
        
        self.color_dpi_var = tk.IntVar(value=self.params.get('color_dpi', 300))
        self.color_dpi_scale = ttk.Scale(
            color_control,
            from_=72, to=600,
            variable=self.color_dpi_var,
            orient=tk.HORIZONTAL,
            length=400,
            command=lambda v: self._on_value_changed()
        )
        self.color_dpi_scale.pack(side=tk.LEFT, padx=(0, 10))
        
        self.color_dpi_label = ttk.Label(color_control, text="300 DPI", width=10)
        self.color_dpi_label.pack(side=tk.LEFT)
        
        # Graustufenbilder
        gray_frame = ttk.Frame(image_tab)
        gray_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(gray_frame, text="Graustufenbilder (DPI):").pack(anchor=tk.W)
        
        gray_control = ttk.Frame(gray_frame)
        gray_control.pack(fill=tk.X)
        
        self.gray_dpi_var = tk.IntVar(value=self.params.get('gray_dpi', 300))
        self.gray_dpi_scale = ttk.Scale(
            gray_control,
            from_=72, to=600,
            variable=self.gray_dpi_var,
            orient=tk.HORIZONTAL,
            length=400,
            command=lambda v: self._on_value_changed()
        )
        self.gray_dpi_scale.pack(side=tk.LEFT, padx=(0, 10))
        
        self.gray_dpi_label = ttk.Label(gray_control, text="300 DPI", width=10)
        self.gray_dpi_label.pack(side=tk.LEFT)
        
        # Schwarz-Weiß-Bilder
        mono_frame = ttk.Frame(image_tab)
        mono_frame.pack(fill=tk.X, pady=(0, 15))
        
        ttk.Label(mono_frame, text="Schwarz-Weiß-Bilder (DPI):").pack(anchor=tk.W)
        
        mono_control = ttk.Frame(mono_frame)
        mono_control.pack(fill=tk.X)
        
        self.mono_dpi_var = tk.IntVar(value=self.params.get('mono_dpi', 600))
        self.mono_dpi_scale = ttk.Scale(
            mono_control,
            from_=150, to=1200,
            variable=self.mono_dpi_var,
            orient=tk.HORIZONTAL,
            length=400,
            command=lambda v: self._on_value_changed()
        )
        self.mono_dpi_scale.pack(side=tk.LEFT, padx=(0, 10))
        
        self.mono_dpi_label = ttk.Label(mono_control, text="600 DPI", width=10)
        self.mono_dpi_label.pack(side=tk.LEFT)
        
        # JPEG-Qualität
        quality_frame = ttk.Frame(image_tab)
        quality_frame.pack(fill=tk.X)
        
        ttk.Label(quality_frame, text="JPEG-Qualität:").pack(anchor=tk.W)
        
        quality_control = ttk.Frame(quality_frame)
        quality_control.pack(fill=tk.X)
        
        self.quality_var = tk.IntVar(value=self.params.get('jpeg_quality', 85))
        self.quality_scale = ttk.Scale(
            quality_control,
            from_=10, to=100,
            variable=self.quality_var,
            orient=tk.HORIZONTAL,
            length=400,
            command=lambda v: self._on_value_changed()
        )
        self.quality_scale.pack(side=tk.LEFT, padx=(0, 10))
        
        self.quality_label = ttk.Label(quality_control, text="85%", width=10)
        self.quality_label.pack(side=tk.LEFT)
        
        # Tab 2: Optimierungen
        optimize_tab = ttk.Frame(self.notebook, padding="10")
        self.notebook.add(optimize_tab, text="Optimierungen")
        
        # Checkboxen
        self.downsample_var = tk.BooleanVar(value=self.params.get('downsample_images', True))
        self.downsample_check = ttk.Checkbutton(
            optimize_tab,
            text="Bilder intelligent herunterskalieren (nur wenn DPI höher als Zielwert)",
            variable=self.downsample_var,
            command=self._on_value_changed
        )
        self.downsample_check.pack(anchor=tk.W, pady=5)
        
        self.subset_fonts_var = tk.BooleanVar(value=self.params.get('subset_fonts', True))
        self.subset_fonts_check = ttk.Checkbutton(
            optimize_tab,
            text="Schriften optimieren (nur verwendete Zeichen einbetten)",
            variable=self.subset_fonts_var,
            command=self._on_value_changed
        )
        self.subset_fonts_check.pack(anchor=tk.W, pady=5)
        
        self.remove_duplicates_var = tk.BooleanVar(value=self.params.get('remove_duplicates', True))
        self.remove_duplicates_check = ttk.Checkbutton(
            optimize_tab,
            text="Duplizierte Bilder erkennen und entfernen",
            variable=self.remove_duplicates_var,
            command=self._on_value_changed
        )
        self.remove_duplicates_check.pack(anchor=tk.W, pady=5)
        
        self.optimize_var = tk.BooleanVar(value=self.params.get('optimize', True))
        self.optimize_check = ttk.Checkbutton(
            optimize_tab,
            text="PDF-Struktur optimieren (Linearisierung)",
            variable=self.optimize_var,
            command=self._on_value_changed
        )
        self.optimize_check.pack(anchor=tk.W, pady=5)
        
        # Test-Bereich
        test_frame = ttk.LabelFrame(main_frame, text="Qualitätstest", padding="10")
        test_frame.pack(fill=tk.X, pady=(0, 15))
        
        test_buttons = ttk.Frame(test_frame)
        test_buttons.pack(fill=tk.X)
        
        self.test_button = ttk.Button(
            test_buttons,
            text="📄 PDF zum Testen auswählen...",
            command=self._test_compression
        )
        self.test_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.preview_button = ttk.Button(
            test_buttons,
            text="👁️ Komprimierte PDF anzeigen",
            command=self._show_preview,
            state=tk.DISABLED
        )
        self.preview_button.pack(side=tk.LEFT)
        
        # Test-Progress
        self.test_progress = ttk.Progressbar(
            test_frame,
            mode='indeterminate',
            length=650
        )
        
        # Test-Ergebnis
        self.test_result_frame = ttk.Frame(test_frame)
        self.test_result_frame.pack(fill=tk.X, pady=(5, 0))
        
        # Profil speichern
        save_frame = ttk.Frame(main_frame)
        save_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.save_profile_button = ttk.Button(
            save_frame,
            text="💾 Profil speichern...",
            command=self._save_profile,
            state=tk.DISABLED
        )
        self.save_profile_button.pack(side=tk.LEFT)
        
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
        self._update_all()
    
    def _check_dependencies(self):
        """Prüft benötigte Abhängigkeiten"""
        errors = []
        
        # Prüfe Ghostscript
        if not self._is_ghostscript_available():
            errors.append("Ghostscript nicht gefunden")
            self.ok_button.config(state=tk.DISABLED)
            self.test_button.config(state=tk.DISABLED)
        
        if errors:
            messagebox.showwarning(
                "Fehlende Abhängigkeiten",
                "Folgende Programme werden benötigt:\n\n" + 
                "\n".join(f"• {err}" for err in errors) +
                "\n\nBitte installieren Sie die fehlenden Programme:\n" +
                "Ghostscript: https://www.ghostscript.com/download/gsdnld.html"
            )
    
    def _is_ghostscript_available(self):
        """Prüft ob Ghostscript verfügbar ist"""
        if os.name == 'nt':
            for cmd in ['gswin64c', 'gswin32c']:
                try:
                    result = subprocess.run([cmd, '--version'], capture_output=True)
                    if result.returncode == 0:
                        return True
                except:
                    continue
        else:
            try:
                result = subprocess.run(['gs', '--version'], capture_output=True)
                return result.returncode == 0
            except:
                pass
        return False
    
    def _set_initial_values(self):
        """Setzt initiale Werte"""
        # Bestimme Profil
        profile = self.params.get('compression_profile', 'Rechnung/Geschäftsdokument')
        
        if profile in self.COMPRESSION_PROFILES:
            self.profile_var.set(profile)
        else:
            # Versuche Profil anhand der Werte zu erkennen
            for name, settings in self.COMPRESSION_PROFILES.items():
                if (settings['color_dpi'] == self.params.get('color_dpi', 0) and
                    settings['jpeg_quality'] == self.params.get('jpeg_quality', 0)):
                    self.profile_var.set(name)
                    break
            else:
                self.profile_var.set('Benutzerdefiniert')
    
    def _on_profile_changed(self, event=None):
        """Wird bei Profiländerung aufgerufen"""
        if self.changing_preset:
            return
        
        profile_name = self.profile_var.get()
        if profile_name in self.COMPRESSION_PROFILES:
            self.changing_preset = True
            
            profile = self.COMPRESSION_PROFILES[profile_name]
            
            # Setze Werte
            self.color_dpi_var.set(profile['color_dpi'])
            self.gray_dpi_var.set(profile['gray_dpi'])
            self.mono_dpi_var.set(profile['mono_dpi'])
            self.quality_var.set(profile['jpeg_quality'])
            self.downsample_var.set(profile['downsample_images'])
            self.subset_fonts_var.set(profile['subset_fonts'])
            self.remove_duplicates_var.set(profile['remove_duplicates'])
            self.optimize_var.set(profile['optimize'])
            
            # Beschreibung aktualisieren
            self.profile_desc_label.config(text=profile['description'])
            
            self.changing_preset = False
            self._update_all()
            
            # Save-Button aktivieren bei Benutzerdefiniert
            self.save_profile_button.config(
                state=tk.NORMAL if profile_name == 'Benutzerdefiniert' else tk.DISABLED
            )
    
    def _on_value_changed(self):
        """Bei Wertänderung"""
        if not self.changing_preset:
            self.profile_var.set('Benutzerdefiniert')
            self.profile_desc_label.config(text=self.COMPRESSION_PROFILES['Benutzerdefiniert']['description'])
            self.save_profile_button.config(state=tk.NORMAL)
        self._update_all()
    
    def _update_all(self):
        """Aktualisiert alle Anzeigen"""
        # DPI-Labels
        self.color_dpi_label.config(text=f"{int(self.color_dpi_var.get())} DPI")
        self.gray_dpi_label.config(text=f"{int(self.gray_dpi_var.get())} DPI")
        self.mono_dpi_label.config(text=f"{int(self.mono_dpi_var.get())} DPI")
        
        # Qualität mit Farbcodierung
        quality = int(self.quality_var.get())
        self.quality_label.config(text=f"{quality}%")
        
        if quality >= 85:
            color = "#2d862d"  # Dunkelgrün
        elif quality >= 70:
            color = "#ff8c00"  # Orange
        else:
            color = "#dc143c"  # Rot
        
        self.quality_label.config(foreground=color)
    
    def _test_compression(self):
        """Testet Komprimierung mit Qualitätsanalyse"""
        if self.test_running:
            return
        
        filename = filedialog.askopenfilename(
            parent=self.dialog,
            title="PDF-Datei zum Testen auswählen",
            filetypes=[("PDF-Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
        )
        
        if not filename:
            return
        
        # Test vorbereiten
        self.test_running = True
        self.test_button.config(state=tk.DISABLED)
        
        # Alte Ergebnisse löschen
        for widget in self.test_result_frame.winfo_children():
            widget.destroy()
        
        self.test_progress.pack(fill=tk.X, pady=(5, 0))
        self.test_progress.start()
        
        # Test in Thread
        thread = threading.Thread(target=self._run_test, args=(filename,), daemon=True)
        thread.start()
    
    def _run_test(self, filename):
        """Führt Komprimierungstest aus"""
        try:
            # Aufräumen
            if self.temp_dir and os.path.exists(self.temp_dir):
                shutil.rmtree(self.temp_dir, ignore_errors=True)
            
            # Temporäre Dateien
            self.temp_dir = tempfile.mkdtemp()
            temp_input = os.path.join(self.temp_dir, "input.pdf")
            temp_output = os.path.join(self.temp_dir, "compressed.pdf")
            
            # Kopiere Original
            shutil.copy2(filename, temp_input)
            
            # Analysiere Original
            original_info = self._analyze_pdf(temp_input)
            
            # Ghostscript-Befehl
            gs_cmd = self._get_gs_command()
            if not gs_cmd:
                raise Exception("Ghostscript nicht gefunden")
            
            # Baue Befehl
            cmd = self._build_gs_command(gs_cmd, temp_input, temp_output, original_info)
            
            # Komprimierung
            start_time = datetime.now()
            result = subprocess.run(cmd, capture_output=True, text=True)
            compression_time = (datetime.now() - start_time).total_seconds()
            
            if result.returncode == 0 and os.path.exists(temp_output):
                # Analysiere komprimierte PDF
                compressed_info = self._analyze_pdf(temp_output)
                
                # Speichere für Vorschau
                self.compressed_pdf_path = temp_output
                
                # Erstelle Ergebnis
                test_result = {
                    'success': True,
                    'original': original_info,
                    'compressed': compressed_info,
                    'time': compression_time
                }
                
                self.dialog.after(0, self._test_complete, test_result)
            else:
                error = result.stderr if result.stderr else "Unbekannter Fehler"
                self.dialog.after(0, self._test_complete, {'success': False, 'error': error})
            
        except Exception as e:
            self.dialog.after(0, self._test_complete, {'success': False, 'error': str(e)})
    
    def _analyze_pdf(self, pdf_path):
        """Analysiert PDF für Qualitätskontrolle"""
        try:
            doc = fitz.open(pdf_path)
            
            info = {
                'size': os.path.getsize(pdf_path),
                'pages': doc.page_count,
                'images': 0,
                'avg_dpi': 0,
                'has_text': False
            }
            
            total_dpi = 0
            dpi_count = 0
            
            # Analysiere erste Seiten
            for page_num in range(min(3, doc.page_count)):
                page = doc[page_num]
                
                # Text prüfen
                if page.get_text().strip():
                    info['has_text'] = True
                
                # Bilder analysieren
                image_list = page.get_images()
                info['images'] += len(image_list)
                
                for img in image_list:
                    xref = img[0]
                    pix = fitz.Pixmap(doc, xref)
                    if pix.width > 0 and pix.height > 0:
                        bbox = page.get_image_bbox(img)
                        if bbox:
                            width_inch = (bbox.x1 - bbox.x0) / 72
                            height_inch = (bbox.y1 - bbox.y0) / 72
                            if width_inch > 0 and height_inch > 0:
                                dpi = (pix.width / width_inch + pix.height / height_inch) / 2
                                total_dpi += dpi
                                dpi_count += 1
                    pix = None
            
            if dpi_count > 0:
                info['avg_dpi'] = int(total_dpi / dpi_count)
            
            doc.close()
            return info
            
        except Exception as e:
            logger.error(f"PDF-Analyse fehlgeschlagen: {e}")
            return {
                'size': os.path.getsize(pdf_path),
                'pages': 0,
                'images': 0,
                'avg_dpi': 0,
                'has_text': False
            }
    
    def _build_gs_command(self, gs_cmd, input_path, output_path, pdf_info):
        """Baut intelligenten Ghostscript-Befehl"""
        cmd = [
            gs_cmd,
            '-sDEVICE=pdfwrite',
            '-dCompatibilityLevel=1.7',
            '-dNOPAUSE',
            '-dBATCH',
            '-dQUIET',
            '-dSAFER',
            f'-sOutputFile={output_path}'
        ]
        
        # DPI-Einstellungen
        cmd.extend([
            f'-dColorImageResolution={self.color_dpi_var.get()}',
            f'-dGrayImageResolution={self.gray_dpi_var.get()}',
            f'-dMonoImageResolution={self.mono_dpi_var.get()}'
        ])
        
        # Intelligentes Downsampling
        if self.downsample_var.get():
            # Nur wenn Original-DPI höher
            if pdf_info.get('avg_dpi', 0) > self.color_dpi_var.get():
                cmd.extend([
                    '-dDownsampleColorImages=true',
                    '-dDownsampleGrayImages=true',
                    '-dDownsampleMonoImages=true',
                    '-dColorImageDownsampleType=/Bicubic',
                    '-dGrayImageDownsampleType=/Bicubic',
                    '-dMonoImageDownsampleType=/Bicubic',
                    '-dColorImageDownsampleThreshold=1.0',
                    '-dGrayImageDownsampleThreshold=1.0',
                    '-dMonoImageDownsampleThreshold=1.0'
                ])
            else:
                cmd.extend([
                    '-dDownsampleColorImages=false',
                    '-dDownsampleGrayImages=false',
                    '-dDownsampleMonoImages=false'
                ])
        
        # Komprimierung
        quality = self.quality_var.get() / 100.0
        cmd.extend([
            '-dAutoFilterColorImages=true',
            '-dAutoFilterGrayImages=true',
            f'-dJPEGQ={quality:.2f}',
            '-dColorImageFilter=/DCTEncode',
            '-dGrayImageFilter=/DCTEncode',
            '-dMonoImageFilter=/CCITTFaxEncode'
        ])
        
        # Weitere Optimierungen
        if self.subset_fonts_var.get():
            cmd.extend(['-dSubsetFonts=true', '-dEmbedAllFonts=true'])
        
        if self.remove_duplicates_var.get():
            cmd.append('-dDetectDuplicateImages=true')
        
        if self.optimize_var.get():
            cmd.extend(['-dOptimize=true', '-dCompressPages=true'])
        
        cmd.append(input_path)
        
        return cmd
    
    def _test_complete(self, result):
        """Zeigt Testergebnis an"""
        self.test_running = False
        self.test_button.config(state=tk.NORMAL)
        self.test_progress.stop()
        self.test_progress.pack_forget()
        
        # Leere Frame
        for widget in self.test_result_frame.winfo_children():
            widget.destroy()
        
        if result['success']:
            # Erfolgreiche Komprimierung
            orig = result['original']
            comp = result['compressed']
            
            # Berechne Statistiken
            size_reduction = (1 - comp['size']/orig['size']) * 100
            size_mb_orig = orig['size'] / (1024*1024)
            size_mb_comp = comp['size'] / (1024*1024)
            
            # Erstelle Ergebnis-UI
            stats_frame = ttk.Frame(self.test_result_frame)
            stats_frame.pack(fill=tk.X)
            
            # Größenvergleich
            size_text = f"Dateigröße: {size_mb_orig:.2f} MB → {size_mb_comp:.2f} MB ({size_reduction:.1f}% kleiner)"
            size_label = ttk.Label(stats_frame, text=size_text)
            size_label.pack(anchor=tk.W)
            
            # Farbcodierung basierend auf Reduktion
            if size_reduction > 50:
                color = "#2d862d"  # Grün
            elif size_reduction > 20:
                color = "#ff8c00"  # Orange
            else:
                color = "#dc143c"  # Rot
            
            size_label.config(foreground=color)
            
            # Weitere Statistiken
            if orig['avg_dpi'] > 0:
                dpi_text = f"Durchschnittliche Bild-DPI: {orig['avg_dpi']} → {comp['avg_dpi']}"
                ttk.Label(stats_frame, text=dpi_text).pack(anchor=tk.W)
            
            time_text = f"Komprimierungszeit: {result['time']:.1f} Sekunden"
            ttk.Label(stats_frame, text=time_text).pack(anchor=tk.W)
            
            # Qualitätswarnung wenn nötig
            if size_reduction > 70:
                warning_frame = ttk.Frame(self.test_result_frame)
                warning_frame.pack(fill=tk.X, pady=(10, 0))
                
                warning_label = ttk.Label(
                    warning_frame,
                    text="⚠️ Hohe Komprimierung - bitte Qualität prüfen!",
                    foreground="#ff8c00"
                )
                warning_label.pack(anchor=tk.W)
            
            # Preview-Button aktivieren
            self.preview_button.config(state=tk.NORMAL)
            self.test_pdf_info = result
            
        else:
            # Fehler
            error_label = ttk.Label(
                self.test_result_frame,
                text=f"❌ Fehler: {result['error']}",
                foreground="#dc143c"
            )
            error_label.pack(anchor=tk.W)
            self.preview_button.config(state=tk.DISABLED)
    
    def _show_preview(self):
        """Öffnet komprimierte PDF"""
        if self.compressed_pdf_path and os.path.exists(self.compressed_pdf_path):
            try:
                if os.name == 'nt':
                    os.startfile(self.compressed_pdf_path)
                elif os.name == 'posix':
                    subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', 
                                  self.compressed_pdf_path])
            except Exception as e:
                messagebox.showerror("Fehler", f"Konnte PDF nicht öffnen: {str(e)}")
    
    def _save_profile(self):
        """Speichert benutzerdefiniertes Profil"""
        name = tk.simpledialog.askstring(
            "Profil speichern",
            "Name für das Profil:",
            parent=self.dialog
        )
        
        if name:
            profile = {
                'color_dpi': self.color_dpi_var.get(),
                'gray_dpi': self.gray_dpi_var.get(),
                'mono_dpi': self.mono_dpi_var.get(),
                'jpeg_quality': self.quality_var.get(),
                'downsample_images': self.downsample_var.get(),
                'subset_fonts': self.subset_fonts_var.get(),
                'remove_duplicates': self.remove_duplicates_var.get(),
                'optimize': self.optimize_var.get()
            }
            
            self.saved_profiles[name] = profile
            self._save_profiles_to_file()
            
            messagebox.showinfo("Profil gespeichert", f"Profil '{name}' wurde gespeichert.")
    
    def _load_saved_profiles(self):
        """Lädt gespeicherte Profile"""
        try:
            profile_file = os.path.join(os.path.dirname(__file__), '..', 'compression_profiles.json')
            if os.path.exists(profile_file):
                with open(profile_file, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except:
            pass
        return {}
    
    def _save_profiles_to_file(self):
        """Speichert Profile in Datei"""
        try:
            profile_file = os.path.join(os.path.dirname(__file__), '..', 'compression_profiles.json')
            with open(profile_file, 'w', encoding='utf-8') as f:
                json.dump(self.saved_profiles, f, indent=2)
        except Exception as e:
            logger.error(f"Fehler beim Speichern der Profile: {e}")
    
    def _get_gs_command(self):
        """Ermittelt Ghostscript-Befehl"""
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
        """Speichert Einstellungen"""
        self.result = {
            'compression_profile': self.profile_var.get(),
            'color_dpi': self.color_dpi_var.get(),
            'gray_dpi': self.gray_dpi_var.get(),
            'mono_dpi': self.mono_dpi_var.get(),
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
        """Schließt Dialog"""
        self._cleanup()
        self.dialog.destroy()
    
    def show(self) -> Optional[Dict]:
        """Zeigt Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result