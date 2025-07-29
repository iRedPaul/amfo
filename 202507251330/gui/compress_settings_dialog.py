"""
Dialog für PDF-Komprimierungseinstellungen - pypdf Version mit Reglern
"""
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from typing import Dict, Optional, Any
import sys
import os
import threading
import shutil
import tempfile
import logging
from datetime import datetime
from pypdf import PdfWriter

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

logger = logging.getLogger(__name__)


class CompressSettingsDialog:
    """Dialog für PDF-Komprimierung mit pypdf und individuellen Reglern"""
    
    def __init__(self, parent, initial_params: Optional[Dict] = None):
        self.parent = parent
        self.result = None
        self.test_running = False
        self.compressed_pdf_path = None
        self.temp_dir = None
        
        # Standard-Parameter
        self.params = {
            'compression_level': 6,
            'image_quality': 70
        }
        
        # Übernehme initiale Parameter
        if initial_params:
            self.params.update(initial_params)
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("PDF-Komprimierung")
        self.dialog.geometry("650x650")  # Etwas breiter für bessere Textanzeige
        self.dialog.resizable(False, False)
        
        # Dialog zentrieren
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() - 650) // 2
        y = (self.dialog.winfo_screenheight() - 650) // 2
        self.dialog.geometry(f"+{x}+{y}")
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # GUI erstellen
        self._create_gui()
        
        # Events
        self.dialog.bind('<Return>', lambda e: self._on_ok())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())
        self.dialog.protocol("WM_DELETE_WINDOW", self._on_cancel)
    
    def _create_gui(self):
        """Erstellt die GUI mit Reglern"""
        # Hauptframe
        main_frame = ttk.Frame(self.dialog, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Titel
        title_label = ttk.Label(
            main_frame, 
            text="PDF-Komprimierungseinstellungen",
            font=('TkDefaultFont', 12, 'bold')
        )
        title_label.pack(pady=(0, 10))
        
        # Info-Label
        info_label = ttk.Label(
            main_frame,
            text="Komprimiert Inhaltsströme und reduziert Bildqualität\nKeine externen Programme erforderlich!",
            font=('TkDefaultFont', 9, 'italic'),
            foreground='gray',
            justify=tk.CENTER
        )
        info_label.pack(pady=(0, 20))
        
        # Einstellungen Frame
        settings_frame = ttk.LabelFrame(main_frame, text="Komprimierungseinstellungen", padding="15")
        settings_frame.pack(fill=tk.X, pady=(0, 20))
        
        # Komprimierungslevel Regler
        compression_frame = ttk.Frame(settings_frame)
        compression_frame.pack(fill=tk.X, pady=(0, 20))
        
        compression_label = ttk.Label(
            compression_frame, 
            text="Komprimierungsstärke (Inhaltsströme):",
            font=('TkDefaultFont', 10)
        )
        compression_label.pack(anchor=tk.W, pady=(0, 5))
        
        compression_desc = ttk.Label(
            compression_frame,
            text="0 = Keine Komprimierung, 9 = Maximale Komprimierung",
            font=('TkDefaultFont', 8, 'italic'),
            foreground='gray'
        )
        compression_desc.pack(anchor=tk.W)
        
        # Regler Frame
        compression_slider_frame = ttk.Frame(compression_frame)
        compression_slider_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.compression_var = tk.IntVar(value=self.params['compression_level'])
        self.compression_slider = ttk.Scale(
            compression_slider_frame,
            from_=0,
            to=9,
            variable=self.compression_var,
            orient=tk.HORIZONTAL,
            command=self._update_compression_label
        )
        self.compression_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        self.compression_value_label = ttk.Label(
            compression_slider_frame,
            text=f"Level: {self.compression_var.get()}",
            width=10
        )
        self.compression_value_label.pack(side=tk.LEFT)
        
        # Bildqualität Regler
        image_frame = ttk.Frame(settings_frame)
        image_frame.pack(fill=tk.X)
        
        image_label = ttk.Label(
            image_frame,
            text="Bildqualität (JPEG):",
            font=('TkDefaultFont', 10)
        )
        image_label.pack(anchor=tk.W, pady=(0, 5))
        
        image_desc = ttk.Label(
            image_frame,
            text="10% = Niedrige Qualität/Kleine Größe, 100% = Beste Qualität/Große Größe",
            font=('TkDefaultFont', 8, 'italic'),
            foreground='gray'
        )
        image_desc.pack(anchor=tk.W)
        
        # Regler Frame
        image_slider_frame = ttk.Frame(image_frame)
        image_slider_frame.pack(fill=tk.X, pady=(5, 0))
        
        self.image_quality_var = tk.IntVar(value=self.params['image_quality'])
        self.image_quality_slider = ttk.Scale(
            image_slider_frame,
            from_=10,
            to=100,
            variable=self.image_quality_var,
            orient=tk.HORIZONTAL,
            command=self._update_image_label
        )
        self.image_quality_slider.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        self.image_quality_label = ttk.Label(
            image_slider_frame,
            text=f"{self.image_quality_var.get()}%",
            width=10
        )
        self.image_quality_label.pack(side=tk.LEFT)
        
        # Vorschläge Frame
        suggestions_frame = ttk.Frame(settings_frame)
        suggestions_frame.pack(fill=tk.X, pady=(20, 0))
        
        suggestions_label = ttk.Label(
            suggestions_frame,
            text="Empfohlene Einstellungen:",
            font=('TkDefaultFont', 9, 'bold')
        )
        suggestions_label.pack(anchor=tk.W, pady=(0, 5))
        
        # Vorschlag-Buttons
        button_frame = ttk.Frame(suggestions_frame)
        button_frame.pack(fill=tk.X)
        
        ttk.Button(
            button_frame,
            text="Hohe Qualität",
            command=lambda: self._set_preset(3, 85)
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            button_frame,
            text="Ausgewogen",
            command=lambda: self._set_preset(6, 70)
        ).pack(side=tk.LEFT, padx=(0, 5))
        
        ttk.Button(
            button_frame,
            text="Kleine Größe",
            command=lambda: self._set_preset(9, 50)
        ).pack(side=tk.LEFT)
        
        # Test-Bereich
        test_frame = ttk.LabelFrame(main_frame, text="Komprimierung testen", padding="15")
        test_frame.pack(fill=tk.X, pady=(0, 20))
        
        test_button_frame = ttk.Frame(test_frame)
        test_button_frame.pack(fill=tk.X)
        
        self.test_button = ttk.Button(
            test_button_frame,
            text="PDF zum Testen auswählen...",
            command=self._test_compression
        )
        self.test_button.pack(side=tk.LEFT, padx=(0, 10))
        
        self.preview_button = ttk.Button(
            test_button_frame,
            text="Ergebnis anzeigen",
            command=self._show_preview,
            state=tk.DISABLED
        )
        self.preview_button.pack(side=tk.LEFT)
        
        # Test-Progress
        self.test_progress = ttk.Progressbar(
            test_frame,
            mode='indeterminate',
            length=600  # Angepasst an breiteres Fenster
        )
        
        # Test-Ergebnis - als Frame für kompaktere Darstellung
        self.test_result_frame = ttk.Frame(test_frame)
        
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
        
        # Initiale Label-Updates
        self._update_compression_label(None)
        self._update_image_label(None)
    
    def _update_compression_label(self, value):
        """Aktualisiert das Komprimierungslevel-Label"""
        level = int(self.compression_var.get())
        self.compression_value_label.config(text=f"Level: {level}")
    
    def _update_image_label(self, value):
        """Aktualisiert das Bildqualität-Label"""
        quality = int(self.image_quality_var.get())
        self.image_quality_label.config(text=f"{quality}%")
        
        # Farbcodierung
        if quality >= 80:
            color = "green"
        elif quality >= 60:
            color = "orange"
        else:
            color = "red"
        
        self.image_quality_label.config(foreground=color)
    
    def _set_preset(self, compression_level, image_quality):
        """Setzt vordefinierte Werte"""
        self.compression_var.set(compression_level)
        self.image_quality_var.set(image_quality)
        self._update_compression_label(None)
        self._update_image_label(None)
    
    def _test_compression(self):
        """Testet Komprimierung"""
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
        
        self.preview_button.config(state=tk.DISABLED)
        
        # Progress anzeigen
        self.test_progress.pack(fill=tk.X, pady=(10, 0))
        self.test_progress.start()
        
        # Test in Thread
        thread = threading.Thread(target=self._run_test, args=(filename,), daemon=True)
        thread.start()
    
    def _run_test(self, filename):
        """Führt Komprimierungstest aus mit pypdf"""
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
            
            # Original-Größe
            original_size = os.path.getsize(temp_input)
            
            # Aktuelle Einstellungen
            compression_level = int(self.compression_var.get())
            image_quality = int(self.image_quality_var.get())
            
            # Komprimierung mit pypdf
            start_time = datetime.now()
            
            try:
                # Öffne PDF mit pypdf
                writer = PdfWriter(clone_from=temp_input)
                
                # Zähle Bilder
                total_images = 0
                compressed_images = 0
                
                # Komprimiere alle Seiten
                page_count = len(writer.pages)
                for i, page in enumerate(writer.pages):
                    # Komprimiere Inhaltsströme nur wenn Level > 0
                    if compression_level > 0:
                        page.compress_content_streams(level=compression_level)
                    
                    # Komprimiere Bilder
                    for img in page.images:
                        total_images += 1
                        try:
                            img.replace(img.image, quality=image_quality)
                            compressed_images += 1
                        except Exception as img_error:
                            logger.debug(f"Bild übersprungen: {img_error}")
                    
                    # Konsolidiere Objekte
                    page.scale_by(1.0)
                
                # Weitere Optimierungen
                try:
                    writer.compress_identical_objects()
                    writer.remove_duplication()
                except:
                    pass  # Ignoriere Fehler bei Optimierungen
                
                # Schreibe komprimierte PDF
                with open(temp_output, 'wb') as output_file:
                    writer.write(output_file)
                
                compression_time = (datetime.now() - start_time).total_seconds()
                
                # Erfolg
                compressed_size = os.path.getsize(temp_output)
                reduction_percent = (1 - compressed_size/original_size) * 100
                
                self.compressed_pdf_path = temp_output
                
                # Ergebnis kompakt darstellen
                result_data = {
                    'success': True,
                    'original_size': original_size,
                    'compressed_size': compressed_size,
                    'reduction_percent': reduction_percent,
                    'compression_time': compression_time,
                    'page_count': page_count,
                    'total_images': total_images,
                    'compressed_images': compressed_images,
                    'compression_level': compression_level,
                    'image_quality': image_quality
                }
                
                self.dialog.after(0, self._test_complete, result_data)
                
            except Exception as e:
                self.dialog.after(0, self._test_complete, {'success': False, 'error': str(e)})
            
        except Exception as e:
            self.dialog.after(0, self._test_complete, {'success': False, 'error': str(e)})
    
    def _test_complete(self, result_data):
        """Zeigt Testergebnis kompakt an"""
        self.test_running = False
        self.test_button.config(state=tk.NORMAL)
        self.test_progress.stop()
        self.test_progress.pack_forget()
        
        # Leere vorherige Ergebnisse
        for widget in self.test_result_frame.winfo_children():
            widget.destroy()
        
        if result_data['success']:
            # Erfolg
            self.test_result_frame.pack(fill=tk.X, pady=(10, 0))
            
            # Größenberechnung
            orig_size = result_data['original_size']
            comp_size = result_data['compressed_size']
            reduction = result_data['reduction_percent']
            
            size_kb_orig = orig_size / 1024
            size_mb_orig = orig_size / (1024*1024)
            size_kb_comp = comp_size / 1024
            size_mb_comp = comp_size / (1024*1024)
            
            # Formatierung
            if size_mb_orig >= 1:
                orig_text = f"{size_mb_orig:.2f} MB"
            else:
                orig_text = f"{size_kb_orig:.0f} KB"
            
            if size_mb_comp >= 1:
                comp_text = f"{size_mb_comp:.2f} MB"
            else:
                comp_text = f"{size_kb_comp:.0f} KB"
            
            # Farbe basierend auf Reduzierung
            if reduction > 30:
                color = "green"
                status_icon = "✓"
            elif reduction > 10:
                color = "orange"
                status_icon = "✓"
            elif reduction > 0:
                color = "dark orange"
                status_icon = "✓"
            else:
                color = "blue"
                status_icon = "ℹ️"
            
            # Kompakte Darstellung in 2-3 Zeilen
            if reduction > 0:
                # Zeile 1: Status und Größenänderung
                line1_frame = ttk.Frame(self.test_result_frame)
                line1_frame.pack(fill=tk.X)
                
                status_label = ttk.Label(
                    line1_frame,
                    text=f"{status_icon} Komprimierung erfolgreich: {orig_text} → {comp_text} ({reduction:.1f}% kleiner)",
                    font=('TkDefaultFont', 10, 'bold'),
                    foreground=color
                )
                status_label.pack(side=tk.LEFT)
                
                # Zeile 2: Details
                line2_frame = ttk.Frame(self.test_result_frame)
                line2_frame.pack(fill=tk.X, pady=(5, 0))
                
                details_text = f"Zeit: {result_data['compression_time']:.1f}s | Seiten: {result_data['page_count']}"
                if result_data['total_images'] > 0:
                    details_text += f" | Bilder: {result_data['compressed_images']}/{result_data['total_images']}"
                details_text += f" | Einstellungen: L{result_data['compression_level']}/Q{result_data['image_quality']}%"
                
                details_label = ttk.Label(
                    line2_frame,
                    text=details_text,
                    font=('TkDefaultFont', 9),
                    foreground='gray'
                )
                details_label.pack(side=tk.LEFT)
            else:
                # PDF bereits optimiert
                line1_frame = ttk.Frame(self.test_result_frame)
                line1_frame.pack(fill=tk.X)
                
                status_label = ttk.Label(
                    line1_frame,
                    text=f"{status_icon} PDF ist bereits optimiert: {orig_text} (keine weitere Reduzierung möglich)",
                    font=('TkDefaultFont', 10),
                    foreground=color
                )
                status_label.pack(side=tk.LEFT)
                
                # Details
                line2_frame = ttk.Frame(self.test_result_frame)
                line2_frame.pack(fill=tk.X, pady=(5, 0))
                
                details_text = f"Zeit: {result_data['compression_time']:.1f}s | Seiten: {result_data['page_count']}"
                if result_data['total_images'] > 0:
                    details_text += f" | Bilder gefunden: {result_data['total_images']}"
                
                details_label = ttk.Label(
                    line2_frame,
                    text=details_text,
                    font=('TkDefaultFont', 9),
                    foreground='gray'
                )
                details_label.pack(side=tk.LEFT)
            
            # Preview-Button aktivieren
            self.preview_button.config(state=tk.NORMAL)
        else:
            # Fehler
            self.test_result_frame.pack(fill=tk.X, pady=(10, 0))
            
            error_label = ttk.Label(
                self.test_result_frame,
                text=f"❌ Fehler bei Komprimierung: {result_data['error']}",
                font=('TkDefaultFont', 10),
                foreground='red',
                wraplength=600
            )
            error_label.pack(side=tk.LEFT)
    
    def _show_preview(self):
        """Öffnet komprimierte PDF"""
        if self.compressed_pdf_path and os.path.exists(self.compressed_pdf_path):
            try:
                if os.name == 'nt':
                    os.startfile(self.compressed_pdf_path)
                elif os.name == 'posix':
                    import subprocess
                    subprocess.run(['open' if sys.platform == 'darwin' else 'xdg-open', 
                                  self.compressed_pdf_path])
            except Exception as e:
                messagebox.showerror("Fehler", f"Konnte PDF nicht öffnen: {str(e)}")
    
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
            'compression_level': int(self.compression_var.get()),
            'image_quality': int(self.image_quality_var.get())
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
