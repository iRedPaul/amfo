"""
OCR-Zonen-Auswahl-Tool
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, Tuple
from PIL import Image, ImageTk
from pdf2image import convert_from_path
import tempfile
import os


class ZoneSelector:
    """Dialog zur grafischen Auswahl einer OCR-Zone"""
    
    def __init__(self, parent, pdf_path: Optional[str] = None, page_num: int = 1):
        self.parent = parent
        self.pdf_path = pdf_path
        self.page_num = page_num
        self.result = None
        self.current_scale = 1.0
        self.image = None
        self.photo_image = None
        
        # Zone-Koordinaten
        self.start_x = None
        self.start_y = None
        self.rect_id = None
        self.zone = None
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("OCR-Zone auswählen")
        self.dialog.geometry("1000x800")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Zentriere relativ zum Parent
        self._center_window()
        
        self._create_widgets()
        self._layout_widgets()
        
        # Lade PDF wenn vorhanden
        if pdf_path and os.path.exists(pdf_path):
            self._load_pdf_page()
        
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
        # Toolbar
        self.toolbar = ttk.Frame(self.dialog)
        
        self.load_button = ttk.Button(self.toolbar, text="PDF laden...", 
                                     command=self._load_pdf)
        
        self.page_label = ttk.Label(self.toolbar, text="Seite:")
        self.page_var = tk.IntVar(value=self.page_num)
        self.page_spinbox = ttk.Spinbox(self.toolbar, from_=1, to=100, 
                                       textvariable=self.page_var, width=5,
                                       command=self._on_page_change)
        
        self.zoom_label = ttk.Label(self.toolbar, text="Zoom:")
        self.zoom_out_button = ttk.Button(self.toolbar, text="-", width=3,
                                         command=self._zoom_out)
        self.zoom_reset_button = ttk.Button(self.toolbar, text="100%", width=6,
                                           command=self._zoom_reset)
        self.zoom_in_button = ttk.Button(self.toolbar, text="+", width=3,
                                        command=self._zoom_in)
        
        self.clear_button = ttk.Button(self.toolbar, text="Zone löschen",
                                      command=self._clear_zone)
        
        # Hauptbereich mit Scrollbars
        self.main_frame = ttk.Frame(self.dialog)
        
        # Canvas mit Scrollbars
        self.canvas_frame = ttk.Frame(self.main_frame)
        self.canvas = tk.Canvas(self.canvas_frame, bg='gray', cursor="cross")
        
        self.h_scrollbar = ttk.Scrollbar(self.canvas_frame, orient="horizontal",
                                        command=self.canvas.xview)
        self.v_scrollbar = ttk.Scrollbar(self.canvas_frame, orient="vertical",
                                        command=self.canvas.yview)
        
        self.canvas.configure(xscrollcommand=self.h_scrollbar.set,
                             yscrollcommand=self.v_scrollbar.set)
        
        # Info-Panel
        self.info_frame = ttk.LabelFrame(self.dialog, text="Zone-Information", 
                                        padding="10")
        self.info_label = ttk.Label(self.info_frame, 
                                   text="Ziehen Sie mit der Maus, um eine Zone auszuwählen")
        self.coords_label = ttk.Label(self.info_frame, text="")
        
        # Buttons
        self.button_frame = ttk.Frame(self.dialog)
        self.save_button = ttk.Button(self.button_frame, text="Übernehmen",
                                     command=self._on_save, state=tk.DISABLED)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen",
                                       command=self._on_cancel)
        
        # Canvas Events
        self.canvas.bind("<ButtonPress-1>", self._on_button_press)
        self.canvas.bind("<B1-Motion>", self._on_mouse_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_button_release)
        self.canvas.bind("<MouseWheel>", self._on_mousewheel)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        # Toolbar
        self.toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        self.load_button.pack(side=tk.LEFT, padx=(0, 10))
        ttk.Separator(self.toolbar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=5)
        self.page_label.pack(side=tk.LEFT)
        self.page_spinbox.pack(side=tk.LEFT, padx=(5, 10))
        ttk.Separator(self.toolbar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=5)
        self.zoom_label.pack(side=tk.LEFT)
        self.zoom_out_button.pack(side=tk.LEFT, padx=2)
        self.zoom_reset_button.pack(side=tk.LEFT, padx=2)
        self.zoom_in_button.pack(side=tk.LEFT, padx=(2, 10))
        ttk.Separator(self.toolbar, orient="vertical").pack(side=tk.LEFT, fill=tk.Y, padx=5)
        self.clear_button.pack(side=tk.LEFT)
        
        # Hauptbereich
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=5)
        
        self.canvas_frame.pack(fill=tk.BOTH, expand=True)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        self.v_scrollbar.grid(row=0, column=1, sticky="ns")
        self.h_scrollbar.grid(row=1, column=0, sticky="ew")
        
        self.canvas_frame.grid_rowconfigure(0, weight=1)
        self.canvas_frame.grid_columnconfigure(0, weight=1)
        
        # Info
        self.info_frame.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        self.info_label.pack(anchor=tk.W)
        self.coords_label.pack(anchor=tk.W)
        
        # Buttons
        self.button_frame.pack(side=tk.BOTTOM, pady=10)
        self.save_button.pack(side=tk.LEFT, padx=5)
        self.cancel_button.pack(side=tk.LEFT)
    
    def _load_pdf(self):
        """Lädt eine PDF-Datei"""
        filename = filedialog.askopenfilename(
            title="PDF auswählen",
            filetypes=[("PDF Dateien", "*.pdf"), ("Alle Dateien", "*.*")]
        )
        
        if filename:
            self.pdf_path = filename
            self._load_pdf_page()
    
    def _load_pdf_page(self):
        """Lädt eine spezifische Seite der PDF"""
        if not self.pdf_path or not os.path.exists(self.pdf_path):
            return
        
        try:
            # Konvertiere PDF-Seite zu Bild
            with tempfile.TemporaryDirectory() as temp_dir:
                poppler_path = os.path.join(os.path.dirname(__file__), '..', 'poppler', 'bin')
                images = convert_from_path(
                    self.pdf_path, 
                    dpi=150,  # Mittlere Qualität für Vorschau
                    first_page=self.page_var.get(),
                    last_page=self.page_var.get(),
                    output_folder=temp_dir,
                    poppler_path=poppler_path
                )
                
                if images:
                    self.image = images[0]
                    self._display_image()
                    
        except Exception as e:
            messagebox.showerror("Fehler", f"Fehler beim Laden der PDF: {e}")
    
    def _display_image(self):
        """Zeigt das Bild auf dem Canvas an"""
        if not self.image:
            return
        
        # Skaliere Bild
        width = int(self.image.width * self.current_scale)
        height = int(self.image.height * self.current_scale)
        
        scaled_image = self.image.resize((width, height), Image.Resampling.LANCZOS)
        self.photo_image = ImageTk.PhotoImage(scaled_image)
        
        # Lösche alten Inhalt
        self.canvas.delete("all")
        
        # Zeige Bild
        self.canvas.create_image(0, 0, anchor=tk.NW, image=self.photo_image)
        
        # Setze Canvas-Größe
        self.canvas.configure(scrollregion=self.canvas.bbox("all"))
        
        # Zeichne Zone neu wenn vorhanden
        if self.zone:
            self._draw_zone()
    
    def _on_page_change(self):
        """Wird aufgerufen wenn die Seitennummer geändert wird"""
        self.page_num = self.page_var.get()
        self._clear_zone()
        self._load_pdf_page()
    
    def _zoom_in(self):
        """Vergrößert die Ansicht"""
        self.current_scale *= 1.2
        self._display_image()
    
    def _zoom_out(self):
        """Verkleinert die Ansicht"""
        self.current_scale /= 1.2
        self._display_image()
    
    def _zoom_reset(self):
        """Setzt den Zoom zurück"""
        self.current_scale = 1.0
        self._display_image()
    
    def _on_button_press(self, event):
        """Mausklick - Start der Zonenauswahl"""
        # Konvertiere Canvas-Koordinaten zu Bild-Koordinaten
        self.start_x = self.canvas.canvasx(event.x)
        self.start_y = self.canvas.canvasy(event.y)
        
        # Lösche vorherige Zone
        if self.rect_id:
            self.canvas.delete(self.rect_id)
    
    def _on_mouse_drag(self, event):
        """Mausbewegung - Zeichne Zone"""
        if self.start_x is None:
            return
        
        # Aktuelle Position
        cur_x = self.canvas.canvasx(event.x)
        cur_y = self.canvas.canvasy(event.y)
        
        # Lösche vorheriges Rechteck
        if self.rect_id:
            self.canvas.delete(self.rect_id)
        
        # Zeichne neues Rechteck
        self.rect_id = self.canvas.create_rectangle(
            self.start_x, self.start_y, cur_x, cur_y,
            outline='red', width=2, dash=(5, 5)
        )
        
        # Aktualisiere Koordinaten-Anzeige
        x1 = min(self.start_x, cur_x) / self.current_scale
        y1 = min(self.start_y, cur_y) / self.current_scale
        x2 = max(self.start_x, cur_x) / self.current_scale
        y2 = max(self.start_y, cur_y) / self.current_scale
        
        width = x2 - x1
        height = y2 - y1
        
        self.coords_label.config(
            text=f"X: {int(x1)}, Y: {int(y1)}, Breite: {int(width)}, Höhe: {int(height)}"
        )
    
    def _on_button_release(self, event):
        """Maus losgelassen - Zone fertig"""
        if self.start_x is None:
            return
        
        # Finale Position
        end_x = self.canvas.canvasx(event.x)
        end_y = self.canvas.canvasy(event.y)
        
        # Berechne Zone in Original-Koordinaten
        x1 = min(self.start_x, end_x) / self.current_scale
        y1 = min(self.start_y, end_y) / self.current_scale
        x2 = max(self.start_x, end_x) / self.current_scale
        y2 = max(self.start_y, end_y) / self.current_scale
        
        width = x2 - x1
        height = y2 - y1
        
        # Mindestgröße prüfen
        if width > 10 and height > 10:
            # DPI-Anpassung (150 DPI für Vorschau, 300 DPI für OCR)
            dpi_factor = 300 / 150
            
            self.zone = (
                int(x1 * dpi_factor),
                int(y1 * dpi_factor),
                int(width * dpi_factor),
                int(height * dpi_factor)
            )
            self.save_button.config(state=tk.NORMAL)
            self.info_label.config(text="Zone ausgewählt. Klicken Sie auf 'Übernehmen'.")
        else:
            # Zone zu klein
            if self.rect_id:
                self.canvas.delete(self.rect_id)
                self.rect_id = None
            self.zone = None
            self.save_button.config(state=tk.DISABLED)
        
        self.start_x = None
        self.start_y = None
    
    def _on_mousewheel(self, event):
        """Mausrad für Zoom"""
        if event.delta > 0:
            self._zoom_in()
        else:
            self._zoom_out()
    
    def draw_zone(self):
        """Öffentliche Methode zum Zeichnen der Zone"""
        self._draw_zone()
    
    def _draw_zone(self):
        """Zeichnet die aktuelle Zone"""
        if not self.zone:
            return
        
        # Konvertiere von OCR-Koordinaten zu Anzeige-Koordinaten
        dpi_factor = 150 / 300  # Umgekehrte Anpassung
        x, y, w, h = self.zone
        
        x1 = x * dpi_factor * self.current_scale
        y1 = y * dpi_factor * self.current_scale
        x2 = (x + w) * dpi_factor * self.current_scale
        y2 = (y + h) * dpi_factor * self.current_scale
        
        self.rect_id = self.canvas.create_rectangle(
            x1, y1, x2, y2,
            outline='red', width=2, dash=(5, 5)
        )
    
    def _clear_zone(self):
        """Löscht die aktuelle Zone"""
        if self.rect_id:
            self.canvas.delete(self.rect_id)
            self.rect_id = None
        
        self.zone = None
        self.save_button.config(state=tk.DISABLED)
        self.coords_label.config(text="")
        self.info_label.config(text="Ziehen Sie mit der Maus, um eine Zone auszuwählen")
    
    def _on_save(self):
        """Speichert die ausgewählte Zone"""
        if self.zone:
            self.result = {
                'zone': self.zone,
                'page_num': self.page_var.get(),
                'pdf_path': self.pdf_path
            }
            self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht ab ohne zu speichern"""
        self.dialog.destroy()
    
    def show(self) -> Optional[dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result