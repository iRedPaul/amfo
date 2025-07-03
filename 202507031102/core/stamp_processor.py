"""
Stempel-Verarbeitung für PDFs
"""
import os
import io
import logging
from typing import List, Dict, Any, Optional, Tuple
from pathlib import Path
import fitz  # PyMuPDF
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter, A4
from reportlab.lib.colors import HexColor
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from PIL import Image, ImageDraw, ImageFont
import tempfile
import math

from models.stamp_config import StampConfig, StampPosition
from core.function_parser import FunctionParser

logger = logging.getLogger(__name__)


class StampProcessor:
    """Verarbeitet und fügt Stempel zu PDFs hinzu"""
    
    def __init__(self):
        self.function_parser = FunctionParser()
        self._register_fonts()
    
    def _register_fonts(self):
        """Registriert Schriftarten für ReportLab"""
        try:
            # Versuche Windows-Schriftarten zu registrieren
            font_paths = {
                'Helvetica': r'C:\Windows\Fonts\arial.ttf',
                'Helvetica-Bold': r'C:\Windows\Fonts\arialbd.ttf',
                'Helvetica-Oblique': r'C:\Windows\Fonts\ariali.ttf',
                'Helvetica-BoldOblique': r'C:\Windows\Fonts\arialbi.ttf',
                'Times-Roman': r'C:\Windows\Fonts\times.ttf',
                'Times-Bold': r'C:\Windows\Fonts\timesbd.ttf',
                'Times-Italic': r'C:\Windows\Fonts\timesi.ttf',
                'Times-BoldItalic': r'C:\Windows\Fonts\timesbi.ttf',
                'Courier': r'C:\Windows\Fonts\cour.ttf',
                'Courier-Bold': r'C:\Windows\Fonts\courbd.ttf',
                'Courier-Oblique': r'C:\Windows\Fonts\couri.ttf',
                'Courier-BoldOblique': r'C:\Windows\Fonts\courbi.ttf'
            }
            
            for font_name, font_path in font_paths.items():
                if os.path.exists(font_path):
                    try:
                        pdfmetrics.registerFont(TTFont(font_name, font_path))
                    except:
                        pass
        except Exception as e:
            logger.debug(f"Font-Registrierung übersprungen: {e}")
    
    def apply_stamps(self, pdf_path: str, stamp_configs: List[StampConfig], 
                     context: Dict[str, Any]) -> Tuple[bool, str]:
        """
        Wendet Stempel auf eine PDF an
        
        Args:
            pdf_path: Pfad zur PDF-Datei
            stamp_configs: Liste der anzuwendenden Stempel
            context: Kontext mit Variablen für Expressions
            
        Returns:
            Tuple[bool, str]: (Erfolg, Fehlermeldung oder Erfolgsmeldung)
        """
        if not stamp_configs:
            return True, "Keine Stempel zu verarbeiten"
        
        # Filtere aktivierte Stempel
        active_stamps = [s for s in stamp_configs if s.enabled]
        if not active_stamps:
            return True, "Keine aktiven Stempel"
        
        try:
            # Erstelle temporäre Datei für Ausgabe
            with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
                temp_path = tmp_file.name
            
            # Öffne Original-PDF mit PyMuPDF
            doc = fitz.open(pdf_path)
            
            # Verarbeite jede Seite
            for page_num in range(doc.page_count):
                page = doc[page_num]
                
                # Wende Stempel auf diese Seite an
                for stamp in active_stamps:
                    if self._should_stamp_page(stamp, page_num + 1, doc.page_count):
                        # Erstelle Stempel-PDF
                        stamp_pdf_bytes = self._create_stamp_pdf(
                            page, stamp, context, page_num + 1
                        )
                        
                        if stamp_pdf_bytes:
                            # Öffne Stempel-PDF
                            stamp_doc = fitz.open("pdf", stamp_pdf_bytes)
                            stamp_page = stamp_doc[0]
                            
                            # Füge Stempel zur Seite hinzu
                            page.show_pdf_page(page.rect, stamp_doc, 0, overlay=True)
                            
                            stamp_doc.close()
            
            # Speichere gestempelte PDF mit optimierten Einstellungen
            doc.save(temp_path, 
                    garbage=4,  # Maximale Garbage Collection
                    deflate=True,  # Komprimierung
                    clean=True,  # Aufräumen
                    pretty=True)  # Schöne Formatierung für Adobe-Konformität
            
            doc.close()
            
            # Ersetze Original mit gestempelter Version
            os.replace(temp_path, pdf_path)
            
            return True, f"{len(active_stamps)} Stempel erfolgreich angewendet"
            
        except Exception as e:
            logger.error(f"Fehler beim Anwenden der Stempel: {e}")
            # Aufräumen
            if 'temp_path' in locals() and os.path.exists(temp_path):
                os.remove(temp_path)
            return False, f"Fehler beim Stempeln: {str(e)}"
    
    def _should_stamp_page(self, stamp: StampConfig, page_num: int, total_pages: int) -> bool:
        """Prüft ob eine Seite gestempelt werden soll"""
        if stamp.apply_to_pages == "all":
            return True
        elif stamp.apply_to_pages == "first":
            return page_num == 1
        elif stamp.apply_to_pages == "last":
            return page_num == total_pages
        elif stamp.apply_to_pages == "custom":
            return page_num in stamp.custom_pages
        return False
    
    def _create_stamp_pdf(self, page: fitz.Page, stamp: StampConfig,
                         context: Dict[str, Any], page_num: int) -> Optional[bytes]:
        """Erstellt eine Stempel-PDF als Bytes"""
        try:
            # Hole Seitengröße
            page_rect = page.rect
            page_width = page_rect.width
            page_height = page_rect.height
            
            # Erstelle In-Memory PDF für Stempel
            packet = io.BytesIO()
            can = canvas.Canvas(packet, pagesize=(page_width, page_height))
            
            # Evaluiere alle Texte
            evaluated_lines = []
            for line in stamp.stamp_lines:
                evaluated_text = self.function_parser.parse_and_evaluate(
                    line['text'], context
                )
                evaluated_line = line.copy()
                evaluated_line['text'] = evaluated_text
                evaluated_lines.append(evaluated_line)
            
            # Berechne Stempel-Größe
            if stamp.auto_size:
                stamp_width, stamp_height = self._calculate_auto_size(
                    can, evaluated_lines, stamp.padding
                )
            else:
                stamp_width = stamp.fixed_width
                stamp_height = stamp.fixed_height
            
            # Berechne Position
            x, y = self._calculate_position(
                stamp, page_width, page_height, stamp_width, stamp_height
            )
            
            # Speichere Canvas-Status
            can.saveState()
            
            # Wende Rotation an
            if stamp.rotation != 0:
                # Rotiere um Stempel-Zentrum
                cx = x + stamp_width / 2
                cy = y + stamp_height / 2
                can.translate(cx, cy)
                can.rotate(stamp.rotation)
                can.translate(-cx, -cy)
            
            # Zeichne Schatten
            if stamp.show_shadow:
                self._draw_shadow(can, x, y, stamp_width, stamp_height, stamp)
            
            # Zeichne Hintergrund
            if stamp.show_background:
                self._draw_background(can, x, y, stamp_width, stamp_height, stamp)
            
            # Zeichne Rahmen
            if stamp.show_border:
                self._draw_border(can, x, y, stamp_width, stamp_height, stamp)
            
            # Zeichne Text
            self._draw_text(can, x, y, stamp_width, stamp_height, evaluated_lines, stamp)
            
            # Restore Canvas-Status
            can.restoreState()
            
            # Finalisiere Canvas
            can.save()
            
            # Hole PDF-Bytes
            packet.seek(0)
            pdf_bytes = packet.read()
            
            return pdf_bytes
            
        except Exception as e:
            logger.error(f"Fehler beim Erstellen des Stempel-PDFs: {e}")
            return None
    
    def _calculate_auto_size(self, canvas_obj: canvas.Canvas, lines: List[Dict], 
                            padding: float) -> Tuple[float, float]:
        """Berechnet die automatische Größe basierend auf dem Inhalt"""
        max_width = 0
        total_height = 0
        
        for i, line in enumerate(lines):
            # Setze Font
            font_name = self._get_font_name(line)
            font_size = line['font_size']
            canvas_obj.setFont(font_name, font_size)
            
            # Berechne Textbreite
            text_width = canvas_obj.stringWidth(line['text'], font_name, font_size)
            max_width = max(max_width, text_width)
            
            # Höhe
            total_height += font_size * 1.2
            if i < len(lines) - 1:
                total_height += font_size * 0.3  # Zeilenabstand
        
        width = max_width + 2 * padding
        height = total_height + 2 * padding
        
        return width, height
    
    def _calculate_position(self, stamp: StampConfig, page_width: float, 
                           page_height: float, stamp_width: float, 
                           stamp_height: float) -> Tuple[float, float]:
        """Berechnet die Position des Stempels auf der Seite"""
        if stamp.position == StampPosition.CUSTOM:
            return stamp.custom_x, page_height - stamp.custom_y - stamp_height
        
        # Automatische Positionen
        margin_x = stamp.margin_x
        margin_y = stamp.margin_y
        
        # X-Position
        if "LEFT" in stamp.position.value:
            x = margin_x
        elif "RIGHT" in stamp.position.value:
            x = page_width - margin_x - stamp_width
        else:  # CENTER
            x = (page_width - stamp_width) / 2
        
        # Y-Position (PDF-Koordinaten: 0 ist unten)
        if "TOP" in stamp.position.value:
            y = page_height - margin_y - stamp_height
        elif "BOTTOM" in stamp.position.value:
            y = margin_y
        else:  # MIDDLE
            y = (page_height - stamp_height) / 2
        
        return x, y
    
    def _draw_shadow(self, can: canvas.Canvas, x: float, y: float, 
                    width: float, height: float, stamp: StampConfig):
        """Zeichnet den Schatten"""
        shadow_color = HexColor(stamp.shadow_color)
        # Setze Transparenz
        shadow_color.alpha = stamp.shadow_opacity
        
        can.setFillColor(shadow_color)
        can.setStrokeColor(shadow_color)
        
        shadow_x = x + stamp.shadow_offset_x
        shadow_y = y - stamp.shadow_offset_y  # Minus weil PDF-Y invertiert
        
        if stamp.border_radius > 0:
            can.roundRect(shadow_x, shadow_y, width, height, 
                         stamp.border_radius, fill=1, stroke=0)
        else:
            can.rect(shadow_x, shadow_y, width, height, fill=1, stroke=0)
    
    def _draw_background(self, can: canvas.Canvas, x: float, y: float, 
                        width: float, height: float, stamp: StampConfig):
        """Zeichnet den Hintergrund"""
        bg_color = HexColor(stamp.background_color)
        bg_color.alpha = stamp.background_opacity
        
        can.setFillColor(bg_color)
        
        if stamp.border_radius > 0:
            can.roundRect(x, y, width, height, stamp.border_radius, fill=1, stroke=0)
        else:
            can.rect(x, y, width, height, fill=1, stroke=0)
    
    def _draw_border(self, can: canvas.Canvas, x: float, y: float, 
                    width: float, height: float, stamp: StampConfig):
        """Zeichnet den Rahmen"""
        can.setStrokeColor(HexColor(stamp.border_color))
        can.setLineWidth(stamp.border_width)
        
        if stamp.border_radius > 0:
            can.roundRect(x, y, width, height, stamp.border_radius, fill=0, stroke=1)
        else:
            can.rect(x, y, width, height, fill=0, stroke=1)
    
    def _draw_text(self, can: canvas.Canvas, x: float, y: float, 
                  width: float, height: float, lines: List[Dict], stamp: StampConfig):
        """Zeichnet die Textzeilen"""
        padding = stamp.padding
        current_y = y + height - padding  # Start von oben
        
        for line in lines:
            # Font
            font_name = self._get_font_name(line)
            font_size = line['font_size']
            can.setFont(font_name, font_size)
            can.setFillColor(HexColor(line['font_color']))
            
            # Text-Position basierend auf Ausrichtung
            text_width = can.stringWidth(line['text'], font_name, font_size)
            
            if line.get('alignment', 'left') == 'center':
                text_x = x + width / 2
            elif line.get('alignment', 'left') == 'right':
                text_x = x + width - padding
            else:  # left
                text_x = x + padding
            
            # Zeichne Text
            current_y -= font_size  # Bewege nach unten für nächste Zeile
            
            if line.get('alignment', 'left') == 'center':
                can.drawCentredString(text_x, current_y, line['text'])
            elif line.get('alignment', 'left') == 'right':
                can.drawRightString(text_x, current_y, line['text'])
            else:
                can.drawString(text_x, current_y, line['text'])
            
            current_y -= font_size * 0.3  # Zeilenabstand
    
    def _get_font_name(self, line: Dict) -> str:
        """Bestimmt den Font-Namen basierend auf Stil"""
        base_font = line['font_name']
        
        # Mapping für Standard-Fonts
        font_map = {
            'Helvetica': {
                'normal': 'Helvetica',
                'bold': 'Helvetica-Bold',
                'italic': 'Helvetica-Oblique',
                'bolditalic': 'Helvetica-BoldOblique'
            },
            'Times-Roman': {
                'normal': 'Times-Roman',
                'bold': 'Times-Bold',
                'italic': 'Times-Italic',
                'bolditalic': 'Times-BoldItalic'
            },
            'Courier': {
                'normal': 'Courier',
                'bold': 'Courier-Bold',
                'italic': 'Courier-Oblique',
                'bolditalic': 'Courier-BoldOblique'
            }
        }
        
        # Bestimme Stil
        style = 'normal'
        if line.get('bold', False) and line.get('italic', False):
            style = 'bolditalic'
        elif line.get('bold', False):
            style = 'bold'
        elif line.get('italic', False):
            style = 'italic'
        
        # Hole Font
        if base_font in font_map and style in font_map[base_font]:
            return font_map[base_font][style]
        
        return base_font