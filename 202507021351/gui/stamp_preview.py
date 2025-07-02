"""
Vorschau-Widget für Stempel
"""
import tkinter as tk
from typing import Optional
import math

from models.stamp_config import StampConfig, StampPosition


class StampPreview:
    """Zeigt eine Vorschau des Stempels"""
    
    def __init__(self, canvas: tk.Canvas):
        self.canvas = canvas
        self.stamp_config = None
        
        # PDF-Seite simulieren
        self.page_width = 400
        self.page_height = 566  # A4 Verhältnis
        self.page_margin = 20
        
        # Standard-Skalierung
        self.scale = 1.0
        self.page_x = 0
        self.page_y = 0
        
        # Warte bis Canvas bereit ist, dann zeichne Seite
        self.canvas.after(100, self._draw_page)
    
    def _draw_page(self):
        """Zeichnet die Seiten-Vorschau"""
        self.canvas.delete("all")
        
        # Canvas-Größe
        canvas_width = self.canvas.winfo_width()
        canvas_height = self.canvas.winfo_height()
        
        if canvas_width <= 1 or canvas_height <= 1:  # Canvas noch nicht bereit
            self.canvas.after(100, self._draw_page)
            return
        
        # Skaliere Seite um in Canvas zu passen
        scale_x = (canvas_width - 2 * self.page_margin) / self.page_width
        scale_y = (canvas_height - 2 * self.page_margin) / self.page_height
        self.scale = min(scale_x, scale_y)
        
        # Zentriere Seite
        self.page_x = (canvas_width - self.page_width * self.scale) / 2
        self.page_y = (canvas_height - self.page_height * self.scale) / 2
        
        # Zeichne Seite mit Schatten
        shadow_offset = 3
        self.canvas.create_rectangle(
            self.page_x + shadow_offset, 
            self.page_y + shadow_offset,
            self.page_x + self.page_width * self.scale + shadow_offset, 
            self.page_y + self.page_height * self.scale + shadow_offset,
            fill="#cccccc", outline="", tags="shadow"
        )
        
        # Zeichne Seite
        self.canvas.create_rectangle(
            self.page_x, self.page_y,
            self.page_x + self.page_width * self.scale, 
            self.page_y + self.page_height * self.scale,
            fill="white", outline="#333333", width=1, tags="page"
        )
        
        # Zeichne Raster (optional)
        self._draw_grid()
        
        # Wenn ein Stempel vorhanden ist, zeichne ihn neu
        if self.stamp_config:
            self.update_stamp(self.stamp_config)
    
    def _draw_grid(self):
        """Zeichnet ein Hilfsraster"""
        grid_size = 50 * self.scale
        
        # Vertikale Linien
        x = self.page_x + grid_size
        while x < self.page_x + self.page_width * self.scale:
            self.canvas.create_line(
                x, self.page_y,
                x, self.page_y + self.page_height * self.scale,
                fill="#f0f0f0", tags="grid"
            )
            x += grid_size
        
        # Horizontale Linien
        y = self.page_y + grid_size
        while y < self.page_y + self.page_height * self.scale:
            self.canvas.create_line(
                self.page_x, y,
                self.page_x + self.page_width * self.scale, y,
                fill="#f0f0f0", tags="grid"
            )
            y += grid_size
    
    def update_stamp(self, config: StampConfig):
        """Aktualisiert die Stempel-Vorschau"""
        self.stamp_config = config
        
        # Lösche alten Stempel
        self.canvas.delete("stamp")
        
        if not config or not config.stamp_lines:
            return
        
        # Warte falls Canvas noch nicht initialisiert
        if self.scale == 1.0 and self.page_x == 0 and self.page_y == 0:
            # Canvas noch nicht bereit, warte
            self.canvas.after(200, lambda: self.update_stamp(config))
            return
        
        # Berechne Stempel-Position
        stamp_x, stamp_y = self._calculate_position(config)
        
        # Berechne Stempel-Größe
        if config.auto_size:
            width, height = self._calculate_auto_size(config)
        else:
            width = config.fixed_width * self.scale
            height = config.fixed_height * self.scale
        
        # Zeichne Stempel
        self._draw_stamp(stamp_x, stamp_y, width, height, config)
    
    def _calculate_position(self, config: StampConfig):
        """Berechnet die Position des Stempels"""
        margin_x = config.margin_x * self.scale
        margin_y = config.margin_y * self.scale
        
        if config.position == StampPosition.CUSTOM:
            x = self.page_x + config.custom_x * self.scale
            y = self.page_y + config.custom_y * self.scale
        else:
            # Automatische Positionen
            if "LEFT" in config.position.value:
                x = self.page_x + margin_x
            elif "RIGHT" in config.position.value:
                # Temporär - wird nach Größenberechnung angepasst
                x = self.page_x + self.page_width * self.scale - margin_x
            else:  # CENTER
                x = self.page_x + self.page_width * self.scale / 2
            
            if "TOP" in config.position.value:
                y = self.page_y + margin_y
            elif "BOTTOM" in config.position.value:
                # Temporär - wird nach Größenberechnung angepasst
                y = self.page_y + self.page_height * self.scale - margin_y
            else:  # MIDDLE
                y = self.page_y + self.page_height * self.scale / 2
        
        return x, y
    
    def _calculate_auto_size(self, config: StampConfig):
        """Berechnet die automatische Größe basierend auf Inhalt"""
        # Simuliere Text-Größen
        padding = config.padding * self.scale
        
        # Berechne maximale Breite und Gesamthöhe
        max_width = 0
        total_height = 0
        
        for i, line in enumerate(config.stamp_lines):
            # Simuliere Textgröße
            font_size = line['font_size'] * self.scale
            text_width = len(line['text']) * font_size * 0.6  # Grobe Schätzung
            text_height = font_size * 1.2
            
            max_width = max(max_width, text_width)
            total_height += text_height
            
            if i < len(config.stamp_lines) - 1:
                total_height += font_size * 0.3  # Zeilenabstand
        
        width = max_width + 2 * padding
        height = total_height + 2 * padding
        
        return width, height
    
    def _draw_stamp(self, x, y, width, height, config: StampConfig):
        """Zeichnet den Stempel"""
        # Anpassung für rechts/unten Positionen
        if "RIGHT" in config.position.value and config.position != StampPosition.CUSTOM:
            x -= width
        elif "CENTER" in config.position.value and config.position != StampPosition.CUSTOM:
            x -= width / 2
        
        if "BOTTOM" in config.position.value and config.position != StampPosition.CUSTOM:
            y -= height
        elif "MIDDLE" in config.position.value and config.position != StampPosition.CUSTOM:
            y -= height / 2
        
        # Speichere Zentrum für Rotation
        center_x = x + width / 2
        center_y = y + height / 2
        
        # Zeichne basierend auf Rotation
        if config.rotation != 0:
            self._draw_rotated_stamp(center_x, center_y, width, height, config.rotation, config)
        else:
            self._draw_normal_stamp(x, y, width, height, config)
    
    def _draw_normal_stamp(self, x, y, width, height, config: StampConfig):
        """Zeichnet einen nicht-rotierten Stempel"""
        # Schatten
        if config.show_shadow:
            shadow_x = config.shadow_offset_x * self.scale
            shadow_y = config.shadow_offset_y * self.scale
            
            # Shadow mit Transparenz simulieren
            self.canvas.create_rectangle(
                x + shadow_x, y + shadow_y,
                x + width + shadow_x, y + height + shadow_y,
                fill=config.shadow_color,
                outline="",
                tags="stamp",
                stipple="gray50"  # Simuliere Transparenz
            )
        
        # Hintergrund
        if config.show_background:
            if config.border_radius > 0:
                # Abgerundete Ecken
                self._create_rounded_rectangle(
                    x, y, x + width, y + height,
                    config.border_radius * self.scale,
                    fill=config.background_color,
                    outline="",
                    tags="stamp"
                )
            else:
                self.canvas.create_rectangle(
                    x, y, x + width, y + height,
                    fill=config.background_color,
                    outline="",
                    tags="stamp"
                )
        
        # Rahmen
        if config.show_border:
            if config.border_radius > 0:
                self._create_rounded_rectangle(
                    x, y, x + width, y + height,
                    config.border_radius * self.scale,
                    fill="",
                    outline=config.border_color,
                    width=config.border_width * self.scale,
                    tags="stamp"
                )
            else:
                self.canvas.create_rectangle(
                    x, y, x + width, y + height,
                    fill="",
                    outline=config.border_color,
                    width=config.border_width * self.scale,
                    tags="stamp"
                )
        
        # Text
        padding = config.padding * self.scale
        text_y = y + padding
        
        for line in config.stamp_lines:
            font_size = int(line['font_size'] * self.scale)
            
            # Font style
            font_style = ""
            if line.get('bold', False):
                font_style += "bold "
            if line.get('italic', False):
                font_style += "italic "
            
            font = (line['font_name'], font_size, font_style.strip())
            
            # Text position basierend auf Ausrichtung
            if line.get('alignment', 'left') == 'center':
                text_x = x + width / 2
                anchor = "n"
            elif line.get('alignment', 'left') == 'right':
                text_x = x + width - padding
                anchor = "ne"
            else:  # left
                text_x = x + padding
                anchor = "nw"
            
            # Zeichne Text
            self.canvas.create_text(
                text_x, text_y,
                text=line['text'],
                font=font,
                fill=line['font_color'],
                anchor=anchor,
                tags="stamp"
            )
            
            text_y += font_size * 1.5
    
    def _draw_rotated_stamp(self, cx, cy, width, height, rotation, config: StampConfig):
        """Zeichnet einen rotierten Stempel mit vollständigem Inhalt"""
        # Speichere Canvas-Status
        self.canvas.create_text(
            cx, cy - height/2 - 20,
            text=f"Rotation: {int(rotation)}°",
            font=("Arial", int(10 * self.scale)),
            fill="#666666",
            anchor="center",
            tags="stamp"
        )
        
        # Berechne rotierte Bounding Box
        rad = math.radians(rotation)
        
        # Die vier Ecken des Rechtecks relativ zum Zentrum
        corners = [
            (-width/2, -height/2),
            (width/2, -height/2),
            (width/2, height/2),
            (-width/2, height/2)
        ]
        
        # Rotiere die Ecken
        rotated_corners = []
        for x, y in corners:
            rx = x * math.cos(rad) - y * math.sin(rad)
            ry = x * math.sin(rad) + y * math.cos(rad)
            rotated_corners.append((cx + rx, cy + ry))
        
        # Schatten
        if config.show_shadow:
            shadow_offset = config.shadow_offset_x * self.scale
            shadow_points = []
            for x, y in rotated_corners:
                shadow_points.extend([x + shadow_offset, y + shadow_offset])
            
            self.canvas.create_polygon(
                shadow_points,
                fill=config.shadow_color,
                outline="",
                stipple="gray50",
                tags="stamp"
            )
        
        # Hintergrund
        if config.show_background:
            points = []
            for x, y in rotated_corners:
                points.extend([x, y])
            
            self.canvas.create_polygon(
                points,
                fill=config.background_color,
                outline="",
                tags="stamp"
            )
        
        # Rahmen
        if config.show_border:
            points = []
            for x, y in rotated_corners:
                points.extend([x, y])
            
            self.canvas.create_polygon(
                points,
                fill="",
                outline=config.border_color,
                width=config.border_width * self.scale,
                tags="stamp"
            )
        
        # Text (rotiert)
        padding = config.padding * self.scale
        
        # Berechne Start-Y relativ zum Zentrum
        total_height = 0
        for line in config.stamp_lines:
            font_size = line['font_size'] * self.scale
            total_height += font_size * 1.2
        
        current_y = -total_height / 2 + padding
        
        for line in config.stamp_lines:
            font_size = int(line['font_size'] * self.scale)
            
            # Font style
            font_style = ""
            if line.get('bold', False):
                font_style += "bold "
            if line.get('italic', False):
                font_style += "italic "
            
            font = (line['font_name'], font_size, font_style.strip())
            
            # Text-Position relativ zum Zentrum
            text_x = 0
            if line.get('alignment', 'left') == 'left':
                text_x = -width/2 + padding
            elif line.get('alignment', 'left') == 'right':
                text_x = width/2 - padding
            
            # Rotiere Text-Position
            rx = text_x * math.cos(rad) - current_y * math.sin(rad)
            ry = text_x * math.sin(rad) + current_y * math.cos(rad)
            
            # Zeichne rotierten Text
            self.canvas.create_text(
                cx + rx, cy + ry,
                text=line['text'],
                font=font,
                fill=line['font_color'],
                anchor="center",
                angle=-rotation,  # Negative Rotation für Text
                tags="stamp"
            )
            
            current_y += font_size * 1.5
    
    def _create_rounded_rectangle(self, x1, y1, x2, y2, radius, **kwargs):
        """Erstellt ein Rechteck mit abgerundeten Ecken"""
        points = []
        
        # Oben links
        points.extend([x1, y1 + radius])
        points.extend([x1, y1 + radius, x1, y1 + radius, x1 + radius, y1])
        
        # Oben rechts
        points.extend([x2 - radius, y1])
        points.extend([x2 - radius, y1, x2 - radius, y1, x2, y1 + radius])
        
        # Unten rechts
        points.extend([x2, y2 - radius])
        points.extend([x2, y2 - radius, x2, y2 - radius, x2 - radius, y2])
        
        # Unten links
        points.extend([x1 + radius, y2])
        points.extend([x1 + radius, y2, x1 + radius, y2, x1, y2 - radius])
        
        return self.canvas.create_polygon(points, smooth=True, **kwargs)