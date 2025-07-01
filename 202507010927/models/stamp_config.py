"""
Datenmodell für Stempel-Konfigurationen
"""
from dataclasses import dataclass, field
from typing import Dict, Any, Optional, List
from enum import Enum
import json


class StampPosition(Enum):
    """Positionierungsoptionen für Stempel"""
    TOP_LEFT = "top_left"
    TOP_CENTER = "top_center"
    TOP_RIGHT = "top_right"
    MIDDLE_LEFT = "middle_left"
    MIDDLE_CENTER = "middle_center"
    MIDDLE_RIGHT = "middle_right"
    BOTTOM_LEFT = "bottom_left"
    BOTTOM_CENTER = "bottom_center"
    BOTTOM_RIGHT = "bottom_right"
    CUSTOM = "custom"


class StampOrientation(Enum):
    """Orientierung des Stempels"""
    HORIZONTAL = "horizontal"
    VERTICAL = "vertical"
    DIAGONAL = "diagonal"


@dataclass
class StampConfig:
    """Konfiguration für einen Eingangsstempel"""
    id: str
    name: str
    enabled: bool = True
    
    # Position
    position: StampPosition = StampPosition.TOP_RIGHT
    custom_x: float = 0  # Nur bei CUSTOM Position
    custom_y: float = 0  # Nur bei CUSTOM Position
    margin_x: float = 20  # Abstand vom Rand
    margin_y: float = 20  # Abstand vom Rand
    
    # Aussehen
    orientation: StampOrientation = StampOrientation.HORIZONTAL
    rotation: float = 0  # Rotation in Grad (0-360)
    
    # Rahmen
    show_border: bool = True
    border_width: float = 2.0
    border_color: str = "#000000"  # Hex-Farbe
    border_radius: float = 5.0  # Eckenradius
    
    # Hintergrund
    show_background: bool = True
    background_color: str = "#FFFFFF"  # Hex-Farbe
    background_opacity: float = 0.8  # 0.0 - 1.0
    
    # Schatten
    show_shadow: bool = True
    shadow_offset_x: float = 2
    shadow_offset_y: float = 2
    shadow_color: str = "#808080"
    shadow_opacity: float = 0.3
    
    # Inhalt
    stamp_lines: List[Dict[str, Any]] = field(default_factory=list)
    # Jede Zeile hat: text, font_name, font_size, font_color, bold, italic, alignment
    
    # Seitenkonfiguration
    apply_to_pages: str = "first"  # "first", "last", "all", "custom"
    custom_pages: List[int] = field(default_factory=list)  # Bei "custom"
    
    # Größe
    auto_size: bool = True  # Automatische Größenanpassung
    fixed_width: float = 200  # Bei auto_size = False
    fixed_height: float = 100  # Bei auto_size = False
    padding: float = 10  # Innenabstand
    
    def to_dict(self) -> dict:
        """Konvertiert die Konfiguration in ein Dictionary"""
        return {
            "id": self.id,
            "name": self.name,
            "enabled": self.enabled,
            "position": self.position.value,
            "custom_x": self.custom_x,
            "custom_y": self.custom_y,
            "margin_x": self.margin_x,
            "margin_y": self.margin_y,
            "orientation": self.orientation.value,
            "rotation": self.rotation,
            "show_border": self.show_border,
            "border_width": self.border_width,
            "border_color": self.border_color,
            "border_radius": self.border_radius,
            "show_background": self.show_background,
            "background_color": self.background_color,
            "background_opacity": self.background_opacity,
            "show_shadow": self.show_shadow,
            "shadow_offset_x": self.shadow_offset_x,
            "shadow_offset_y": self.shadow_offset_y,
            "shadow_color": self.shadow_color,
            "shadow_opacity": self.shadow_opacity,
            "stamp_lines": self.stamp_lines,
            "apply_to_pages": self.apply_to_pages,
            "custom_pages": self.custom_pages,
            "auto_size": self.auto_size,
            "fixed_width": self.fixed_width,
            "fixed_height": self.fixed_height,
            "padding": self.padding
        }
    
    @classmethod
    def from_dict(cls, data: dict) -> 'StampConfig':
        """Erstellt eine StampConfig aus einem Dictionary"""
        position = StampPosition(data.get("position", StampPosition.TOP_RIGHT.value))
        orientation = StampOrientation(data.get("orientation", StampOrientation.HORIZONTAL.value))
        
        return cls(
            id=data["id"],
            name=data["name"],
            enabled=data.get("enabled", True),
            position=position,
            custom_x=data.get("custom_x", 0),
            custom_y=data.get("custom_y", 0),
            margin_x=data.get("margin_x", 20),
            margin_y=data.get("margin_y", 20),
            orientation=orientation,
            rotation=data.get("rotation", 0),
            show_border=data.get("show_border", True),
            border_width=data.get("border_width", 2.0),
            border_color=data.get("border_color", "#000000"),
            border_radius=data.get("border_radius", 5.0),
            show_background=data.get("show_background", True),
            background_color=data.get("background_color", "#FFFFFF"),
            background_opacity=data.get("background_opacity", 0.8),
            show_shadow=data.get("show_shadow", True),
            shadow_offset_x=data.get("shadow_offset_x", 2),
            shadow_offset_y=data.get("shadow_offset_y", 2),
            shadow_color=data.get("shadow_color", "#808080"),
            shadow_opacity=data.get("shadow_opacity", 0.3),
            stamp_lines=data.get("stamp_lines", []),
            apply_to_pages=data.get("apply_to_pages", "first"),
            custom_pages=data.get("custom_pages", []),
            auto_size=data.get("auto_size", True),
            fixed_width=data.get("fixed_width", 200),
            fixed_height=data.get("fixed_height", 100),
            padding=data.get("padding", 10)
        )