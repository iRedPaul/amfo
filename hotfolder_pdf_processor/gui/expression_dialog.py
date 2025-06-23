"""
Dialog zur Bearbeitung von Ausdrücken mit Variablen und Funktionen
"""
import tkinter as tk
from tkinter import ttk
from typing import Optional, List, Dict
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gui.expression_editor_base import ExpressionEditorBase


class ExpressionDialog(ExpressionEditorBase):
    """Dialog zur Bearbeitung von Ausdrücken - erweitert die Basis-Klasse"""
    
    def __init__(self, parent, title: str = "Ausdruck bearbeiten", 
                 expression: str = "", description: str = "",
                 xml_field_mappings: List[Dict] = None):
        # Initialisiere mit der Basis-Klasse
        super().__init__(
            parent=parent,
            title=title,
            expression=expression,
            description=description,
            ocr_zones=[],  # Keine OCR-Zonen für einfache Expression-Dialoge
            geometry="1200x800",
            xml_field_mappings=xml_field_mappings  # WICHTIG: Übergebe XML-Feld-Mappings
        )
        
        # Fokus auf Expression-Text
        self.expr_text.focus()
    
    # Alle anderen Methoden werden von der Basis-Klasse übernommen
    # Keine zusätzliche Implementierung nötig!