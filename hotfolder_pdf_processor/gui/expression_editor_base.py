"""
Basis-Klasse für Expression-Editoren mit Variablen, Funktionen und Hilfe
"""
import tkinter as tk
from tkinter import ttk
from typing import Optional, List, Dict
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.xml_field_processor import XMLFieldProcessor


class ExpressionEditorBase:
    """Basis-Klasse für Expression-Editoren mit integrierter Hilfe"""
    
    def __init__(self, parent, title: str = "Ausdruck bearbeiten", 
                 expression: str = "", description: str = "", 
                 ocr_zones: List[Dict] = None, geometry: str = "1200x800"):
        self.parent = parent
        self.expression = expression
        self.result = None
        self.xml_processor = XMLFieldProcessor()
        self.ocr_zones = ocr_zones or []
        
        # Dialog erstellen
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(title)
        self.dialog.geometry(geometry)
        self.dialog.resizable(True, True)
        self.dialog.minsize(1000, 600)
        
        # Zentriere Dialog
        self._center_window()
        
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        self.description = description
        
        # Hilfe-Inhalte Dictionary
        self.help_content = self._create_help_content()
        
        self._create_widgets()
        self._layout_widgets()
        
        # Setze initialen Ausdruck
        if expression:
            self.expr_text.insert("1.0", expression)
        
        # Zeige initiale Hilfe
        self._show_initial_help()
        
        # Bind Events
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
        """Erstellt alle Widgets - kann in Subklassen erweitert werden"""
        # Hauptframe
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Beschreibung (optional)
        if self.description:
            self.desc_frame = ttk.LabelFrame(self.main_frame, text="Information", padding="10")
            self.desc_label = ttk.Label(self.desc_frame, text=self.description, wraplength=800)
        
        # Erstelle zusätzliche Widgets für Subklassen
        self._create_additional_widgets()
        
        # Expression Builder
        self.expr_frame = ttk.LabelFrame(self.main_frame, text="Ausdruck erstellen", padding="10")
        
        # Expression-Eingabe
        self.expr_input_frame = ttk.Frame(self.expr_frame)
        self.expr_label = ttk.Label(self.expr_input_frame, text="Ausdruck:")
        self.expr_text = tk.Text(self.expr_input_frame, height=4, width=50)
        
        # Unterer Bereich: Variablen/Funktionen und Hilfe nebeneinander
        self.bottom_frame = ttk.Frame(self.expr_frame)
        
        # Linke Seite: Variablen und Funktionen
        self.left_panel = ttk.LabelFrame(self.bottom_frame, text="Variablen und Funktionen", padding="5")
        
        # Tree mit Scrollbar
        self.tree_container = ttk.Frame(self.left_panel)
        self.var_func_tree = ttk.Treeview(self.tree_container, height=20)
        self.var_func_tree.heading("#0", text="Verfügbare Elemente")
        
        self.tree_scroll = ttk.Scrollbar(self.tree_container, orient="vertical",
                                        command=self.var_func_tree.yview)
        self.var_func_tree.configure(yscrollcommand=self.tree_scroll.set)
        
        # Rechte Seite: Hilfe-Panel
        self.right_panel = ttk.LabelFrame(self.bottom_frame, text="Hilfe", padding="5")
        
        self.help_text = tk.Text(self.right_panel, height=20, width=50, wrap=tk.WORD, state=tk.DISABLED)
        self.help_scroll = ttk.Scrollbar(self.right_panel, orient="vertical",
                                        command=self.help_text.yview)
        self.help_text.configure(yscrollcommand=self.help_scroll.set)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self._create_buttons()
        
        # Events
        self.var_func_tree.bind("<Double-Button-1>", self._on_tree_double_click)
        self.var_func_tree.bind("<<TreeviewSelect>>", self._on_tree_selection_changed)
        
        # Lade Variablen und Funktionen
        self._load_variables_and_functions()
    
    def _create_additional_widgets(self):
        """Überschreibbar für zusätzliche Widgets in Subklassen"""
        pass
    
    def _create_buttons(self):
        """Erstellt die Standard-Buttons - kann überschrieben werden"""
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
        self.ok_button = ttk.Button(self.button_frame, text="OK", 
                                   command=self._on_ok, default=tk.ACTIVE)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets - kann in Subklassen erweitert werden"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Beschreibung
        if hasattr(self, 'desc_frame'):
            self.desc_frame.pack(fill=tk.X, pady=(0, 10))
            self.desc_label.pack(anchor=tk.W)
        
        # Zusätzliche Widgets für Subklassen
        self._layout_additional_widgets()
        
        # Expression Builder
        self.expr_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        
        # Expression-Eingabe
        self.expr_input_frame.pack(fill=tk.X, pady=(0, 10))
        self.expr_label.pack(anchor=tk.W)
        self.expr_text.pack(fill=tk.X, pady=(5, 0))
        
        # Unterer Bereich
        self.bottom_frame.pack(fill=tk.BOTH, expand=True)
        
        # Linkes Panel: Variablen/Funktionen
        self.left_panel.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=(0, 5))
        self.tree_container.pack(fill=tk.BOTH, expand=True)
        self.var_func_tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Rechtes Panel: Hilfe
        self.right_panel.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True, padx=(5, 0))
        self.help_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        self.help_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Buttons
        self.button_frame.pack(fill=tk.X)
        self._layout_buttons()
    
    def _layout_additional_widgets(self):
        """Überschreibbar für zusätzliche Widget-Layout in Subklassen"""
        pass
    
    def _layout_buttons(self):
        """Layoutet die Buttons - kann überschrieben werden"""
        self.cancel_button.pack(side=tk.RIGHT, padx=(5, 0))
        self.ok_button.pack(side=tk.RIGHT)
    
    def _load_variables_and_functions(self):
        """Lädt verfügbare Variablen und Funktionen"""
        # Variablen-Hauptknoten
        var_root = self.var_func_tree.insert("", "end", text="Variablen", open=True, tags=("category",))
        
        # Standard-Variablen
        std_node = self.var_func_tree.insert(var_root, "end", text="Standard", open=True, tags=("category",))
        std_vars = ["Date", "Time", "DateTime", "Year", "Month", "Day", 
                   "Hour", "Minute", "Second", "Weekday", "WeekNumber"]
        for var in std_vars:
            self.var_func_tree.insert(std_node, "end", text=var, tags=("variable",))
        
        # Datei-Variablen
        file_node = self.var_func_tree.insert(var_root, "end", text="Datei", open=True, tags=("category",))
        file_vars = ["FileName", "FileExtension", "FullFileName", "FilePath", 
                    "FullPath", "FileSize"]
        for var in file_vars:
            self.var_func_tree.insert(file_node, "end", text=var, tags=("variable",))
        
        # OCR-Variablen
        ocr_node = self.var_func_tree.insert(var_root, "end", text="OCR", open=True, tags=("category",))
        self.var_func_tree.insert(ocr_node, "end", text="OCR_FullText", tags=("variable",))
        
        # OCR-Zonen vom Hotfolder
        if self.ocr_zones:
            zones_node = self.var_func_tree.insert(ocr_node, "end", text="Definierte Zonen", open=True, tags=("category",))
            for i, zone_info in enumerate(self.ocr_zones):
                zone_name = zone_info.get('name', f'Zone_{i+1}')
                self.var_func_tree.insert(zones_node, "end", text=zone_name, tags=("variable",))
        
        # XML-Variablen (falls verfügbar)
        xml_node = self.var_func_tree.insert(var_root, "end", text="XML", open=False, tags=("category",))
        xml_vars = ["XML_Beispielfeld1", "XML_Beispielfeld2"]  # Beispiel-Felder
        for var in xml_vars:
            self.var_func_tree.insert(xml_node, "end", text=var, tags=("variable",))
        
        # Funktionen-Hauptknoten
        func_root = self.var_func_tree.insert("", "end", text="Funktionen", open=True, tags=("category",))
        
        # String-Funktionen
        string_node = self.var_func_tree.insert(func_root, "end", text="String", open=True, tags=("category",))
        string_funcs = [
            ("FORMAT", 'FORMAT("<VAR>","<FORMATSTRING>")'),
            ("TRIM", 'TRIM("<VAR>")'),
            ("LEFT", 'LEFT("<VAR>",<LENGTH>)'),
            ("RIGHT", 'RIGHT("<VAR>",<LENGTH>)'),
            ("MID", 'MID("<VAR>",<START>,<LENGTH>)'),
            ("TOUPPER", 'TOUPPER("<VAR>")'),
            ("TOLOWER", 'TOLOWER("<VAR>")'),
            ("LEN", 'LEN("<VAR>")'),
            ("INDEXOF", 'INDEXOF(<START>,"<STRING>","<SEARCH>",<CASE>)')
        ]
        for name, syntax in string_funcs:
            self.var_func_tree.insert(string_node, "end", text=name, 
                                     tags=("function", syntax))
        
        # Datumsfunktionen
        date_node = self.var_func_tree.insert(func_root, "end", text="Datum", open=True, tags=("category",))
        self.var_func_tree.insert(date_node, "end", text="FORMATDATE", 
                                 tags=("function", 'FORMATDATE("<FORMAT>")'))
        
        # Bedingungen
        cond_node = self.var_func_tree.insert(func_root, "end", text="Bedingungen", open=True, tags=("category",))
        self.var_func_tree.insert(cond_node, "end", text="IF", 
                                 tags=("function", 'IF("<VAR>","<OP>","<VALUE>","<TRUE>","<FALSE>")'))
        
        # RegEx
        regex_node = self.var_func_tree.insert(func_root, "end", text="Reguläre Ausdrücke", open=True, tags=("category",))
        regex_funcs = [
            ("REGEXP.MATCH", 'REGEXP.MATCH("<VAR>","<PATTERN>",<INDEX>)'),
            ("REGEXP.REPLACE", 'REGEXP.REPLACE("<VAR>","<PATTERN>","<REPLACE>")')
        ]
        for name, syntax in regex_funcs:
            self.var_func_tree.insert(regex_node, "end", text=name, 
                                     tags=("function", syntax))
    
    def _on_tree_double_click(self, event):
        """Doppelklick fügt Variable oder Funktion ein"""
        selection = self.var_func_tree.selection()
        if selection:
            item = self.var_func_tree.item(selection[0])
            tags = item.get('tags', [])
            
            if tags and tags[0] == "variable":
                # Variable einfügen
                self.expr_text.insert(tk.INSERT, f"<{item['text']}>")
            elif tags and tags[0] == "function" and len(tags) > 1:
                # Funktion einfügen
                self.expr_text.insert(tk.INSERT, tags[1])
    
    def _on_tree_selection_changed(self, event):
        """Wird aufgerufen wenn die Tree-Auswahl sich ändert"""
        selection = self.var_func_tree.selection()
        if selection:
            item = self.var_func_tree.item(selection[0])
            item_text = item['text']
            tags = item.get('tags', [])
            
            # Zeige entsprechende Hilfe
            if tags and len(tags) > 0:
                self._show_help_for_item(item_text, tags)
    
    def _show_help_for_item(self, item_text: str, tags: list):
        """Zeigt Hilfe für das ausgewählte Element"""
        help_text = ""
        
        if tags[0] == "category":
            help_text = self.help_content.get("categories", {}).get(item_text, "")
        elif tags[0] == "variable":
            help_text = self.help_content.get("variables", {}).get(item_text, "")
        elif tags[0] == "function":
            help_text = self.help_content.get("functions", {}).get(item_text, "")
        
        # Aktualisiere Hilfe-Text
        self.help_text.config(state=tk.NORMAL)
        self.help_text.delete("1.0", tk.END)
        self.help_text.insert("1.0", help_text)
        self.help_text.config(state=tk.DISABLED)
    
    def _show_initial_help(self):
        """Zeigt die initiale Hilfe"""
        initial_help = self.help_content.get("initial", "")
        self.help_text.config(state=tk.NORMAL)
        self.help_text.delete("1.0", tk.END)
        self.help_text.insert("1.0", initial_help)
        self.help_text.config(state=tk.DISABLED)
    
    def _create_help_content(self):
        """Erstellt den Hilfe-Inhalt"""
        return {
            "initial": """WILLKOMMEN ZUM AUSDRUCK-EDITOR

Hier können Sie dynamische Ausdrücke erstellen.

GRUNDLAGEN:
• Variablen werden in spitze Klammern gesetzt: <VariablenName>
• Funktionen haben Parameter: FUNKTION("Parameter")
• Kombinieren Sie beides für komplexe Ausdrücke

BEDIENUNG:
• Wählen Sie links eine Variable oder Funktion aus
• Doppelklicken Sie, um sie einzufügen
• Diese Hilfe zeigt Details zum ausgewählten Element

TIPP: Beginnen Sie mit einfachen Variablen wie <FileName> und erweitern Sie schrittweise.""",
            
            "categories": {
                "Variablen": """VARIABLEN

Variablen enthalten dynamische Werte, die zur Laufzeit ersetzt werden.

Syntax: <VariablenName>

Beispiel: <FileName> wird durch den aktuellen Dateinamen ersetzt.

Kategorien:
• Standard: Datum, Zeit, etc.
• Datei: Informationen zur aktuellen Datei
• OCR: Erkannter Text aus der PDF
• XML: Felder aus der XML-Datei""",
                
                "Funktionen": """FUNKTIONEN

Funktionen verarbeiten und transformieren Werte.

Syntax: FUNKTION("Parameter")

Beispiel: TOUPPER(<FileName>) wandelt den Dateinamen in Großbuchstaben um.

Kategorien:
• String: Textverarbeitung
• Datum: Datumsformatierung
• Bedingungen: Wenn-Dann-Logik
• RegEx: Mustersuche und -ersetzung""",
                
                "Standard": """STANDARD-VARIABLEN

Vordefinierte Werte für Datum, Zeit und System-Informationen.

Verfügbar ohne weitere Konfiguration.
Werden automatisch mit aktuellen Werten gefüllt.""",
                
                "Datei": """DATEI-VARIABLEN

Informationen über die aktuell verarbeitete PDF-Datei.

Beispiele:
• <FileName>: "Rechnung_2024_001"
• <FileSize>: Größe in Bytes
• <FilePath>: Ordnerpfad der Datei""",
                
                "OCR": """OCR-VARIABLEN

Text, der aus der PDF mittels OCR erkannt wurde.

<OCR_FullText>: Kompletter erkannter Text

OCR-Zonen werden im Hotfolder-Dialog definiert.""",
                
                "Definierte Zonen": """DEFINIERTE OCR-ZONEN

Im Hotfolder wurden OCR-Zonen definiert.

Diese können als Variablen verwendet werden:
<ZonenName>

Die Zonen enthalten den erkannten Text aus den definierten Bereichen der PDF.""",
                
                "XML": """XML-VARIABLEN

Felder aus der zugehörigen XML-Datei.

Syntax: <XML_Feldname>

Nur verfügbar wenn eine XML-Datei vorhanden ist und das entsprechende Feld existiert.""",
                
                "String": """STRING-FUNKTIONEN

Funktionen zur Textverarbeitung und -manipulation.

Ermöglichen das Formatieren, Kürzen, Umwandeln und Durchsuchen von Text.

Doppelklicken Sie auf eine Funktion um die Syntax zu sehen.""",
                
                "Datum": """DATUM-FUNKTIONEN

Funktionen zur Formatierung von Datum und Zeit.

FORMATDATE ermöglicht benutzerdefinierte Datumsformate.""",
                
                "Bedingungen": """BEDINGUNGS-FUNKTIONEN

Ermöglichen Wenn-Dann-Logik in Ausdrücken.

IF-Funktion prüft Bedingungen und gibt entsprechende Werte zurück.""",
                
                "Reguläre Ausdrücke": """REGULÄRE AUSDRÜCKE

Erweiterte Mustersuche und -ersetzung in Text.

REGEXP.MATCH: Findet Muster
REGEXP.REPLACE: Ersetzt Muster

Erfordert Kenntnisse regulärer Ausdrücke."""
            },
            
            "variables": {
                "Date": """VARIABLE: Date

Aktuelles Datum im Format dd.mm.yyyy

Beispiel: 20.06.2025

Verwendung: <Date>

Für andere Formate verwenden Sie FORMATDATE()""",
                
                "Time": """VARIABLE: Time

Aktuelle Zeit im Format hh:mm:ss

Beispiel: 14:30:25

Verwendung: <Time>""",
                
                "DateTime": """VARIABLE: DateTime

Aktuelles Datum und Zeit

Beispiel: 20.06.2025 14:30:25

Verwendung: <DateTime>""",
                
                "Year": """VARIABLE: Year

Aktuelles Jahr vierstellig

Beispiel: 2025

Verwendung: <Year>""",
                
                "Month": """VARIABLE: Month

Aktueller Monat zweistellig

Beispiel: 06

Verwendung: <Month>""",
                
                "Day": """VARIABLE: Day

Aktueller Tag zweistellig

Beispiel: 20

Verwendung: <Day>""",
                
                "Hour": """VARIABLE: Hour

Aktuelle Stunde zweistellig (24h)

Beispiel: 14

Verwendung: <Hour>""",
                
                "Minute": """VARIABLE: Minute

Aktuelle Minute zweistellig

Beispiel: 30

Verwendung: <Minute>""",
                
                "Second": """VARIABLE: Second

Aktuelle Sekunde zweistellig

Beispiel: 25

Verwendung: <Second>""",
                
                "Weekday": """VARIABLE: Weekday

Name des aktuellen Wochentags

Beispiel: Freitag

Verwendung: <Weekday>""",
                
                "WeekNumber": """VARIABLE: WeekNumber

Aktuelle Kalenderwoche

Beispiel: 25

Verwendung: <WeekNumber>""",
                
                "FileName": """VARIABLE: FileName

Dateiname ohne Erweiterung

Beispiel: "Rechnung_2024_001"
(für Datei "Rechnung_2024_001.pdf")

Verwendung: <FileName>

Häufig verwendete Variable für Dateibenennungen.""",
                
                "FileExtension": """VARIABLE: FileExtension

Dateierweiterung mit Punkt

Beispiel: ".pdf"

Verwendung: <FileExtension>""",
                
                "FullFileName": """VARIABLE: FullFileName

Vollständiger Dateiname mit Erweiterung

Beispiel: "Rechnung_2024_001.pdf"

Verwendung: <FullFileName>""",
                
                "FilePath": """VARIABLE: FilePath

Ordnerpfad der Datei ohne Dateiname

Beispiel: "C:\\Input\\Rechnungen"

Verwendung: <FilePath>""",
                
                "FullPath": """VARIABLE: FullPath

Vollständiger Pfad inklusive Dateiname

Beispiel: "C:\\Input\\Rechnungen\\Rechnung_2024_001.pdf"

Verwendung: <FullPath>""",
                
                "FileSize": """VARIABLE: FileSize

Dateigröße in Bytes

Beispiel: "1048576"

Verwendung: <FileSize>

Kann in Bedingungen verwendet werden.""",
                
                "OCR_FullText": """VARIABLE: OCR_FullText

Kompletter Text, der aus der PDF erkannt wurde

Enthält den gesamten OCR-Text aller Seiten.

Verwendung: <OCR_FullText>

Kann mit REGEXP.MATCH durchsucht werden."""
            },
            
            "functions": {
                "FORMAT": """FUNKTION: FORMAT

Formatiert eine Variable nach einem Muster.

Syntax: FORMAT("<Variable>","<Formatstring>")

Parameter:
• Variable: Zu formatierende Variable
• Formatstring: Formatierungsmuster (z.B. "####" für führende Nullen)

Beispiel:
FORMAT("<Month>","##") → "06" (für Monat 6)

Nützlich für einheitliche Zahlenformate.""",
                
                "TRIM": """FUNKTION: TRIM

Entfernt Leerzeichen am Anfang und Ende.

Syntax: TRIM("<Variable>")

Parameter:
• Variable: Text-Variable

Beispiel:
TRIM("  Hallo  ") → "Hallo"

Besonders nützlich bei OCR-Text, der oft Leerzeichen enthält.""",
                
                "LEFT": """FUNKTION: LEFT

Gibt die ersten n Zeichen zurück.

Syntax: LEFT("<Variable>",<Anzahl>)

Parameter:
• Variable: Text-Variable
• Anzahl: Anzahl Zeichen

Beispiel:
LEFT("<FileName>",8) → "Rechnung" (für "Rechnung_001")""",
                
                "RIGHT": """FUNKTION: RIGHT

Gibt die letzten n Zeichen zurück.

Syntax: RIGHT("<Variable>",<Anzahl>)

Parameter:
• Variable: Text-Variable
• Anzahl: Anzahl Zeichen

Beispiel:
RIGHT("<FileName>",3) → "001" (für "Rechnung_001")""",
                
                "MID": """FUNKTION: MID

Gibt Zeichen aus der Mitte zurück.

Syntax: MID("<Variable>",<Start>,<Länge>)

Parameter:
• Variable: Text-Variable
• Start: Startposition (1-basiert)
• Länge: Anzahl Zeichen

Beispiel:
MID("<FileName>",10,3) → "001" (ab Position 10, 3 Zeichen)""",
                
                "TOUPPER": """FUNKTION: TOUPPER

Wandelt Text in Großbuchstaben um.

Syntax: TOUPPER("<Variable>")

Parameter:
• Variable: Text-Variable

Beispiel:
TOUPPER("<FileName>") → "RECHNUNG_001"

Nützlich für einheitliche Formatierung.""",
                
                "TOLOWER": """FUNKTION: TOLOWER

Wandelt Text in Kleinbuchstaben um.

Syntax: TOLOWER("<Variable>")

Parameter:
• Variable: Text-Variable

Beispiel:
TOLOWER("<FileName>") → "rechnung_001\"""",
                
                "LEN": """FUNKTION: LEN

Gibt die Länge des Texts zurück.

Syntax: LEN("<Variable>")

Parameter:
• Variable: Text-Variable

Beispiel:
LEN("<FileName>") → "13" (für "Rechnung_001")

Kann in Bedingungen verwendet werden.""",
                
                "INDEXOF": """FUNKTION: INDEXOF

Findet die Position eines Teilstrings.

Syntax: INDEXOF(<Start>,"<Text>","<Suchtext>",<CaseSensitive>)

Parameter:
• Start: Startposition der Suche
• Text: Zu durchsuchender Text
• Suchtext: Gesuchter Text
• CaseSensitive: "true" oder "false"

Beispiel:
INDEXOF(1,"<FileName>","_","true") → Position des ersten Unterstrichs""",
                
                "FORMATDATE": """FUNKTION: FORMATDATE

Formatiert das aktuelle Datum.

Syntax: FORMATDATE("<Format>")

Parameter:
• Format: Formatstring

Formatcodes:
• dd: Tag (01-31)
• mm: Monat (01-12)
• yyyy: Jahr (2025)
• hh: Stunde (00-23)
• MM: Minute (00-59)
• ss: Sekunde (00-59)

Beispiel:
FORMATDATE("yyyy-mm-dd") → "2025-06-20\"""",
                
                "IF": """FUNKTION: IF

Bedingte Auswertung (Wenn-Dann-Sonst).

Syntax: IF("<Variable>","<Operator>","<Wert>","<Dann>","<Sonst>")

Parameter:
• Variable: Zu prüfende Variable
• Operator: Vergleichsoperator (=, !=, >, <, >=, <=, contains, startswith, endswith)
• Wert: Vergleichswert
• Dann: Rückgabe wenn Bedingung wahr
• Sonst: Rückgabe wenn Bedingung falsch

Beispiel:
IF("<FileSize>",">","1000000","Große Datei","Kleine Datei")""",
                
                "REGEXP.MATCH": """FUNKTION: REGEXP.MATCH

Findet Text mit regulären Ausdrücken.

Syntax: REGEXP.MATCH("<Variable>","<Muster>",<Index>)

Parameter:
• Variable: Zu durchsuchender Text
• Muster: Regulärer Ausdruck
• Index: Index der Übereinstimmung (0 = erste)

Beispiel:
REGEXP.MATCH("<OCR_FullText>","\\\\d{4}-\\\\d{2}-\\\\d{2}",0)
→ Findet Datum im Format YYYY-MM-DD

Erfordert Kenntnisse regulärer Ausdrücke.""",
                
                "REGEXP.REPLACE": """FUNKTION: REGEXP.REPLACE

Ersetzt Text mit regulären Ausdrücken.

Syntax: REGEXP.REPLACE("<Variable>","<Muster>","<Ersatz>")

Parameter:
• Variable: Zu bearbeitender Text
• Muster: Regulärer Ausdruck
• Ersatz: Ersetzungstext

Beispiel:
REGEXP.REPLACE("<OCR_FullText>","\\\\s+","_")
→ Ersetzt alle Leerzeichen durch Unterstriche

Erfordert Kenntnisse regulärer Ausdrücke."""
            }
        }
    
    def _on_ok(self):
        """Standard OK-Handler - kann überschrieben werden"""
        self.result = self.expr_text.get("1.0", tk.END).strip()
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Standard Cancel-Handler"""
        self.dialog.destroy()
    
    def show(self) -> Optional[str]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result