"""
Dialog zum Erstellen und Bearbeiten von Hotfoldern
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from typing import Optional, List, Dict
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.hotfolder_config import HotfolderConfig, ProcessingAction
from gui.xml_field_dialog import XMLFieldDialog


class HotfolderDialog:
    """Dialog zum Erstellen/Bearbeiten eines Hotfolders"""
    
    def __init__(self, parent, hotfolder: Optional[HotfolderConfig] = None):
        self.parent = parent
        self.hotfolder = hotfolder
        self.result = None
        
        # Erstelle Dialog
        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Hotfolder bearbeiten" if hotfolder else "Neuer Hotfolder")
        self.dialog.geometry("600x700")
        self.dialog.resizable(True, True)
        
        # Zentriere Dialog
        self.dialog.transient(parent)
        self.dialog.grab_set()
        
        # Variablen
        self.name_var = tk.StringVar(value=hotfolder.name if hotfolder else "")
        self.input_path_var = tk.StringVar(value=hotfolder.input_path if hotfolder else "")
        self.output_path_var = tk.StringVar(value=hotfolder.output_path if hotfolder else "")
        self.process_pairs_var = tk.BooleanVar(value=hotfolder.process_pairs if hotfolder else True)
        
        # Action Variablen
        self.action_vars = {}
        for action in ProcessingAction:
            is_selected = hotfolder and action in hotfolder.actions
            self.action_vars[action] = tk.BooleanVar(value=is_selected)
        
        # Action Parameter
        self.action_params = hotfolder.action_params.copy() if hotfolder else {}
        
        # XML-Feld-Mappings
        self.xml_field_mappings = hotfolder.xml_field_mappings.copy() if hotfolder else []
        
        self._create_widgets()
        self._layout_widgets()
        
        # Initialisiere XML-Felder-Button Text wenn Felder vorhanden
        if self.xml_field_mappings:
            count = len(self.xml_field_mappings)
            self.xml_fields_button.config(text=f"XML-Felder konfigurieren... ({count} Felder)")
        
        # Fokus auf erstes Eingabefeld
        self.name_entry.focus()
        
        # Bind Enter-Taste
        self.dialog.bind('<Return>', lambda e: self._on_save())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Hauptframe
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        
        # Name
        self.name_label = ttk.Label(self.main_frame, text="Name:")
        self.name_entry = ttk.Entry(self.main_frame, textvariable=self.name_var, width=50)
        
        # Input-Pfad
        self.input_label = ttk.Label(self.main_frame, text="Input-Ordner:")
        self.input_frame = ttk.Frame(self.main_frame)
        self.input_entry = ttk.Entry(self.input_frame, textvariable=self.input_path_var, width=40)
        self.input_button = ttk.Button(self.input_frame, text="Durchsuchen...", 
                                      command=self._browse_input)
        
        # Output-Pfad
        self.output_label = ttk.Label(self.main_frame, text="Output-Ordner:")
        self.output_frame = ttk.Frame(self.main_frame)
        self.output_entry = ttk.Entry(self.output_frame, textvariable=self.output_path_var, width=40)
        self.output_button = ttk.Button(self.output_frame, text="Durchsuchen...", 
                                       command=self._browse_output)
        
        # Optionen
        self.options_frame = ttk.LabelFrame(self.main_frame, text="Optionen", padding="10")
        self.process_pairs_check = ttk.Checkbutton(
            self.options_frame,
            text="PDF-XML Paare verarbeiten",
            variable=self.process_pairs_var,
            command=self._on_process_pairs_toggle
        )
        
        # XML-Felder Button
        self.xml_fields_button = ttk.Button(
            self.options_frame,
            text="XML-Felder konfigurieren...",
            command=self._configure_xml_fields,
            state=tk.NORMAL if self.process_pairs_var.get() else tk.DISABLED
        )
        
        # Aktionen
        self.actions_frame = ttk.LabelFrame(self.main_frame, text="PDF-Bearbeitungsaktionen", 
                                           padding="10")
        
        self.action_checks = {}
        action_descriptions = {
            ProcessingAction.COMPRESS: "PDF komprimieren",
            ProcessingAction.SPLIT: "In Einzelseiten aufteilen",
            ProcessingAction.OCR: "OCR durchführen (durchsuchbare PDF)",
            ProcessingAction.PDF_A: "In PDF/A konvertieren"
        }
        
        for action, description in action_descriptions.items():
            frame = ttk.Frame(self.actions_frame)
            check = ttk.Checkbutton(frame, text=description, 
                                   variable=self.action_vars[action])
            self.action_checks[action] = check
            check.pack(anchor=tk.W)
            frame.pack(anchor=tk.W, pady=2, fill=tk.X)
        
        # Info-Label
        self.info_label = ttk.Label(self.main_frame, 
                                   text="Hinweis: Die ausgewählten Aktionen werden in der "
                                        "angegebenen Reihenfolge ausgeführt.",
                                   wraplength=550)
        
        # Buttons
        self.button_frame = ttk.Frame(self.main_frame)
        self.save_button = ttk.Button(self.button_frame, text="Speichern", 
                                     command=self._on_save, default=tk.ACTIVE)
        self.cancel_button = ttk.Button(self.button_frame, text="Abbrechen", 
                                       command=self._on_cancel)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Name
        self.name_label.grid(row=0, column=0, sticky=tk.W, pady=(0, 5))
        self.name_entry.grid(row=0, column=1, sticky="we", pady=(0, 5))
        
        # Input
        self.input_label.grid(row=1, column=0, sticky=tk.W, pady=(0, 5))
        self.input_frame.grid(row=1, column=1, sticky="we", pady=(0, 5))
        self.input_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.input_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Output
        self.output_label.grid(row=2, column=0, sticky=tk.W, pady=(0, 5))
        self.output_frame.grid(row=2, column=1, sticky="we", pady=(0, 5))
        self.output_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
        self.output_button.pack(side=tk.LEFT, padx=(5, 0))
        
        # Optionen
        self.options_frame.grid(row=3, column=0, columnspan=2, sticky="we", 
                               pady=(10, 5))
        self.process_pairs_check.pack(anchor=tk.W)
        self.xml_fields_button.pack(anchor=tk.W, pady=(5, 0))
        
        # Aktionen
        self.actions_frame.grid(row=4, column=0, columnspan=2, sticky="we", 
                               pady=(10, 5))
        
        # Info
        self.info_label.grid(row=5, column=0, columnspan=2, sticky=tk.W, pady=(10, 5))
        
        # Buttons
        self.button_frame.grid(row=6, column=0, columnspan=2, pady=(20, 0))
        self.save_button.pack(side=tk.LEFT, padx=(0, 5))
        self.cancel_button.pack(side=tk.LEFT)
        
        # Spalten-Konfiguration
        self.main_frame.columnconfigure(1, weight=1)
    
    def _browse_input(self):
        """Öffnet Dialog zur Auswahl des Input-Ordners"""
        folder = filedialog.askdirectory(
            title="Input-Ordner auswählen",
            initialdir=self.input_path_var.get() or os.path.expanduser("~")
        )
        if folder:
            self.input_path_var.set(folder)
    
    def _browse_output(self):
        """Öffnet Dialog zur Auswahl des Output-Ordners"""
        folder = filedialog.askdirectory(
            title="Output-Ordner auswählen",
            initialdir=self.output_path_var.get() or os.path.expanduser("~")
        )
        if folder:
            self.output_path_var.set(folder)
    
    def _validate(self) -> bool:
        """Validiert die Eingaben"""
        # Name prüfen
        if not self.name_var.get().strip():
            messagebox.showerror("Fehler", "Bitte geben Sie einen Namen ein.")
            self.name_entry.focus()
            return False
        
        # Input-Pfad prüfen
        if not self.input_path_var.get().strip():
            messagebox.showerror("Fehler", "Bitte wählen Sie einen Input-Ordner.")
            self.input_entry.focus()
            return False
        
        # Output-Pfad prüfen
        if not self.output_path_var.get().strip():
            messagebox.showerror("Fehler", "Bitte wählen Sie einen Output-Ordner.")
            self.output_entry.focus()
            return False
        
        # Pfade dürfen nicht identisch sein
        if os.path.abspath(self.input_path_var.get()) == os.path.abspath(self.output_path_var.get()):
            messagebox.showerror("Fehler", "Input- und Output-Ordner dürfen nicht identisch sein.")
            return False
        
        # Mindestens eine Aktion
        if not any(var.get() for var in self.action_vars.values()):
            messagebox.showerror("Fehler", "Bitte wählen Sie mindestens eine Bearbeitungsaktion.")
            return False
        
        return True
    
    def _on_save(self):
        """Speichert die Eingaben"""
        if not self._validate():
            return
        
        # Sammle ausgewählte Aktionen
        selected_actions = [action.value for action, var in self.action_vars.items() 
                           if var.get()]
        
        self.result = {
            "name": self.name_var.get().strip(),
            "input_path": self.input_path_var.get().strip(),
            "output_path": self.output_path_var.get().strip(),
            "process_pairs": self.process_pairs_var.get(),
            "actions": selected_actions,
            "action_params": self.action_params,
            "xml_field_mappings": self.xml_field_mappings
        }
        
        self.dialog.destroy()
    
    def _on_cancel(self):
        """Bricht den Dialog ab"""
        self.dialog.destroy()
    
    def _on_process_pairs_toggle(self):
        """Wird aufgerufen wenn Process Pairs Checkbox geändert wird"""
        if self.process_pairs_var.get():
            self.xml_fields_button.config(state=tk.NORMAL)
        else:
            self.xml_fields_button.config(state=tk.DISABLED)
    
    def _configure_xml_fields(self):
        """Öffnet Dialog zur Konfiguration der XML-Felder"""
        dialog = XMLFieldDialog(self.dialog, self.xml_field_mappings)
        result = dialog.show()
        
        if result is not None:
            self.xml_field_mappings = result
            count = len(self.xml_field_mappings)
            if count > 0:
                self.xml_fields_button.config(text=f"XML-Felder konfigurieren... ({count} Felder)")
            else:
                self.xml_fields_button.config(text="XML-Felder konfigurieren...")
    
    def show(self) -> Optional[Dict]:
        """Zeigt den Dialog und wartet auf Ergebnis"""
        self.dialog.wait_window()
        return self.result