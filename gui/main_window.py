"""
Hauptfenster der Hotfolder-Anwendung
"""
import tkinter as tk
from tkinter import ttk, messagebox
import sys
import os
from typing import Optional

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from gui.hotfolder_dialog import HotfolderDialog
from core.hotfolder_manager import HotfolderManager
from models.hotfolder_config import ProcessingAction


class MainWindow:
    """Hauptfenster der Anwendung"""
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Hotfolder PDF Processor")
        self.root.geometry("900x600")
        
        # Manager
        self.manager = HotfolderManager()
        
        # Style
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Erstelle UI
        self._create_menu()
        self._create_widgets()
        self._layout_widgets()
        
        # Lade Hotfolder
        self._refresh_list()
        
        # Starte Manager
        self.manager.start()
        
        # Bind Events
        self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
        
        # Statusbar Update
        self._update_status()
    
    def _create_menu(self):
        """Erstellt die Menüleiste"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Datei-Menü
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Datei", menu=file_menu)
        file_menu.add_command(label="Neuer Hotfolder", command=self._new_hotfolder,
                             accelerator="Ctrl+N")
        file_menu.add_separator()
        file_menu.add_command(label="Beenden", command=self._on_closing,
                             accelerator="Alt+F4")
        
        # Bearbeiten-Menü
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Bearbeiten", menu=edit_menu)
        edit_menu.add_command(label="Hotfolder bearbeiten", command=self._edit_hotfolder,
                             accelerator="F2")
        edit_menu.add_command(label="Hotfolder löschen", command=self._delete_hotfolder,
                             accelerator="Del")
        
        # Hilfe-Menü
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Hilfe", menu=help_menu)
        help_menu.add_command(label="Über", command=self._show_about)
        
        # Keyboard Shortcuts
        self.root.bind("<Control-n>", lambda e: self._new_hotfolder())
        self.root.bind("<F2>", lambda e: self._edit_hotfolder())
        self.root.bind("<Delete>", lambda e: self._delete_hotfolder())
    
    def _create_widgets(self):
        """Erstellt alle Widgets"""
        # Toolbar
        self.toolbar = ttk.Frame(self.root)
        self.new_button = ttk.Button(self.toolbar, text="➕ Neuer Hotfolder", 
                                    command=self._new_hotfolder)
        self.edit_button = ttk.Button(self.toolbar, text="✏️ Bearbeiten", 
                                     command=self._edit_hotfolder, state=tk.DISABLED)
        self.delete_button = ttk.Button(self.toolbar, text="🗑️ Löschen", 
                                       command=self._delete_hotfolder, state=tk.DISABLED)
        self.refresh_button = ttk.Button(self.toolbar, text="🔄 Aktualisieren", 
                                        command=self._refresh_list)
        
        # Hauptbereich
        self.main_frame = ttk.Frame(self.root)
        
        # Treeview für Hotfolder-Liste
        self.tree_frame = ttk.Frame(self.main_frame)
        self.tree = ttk.Treeview(self.tree_frame, columns=("Name", "Input", "Output", 
                                                          "Aktionen", "Status"),
                                show="tree headings")
        
        # Spalten konfigurieren
        self.tree.heading("#0", text="")
        self.tree.heading("Name", text="Name")
        self.tree.heading("Input", text="Input-Ordner")
        self.tree.heading("Output", text="Output-Ordner")
        self.tree.heading("Aktionen", text="Aktionen")
        self.tree.heading("Status", text="Status")
        
        self.tree.column("#0", width=50)
        self.tree.column("Name", width=150)
        self.tree.column("Input", width=200)
        self.tree.column("Output", width=200)
        self.tree.column("Aktionen", width=150)
        self.tree.column("Status", width=100)
        
        # Scrollbars
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.hsb = ttk.Scrollbar(self.tree_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=self.vsb.set, xscrollcommand=self.hsb.set)
        
        # Kontextmenü
        self.context_menu = tk.Menu(self.root, tearoff=0)
        self.context_menu.add_command(label="Bearbeiten", command=self._edit_hotfolder)
        self.context_menu.add_command(label="Löschen", command=self._delete_hotfolder)
        self.context_menu.add_separator()
        self.context_menu.add_command(label="Aktivieren/Deaktivieren", 
                                     command=self._toggle_hotfolder)
        
        # Statusbar
        self.statusbar = ttk.Frame(self.root)
        self.status_label = ttk.Label(self.statusbar, text="Bereit")
        
        # Events
        self.tree.bind("<Double-Button-1>", lambda e: self._edit_hotfolder())
        self.tree.bind("<Button-3>", self._show_context_menu)
        self.tree.bind("<<TreeviewSelect>>", self._on_selection_changed)
    
    def _layout_widgets(self):
        """Layoutet alle Widgets"""
        # Toolbar
        self.toolbar.pack(side=tk.TOP, fill=tk.X, padx=5, pady=5)
        self.new_button.pack(side=tk.LEFT, padx=(0, 5))
        self.edit_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_button.pack(side=tk.LEFT, padx=(0, 5))
        self.refresh_button.pack(side=tk.RIGHT)
        
        # Hauptbereich
        self.main_frame.pack(fill=tk.BOTH, expand=True, padx=5, pady=(0, 5))
        
        # Treeview
        self.tree_frame.pack(fill=tk.BOTH, expand=True)
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.hsb.grid(row=1, column=0, sticky="ew")
        
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.tree_frame.grid_columnconfigure(0, weight=1)
        
        # Statusbar
        self.statusbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.status_label.pack(side=tk.LEFT, padx=5)
    
    def _refresh_list(self):
        """Aktualisiert die Hotfolder-Liste"""
        # Lösche aktuelle Einträge
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Füge Hotfolder hinzu
        for hotfolder in self.manager.get_hotfolders():
            # Checkbox-Symbol
            checkbox = "☑" if hotfolder.enabled else "☐"
            
            # Aktionen als Text
            action_names = {
                ProcessingAction.COMPRESS: "Komprimieren",
                ProcessingAction.SPLIT: "Aufteilen",
                ProcessingAction.OCR: "OCR",
                ProcessingAction.PDF_A: "PDF/A"
            }
            
            actions = ", ".join([action_names.get(a, a.value) for a in hotfolder.actions])
            
            # Status
            status = "Aktiv" if hotfolder.enabled else "Inaktiv"
            
            # Füge zur Tabelle hinzu
            item = self.tree.insert("", "end", text=checkbox,
                                   values=(hotfolder.name, 
                                          hotfolder.input_path,
                                          hotfolder.output_path,
                                          actions,
                                          status),
                                   tags=(hotfolder.id,))
            
            # Färbe inaktive Einträge grau
            if not hotfolder.enabled:
                self.tree.item(item, tags=(hotfolder.id, "disabled"))
        
        # Style für disabled
        self.tree.tag_configure("disabled", foreground="gray")
    
    def _on_selection_changed(self, event):
        """Wird aufgerufen wenn die Auswahl sich ändert"""
        selection = self.tree.selection()
        
        if selection:
            self.edit_button.config(state=tk.NORMAL)
            self.delete_button.config(state=tk.NORMAL)
        else:
            self.edit_button.config(state=tk.DISABLED)
            self.delete_button.config(state=tk.DISABLED)
    
    def _get_selected_hotfolder_id(self) -> Optional[str]:
        """Gibt die ID des ausgewählten Hotfolders zurück"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            tags = self.tree.item(item)["tags"]
            if tags:
                return tags[0]  # Erste Tag ist die ID
        return None
    
    def _new_hotfolder(self):
        """Erstellt einen neuen Hotfolder"""
        dialog = HotfolderDialog(self.root)
        result = dialog.show()
        
        if result:
            success, message = self.manager.create_hotfolder(
                name=result["name"],
                input_path=result["input_path"],
                output_path=result["output_path"],
                actions=result["actions"],
                action_params=result["action_params"],
                xml_field_mappings=result.get("xml_field_mappings", []),
                output_filename_expression=result.get("output_filename_expression", "<FileName>"),
                ocr_zones=result.get("ocr_zones", [])
            )
            
            if success:
                self._refresh_list()
                self._update_status("Hotfolder erstellt")
            else:
                messagebox.showerror("Fehler", message)
    
    def _edit_hotfolder(self):
        """Bearbeitet den ausgewählten Hotfolder"""
        hotfolder_id = self._get_selected_hotfolder_id()
        if not hotfolder_id:
            return
        
        hotfolder = self.manager.get_hotfolder(hotfolder_id)
        if not hotfolder:
            return
        
        dialog = HotfolderDialog(self.root, hotfolder)
        result = dialog.show()
        
        if result:
            success, message = self.manager.update_hotfolder(
                hotfolder_id,
                name=result["name"],
                input_path=result["input_path"],
                output_path=result["output_path"],
                process_pairs=result["process_pairs"],
                actions=result["actions"],
                action_params=result["action_params"],
                xml_field_mappings=result.get("xml_field_mappings", []),
                output_filename_expression=result.get("output_filename_expression", "<FileName>"),
                ocr_zones=result.get("ocr_zones", [])
            )
            
            if success:
                self._refresh_list()
                self._update_status("Hotfolder aktualisiert")
            else:
                messagebox.showerror("Fehler", message)
    
    def _delete_hotfolder(self):
        """Löscht den ausgewählten Hotfolder"""
        hotfolder_id = self._get_selected_hotfolder_id()
        if not hotfolder_id:
            return
        
        hotfolder = self.manager.get_hotfolder(hotfolder_id)
        if not hotfolder:
            return
        
        # Bestätigung
        result = messagebox.askyesno(
            "Hotfolder löschen",
            f"Möchten Sie den Hotfolder '{hotfolder.name}' wirklich löschen?"
        )
        
        if result:
            success, message = self.manager.delete_hotfolder(hotfolder_id)
            
            if success:
                self._refresh_list()
                self._update_status("Hotfolder gelöscht")
            else:
                messagebox.showerror("Fehler", message)
    
    def _toggle_hotfolder(self):
        """Aktiviert/Deaktiviert den ausgewählten Hotfolder"""
        hotfolder_id = self._get_selected_hotfolder_id()
        if not hotfolder_id:
            return
        
        hotfolder = self.manager.get_hotfolder(hotfolder_id)
        if not hotfolder:
            return
        
        new_state = not hotfolder.enabled
        success, message = self.manager.toggle_hotfolder(hotfolder_id, new_state)
        
        if success:
            self._refresh_list()
            self._update_status(message)
        else:
            messagebox.showerror("Fehler", message)
    
    def _show_context_menu(self, event):
        """Zeigt das Kontextmenü"""
        # Wähle Item unter Mauszeiger
        item = self.tree.identify_row(event.y)
        if item:
            self.tree.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)
    
    def _show_about(self):
        """Zeigt den Über-Dialog"""
        messagebox.showinfo(
            "Über Hotfolder PDF Processor",
            "Hotfolder PDF Processor v1.0\n\n"
            "Ein skalierbares Tool zur automatischen PDF-Verarbeitung.\n\n"
            "Unterstützt:\n"
            "• Mehrere Hotfolder\n"
            "• Verschiedene PDF-Bearbeitungen\n"
            "• PDF-XML Dokumentenpaare\n"
            "• Automatische Ordnerüberwachung"
        )
    
    def _update_status(self, message: str = None):
        """Aktualisiert die Statusleiste"""
        if message:
            self.status_label.config(text=message)
        else:
            active_count = len([h for h in self.manager.get_hotfolders() if h.enabled])
            total_count = len(self.manager.get_hotfolders())
            self.status_label.config(text=f"{active_count} von {total_count} Hotfoldern aktiv")
        
        # Reset nach 3 Sekunden
        if message:
            self.root.after(3000, lambda: self._update_status())
    
    def _on_closing(self):
        """Wird beim Schließen des Fensters aufgerufen"""
        # Stoppe Manager
        self.manager.stop()
        
        # Schließe Fenster
        self.root.destroy()
    
    def run(self):
        """Startet die Anwendung"""
        self.root.mainloop()