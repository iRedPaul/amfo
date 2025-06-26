import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional
import sys
import os

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from core.counter_manager import get_counter_manager
except ImportError:
    def get_counter_manager():
        return None


class CounterManagementDialog:
    def __init__(self, parent):
        self.parent = parent
        self.counter_manager = get_counter_manager()
        if self.counter_manager is None:
            messagebox.showerror("Fehler", "Counter-Manager nicht verfügbar!")
            return

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Auto-Increment Counter verwalten")
        self.dialog.geometry("600x500")
        self.dialog.resizable(True, True)
        self._center_window()
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self._create_widgets()
        self._layout_widgets()
        self._refresh_counter_list()
        self.dialog.bind('<Escape>', lambda e: self.dialog.destroy())

    def _center_window(self):
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() - width) // 2
        y = (self.dialog.winfo_screenheight() - height) // 2
        self.dialog.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        self.main_frame = ttk.Frame(self.dialog, padding="10")
        self.info_label = ttk.Label(
            self.main_frame, 
            text="Hier können Sie alle Auto-Increment Counter verwalten.",
            justify=tk.LEFT
        )
        self.toolbar = ttk.Frame(self.main_frame)
        self.refresh_button = ttk.Button(self.toolbar, text="🔄 Aktualisieren", command=self._refresh_counter_list)
        self.edit_button = ttk.Button(self.toolbar, text="✏️ Bearbeiten", command=self._edit_counter, state=tk.DISABLED)
        self.reset_button = ttk.Button(self.toolbar, text="🔄 Zurücksetzen", command=self._reset_counter, state=tk.DISABLED)
        self.delete_button = ttk.Button(self.toolbar, text="🗑️ Löschen", command=self._delete_counter, state=tk.DISABLED)

        self.tree_frame = ttk.Frame(self.main_frame)
        self.tree = ttk.Treeview(self.tree_frame, columns=("Name", "Wert"), show="tree headings", height=15)
        self.tree.heading("#0", text="")
        self.tree.heading("Name", text="Counter-Name")
        self.tree.heading("Wert", text="Nächster Wert")
        self.tree.column("#0", width=30)
        self.tree.column("Name", width=300)
        self.tree.column("Wert", width=120)
        self.vsb = ttk.Scrollbar(self.tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=self.vsb.set)

        self.button_frame = ttk.Frame(self.main_frame)
        self.clear_all_button = ttk.Button(self.button_frame, text="Alle löschen", command=self._clear_all_counters)
        self.close_button = ttk.Button(self.button_frame, text="Schließen", command=self.dialog.destroy)

        self.tree.bind("<<TreeviewSelect>>", self._on_selection_changed)
        self.tree.bind("<Double-Button-1>", lambda e: self._edit_counter())

    def _layout_widgets(self):
        self.main_frame.pack(fill=tk.BOTH, expand=True)
        self.info_label.pack(fill=tk.X, pady=(0, 10))
        self.toolbar.pack(fill=tk.X, pady=(0, 10))
        self.refresh_button.pack(side=tk.LEFT, padx=(0, 5))
        self.edit_button.pack(side=tk.LEFT, padx=(0, 5))
        self.reset_button.pack(side=tk.LEFT, padx=(0, 5))
        self.delete_button.pack(side=tk.LEFT)
        self.tree_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))
        self.tree.grid(row=0, column=0, sticky="nsew")
        self.vsb.grid(row=0, column=1, sticky="ns")
        self.tree_frame.grid_columnconfigure(0, weight=1)
        self.tree_frame.grid_rowconfigure(0, weight=1)
        self.button_frame.pack(fill=tk.X)
        self.clear_all_button.pack(side=tk.LEFT)
        self.close_button.pack(side=tk.RIGHT)

    def _refresh_counter_list(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        counters = self.counter_manager.list_counters()
        if not counters:
            self.tree.insert("", "end", text="ℹ️", values=("Keine Counter vorhanden", ""))
        else:
            for counter_name, value in sorted(counters.items()):
                display_name = counter_name.replace("auto_", "") if counter_name.startswith("auto_") else counter_name
                self.tree.insert("", "end", text="🔢", values=(display_name, str(value)))

    def _on_selection_changed(self, event):
        selection = self.tree.selection()
        state = tk.DISABLED
        if selection:
            item = self.tree.item(selection[0])
            values = item.get("values", [])
            if values and values[0] != "Keine Counter vorhanden":
                state = tk.NORMAL
        for button in [self.edit_button, self.reset_button, self.delete_button]:
            button.config(state=state)

    def _get_selected_counter_name(self) -> Optional[str]:
        selection = self.tree.selection()
        if not selection:
            return None
        item = self.tree.item(selection[0])
        values = item.get("values", [])
        if not values or values[0] == "Keine Counter vorhanden":
            return None
        return f"auto_{values[0]}"

    def _edit_counter(self):
        counter_name = self._get_selected_counter_name()
        if not counter_name:
            return
        current_value = self.counter_manager.get_counter(counter_name, 0)
        dialog = CounterEditDialog(self.dialog, counter_name, current_value)
        result = dialog.show()
        if result is not None:
            self.counter_manager.set_counter(counter_name, result)
            self._refresh_counter_list()

    def _reset_counter(self):
        counter_name = self._get_selected_counter_name()
        if not counter_name:
            return
        display_name = counter_name.replace("auto_", "")
        if messagebox.askyesno("Counter zurücksetzen", f"Möchten Sie den Counter '{display_name}' auf 1 zurücksetzen?"):
            self.counter_manager.reset_counter(counter_name, 1)
            self._refresh_counter_list()

    def _delete_counter(self):
        counter_name = self._get_selected_counter_name()
        if not counter_name:
            return
        display_name = counter_name.replace("auto_", "")
        if messagebox.askyesno("Counter löschen", f"Möchten Sie den Counter '{display_name}' wirklich löschen?"):
            self.counter_manager.delete_counter(counter_name)
            self._refresh_counter_list()

    def _clear_all_counters(self):
        counters = self.counter_manager.list_counters()
        if not counters:
            messagebox.showinfo("Keine Counter", "Es sind keine Counter vorhanden.")
            return
        if messagebox.askyesno("Alle Counter löschen", "Möchten Sie wirklich ALLE Counter löschen? Diese Aktion kann nicht rückgängig gemacht werden!"):
            self.counter_manager.clear_all_counters()
            self._refresh_counter_list()

    def show(self):
        self.dialog.wait_window()


class CounterEditDialog:
    def __init__(self, parent, counter_name: str, current_value: int):
        self.parent = parent
        self.counter_name = counter_name
        self.current_value = current_value
        self.result = None

        self.dialog = tk.Toplevel(parent)
        self.dialog.title("Counter bearbeiten")
        self.dialog.geometry("400x125")
        self.dialog.resizable(False, False)
        self._center_window()
        self.dialog.transient(parent)
        self.dialog.grab_set()

        self._create_widgets()
        self.value_entry.focus()
        self.value_entry.select_range(0, tk.END)
        self.dialog.bind('<Return>', lambda e: self._on_save())
        self.dialog.bind('<Escape>', lambda e: self._on_cancel())

    def _center_window(self):
        self.dialog.update_idletasks()
        width = self.dialog.winfo_width()
        height = self.dialog.winfo_height()
        x = (self.dialog.winfo_screenwidth() - width) // 2
        y = (self.dialog.winfo_screenheight() - height) // 2
        self.dialog.geometry(f"+{x}+{y}")

    def _create_widgets(self):
        self.main_frame = ttk.Frame(self.dialog, padding="20")
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        display_name = self.counter_name.replace("auto_", "")
        ttk.Label(self.main_frame, text=f"Counter '{display_name}' bearbeiten:").grid(row=0, column=0, columnspan=2, sticky="w", pady=(0, 15))

        ttk.Label(self.main_frame, text="Neuer Wert:").grid(row=1, column=0, sticky="w")
        self.value_var = tk.StringVar(value=str(self.current_value))
        self.value_entry = ttk.Entry(self.main_frame, textvariable=self.value_var, width=20)
        self.value_entry.grid(row=1, column=1, sticky="ew", pady=5)

        self.main_frame.columnconfigure(1, weight=1)

        button_frame = ttk.Frame(self.main_frame)
        button_frame.grid(row=3, column=0, columnspan=2, sticky="e")
        self.cancel_button = ttk.Button(button_frame, text="Abbrechen", command=self._on_cancel)
        self.cancel_button.pack(side=tk.RIGHT)
        self.save_button = ttk.Button(button_frame, text="Speichern", command=self._on_save)
        self.save_button.pack(side=tk.RIGHT, padx=(0, 5))

    def _on_save(self):
        try:
            new_value = int(self.value_var.get())
            self.result = new_value
            self.dialog.destroy()
        except ValueError:
            messagebox.showerror("Ungültiger Wert", "Bitte geben Sie eine gültige Zahl ein.")
            self.value_entry.focus()

    def _on_cancel(self):
        self.dialog.destroy()

    def show(self) -> Optional[int]:
        self.dialog.wait_window()
        return self.result