"""
belegpilot - Konfigurations-Tool
"""
import sys
import os
import logging
import tkinter as tk
from tkinter import messagebox

# Füge das aktuelle Verzeichnis zum Python-Pfad hinzu
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.main_window import MainWindow
from core.logging_config import initialize_logging, cleanup_logging
from core.license_manager import get_license_manager
from core.config_manager import ensure_config_directory

# Logger für dieses Modul
logger = logging.getLogger(__name__)


def show_license_warning():
    """Zeigt eine Warnung wenn keine Lizenz vorhanden ist"""
    try:
        license_manager = get_license_manager()
        valid, license_info, message = license_manager.validate_license()
        
        if not valid:
            root = tk.Tk()
            root.withdraw()
            
            messagebox.showwarning(
                "Keine gültige Lizenz",
                f"{message}\n\n"
                "Das Programm läuft im Demo-Modus.\n"
                "Sie können keine Hotfolder aktivieren.\n\n"
                "Bitte installieren Sie eine gültige Lizenz über die Einstellungen."
            )
            root.destroy()
            logger.warning("Keine gültige Lizenz - Demo-Modus aktiv")
        else:
            # Lizenz ist gültig - prüfe ob sie bald abläuft
            if "days_remaining" in license_info:
                days = license_info["days_remaining"]
                if days <= 7:  # Warnung bei weniger als 7 Tagen
                    root = tk.Tk()
                    root.withdraw()
                    
                    license_type = license_info.get("type", "").upper()
                    messagebox.showwarning(
                        "Lizenz läuft ab",
                        f"Ihre {license_type}-Lizenz läuft in {days} Tagen ab.\n\n"
                        "Bitte erneuern Sie rechtzeitig Ihre Lizenz."
                    )
                    root.destroy()
        
    except Exception as e:
        logger.exception("Fehler bei der Lizenzprüfung")


def main():
    """Hauptfunktion der Anwendung"""
    # Initialisiere Logging
    initialize_logging()
    
    try:
        # Stelle sicher, dass config Verzeichnis existiert
        ensure_config_directory()
        
        # Zeige Lizenzwarnung falls nötig
        show_license_warning()
        
        # WICHTIGE ÄNDERUNG:
        # Der HotfolderManager wird hier nicht mehr gestartet.
        # Die MainWindow dient nur noch der Konfiguration des Dienstes.
        
        # Erstelle und starte die Anwendung
        app = MainWindow()
        app.run()
    except Exception as e:
        logger.exception("Fehler beim Starten der Anwendung")
        # Für kritische Fehler beim Start einen Fallback verwenden
        try:
            import tkinter
            import tkinter.messagebox as messagebox
            root = tkinter.Tk()
            root.withdraw()  # Verstecke das Hauptfenster
            messagebox.showerror("Fehler", f"Fehler beim Starten der Anwendung:\n{e}")
        except:
            # Falls auch Tkinter fehlschlägt, als letzter Ausweg
            input(f"FEHLER: {e}\nDrücken Sie Enter zum Beenden...")
    finally:
        # Aufräumen
        cleanup_logging()


if __name__ == "__main__":
    main()