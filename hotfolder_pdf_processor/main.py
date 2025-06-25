"""
Hotfolder PDF Processor
Hauptprogramm
"""
import sys
import os
import logging

# Füge das aktuelle Verzeichnis zum Python-Pfad hinzu
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.main_window import MainWindow
from core.logging_config import initialize_logging, cleanup_logging

# Logger für dieses Modul
logger = logging.getLogger(__name__)


def main():
    """Hauptfunktion der Anwendung"""
    # Initialisiere Logging
    initialize_logging()
    
    try:
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
