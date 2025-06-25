"""
Hotfolder PDF Processor
Hauptprogramm
"""
import sys
import os

# Füge das aktuelle Verzeichnis zum Python-Pfad hinzu
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from gui.main_window import MainWindow
from core.log_handler import initialize_logging, cleanup_logging


def main():
    """Hauptfunktion der Anwendung"""
    # Initialisiere Logging
    log_handler = initialize_logging()
    
    try:
        # Erstelle und starte die Anwendung
        app = MainWindow()
        app.run()
    except Exception as e:
        print(f"Fehler beim Starten der Anwendung: {e}")
        import traceback
        traceback.print_exc()
        input("Drücken Sie Enter zum Beenden...")
    finally:
        # Aufräumen
        cleanup_logging()


if __name__ == "__main__":
    main()