"""
Überwacht Hotfolder auf neue Dateien
"""
import os
import time
from pathlib import Path
from typing import Dict, List, Optional
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler, FileCreatedEvent
import sys

sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from models.hotfolder_config import HotfolderConfig, DocumentPair
from core.pdf_processor import PDFProcessor


class HotfolderHandler(FileSystemEventHandler):
    """Handler für Dateisystem-Events in einem Hotfolder"""
    
    def __init__(self, hotfolder_config: HotfolderConfig, processor: PDFProcessor):
        self.config = hotfolder_config
        self.processor = processor
        self.pending_files: Dict[str, float] = {}  # Dateiname -> Zeitstempel
        self.processing_files: List[str] = []  # Liste der aktuell verarbeiteten Dateien
    
    def on_created(self, event):
        """Wird aufgerufen wenn eine neue Datei erstellt wurde"""
        if event.is_directory:
            return
        
        file_path = event.src_path
        
        # Prüfe ob Datei zum Pattern passt
        if self._matches_pattern(file_path):
            print(f"Neue Datei erkannt: {file_path}")
            # Warte bis Datei vollständig geschrieben ist
            self.pending_files[file_path] = time.time()
    
    def on_modified(self, event):
        """Wird aufgerufen wenn eine Datei modifiziert wurde"""
        if event.is_directory:
            return
        
        file_path = event.src_path
        
        # Aktualisiere Zeitstempel wenn Datei noch geschrieben wird
        if file_path in self.pending_files:
            self.pending_files[file_path] = time.time()
    
    def process_pending_files(self):
        """Verarbeitet Dateien die bereit sind"""
        current_time = time.time()
        files_to_process = []
        
        # Finde Dateien die seit 2 Sekunden nicht mehr verändert wurden
        for file_path, last_modified in list(self.pending_files.items()):
            if current_time - last_modified > 2.0:
                if os.path.exists(file_path) and file_path not in self.processing_files:
                    files_to_process.append(file_path)
                del self.pending_files[file_path]
        
        # Verarbeite gefundene Dateien
        for file_path in files_to_process:
            self._process_file(file_path)
    
    def _process_file(self, file_path: str):
        """Verarbeitet eine einzelne Datei"""
        if not os.path.exists(file_path):
            return
        
        self.processing_files.append(file_path)
        
        try:
            # Erstelle DocumentPair
            doc_pair = self._create_document_pair(file_path)
            
            if doc_pair:
                # Verarbeite Dokument
                success = self.processor.process_document(doc_pair, self.config)
                if success:
                    print(f"Datei erfolgreich verarbeitet: {os.path.basename(file_path)}")
                else:
                    print(f"Fehler bei der Verarbeitung: {os.path.basename(file_path)}")
            
        except Exception as e:
            print(f"Fehler bei der Dateiverarbeitung: {e}")
        finally:
            if file_path in self.processing_files:
                self.processing_files.remove(file_path)
    
    def _create_document_pair(self, file_path: str) -> Optional[DocumentPair]:
        """Erstellt ein DocumentPair Objekt"""
        if file_path.lower().endswith('.pdf'):
            # Suche nach zugehöriger XML-Datei
            xml_path = None
            if self.config.process_pairs:
                potential_xml = file_path[:-4] + '.xml'
                if os.path.exists(potential_xml):
                    xml_path = potential_xml
                    # Markiere XML als in Verarbeitung
                    if xml_path not in self.processing_files:
                        self.processing_files.append(xml_path)
                else:
                    # Wenn process_pairs aktiv ist, aber keine XML existiert, ignoriere PDF
                    print(f"PDF ohne zugehörige XML gefunden (wird ignoriert): {os.path.basename(file_path)}")
                    return None
            
            return DocumentPair(pdf_path=file_path, xml_path=xml_path)
        
        elif file_path.lower().endswith('.xml') and self.config.process_pairs:
            # Prüfe ob zugehörige PDF existiert
            potential_pdf = file_path[:-4] + '.pdf'
            if os.path.exists(potential_pdf):
                # PDF wird separat verarbeitet, ignoriere XML
                return None
            else:
                # Keine zugehörige PDF, ignoriere XML
                print(f"XML ohne zugehörige PDF gefunden: {os.path.basename(file_path)}")
                return None
        
        return None
    
    def _matches_pattern(self, file_path: str) -> bool:
        """Prüft ob Datei zu den konfigurierten Patterns passt"""
        from fnmatch import fnmatch
        
        file_name = os.path.basename(file_path)
        for pattern in self.config.file_patterns:
            if fnmatch(file_name.lower(), pattern.lower()):
                return True
        return False


class FileWatcher:
    """Verwaltet die Überwachung mehrerer Hotfolder"""
    
    def __init__(self):
        self.observers: Dict[str, Observer] = {}
        self.handlers: Dict[str, HotfolderHandler] = {}
        self.processor = PDFProcessor()
        self._running = False
        self._last_cleanup = time.time()
        self._cleanup_interval = 3600  # Cleanup alle Stunde
    
    def start_watching(self, hotfolder: HotfolderConfig):
        """Startet die Überwachung eines Hotfolders"""
        if hotfolder.id in self.observers:
            print(f"Hotfolder {hotfolder.name} wird bereits überwacht")
            return
        
        if not os.path.exists(hotfolder.input_path):
            print(f"Input-Pfad existiert nicht: {hotfolder.input_path}")
            return
        
        # Erstelle Handler und Observer
        handler = HotfolderHandler(hotfolder, self.processor)
        observer = Observer()
        observer.schedule(handler, hotfolder.input_path, recursive=False)
        
        # Starte Observer
        observer.start()
        
        # Speichere Referenzen
        self.observers[hotfolder.id] = observer
        self.handlers[hotfolder.id] = handler
        
        print(f"Überwachung gestartet für: {hotfolder.name}")
    
    def stop_watching(self, hotfolder_id: str):
        """Stoppt die Überwachung eines Hotfolders"""
        if hotfolder_id in self.observers:
            observer = self.observers[hotfolder_id]
            observer.stop()
            observer.join(timeout=5)
            
            del self.observers[hotfolder_id]
            del self.handlers[hotfolder_id]
            
            print(f"Überwachung gestoppt für Hotfolder ID: {hotfolder_id}")
    
    def stop_all(self):
        """Stoppt alle Überwachungen"""
        self._running = False
        for hotfolder_id in list(self.observers.keys()):
            self.stop_watching(hotfolder_id)
        
        # Führe finales Cleanup durch
        self.processor.cleanup_temp_dir()
    
    def process_pending_files(self):
        """Verarbeitet ausstehende Dateien in allen Hotfoldern"""
        for handler in self.handlers.values():
            handler.process_pending_files()
        
        # Prüfe ob Cleanup notwendig ist
        current_time = time.time()
        if current_time - self._last_cleanup > self._cleanup_interval:
            self.processor.cleanup_temp_dir()
            self._last_cleanup = current_time
    
    def scan_existing_files(self, hotfolder: HotfolderConfig):
        """Scannt und verarbeitet bereits vorhandene Dateien in einem Hotfolder"""
        if not os.path.exists(hotfolder.input_path):
            return
        
        handler = self.handlers.get(hotfolder.id)
        if not handler:
            return
        
        # Finde alle passenden Dateien
        for file_name in os.listdir(hotfolder.input_path):
            file_path = os.path.join(hotfolder.input_path, file_name)
            
            if os.path.isfile(file_path) and handler._matches_pattern(file_path):
                # Füge Datei zur Verarbeitung hinzu
                handler.pending_files[file_path] = time.time() - 3  # Markiere als bereit