"""
Überwacht Hotfolder auf neue Dateien
"""
import os
import time
from pathlib import Path
from typing import Dict, List, Optional, Set
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
        self.waiting_for_partner: Dict[str, float] = {}  # Dateien die auf ihren Partner warten
    
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
            
            # Prüfe ob Partner bereits wartet
            self._check_waiting_partner(file_path)
    
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
    
    def _check_waiting_partner(self, new_file_path: str):
        """Prüft ob ein Partner auf diese Datei wartet und verarbeitet das Paar"""
        if not self.config.process_pairs:
            return
        
        new_file_lower = new_file_path.lower()
        
        # Bestimme Partner-Datei
        if new_file_lower.endswith('.pdf'):
            partner_path = new_file_path[:-4] + '.xml'
            pdf_path = new_file_path
            xml_path = partner_path
        elif new_file_lower.endswith('.xml'):
            partner_path = new_file_path[:-4] + '.pdf'
            pdf_path = partner_path
            xml_path = new_file_path
        else:
            return
        
        # Prüfe ob Partner in der Warteliste ist
        if partner_path in self.waiting_for_partner:
            print(f"Partner gefunden! Verarbeite Paar: {os.path.basename(pdf_path)} + {os.path.basename(xml_path)}")
            
            # Entferne beide aus allen Listen
            del self.waiting_for_partner[partner_path]
            if new_file_path in self.pending_files:
                del self.pending_files[new_file_path]
            if partner_path in self.pending_files:
                del self.pending_files[partner_path]
            
            # Verarbeite das Paar sofort
            if os.path.exists(pdf_path) and os.path.exists(xml_path):
                self._process_document_pair(pdf_path, xml_path)
    
    def _process_document_pair(self, pdf_path: str, xml_path: str):
        """Verarbeitet ein PDF-XML Dokumentenpaar"""
        # Markiere beide Dateien als in Verarbeitung
        self.processing_files.append(pdf_path)
        self.processing_files.append(xml_path)
        
        try:
            doc_pair = DocumentPair(pdf_path=pdf_path, xml_path=xml_path)
            
            # Verarbeite Dokument
            success = self.processor.process_document(doc_pair, self.config)
            if success:
                print(f"Dokumentenpaar erfolgreich verarbeitet: {os.path.basename(pdf_path)}")
            else:
                print(f"Fehler bei der Verarbeitung: {os.path.basename(pdf_path)}")
        
        except Exception as e:
            print(f"Fehler bei der Dateiverarbeitung: {e}")
        finally:
            # Entferne aus processing_files
            if pdf_path in self.processing_files:
                self.processing_files.remove(pdf_path)
            if xml_path in self.processing_files:
                self.processing_files.remove(xml_path)
    
    def _process_file(self, file_path: str):
        """Verarbeitet eine einzelne Datei"""
        if not os.path.exists(file_path):
            return
        
        # Erstelle DocumentPair
        doc_pair = self._create_document_pair(file_path)
        
        if doc_pair:
            # Markiere Dateien als in Verarbeitung
            self.processing_files.append(file_path)
            if doc_pair.xml_path:
                self.processing_files.append(doc_pair.xml_path)
            
            try:
                # Verarbeite Dokument
                success = self.processor.process_document(doc_pair, self.config)
                if success:
                    print(f"Datei erfolgreich verarbeitet: {os.path.basename(file_path)}")
                else:
                    print(f"Fehler bei der Verarbeitung: {os.path.basename(file_path)}")
            
            except Exception as e:
                print(f"Fehler bei der Dateiverarbeitung: {e}")
            finally:
                # Entferne aus processing_files
                if file_path in self.processing_files:
                    self.processing_files.remove(file_path)
                if doc_pair.xml_path and doc_pair.xml_path in self.processing_files:
                    self.processing_files.remove(doc_pair.xml_path)
    
    def _create_document_pair(self, file_path: str) -> Optional[DocumentPair]:
        """Erstellt ein DocumentPair Objekt"""
        file_path_lower = file_path.lower()
        
        if file_path_lower.endswith('.pdf'):
            if self.config.process_pairs:
                # Suche nach zugehöriger XML-Datei
                xml_path = file_path[:-4] + '.xml'
                if os.path.exists(xml_path):
                    # XML gefunden - verarbeite als Paar
                    # Markiere XML als in Verarbeitung
                    if xml_path not in self.processing_files:
                        self.processing_files.append(xml_path)
                    return DocumentPair(pdf_path=file_path, xml_path=xml_path)
                else:
                    # Keine XML gefunden - warte auf Partner (ohne Timeout)
                    print(f"PDF gefunden, warte auf zugehörige XML: {os.path.basename(file_path)}")
                    self.waiting_for_partner[file_path] = time.time()
                    return None
            else:
                # process_pairs ist deaktiviert - verarbeite PDF allein
                return DocumentPair(pdf_path=file_path, xml_path=None)
        
        elif file_path_lower.endswith('.xml') and self.config.process_pairs:
            # Prüfe ob zugehörige PDF existiert
            pdf_path = file_path[:-4] + '.pdf'
            if os.path.exists(pdf_path):
                # PDF gefunden - verarbeite als Paar
                # Markiere PDF als in Verarbeitung
                if pdf_path not in self.processing_files:
                    self.processing_files.append(pdf_path)
                return DocumentPair(pdf_path=pdf_path, xml_path=file_path)
            else:
                # Keine PDF gefunden - warte auf Partner (ohne Timeout)
                print(f"XML gefunden, warte auf zugehörige PDF: {os.path.basename(file_path)}")
                self.waiting_for_partner[file_path] = time.time()
                return None
        
        return None
    
    def _matches_pattern(self, file_path: str) -> bool:
        """Prüft ob Datei zu den konfigurierten Patterns passt"""
        from fnmatch import fnmatch
        
        file_name = os.path.basename(file_path)
        
        # Bei aktiviertem process_pairs prüfe auch XML-Dateien
        if self.config.process_pairs:
            # Akzeptiere sowohl PDFs als auch XMLs
            if file_name.lower().endswith('.xml'):
                # Prüfe ob es ein entsprechendes PDF-Pattern gibt
                for pattern in self.config.file_patterns:
                    if pattern.lower().endswith('.pdf'):
                        xml_pattern = pattern[:-4] + '.xml'
                        if fnmatch(file_name.lower(), xml_pattern.lower()):
                            return True
                # Falls kein spezifisches Pattern, akzeptiere alle XMLs
                if '*.pdf' in [p.lower() for p in self.config.file_patterns]:
                    return True
        
        # Normale Pattern-Prüfung für PDFs
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
        
        # Sammle alle Dateien
        all_files = []
        processed_files: Set[str] = set()  # Um doppelte Verarbeitung zu vermeiden
        
        # Finde alle passenden Dateien
        for file_name in os.listdir(hotfolder.input_path):
            file_path = os.path.join(hotfolder.input_path, file_name)
            
            if os.path.isfile(file_path) and handler._matches_pattern(file_path):
                all_files.append(file_path)
        
        # Bei process_pairs: Versuche Paare zu bilden
        if hotfolder.process_pairs:
            # Gruppiere PDFs und XMLs
            pdf_files = [f for f in all_files if f.lower().endswith('.pdf')]
            xml_files = [f for f in all_files if f.lower().endswith('.xml')]
            
            # Verarbeite Paare
            for pdf_file in pdf_files:
                if pdf_file in processed_files:
                    continue
                
                xml_file = pdf_file[:-4] + '.xml'
                if xml_file in xml_files:
                    # Paar gefunden - füge beide zur Verarbeitung hinzu
                    handler.pending_files[pdf_file] = time.time() - 3  # Markiere als bereit
                    processed_files.add(pdf_file)
                    processed_files.add(xml_file)
                    print(f"Existierendes Paar gefunden: {os.path.basename(pdf_file)} + {os.path.basename(xml_file)}")
                else:
                    # Kein Partner - zur Warteliste hinzufügen (ohne Timeout)
                    handler.waiting_for_partner[pdf_file] = time.time()
                    processed_files.add(pdf_file)
                    print(f"Existierende PDF ohne XML gefunden, warte auf Partner: {os.path.basename(pdf_file)}")
            
            # Prüfe verbleibende XMLs
            for xml_file in xml_files:
                if xml_file in processed_files:
                    continue
                
                # XML ohne PDF - zur Warteliste hinzufügen (ohne Timeout)
                handler.waiting_for_partner[xml_file] = time.time()
                processed_files.add(xml_file)
                print(f"Existierende XML ohne PDF gefunden, warte auf Partner: {os.path.basename(xml_file)}")
        else:
            # Ohne process_pairs: Verarbeite nur PDFs
            for file_path in all_files:
                if file_path.lower().endswith('.pdf'):
                    handler.pending_files[file_path] = time.time() - 3  # Markiere als bereit
                    
    def rescan_hotfolder(self, hotfolder: HotfolderConfig):
        """Rescan: prüft alle Dateien erneut unabhängig vom bisherigen Zustand"""
        if not os.path.exists(hotfolder.input_path):
            return

        handler = self.handlers.get(hotfolder.id)
        if not handler:
            return

        print(f"Rescan: Überprüfe alle Dateien im Hotfolder {hotfolder.name}")

        for file_name in os.listdir(hotfolder.input_path):
            file_path = os.path.join(hotfolder.input_path, file_name)

            if not os.path.isfile(file_path):
                continue

            if not handler._matches_pattern(file_path):
                continue

            # Prüfe ob Datei aktuell schon verarbeitet wird
            if file_path in handler.processing_files:
                print(f"Datei wird gerade verarbeitet, überspringe: {file_name}")
                continue

            # Prüfe ob Datei evtl. schon auf Partner wartet (nur bei Paaren relevant)
            if hotfolder.process_pairs:
                file_lower = file_path.lower()

                if file_lower.endswith('.pdf'):
                    partner_path = file_path[:-4] + '.xml'
                    if os.path.exists(partner_path):
                        # Paar direkt verarbeiten
                        if partner_path not in handler.processing_files:
                            handler._process_document_pair(file_path, partner_path)
                        else:
                            print(f"Partner wird noch verarbeitet: {partner_path}")
                    else:
                        # PDF ohne Partner: wieder auf Partner warten
                        handler.waiting_for_partner[file_path] = time.time()
                        print(f"Rescan: PDF gefunden ohne XML, warte auf Partner: {file_name}")
                elif file_lower.endswith('.xml'):
                    partner_path = file_path[:-4] + '.pdf'
                    if os.path.exists(partner_path):
                        if partner_path not in handler.processing_files:
                            handler._process_document_pair(partner_path, file_path)
                        else:
                            print(f"Partner wird noch verarbeitet: {partner_path}")
                    else:
                        handler.waiting_for_partner[file_path] = time.time()
                        print(f"Rescan: XML gefunden ohne PDF, warte auf Partner: {file_name}")
            else:
                # Kein Paarmodus: Einzeldokumente erneut in pending queue legen
                handler.pending_files[file_path] = time.time() - 3  # direkt bereit zur Verarbeitung
                print(f"Rescan: Einzeldatei erneut vorgemerkt: {file_name}")
