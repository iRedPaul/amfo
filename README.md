# Hotfolder


Ein fortschrittliches Tool zur automatischen PDF-Verarbeitung mit flexiblen Export-Optionen.

### Flexible Export-Konfigurationen
- **Mehrere Export-Ziele**: Konfigurieren Sie beliebig viele Export-Ziele pro Hotfolder
- **Verschiedene Export-Typen**: 
  - PDF-Dateien in Ordner
  - ZIP-Archive erstellen
  - E-Mail-Versand mit Anhängen
  - FTP/SFTP-Upload
  - Netzwerk-Freigaben

### Erweiterte Export-Optionen
- **Dynamische Pfade**: Verwenden Sie Variablen und Funktionen für Ordnerpfade
- **Bedingte Exports**: Definieren Sie Bedingungen, wann ein Export ausgeführt werden soll
- **Export-Bedingungen**: UND/ODER-Verknüpfung von Bedingungen basierend auf:
  - Dateigrößen
  - OCR-Textinhalten  
  - Dateinamen-Mustern
  - XML-Feldinhalten
  - Datum/Zeit-Kriterien

### E-Mail-Integration
- **SMTP-Unterstützung**: Vollständige E-Mail-Konfiguration
- **Mehrere Empfänger**: To, CC, BCC-Listen
- **Dynamische Inhalte**: Betreff und Nachrichtentext mit Variablen
- **Flexible Anhänge**: PDF und/oder XML-Dateien
- **TLS/SSL-Unterstützung**: Sichere E-Mail-Übertragung

### FTP/SFTP-Upload
- **Mehrere Server**: Verschiedene FTP-Server pro Export
- **Sichere Übertragung**: SSL/TLS-Unterstützung
- **Dynamische Pfade**: Server-Pfade mit Variablen
- **Passiver/Aktiver Modus**: Flexible Verbindungsoptionen

### Erweiterte Verarbeitung
- **Parallele Verarbeitung**: Mehrere Dateien gleichzeitig verarbeiten
- **Retry-Mechanismus**: Automatische Wiederholung bei Fehlern
- **Archivierung**: Automatische Archivierung verarbeiteter Dateien
- **Fehler-Benachrichtigungen**: E-Mail-Benachrichtigung bei Fehlern

### Neue Funktions-Sprache
- **AUTOINCREMENT**: Persistente Counter für fortlaufende Nummerierung
- **FORMATDATE**: Erweiterte Datumsformatierung
- **IF-Bedingungen**: Komplexe bedingte Logik
- **Reguläre Ausdrücke**: Mustersuche und -ersetzung
- **String-Funktionen**: LEFT, RIGHT, MID, TRIM, etc.

## 📁 Dateistruktur

```
hotfolder_pdf_processor/
├── core/
│   ├── config_manager.py          # Konfigurationsverwaltung
│   ├── counter_manager.py          # Persistente Counter
│   ├── export_engine.py           # 🆕 Export-Verarbeitung
│   ├── file_watcher.py             # Dateiüberwachung
│   ├── function_parser.py          # Erweiterte Funktions-Parser
│   ├── hotfolder_manager.py        # 🔄 Erweiterte Hotfolder-Verwaltung
│   ├── ocr_processor.py            # OCR-Verarbeitung
│   ├── pdf_processor.py            # 🔄 Erweiterte PDF-Verarbeitung
│   └── xml_field_processor.py      # XML-Feldverarbeitung
├── models/
│   ├── export_config.py            # 🆕 Export-Datenmodelle
│   └── hotfolder_config.py         # 🔄 Erweiterte Hotfolder-Modelle
├── gui/
│   ├── counter_management_dialog.py # Counter-Verwaltung
│   ├── email_ftp_dialogs.py        # 🆕 E-Mail/FTP-Konfiguration
│   ├── export_config_dialog.py     # 🆕 Export-Konfiguration
│   ├── export_conditions_dialog.py # 🆕 Export-Bedingungen
│   ├── expression_dialog.py        # Ausdruck-Editor
│   ├── expression_editor_base.py   # Basis-Editor
│   ├── hotfolder_dialog.py         # 🔄 Erweiterter Hotfolder-Dialog
│   ├── main_window.py              # 🔄 Erweitertes Hauptfenster
│   ├── xml_field_dialog.py         # XML-Feld-Editor
│   └── zone_selector.py            # OCR-Zonen-Auswahl
├── main.py                         # Hauptprogramm
├── requirements.txt                # Abhängigkeiten
└── README.md                       # Diese Datei
```

## 🚀 Installation

### Voraussetzungen
- Python 3.8 oder höher
- Windows 10/11 (für Service-Installation)

### Abhängigkeiten installieren
```bash
pip install -r requirements.txt
```

### Zusätzliche Software (optional)
- **Tesseract OCR**: Für OCR-Funktionalität
  - Download: https://github.com/UB-Mannheim/tesseract/wiki
  - Fügen Sie Tesseract zum PATH hinzu
- **Poppler**: Für PDF-zu-Bild-Konvertierung
  - Download: https://poppler.freedesktop.org/
- **OCRmyPDF**: Für erweiterte OCR-Features
  ```bash
  pip install ocrmypdf
  ```

## 🔧 Konfiguration

### Hotfolder erstellen
1. Starten Sie die Anwendung: `python main.py`
2. Klicken Sie auf "➕ Neuer Hotfolder"
3. Konfigurieren Sie:
   - **Name**: Eindeutiger Name für den Hotfolder
   - **Input-Ordner**: Überwachter Ordner für neue Dateien
   - **Verarbeitungsaktionen**: PDF-Komprimierung, OCR, PDF/A, etc.
   - **Export-Ziele**: Ein oder mehrere Export-Konfigurationen

### Export-Konfigurationen
Jeder Hotfolder kann mehrere Export-Ziele haben:

#### 1. PDF in Ordner
```
Typ: PDF-Datei
Ausgabe-Pfad: C:\Output\<Year>\<Month>
Dateiname: <FileName>_<Date>
```

#### 2. E-Mail-Versand
```
Typ: E-Mail versenden
SMTP-Server: smtp.gmail.com:587
Betreff: Dokument verarbeitet: <FileName>
Anhänge: PDF und XML
```

#### 3. FTP-Upload
```
Typ: FTP-Upload
Server: ftp.example.com
Remote-Pfad: /upload/<Year>/<Month>
```

### Export-Bedingungen
Definieren Sie, wann ein Export ausgeführt werden soll:

```
Bedingung 1: FileSize > 1000000 (größer als 1MB)
Bedingung 2: OCR_FullText contains "Rechnung"
Verknüpfung: UND (beide müssen erfüllt sein)
```

## 📝 Variablen und Funktionen

### Standard-Variablen
- `<Date>` - Aktuelles Datum (dd.mm.yyyy)
- `<Time>` - Aktuelle Zeit (hh:mm:ss)
- `<Year>`, `<Month>`, `<Day>` - Datums-Komponenten
- `<FileName>` - Dateiname ohne Erweiterung
- `<FileSize>` - Dateigröße in Bytes
- `<OCR_FullText>` - Kompletter OCR-Text

### Datums-Funktionen
```
FORMATDATE("d.m.yyyy hh:MM:ss") → "20.6.2025 14:30:25"
FORMATDATE("dddd, d. mmmm yyyy") → "Freitag, 20. Juni 2025"
```

### String-Funktionen
```
LEFT("<FileName>", 8) → Erste 8 Zeichen
RIGHT("<FileName>", 3) → Letzte 3 Zeichen
TOUPPER("<FileName>") → Großbuchstaben
TRIM("<OCR_Text>") → Leerzeichen entfernen
```

### Auto-Increment Counter
```
AUTOINCREMENT("Rechnung", 1000, 1) → 1000, 1001, 1002, ...
AUTOINCREMENT("Monat", 1, 1) → 1, 2, 3, ...
```

### Bedingungen
```
IF("<FileSize>", ">", "1000000", "Große Datei", "Kleine Datei")
```

### Reguläre Ausdrücke
```
REGEXP.MATCH("<OCR_FullText>", "\\d{2}\\.\\d{2}\\.\\d{4}", 0) → Findet Datum
REGEXP.REPLACE("<FileName>", "\\s+", "_") → Ersetzt Leerzeichen
```

## 🛠️ Service-Installation (Windows)

### Als Windows-Service installieren
```bash
# Als Administrator ausführen
python windows_service.py install
python windows_service.py start
```

### Service-Verwaltung
```bash
python windows_service.py status   # Status anzeigen
python windows_service.py stop     # Service stoppen
python windows_service.py remove   # Service entfernen
```

## 📊 Monitoring und Statistiken

Das erweiterte Hauptfenster zeigt:
- **Hotfolder-Status**: Aktive/Inaktive Hotfolder
- **Export-Statistiken**: Anzahl und Typen der Exports
- **Konfigurationsfehler**: Automatische Validierung
- **Details-Panel**: Detaillierte Informationen pro Hotfolder

## 🔄 Migration von v1.0

Bestehende Konfigurationen werden automatisch migriert:
- **Legacy Output-Pfad** → **Standard-Export-Konfiguration**
- **Output-Filename-Expression** → **Export-Dateiname**
- **Alte XML-Mappings** → **Neue Expression-Syntax**

## 🧪 Testen

### Export-Konfigurationen testen
1. Wählen Sie einen Hotfolder aus
2. Klicken Sie auf "🧪 Testen"
3. Das System testet:
   - E-Mail-Verbindungen
   - FTP-Verbindungen
   - Pfad-Ausdrücke
   - Bedingungen

### Konfiguration validieren
- **Menü → Extras → Konfiguration validieren**
- Prüft alle Hotfolder auf Fehler
- Zeigt Warnungen und Empfehlungen

## 📈 Performance-Optimierung

### Parallele Verarbeitung
```
Parallele Verarbeitung: ✓ Aktiviert
Max. parallele Jobs: 4
```

### Retry-Mechanismus
```
Wiederholung bei Fehlern: ✓ Aktiviert
Max. Wiederholungen: 3
```

### OCR-Cache
- Automatisches Caching von OCR-Ergebnissen
- Wiederverwendung für mehrere Export-Ziele

## 🔧 Erweiterte Konfiguration

### E-Mail-Einstellungen
```json
{
  "smtp_server": "smtp.gmail.com",
  "smtp_port": 587,
  "use_tls": true,
  "username": "ihr.email@gmail.com",
  "password": "app-password",
  "from_address": "system@ihrefirma.de",
  "subject_expression": "Dokument verarbeitet: <FileName>",
  "body_expression": "Anbei finden Sie das verarbeitete Dokument vom <Date>."
}
```

### FTP-Einstellungen
```json
{
  "server": "ftp.ihrefirma.de",
  "port": 21,
  "username": "upload_user",
  "password": "secure_password",
  "remote_path_expression": "/uploads/<Year>/<Month>",
  "use_passive": true,
  "use_ssl": false
}
```

### Export-Bedingungen
```json
{
  "conditions": [
    {
      "variable": "FileSize",
      "operator": "greater_than",
      "value": "1000000",
      "enabled": true
    },
    {
      "variable": "OCR_FullText",
      "operator": "contains",
      "value": "Rechnung",
      "enabled": true
    }
  ],
  "condition_logic": "AND"
}
```

## 🐛 Fehlerbehebung

### Häufige Probleme

#### E-Mail-Versand funktioniert nicht
- Prüfen Sie SMTP-Server und Port
- Verwenden Sie App-Passwörter für Gmail
- Testen Sie die Verbindung über "🧪 Testen"

#### FTP-Upload schlägt fehl
- Prüfen Sie Firewall-Einstellungen
- Versuchen Sie passiven Modus
- Testen Sie Anmeldedaten

#### OCR erkennt keinen Text
- Installieren Sie Tesseract OCR
- Prüfen Sie PDF-Qualität
- Verwenden Sie höhere DPI-Einstellungen

#### Counter funktionieren nicht
- Prüfen Sie Schreibrechte im Programm-Ordner
- Löschen Sie `counters.json` für Neustart

### Log-Dateien
- **Service-Logs**: `C:\ProgramData\HotfolderPDFProcessor\service.log`
- **Anwendungs-Logs**: Console-Output

## 🔐 Sicherheit

### E-Mail-Sicherheit
- Verwenden Sie App-Passwörter statt Haupt-Passwörter
- Aktivieren Sie TLS/SSL für SMTP
- Verschlüsseln Sie gespeicherte Passwörter

### FTP-Sicherheit
- Verwenden Sie SFTP wenn möglich
- Erstellen Sie dedizierte Upload-Benutzer
- Beschränken Sie FTP-Berechtigungen

### Datei-Sicherheit
- Überwachen Sie Input-Ordner-Berechtigungen
- Verwenden Sie sichere Output-Pfade
- Archivieren Sie verarbeitete Dateien

## 📞 Support

Bei Fragen oder Problemen:
1. Prüfen Sie diese Dokumentation
2. Validieren Sie Ihre Konfiguration
3. Testen Sie Export-Konfigurationen
4. Prüfen Sie Log-Dateien

## 🚀 Roadmap

Geplante Features für zukünftige Versionen:
- **Cloud-Integration**: OneDrive, Google Drive, Dropbox
- **Webhook-Support**: HTTP-POST-Benachrichtigungen
- **Erweiterte OCR**: Formular-Erkennung, Tabellen-Extraktion
- **API-Interface**: REST-API für externe Integration
- **Dashboard**: Web-basierte Überwachung
- **Batch-Verarbeitung**: Verarbeitung bestehender Dateien

## 📄 Lizenz

Dieses Projekt ist unter der MIT-Lizenz lizenziert.
