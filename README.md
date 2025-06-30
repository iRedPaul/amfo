# 📁 Hotfolder PDF Processor

Der Hotfolder PDF Processor ist eine leistungsstarke Anwendung zur automatischen Verarbeitung von PDF-Dateien. Er überwacht konfigurierte Eingangsordner (Hotfolder), führt eine Reihe von Aktionen auf eingehenden PDFs aus und exportiert die Ergebnisse in verschiedenen Formaten. Das Programm kann als Windows-Dienst ausgeführt werden und bietet eine grafische Benutzeroberfläche zur einfachen Verwaltung und Konfiguration.

## ✨ Funktionen

* **Hotfolder-Überwachung**: Automatische Erkennung und Verarbeitung neuer Dateien in definierten Eingangsordnern.
* **Optische Zeichenerkennung (OCR)**:
    * Erkennung von Text in PDFs.
    * Definierbare OCR-Zonen zur Extraktion spezifischer Informationen aus Dokumenten.
* **XML-Feldzuordnung**: Automatische Verarbeitung und Zuordnung von Daten aus XML-Dateien, die PDF-Paaren zugeordnet sind.
* **Datenbankintegration**: Unterstützung für SQL-Abfragen, um Daten aus Datenbanken (z.B. MariaDB) abzurufen und in der Verarbeitung zu nutzen.
* **Flexible Exportoptionen**:
    * Export als durchsuchbare PDF/A.
    * Export von Metadaten und extrahierten Daten als XML, JSON oder CSV.
* **Dynamische Dateinamen und Pfade**: Verwendung von Ausdrücken mit Variablen (Datei, Datum, Zeit, OCR, XML, Ordnerstruktur) und Funktionen zur flexiblen Benennung und Ablage von Ausgabedateien.
    * **Verfügbare Variablen**:
        * `Date`, `Time`, `Now`
        * `FileName`, `FileExtension`, `FullFileName`, `FilePath`, `FullPath`, `FileSize`
        * `level0` bis `level5` für Ordnerstrukturen
        * `OCR_FullText` und benannte OCR-Zonen
        * `XML_` Felder aus XML-Dokumenten
    * **Verfügbare Funktionen**: String-Manipulation, Datumsformatierung, numerische Operationen, bedingte Logik, reguläre Ausdrücke (RegEx), externe Skripte (BAT, VBS) und Datenbankabfragen.
* **Aktionsbasierte Verarbeitung**: Unterstützung für verschiedene Verarbeitungsaktionen wie Komprimierung von PDFs.
* **Windows Service**: Kann als Hintergrunddienst installiert und verwaltet werden, um eine kontinuierliche Verarbeitung zu gewährleisten.
* **Grafische Benutzeroberfläche (GUI)**: Eine intuitive Benutzeroberfläche zur Konfiguration von Hotfoldern, Datenbankverbindungen, Exporteinstellungen und weiteren Optionen.
* **E-Mail-Benachrichtigungen**: Konfigurierbare SMTP-Einstellungen mit Unterstützung für Basis- und OAuth2-Authentifizierung.
