@echo off
title Hotfolder PDF Processor
echo Starte Hotfolder PDF Processor...
python main.py
if errorlevel 1 (
    echo.
    echo Fehler beim Starten des Programms!
    echo Stellen Sie sicher, dass Python installiert ist und alle Abhängigkeiten installiert wurden.
    echo.
    echo Führen Sie folgenden Befehl aus: pip install -r requirements.txt
    echo.
    pause
)