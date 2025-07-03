@echo off
title Hotfolder PDF Processor
echo Starte Hotfolder PDF Processor...
python main.py
if errorlevel 1 (
    echo.
    echo Fehler beim Starten des Programms!
    echo.
    pause
)