@echo off
title Hotfolder PDF Processor - Service Installation
echo ========================================
echo Hotfolder PDF Processor Service Setup
echo ========================================
echo.

REM Verzeichnis des Skripts ermitteln
set SCRIPT_DIR=%~dp0
REM Backslash am Ende entfernen (nur zur Sicherheit)
if "%SCRIPT_DIR:~-1%"=="\" set SCRIPT_DIR=%SCRIPT_DIR:~0,-1%

REM Prüfe ob als Administrator ausgeführt
net session >nul 2>&1
if %errorLevel% neq 0 (
    echo FEHLER: Dieses Skript muss als Administrator ausgefuehrt werden!
    echo.
    echo Rechtsklick auf die Datei und "Als Administrator ausfuehren" waehlen.
    echo.
    pause
    exit /b 1
)

:menu
echo Was moechten Sie tun?
echo.
echo 1. Service installieren
echo 2. Service starten
echo 3. Service stoppen
echo 4. Service-Status anzeigen
echo 5. Service deinstallieren
echo 6. Beenden
echo.
set /p choice="Ihre Wahl (1-6): "

if "%choice%"=="1" goto install
if "%choice%"=="2" goto start
if "%choice%"=="3" goto stop
if "%choice%"=="4" goto status
if "%choice%"=="5" goto remove
if "%choice%"=="6" goto end

echo Ungueltige Auswahl!
echo.
goto menu

:install
echo.
echo Installiere Service...
python "%SCRIPT_DIR%\windows_service.py" install
if %errorLevel% equ 0 (
    echo Service erfolgreich installiert!
    echo.
    echo Moechten Sie den Service jetzt starten? (J/N)
    set /p start_now=""
    if /i "%start_now%"=="J" goto start
) else (
    echo Fehler bei der Installation!
)
echo.
pause
goto menu

:start
echo.
echo Starte Service...
python "%SCRIPT_DIR%\windows_service.py" start
if %errorLevel% equ 0 (
    echo Service erfolgreich gestartet!
) else (
    echo Fehler beim Starten!
)
echo.
pause
goto menu

:stop
echo.
echo Stoppe Service...
python "%SCRIPT_DIR%\windows_service.py" stop
if %errorLevel% equ 0 (
    echo Service erfolgreich gestoppt!
) else (
    echo Fehler beim Stoppen!
)
echo.
pause
goto menu

:status
echo.
echo Pruefe Service-Status...
python "%SCRIPT_DIR%\windows_service.py" status
echo.
pause
goto menu

:remove
echo.
echo WARNUNG: Service wird deinstalliert!
echo.
set /p confirm="Sind Sie sicher? (J/N): "
if /i "%confirm%"=="J" (
    python "%SCRIPT_DIR%\windows_service.py" stop >nul 2>&1
    python "%SCRIPT_DIR%\windows_service.py" remove
    if %errorLevel% equ 0 (
        echo Service erfolgreich entfernt!
    ) else (
        echo Fehler beim Entfernen!
    )
) else (
    echo Abgebrochen.
)
echo.
pause
goto menu

:end
exit /b 0
