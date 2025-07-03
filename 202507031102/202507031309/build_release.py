#!/usr/bin/env python
"""
Build-System für Hotfolder PDF Processor mit Nuitka
Konfigurieren Sie die Pfade zu den Build-Dateien unten
"""
import os
import sys
import shutil
import subprocess
import hashlib
from pathlib import Path
from datetime import datetime

# ========================================================================
# KONFIGURATION - Passen Sie diese Pfade an Ihre Umgebung an
# ========================================================================

# Pfad zur Nuitka Build-Konfiguration
BUILD_CONFIG = "nuitka_build_config.py"

# Pfad zur Inno Setup .iss Datei  
ISS_FILE = "hotfolder_pdf_processor.iss"

# Pfad zum Inno Setup Compiler (ISCC.exe)
INNO_COMPILER_PATH = r"C:\Program Files (x86)\Inno Setup 6\ISCC.exe"

# ========================================================================
# AB HIER NICHTS ÄNDERN
# ========================================================================

class ReleaseBuilder:
    def __init__(self):
        self.root_dir = Path(__file__).parent
        self.build_dir = self.root_dir / "build"
        self.dist_dir = self.root_dir / "dist"
        self.installer_dir = self.root_dir / "installer"
        self.config_file = self.root_dir / BUILD_CONFIG
        self.iss_file = self.root_dir / ISS_FILE
        self.inno_compiler = INNO_COMPILER_PATH
        
    def clean_build(self):
        """Bereinigt alle Build-Verzeichnisse"""
        print("Bereinige Build-Verzeichnisse...")
        for directory in [self.build_dir, self.dist_dir]:
            if directory.exists():
                shutil.rmtree(directory)
        # Installer-Verzeichnis nur leeren
        if self.installer_dir.exists():
            for file in self.installer_dir.glob("*"):
                if file.is_file():
                    file.unlink()
        
        # Nuitka-spezifische Verzeichnisse
        nuitka_dirs = ["main.build", "main.dist", "main.onefile-build", 
                      "windows_service.build", "windows_service.dist"]
        for dir_name in nuitka_dirs:
            dir_path = self.root_dir / dir_name
            if dir_path.exists():
                shutil.rmtree(dir_path)
                
        print("✓ Build-Verzeichnisse bereinigt")
    
    def find_msvc_compiler(self):
        """Sucht nach Microsoft Visual C++ Compiler"""
        # Typische Installationspfade für VS 2022/2019/2017
        vs_paths = [
            r"C:\Program Files\Microsoft Visual Studio\2022",
            r"C:\Program Files (x86)\Microsoft Visual Studio\2022",
            r"C:\Program Files\Microsoft Visual Studio\2019",
            r"C:\Program Files (x86)\Microsoft Visual Studio\2019",
            r"C:\Program Files\Microsoft Visual Studio\2017",
            r"C:\Program Files (x86)\Microsoft Visual Studio\2017",
        ]
        
        editions = ["Enterprise", "Professional", "Community", "BuildTools"]
        
        for vs_path in vs_paths:
            for edition in editions:
                # VS 2022/2019 Pfad
                vcvars_path = Path(vs_path) / edition / "VC" / "Auxiliary" / "Build" / "vcvars64.bat"
                if vcvars_path.exists():
                    return str(vcvars_path)
                
                # VS 2017 Pfad
                vcvars_path = Path(vs_path) / edition / "VC" / "Auxiliary" / "Build" / "vcvars64.bat"
                if vcvars_path.exists():
                    return str(vcvars_path)
        
        # Prüfe auch Visual Studio Build Tools direkt
        buildtools_paths = [
            r"C:\Program Files (x86)\Microsoft Visual Studio\2022\BuildTools\VC\Auxiliary\Build\vcvars64.bat",
            r"C:\Program Files (x86)\Microsoft Visual Studio\2019\BuildTools\VC\Auxiliary\Build\vcvars64.bat",
        ]
        
        for path in buildtools_paths:
            if os.path.exists(path):
                return path
                
        return None
    
    def setup_msvc_environment(self):
        """Richtet die MSVC-Umgebung ein"""
        vcvars_path = self.find_msvc_compiler()
        if not vcvars_path:
            return False
            
        print(f"  Gefunden: {vcvars_path}")
        
        # Führe vcvars64.bat aus und capture die Umgebungsvariablen
        cmd = f'"{vcvars_path}" && set'
        try:
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            if result.returncode != 0:
                return False
                
            # Parse die Umgebungsvariablen
            for line in result.stdout.split('\n'):
                if '=' in line:
                    key, value = line.split('=', 1)
                    os.environ[key] = value
                    
            return True
        except:
            return False
    
    def check_requirements(self):
        """Prüft ob alle Abhängigkeiten vorhanden sind"""
        print("\nPrüfe Abhängigkeiten...")
        
        errors = []
        
        # Prüfe Nuitka
        try:
            result = subprocess.run([sys.executable, '-m', 'nuitka', '--version'], 
                                  capture_output=True, text=True)
            if result.returncode == 0:
                print(f"✓ Nuitka gefunden (Version: {result.stdout.strip()})")
            else:
                errors.append("Nuitka nicht gefunden")
        except:
            errors.append("Nuitka nicht gefunden")
        
        # Prüfe C++ Compiler
        print("\nPrüfe C++ Compiler...")
        compiler_found = False
        
        # Versuche MSVC zu finden und einzurichten
        if self.setup_msvc_environment():
            print("✓ Microsoft Visual C++ Compiler eingerichtet")
            compiler_found = True
        else:
            # Prüfe ob cl.exe direkt verfügbar ist
            try:
                result = subprocess.run(['cl'], capture_output=True, text=True, shell=True)
                if "Microsoft" in result.stderr:
                    print("✓ Microsoft C++ Compiler gefunden (bereits im PATH)")
                    compiler_found = True
            except:
                pass
                
        if not compiler_found:
            # Prüfe MinGW als Alternative
            try:
                result = subprocess.run(['gcc', '--version'], capture_output=True, text=True)
                if result.returncode == 0:
                    print("✓ GCC Compiler gefunden (MinGW)")
                    compiler_found = True
            except:
                pass
                
        if not compiler_found:
            errors.append("Kein C++ Compiler gefunden (MSVC oder MinGW benötigt)")
            print("\n  Hinweise zur Installation:")
            print("  1. Visual Studio Installer öffnen")
            print("  2. 'Desktop development with C++' Workload installieren")
            print("  3. Oder MinGW64 installieren: https://www.mingw-w64.org/")
            
        # Prüfe Build-Config
        if not self.config_file.exists():
            errors.append(f"Build-Config nicht gefunden: {self.config_file}")
        else:
            print(f"✓ Build-Config gefunden: {self.config_file.name}")
            
        # Prüfe iss-Datei
        if not self.iss_file.exists():
            errors.append(f"ISS-Datei nicht gefunden: {self.iss_file}")
        else:
            print(f"✓ ISS-Datei gefunden: {self.iss_file.name}")
        
        # Prüfe Inno Setup
        if not os.path.exists(self.inno_compiler):
            errors.append(f"Inno Setup nicht gefunden: {self.inno_compiler}")
        else:
            print(f"✓ Inno Setup gefunden: {self.inno_compiler}")
        
        if errors:
            print("\n✗ Fehlende Abhängigkeiten:")
            for error in errors:
                print(f"  - {error}")
            sys.exit(1)
        
        print("\n✓ Alle Abhängigkeiten vorhanden")
        return True
    
    def build_with_nuitka(self):
        """Führt den Nuitka Build aus"""
        print("\nStarte Nuitka Build...")
        
        # Lade Build-Konfiguration
        spec = {}
        with open(self.config_file, 'r', encoding='utf-8') as f:
            exec(f.read(), spec)
        
        builds = spec.get('NUITKA_BUILDS', [])
        
        for build in builds:
            print(f"\nBaue {build['name']}...")
            
            cmd = [sys.executable, '-m', 'nuitka'] + build['options']
            
            # Zeige Kommando (gekürzt)
            print(f"Führe aus: {' '.join(cmd[:10])}...")
            
            result = subprocess.run(cmd, cwd=self.root_dir)
            
            if result.returncode != 0:
                print(f"✗ Nuitka Build fehlgeschlagen für {build['name']}!")
                sys.exit(1)
            
            print(f"✓ {build['name']} erfolgreich gebaut")
        
        # Verschiebe Build-Ergebnisse
        self._organize_build_output()
        
        print("\n✓ Nuitka Build erfolgreich abgeschlossen")
    
    def _organize_build_output(self):
        """Organisiert die Build-Ausgabe für Inno Setup"""
        print("\nOrganisiere Build-Ausgabe...")
        
        # Erstelle dist-Verzeichnisstruktur
        output_dir = self.dist_dir / "Hotfolder PDF Processor"
        output_dir.mkdir(parents=True, exist_ok=True)
        
        # Verschiebe Hauptanwendung
        main_exe = self.root_dir / "Hotfolder PDF Processor.exe"
        if main_exe.exists():
            shutil.move(str(main_exe), str(output_dir / main_exe.name))
        
        # Verschiebe Service
        service_exe = self.root_dir / "HotfolderPDFService.exe"
        if service_exe.exists():
            shutil.move(str(service_exe), str(output_dir / service_exe.name))
        
        # Kopiere Daten-Dateien
        data_files = ['settings.json', 'counters.json', 'hotfolders.json', 'icon.ico']
        for file in data_files:
            src = self.root_dir / file
            if src.exists():
                shutil.copy2(str(src), str(output_dir / file))
        
        # Kopiere Poppler
        poppler_dir = self.root_dir / "poppler"
        if poppler_dir.exists():
            shutil.copytree(str(poppler_dir), str(output_dir / "poppler"))
        
        print("✓ Build-Ausgabe organisiert")
    
    def build_installer(self):
        """Erstellt den Installer mit Inno Setup"""
        print("\nErstelle Installer mit Inno Setup...")
        print(f"Verwende ISS-Datei: {self.iss_file.name}")
        
        # Stelle sicher dass installer Verzeichnis existiert
        self.installer_dir.mkdir(exist_ok=True)
        
        cmd = [
            self.inno_compiler,
            '/Q',  # Quiet mode
            '/O' + str(self.installer_dir),  # Output directory
            str(self.iss_file)
        ]
        
        result = subprocess.run(cmd)
        
        if result.returncode != 0:
            print("✗ Inno Setup Build fehlgeschlagen!")
            return False
        
        print("✓ Installer erfolgreich erstellt")
        return True
    
    def create_checksums(self):
        """Erstellt Checksummen für die Installer-Datei"""
        print("\nErstelle Checksummen...")
        
        installer_files = list(self.installer_dir.glob("*.exe"))
        if not installer_files:
            print("✗ Keine Installer-Datei gefunden!")
            return
        
        installer_file = installer_files[0]
        
        # Berechne Hashes
        sha256_hash = hashlib.sha256()
        md5_hash = hashlib.md5()
        
        with open(installer_file, "rb") as f:
            while chunk := f.read(8192):
                sha256_hash.update(chunk)
                md5_hash.update(chunk)
        
        # Schreibe Checksummen
        checksum_file = installer_file.with_suffix('.checksums.txt')
        with open(checksum_file, 'w') as f:
            f.write(f"Datei: {installer_file.name}\n")
            f.write(f"Größe: {installer_file.stat().st_size:,} Bytes\n")
            f.write(f"Erstellt: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
            f.write(f"SHA256: {sha256_hash.hexdigest()}\n")
            f.write(f"MD5:    {md5_hash.hexdigest()}\n")
        
        print(f"✓ Checksummen erstellt: {checksum_file.name}")
    
    def build_release(self):
        """Führt den kompletten Release-Build aus"""
        print("="*60)
        print("Hotfolder PDF Processor - Release Build (Nuitka)")
        print("="*60)
        
        start_time = datetime.now()
        
        try:
            # 1. Prüfe Abhängigkeiten
            self.check_requirements()
            
            # 2. Bereinige alte Builds
            self.clean_build()
            
            # 3. Nuitka Build
            self.build_with_nuitka()
            
            # 4. Erstelle Installer
            if self.build_installer():
                # 5. Erstelle Checksummen
                self.create_checksums()
            
            # Erfolgsmeldung
            duration = datetime.now() - start_time
            print("\n" + "="*60)
            print("✓ Release Build erfolgreich abgeschlossen!")
            print(f"  Dauer: {duration.total_seconds():.1f} Sekunden")
            print(f"  Ausgabe: {self.installer_dir}")
            print("="*60)
            
        except Exception as e:
            print("\n" + "="*60)
            print(f"✗ Build fehlgeschlagen: {e}")
            print("="*60)
            sys.exit(1)


if __name__ == "__main__":
    builder = ReleaseBuilder()
    builder.build_release()