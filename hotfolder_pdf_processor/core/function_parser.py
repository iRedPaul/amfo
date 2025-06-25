"""
Function Parser für Variablen und Funktionen
"""
import re
import os
from datetime import datetime
from typing import Any, Dict, List, Optional, Union, Tuple
import xml.etree.ElementTree as ET
import sys

# Füge das Hauptverzeichnis zum Python-Pfad hinzu falls nötig
sys.path.append(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    from core.counter_manager import get_counter_manager
except ImportError:
    # Fallback falls counter_manager nicht verfügbar
    def get_counter_manager():
        return None


class FunctionParser:
    """Parser für Funktionen und Variablen"""
    
    def __init__(self):
        self.variables = {}
        self.functions = {
            # String-Funktionen
            'FORMAT': self._format,
            'TRIM': self._trim,
            'LEFT': self._left,
            'RIGHT': self._right,
            'MID': self._mid,
            'TOUPPER': self._toupper,
            'TOLOWER': self._tolower,
            'LEN': self._len,
            'INDEXOF': self._indexof,
            
            # Datumsfunktionen
            'FORMATDATE': self._formatdate,
            
            # Numerische Funktionen
            'AUTOINCREMENT': self._autoincrement,
            
            # Bedingungen
            'IF': self._if,
            
            # Reguläre Ausdrücke
            'REGEXP.MATCH': self._regexp_match,
            'REGEXP.REPLACE': self._regexp_replace,
            
            # Scripting
            'SCRIPTING': self._scripting,
        }
        
        # Counter Manager für persistente Counter
        self.counter_manager = get_counter_manager()
    
    def parse_and_evaluate(self, expression: str, context: Dict[str, Any]) -> str:
        """
        Parst und evaluiert einen Ausdruck mit Funktionen und Variablen
        
        Args:
            expression: Der zu parsende Ausdruck
            context: Dictionary mit verfügbaren Variablen
            
        Returns:
            Das Ergebnis der Evaluation als String
        """
        self.variables = context.copy()
        
        # Wenn kein Funktionsaufruf, direkt zurückgeben
        if not self._contains_function(expression):
            return self._replace_variables(expression)
        
        # Funktionen von innen nach außen evaluieren
        result = self._evaluate_expression(expression)
        
        # Finale Variablenersetzung
        return self._replace_variables(str(result))
    
    def _contains_function(self, expression: str) -> bool:
        """Prüft ob der Ausdruck eine Funktion enthält"""
        for func_name in self.functions:
            if func_name + '(' in expression:
                return True
        return False
    
    def _evaluate_expression(self, expression: str) -> str:
        """Evaluiert einen Ausdruck rekursiv"""
        # Finde die innerste Funktion
        pattern = r'([A-Z]+(?:\.[A-Z]+)?)\s*\(((?:[^()]*|\([^()]*\))*)\)'
        
        while True:
            match = re.search(pattern, expression)
            if not match:
                break
            
            func_name = match.group(1)
            args_str = match.group(2)
            
            if func_name in self.functions:
                # Parse Argumente
                args = self._parse_arguments(args_str)
                
                # Evaluiere verschachtelte Funktionen in Argumenten
                evaluated_args = []
                for arg in args:
                    if self._contains_function(arg):
                        evaluated_args.append(self._evaluate_expression(arg))
                    else:
                        evaluated_args.append(self._replace_variables(arg))
                
                # Führe Funktion aus
                try:
                    result = self.functions[func_name](*evaluated_args)
                except Exception as e:
                    print(f"Fehler bei Funktion {func_name}: {e}")
                    result = ""
                
                # Ersetze Funktionsaufruf mit Ergebnis
                expression = expression[:match.start()] + str(result) + expression[match.end():]
            else:
                # Unbekannte Funktion, entferne sie
                expression = expression[:match.start()] + expression[match.end():]
        
        return expression
    
    def _parse_arguments(self, args_str: str) -> List[str]:
        """Parst Funktionsargumente"""
        if not args_str.strip():
            return []
        
        args = []
        current_arg = ""
        paren_level = 0
        in_quotes = False
        quote_char = None
        
        for char in args_str:
            if char in ('"', "'") and not in_quotes:
                in_quotes = True
                quote_char = char
                current_arg += char
            elif char == quote_char and in_quotes:
                in_quotes = False
                quote_char = None
                current_arg += char
            elif char == '(' and not in_quotes:
                paren_level += 1
                current_arg += char
            elif char == ')' and not in_quotes:
                paren_level -= 1
                current_arg += char
            elif char == ',' and paren_level == 0 and not in_quotes:
                args.append(current_arg.strip())
                current_arg = ""
            else:
                current_arg += char
        
        if current_arg.strip():
            args.append(current_arg.strip())
        
        # Entferne Quotes von String-Argumenten
        cleaned_args = []
        for arg in args:
            if arg.startswith('"') and arg.endswith('"'):
                cleaned_args.append(arg[1:-1])
            elif arg.startswith("'") and arg.endswith("'"):
                cleaned_args.append(arg[1:-1])
            else:
                cleaned_args.append(arg)
        
        return cleaned_args
    
    def _replace_variables(self, text: str) -> str:
        """Ersetzt Variablen im Text"""
        # Ersetze <Variable> Syntax
        pattern = r'<([^>]+)>'
        
        def replace_var(match):
            var_name = match.group(1)
            if var_name in self.variables:
                return str(self.variables[var_name])
            return match.group(0)
        
        return re.sub(pattern, replace_var, text)
    
    # String-Funktionen
    def _format(self, var: str, format_string: str) -> str:
        """FORMAT Funktion"""
        try:
            # Implementiere benutzerdefinierte Formatierung
            if '#' in format_string:
                # Zahlenformatierung
                num_digits = format_string.count('#')
                return var.zfill(num_digits)
            return var
        except:
            return var
    
    def _trim(self, var: str) -> str:
        """TRIM Funktion"""
        return var.strip()
    
    def _left(self, var: str, length: str) -> str:
        """LEFT Funktion"""
        try:
            return var[:int(length)]
        except:
            return var
    
    def _right(self, var: str, length: str) -> str:
        """RIGHT Funktion"""
        try:
            return var[-int(length):]
        except:
            return var
    
    def _mid(self, var: str, start: str, length: str = None) -> str:
        """MID Funktion"""
        try:
            start_idx = int(start) - 1  # 1-basiert zu 0-basiert
            if length:
                return var[start_idx:start_idx + int(length)]
            else:
                return var[start_idx:]
        except:
            return var
    
    def _toupper(self, var: str) -> str:
        """TOUPPER Funktion"""
        return var.upper()
    
    def _tolower(self, var: str) -> str:
        """TOLOWER Funktion"""
        return var.lower()
    
    def _len(self, var: str) -> str:
        """LEN Funktion"""
        return str(len(var))
    
    def _indexof(self, start_index: str, string_to_search: str, 
                 string_to_find: str, case_sensitive: str = "true") -> str:
        """INDEXOF Funktion"""
        try:
            start_idx = int(start_index)
            case_sensitive_bool = case_sensitive.lower() == "true"
            
            if not case_sensitive_bool:
                search_in = string_to_search.lower()
                search_for = string_to_find.lower()
            else:
                search_in = string_to_search
                search_for = string_to_find
            
            idx = search_in.find(search_for, start_idx)
            return str(idx + 1) if idx >= 0 else "0"  # 1-basiert
        except:
            return "0"
    
    # Datumsfunktionen
    def _formatdate(self, format_string: str) -> str:
        """FORMATDATE Funktion - KORRIGIERT und vereinfacht"""
        try:
            now = datetime.now()
            
            # Erstelle eine Kopie des Format-Strings
            result = format_string
            original_format = format_string  # Behalte das Original für die Prüfungen
            
            # Ersetze benutzerdefinierte Patterns mit Python strftime Patterns
            # WICHTIG: Längere Patterns zuerst ersetzen!
            replacements = [
                # Jahr
                ('yyyy', '%Y'),  # 4-stelliges Jahr
                ('yy', '%y'),    # 2-stelliges Jahr
                
                # Monat
                ('mmmm', '%B'),  # Vollständiger Monatsname
                ('mmm', '%b'),   # Abgekürzter Monatsname
                ('mm', '%m'),    # Monat mit führender Null
                
                # Tag
                ('dddd', '%A'),  # Vollständiger Wochentag
                ('ddd', '%a'),   # Abgekürzter Wochentag
                ('dd', '%d'),    # Tag mit führender Null
                
                # Zeit
                ('hh', '%H'),    # Stunde 24h mit führender Null
                ('MM', '%M'),    # Minute mit führender Null (GROSS M für Minute!)
                ('ss', '%S'),    # Sekunde mit führender Null
                
                # Spezielle Patterns
                ('ww', '%V'),    # Kalenderwoche
                ('tt', lambda: 'AM' if now.hour < 12 else 'PM'),  # AM/PM
                ('t', lambda: 'A' if now.hour < 12 else 'P'),     # A/P
            ]
            
            # Ersetze Patterns in der richtigen Reihenfolge (längste zuerst)
            for pattern, replacement in replacements:
                if callable(replacement):
                    # Für spezielle Funktionen wie AM/PM
                    result = result.replace(pattern, replacement())
                else:
                    # Für normale strftime Patterns
                    result = result.replace(pattern, replacement)
            
            # Jetzt behandle einzelne Zeichen BEVOR strftime
            # Prüfe ob einzelne Zeichen (nicht als Teil von Doppel-Zeichen) im Original-Format stehen
            
            # Einzelne 'd' ersetzen (nur wenn nicht bereits als 'dd' ersetzt)
            if self._contains_single_char(original_format, 'd'):
                result = result.replace('d', str(now.day))
            
            # Einzelne 'm' ersetzen (nur wenn nicht bereits als 'mm' oder länger ersetzt)
            if self._contains_single_char(original_format, 'm'):
                result = result.replace('m', str(now.month))
            
            # Einzelne 'h' ersetzen (nur wenn nicht bereits als 'hh' ersetzt)
            if self._contains_single_char(original_format, 'h'):
                result = result.replace('h', str(now.hour))
            
            # Einzelne 'M' ersetzen (nur wenn nicht bereits als 'MM' ersetzt)
            if self._contains_single_char(original_format, 'M'):
                result = result.replace('M', str(now.minute))
            
            # Einzelne 's' ersetzen (nur wenn nicht bereits als 'ss' ersetzt)
            if self._contains_single_char(original_format, 's'):
                result = result.replace('s', str(now.second))
            
            # Einzelne 'y' ersetzen (nur wenn nicht bereits als 'yy' oder 'yyyy' ersetzt)
            if self._contains_single_char(original_format, 'y'):
                result = result.replace('y', str(now.year % 100))
            
            # Jetzt formatiere mit strftime (für die übrigen %-Patterns)
            formatted_result = now.strftime(result)
            
            return formatted_result
                
        except Exception as e:
            print(f"Fehler bei FORMATDATE: {e}")
            print(f"Format-String: {format_string}")
            import traceback
            traceback.print_exc()
            return datetime.now().strftime("%d.%m.%Y")
    
    def _contains_single_char(self, format_string: str, char: str) -> bool:
        """Prüft ob ein einzelnes Zeichen (nicht als Teil eines Doppel-Zeichens) vorkommt"""
        import re
        
        if char == 'd':
            # Prüfe ob 'd' vorkommt, aber nicht als 'dd', 'ddd', oder 'dddd'
            pattern = r'(?<!d)d(?!d)'
            return bool(re.search(pattern, format_string))
        elif char == 'm':
            # Prüfe ob 'm' vorkommt, aber nicht als 'mm', 'mmm', oder 'mmmm'
            pattern = r'(?<!m)m(?!m)'
            return bool(re.search(pattern, format_string))
        elif char == 'y':
            # Prüfe ob 'y' vorkommt, aber nicht als 'yy', 'yyy', oder 'yyyy'
            pattern = r'(?<!y)y(?!y)'
            return bool(re.search(pattern, format_string))
        elif char == 'h':
            # Prüfe ob 'h' vorkommt, aber nicht als 'hh'
            pattern = r'(?<!h)h(?!h)'
            return bool(re.search(pattern, format_string))
        elif char == 'M':
            # Prüfe ob 'M' vorkommt, aber nicht als 'MM'
            pattern = r'(?<!M)M(?!M)'
            return bool(re.search(pattern, format_string))
        elif char == 's':
            # Prüfe ob 's' vorkommt, aber nicht als 'ss'
            pattern = r'(?<!s)s(?!s)'
            return bool(re.search(pattern, format_string))
        
        return False
    
    # Numerische Funktionen
    def _autoincrement(self, counter_name: str, start_value: str = "1", step: str = "1") -> str:
        """AUTOINCREMENT Funktion - Persistente Counter"""
        try:
            if self.counter_manager is None:
                # Fallback auf alte lokale Implementierung
                counter_key = f"auto_{counter_name}"
                if not hasattr(self, 'counters'):
                    self.counters = {}
                
                if counter_key not in self.counters:
                    try:
                        self.counters[counter_key] = int(start_value)
                    except:
                        self.counters[counter_key] = 1
                
                current_value = self.counters[counter_key]
                self.counters[counter_key] += int(step)
                return str(current_value)
            
            # Neue persistente Implementierung
            try:
                start_val = int(start_value)
            except (ValueError, TypeError):
                start_val = 1
                
            try:
                step_val = int(step)
            except (ValueError, TypeError):
                step_val = 1
            
            # Verwende den Counter-Manager für persistente Speicherung
            current_value = self.counter_manager.get_and_increment(
                counter_name=f"auto_{counter_name}",
                start_value=start_val,
                step=step_val
            )
            
            print(f"AUTOINCREMENT '{counter_name}': {current_value} (nächster: {current_value + step_val})")
            return str(current_value)
            
        except Exception as e:
            print(f"Fehler bei AUTOINCREMENT: {e}")
            return start_value
    
    # Bedingungen
    def _if(self, var: str, operator: str, value: str, 
            true_result: str, false_result: str, 
            case_sensitive: str = "true") -> str:
        """IF Funktion"""
        try:
            case_sensitive_bool = case_sensitive.lower() == "true"
            
            if not case_sensitive_bool:
                var = var.lower()
                value = value.lower()
            
            # Evaluiere Bedingung
            result = False
            
            if operator == "==" or operator == "=":
                result = var == value
            elif operator == "!=":
                result = var != value
            elif operator == ">":
                try:
                    result = float(var) > float(value)
                except:
                    result = var > value
            elif operator == "<":
                try:
                    result = float(var) < float(value)
                except:
                    result = var < value
            elif operator == ">=":
                try:
                    result = float(var) >= float(value)
                except:
                    result = var >= value
            elif operator == "<=":
                try:
                    result = float(var) <= float(value)
                except:
                    result = var <= value
            elif operator == "contains":
                result = value in var
            elif operator == "startswith":
                result = var.startswith(value)
            elif operator == "endswith":
                result = var.endswith(value)
            
            return true_result if result else false_result
            
        except Exception as e:
            print(f"Fehler bei IF: {e}")
            return false_result
    
    # Reguläre Ausdrücke
    def _regexp_match(self, var: str, pattern: str, 
                      submatch_index: str = "0") -> str:
        """REGEXP.MATCH Funktion"""
        try:
            matches = re.findall(pattern, var)
            if matches:
                idx = int(submatch_index)
                if isinstance(matches[0], tuple):
                    # Gruppen-Match
                    if idx < len(matches[0]):
                        return matches[0][idx]
                else:
                    # Einfacher Match
                    if idx == 0 and len(matches) > 0:
                        return matches[0]
            return ""
        except Exception as e:
            print(f"Fehler bei REGEXP.MATCH: {e}")
            return ""
    
    def _regexp_replace(self, var: str, pattern: str, replacement: str) -> str:
        """REGEXP.REPLACE Funktion"""
        try:
            return re.sub(pattern, replacement, var)
        except Exception as e:
            print(f"Fehler bei REGEXP.REPLACE: {e}")
            return var
    
    # Scripting
    def _scripting(self, script_path: str, *args) -> str:
        """SCRIPTING Funktion"""
        try:
            import subprocess
            
            if not os.path.exists(script_path):
                print(f"Script nicht gefunden: {script_path}")
                return ""
            
            # Führe Script aus
            if script_path.endswith('.bat'):
                cmd = [script_path] + list(args)
                result = subprocess.run(cmd, capture_output=True, text=True, shell=True)
            elif script_path.endswith('.vbs'):
                cmd = ['cscript', '//NoLogo', script_path] + list(args)
                result = subprocess.run(cmd, capture_output=True, text=True)
            else:
                print(f"Unbekannter Script-Typ: {script_path}")
                return ""
            
            if result.returncode == 0:
                return result.stdout.strip()
            else:
                print(f"Script-Fehler: {result.stderr}")
                return ""
                
        except Exception as e:
            print(f"Fehler bei SCRIPTING: {e}")
            return ""


class VariableExtractor:
    """Extrahiert verfügbare Variablen aus verschiedenen Quellen"""
    
    @staticmethod
    def get_standard_variables() -> Dict[str, Any]:
        """Gibt Standard-Variablen zurück"""
        now = datetime.now()
        return {
            'Date': now.strftime('%d.%m.%Y'),
            'Time': now.strftime('%H:%M:%S'),
            'DateTime': now.strftime('%d.%m.%Y %H:%M:%S'),
            'Year': str(now.year),
            'Month': str(now.month).zfill(2),
            'Day': str(now.day).zfill(2),
            'Hour': str(now.hour).zfill(2),
            'Minute': str(now.minute).zfill(2),
            'Second': str(now.second).zfill(2),
            'Weekday': now.strftime('%A'),
            'WeekNumber': now.strftime('%V'),
        }
    
    @staticmethod
    def get_ocr_variables(ocr_text: str, zones: Dict[str, str] = None) -> Dict[str, Any]:
        """Gibt OCR-Variablen zurück"""
        variables = {
            'OCR_FullText': ocr_text
        }
        
        # Füge Zonen-Variablen hinzu
        if zones:
            for zone_name, zone_text in zones.items():
                variables[f'OCR_{zone_name}'] = zone_text
        
        return variables
    
    @staticmethod
    def get_file_variables(file_path: str) -> Dict[str, Any]:
        """Gibt Datei-bezogene Variablen zurück"""
        from pathlib import Path
        
        path = Path(file_path)
        return {
            'FileName': path.stem,
            'FileExtension': path.suffix,
            'FullFileName': path.name,
            'FilePath': str(path.parent),
            'FullPath': str(path),
            'FileSize': str(path.stat().st_size) if path.exists() else '0',
        }
    
    @staticmethod
    def get_xml_variables(xml_path: str) -> Dict[str, Any]:
        """Extrahiert Variablen aus XML-Datei"""
        variables = {}
        
        try:
            tree = ET.parse(xml_path)
            root = tree.getroot()
            
            # Extrahiere alle Felder aus dem Fields-Element
            fields_elem = root.find(".//Fields")
            if fields_elem is not None:
                for field in fields_elem:
                    field_name = field.tag
                    field_value = field.text or ""
                    variables[f'XML_{field_name}'] = field_value
            
            # Extrahiere Document-Attribute
            doc_elem = root.find(".//Document")
            if doc_elem is not None:
                for attr_name, attr_value in doc_elem.attrib.items():
                    variables[f'XML_Doc_{attr_name}'] = attr_value
                    
        except Exception as e:
            print(f"Fehler beim Extrahieren von XML-Variablen: {e}")
        
        return variables