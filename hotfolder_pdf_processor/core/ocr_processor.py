"""
OCR Processor Module

Handles optical character recognition operations for various image formats.
Uses the centralized logging system for consistent logging across the application.
"""

import os
import tempfile
from typing import Dict, List, Optional, Tuple, Any
from PIL import Image
import pytesseract
import cv2
import numpy as np
from pathlib import Path

from core.logger import Logger


class OCRProcessor:
    """
    Handles OCR processing operations for various image formats.
    """
    
    def __init__(self, config: Dict[str, Any]):
        """
        Initialize the OCR processor.
        
        Args:
            config: Configuration dictionary containing OCR settings
        """
        self.logger = Logger()
        self.config = config
        self.supported_formats = ['.jpg', '.jpeg', '.png', '.tiff', '.tif', '.bmp', '.gif']
        
        # OCR Engine-Konfiguration
        self.tesseract_config = config.get('tesseract_config', '--psm 3')
        self.language = config.get('ocr_language', 'deu+eng')
        self.confidence_threshold = config.get('confidence_threshold', 60)
        
        self.logger.info(
            "OCR Processor initialisiert",
            extra={
                'supported_formats': self.supported_formats,
                'language': self.language,
                'confidence_threshold': self.confidence_threshold
            }
        )
    
    def is_supported_format(self, file_path: str) -> bool:
        """
        Check if the file format is supported for OCR processing.
        
        Args:
            file_path: Path to the file
            
        Returns:
            bool: True if format is supported
        """
        file_ext = Path(file_path).suffix.lower()
        supported = file_ext in self.supported_formats
        
        if not supported:
            self.logger.warning(
                "Nicht unterstütztes Dateiformat für OCR",
                extra={
                    'file_path': file_path,
                    'file_extension': file_ext,
                    'supported_formats': self.supported_formats
                }
            )
        
        return supported
    
    def preprocess_image(self, image_path: str) -> Optional[np.ndarray]:
        """
        Preprocess image for better OCR results.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Preprocessed image as numpy array or None if error
        """
        try:
            self.logger.debug(
                "Starte Bildvorverarbeitung für OCR",
                extra={'image_path': image_path}
            )
            
            # Load image
            image = cv2.imread(image_path)
            if image is None:
                raise ValueError(f"Konnte Bild nicht laden: {image_path}")
            
            # Convert to grayscale
            gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
            
            # Apply preprocessing techniques
            # 1. Noise reduction
            denoised = cv2.medianBlur(gray, 3)
            
            # 2. Contrast enhancement
            clahe = cv2.createCLAHE(clipLimit=2.0, tileGridSize=(8, 8))
            enhanced = clahe.apply(denoised)
            
            # 3. Thresholding
            _, thresh = cv2.threshold(enhanced, 0, 255, cv2.THRESH_BINARY + cv2.THRESH_OTSU)
            
            self.logger.debug(
                "Bildvorverarbeitung abgeschlossen",
                extra={
                    'image_path': image_path,
                    'original_shape': image.shape,
                    'processed_shape': thresh.shape
                }
            )
            
            return thresh
            
        except Exception as e:
            self.logger.error(
                "Fehler bei Bildvorverarbeitung",
                extra={
                    'image_path': image_path,
                    'error': str(e)
                },
                exc_info=True
            )
            return None
    
    def extract_text_with_confidence(self, image_path: str) -> Tuple[str, float]:
        """
        Extract text from image with confidence scoring.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Tuple of (extracted_text, average_confidence)
        """
        try:
            self.logger.debug(
                "Starte OCR-Textextraktion",
                extra={'image_path': image_path}
            )
            
            # Preprocess image
            processed_image = self.preprocess_image(image_path)
            if processed_image is None:
                return "", 0.0
            
            # Perform OCR with detailed data
            ocr_data = pytesseract.image_to_data(
                processed_image,
                lang=self.language,
                config=self.tesseract_config,
                output_type=pytesseract.Output.DICT
            )
            
            # Extract text and calculate confidence
            text_parts = []
            confidences = []
            
            for i, confidence in enumerate(ocr_data['conf']):
                if int(confidence) >= self.confidence_threshold:
                    text = ocr_data['text'][i].strip()
                    if text:
                        text_parts.append(text)
                        confidences.append(int(confidence))
            
            extracted_text = ' '.join(text_parts)
            avg_confidence = sum(confidences) / len(confidences) if confidences else 0.0
            
            self.logger.info(
                "OCR-Textextraktion abgeschlossen",
                extra={
                    'image_path': image_path,
                    'text_length': len(extracted_text),
                    'average_confidence': avg_confidence,
                    'high_confidence_words': len(confidences),
                    'total_words': len([t for t in ocr_data['text'] if t.strip()])
                }
            )
            
            return extracted_text, avg_confidence
            
        except Exception as e:
            self.logger.error(
                "Fehler bei OCR-Textextraktion",
                extra={
                    'image_path': image_path,
                    'error': str(e)
                },
                exc_info=True
            )
            return "", 0.0
    
    def extract_text_simple(self, image_path: str) -> str:
        """
        Simple text extraction without confidence scoring.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Extracted text
        """
        try:
            self.logger.debug(
                "Starte einfache OCR-Textextraktion",
                extra={'image_path': image_path}
            )
            
            # Check if file exists
            if not os.path.exists(image_path):
                raise FileNotFoundError(f"Bilddatei nicht gefunden: {image_path}")
            
            # Perform OCR
            extracted_text = pytesseract.image_to_string(
                image_path,
                lang=self.language,
                config=self.tesseract_config
            ).strip()
            
            self.logger.info(
                "Einfache OCR-Textextraktion abgeschlossen",
                extra={
                    'image_path': image_path,
                    'text_length': len(extracted_text)
                }
            )
            
            return extracted_text
            
        except Exception as e:
            self.logger.error(
                "Fehler bei einfacher OCR-Textextraktion",
                extra={
                    'image_path': image_path,
                    'error': str(e)
                },
                exc_info=True
            )
            return ""
    
    def process_multiple_images(self, image_paths: List[str]) -> Dict[str, Dict[str, Any]]:
        """
        Process multiple images with OCR.
        
        Args:
            image_paths: List of image file paths
            
        Returns:
            Dictionary with results for each image
        """
        results = {}
        
        self.logger.info(
            "Starte Batch-OCR-Verarbeitung",
            extra={
                'total_images': len(image_paths),
                'supported_formats': self.supported_formats
            }
        )
        
        successful_count = 0
        failed_count = 0
        
        for image_path in image_paths:
            try:
                if not self.is_supported_format(image_path):
                    results[image_path] = {
                        'success': False,
                        'error': 'Nicht unterstütztes Dateiformat',
                        'text': '',
                        'confidence': 0.0
                    }
                    failed_count += 1
                    continue
                
                text, confidence = self.extract_text_with_confidence(image_path)
                
                results[image_path] = {
                    'success': True,
                    'text': text,
                    'confidence': confidence,
                    'error': None
                }
                successful_count += 1
                
            except Exception as e:
                self.logger.error(
                    "Fehler bei Batch-OCR für Bild",
                    extra={
                        'image_path': image_path,
                        'error': str(e)
                    },
                    exc_info=True
                )
                
                results[image_path] = {
                    'success': False,
                    'error': str(e),
                    'text': '',
                    'confidence': 0.0
                }
                failed_count += 1
        
        self.logger.info(
            "Batch-OCR-Verarbeitung abgeschlossen",
            extra={
                'total_images': len(image_paths),
                'successful': successful_count,
                'failed': failed_count
            }
        )
        
        return results
    
    def get_image_info(self, image_path: str) -> Optional[Dict[str, Any]]:
        """
        Get detailed information about an image file.
        
        Args:
            image_path: Path to the image file
            
        Returns:
            Dictionary with image information or None if error
        """
        try:
            self.logger.debug(
                "Sammle Bildinformationen",
                extra={'image_path': image_path}
            )
            
            # Check if file exists
            if not os.path.exists(image_path):
                raise FileNotFoundError(f"Bilddatei nicht gefunden: {image_path}")
            
            # Get file info
            file_stats = os.stat(image_path)
            
            # Get image info using PIL
            with Image.open(image_path) as img:
                info = {
                    'file_path': image_path,
                    'file_size': file_stats.st_size,
                    'format': img.format,
                    'mode': img.mode,
                    'size': img.size,
                    'width': img.width,
                    'height': img.height,
                    'has_transparency': img.mode in ('RGBA', 'LA') or 'transparency' in img.info
                }
            
            self.logger.debug(
                "Bildinformationen gesammelt",
                extra={
                    'image_path': image_path,
                    'image_info': info
                }
            )
            
            return info
            
        except Exception as e:
            self.logger.error(
                "Fehler beim Sammeln von Bildinformationen",
                extra={
                    'image_path': image_path,
                    'error': str(e)
                },
                exc_info=True
            )
            return None
    
    def validate_ocr_setup(self) -> bool:
        """
        Validate that OCR setup is working correctly.
        
        Returns:
            bool: True if OCR setup is valid
        """
        try:
            self.logger.info("Validiere OCR-Setup")
            
            # Test Tesseract installation
            version = pytesseract.get_tesseract_version()
            self.logger.info(
                "Tesseract-Version gefunden",
                extra={'version': str(version)}
            )
            
            # Test language availability
            languages = pytesseract.get_languages()
            required_langs = self.language.split('+')
            missing_langs = [lang for lang in required_langs if lang not in languages]
            
            if missing_langs:
                self.logger.error(
                    "Erforderliche OCR-Sprachen nicht verfügbar",
                    extra={
                        'missing_languages': missing_langs,
                        'available_languages': languages,
                        'required_languages': required_langs
                    }
                )
                return False
            
            # Create test image for basic functionality test
            test_image = Image.new('RGB', (200, 50), color='white')
            
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as tmp_file:
                test_image.save(tmp_file.name)
                
                # Test basic OCR functionality
                result = pytesseract.image_to_string(tmp_file.name, lang=self.language)
                
                # Clean up
                os.unlink(tmp_file.name)
            
            self.logger.info(
                "OCR-Setup erfolgreich validiert",
                extra={
                    'tesseract_version': str(version),
                    'available_languages': languages,
                    'configured_language': self.language
                }
            )
            
            return True
            
        except Exception as e:
            self.logger.error(
                "OCR-Setup-Validierung fehlgeschlagen",
                extra={'error': str(e)},
                exc_info=True
            )
            return False
    
    def cleanup_temp_files(self):
        """
        Clean up any temporary files created during OCR processing.
        """
        try:
            self.logger.debug("Bereinige temporäre OCR-Dateien")
            
            # This method can be extended to clean up specific temp files
            # created during OCR processing if needed
            
            self.logger.debug("Temporäre OCR-Dateien bereinigt")
            
        except Exception as e:
            self.logger.error(
                "Fehler bei der Bereinigung temporärer OCR-Dateien",
                extra={'error': str(e)},
                exc_info=True
            )
