"""
PDF to SVG conversion utility.

This module provides functionality to convert PDF pages to high-quality SVG format
for insertion into Word documents while preserving vector graphics quality.
"""

import os
import fitz  # PyMuPDF
from pathlib import Path
from typing import Optional

from .logging_config import get_module_logger


class PdfToSvgConverter:
    """Converts PDF pages to SVG format with high quality preservation."""
    
    def __init__(self):
        """Initialize the PDF to SVG converter."""
        self.logger = get_module_logger(__name__)
    
    def convert_page_to_svg(self, pdf_path: str, page_number: int, output_svg_path: str) -> bool:
        """
        Convert a specific PDF page to SVG format.
        
        Args:
            pdf_path: Path to the input PDF file.
            page_number: Page number to convert (1-based indexing).
            output_svg_path: Path where the SVG file will be saved.
            
        Returns:
            bool: True if conversion was successful, False otherwise.
        """
        try:
            # Validate inputs
            if not os.path.exists(pdf_path):
                self.logger.error(f"PDF file not found: {pdf_path}")
                return False
            
            # Ensure output directory exists
            output_dir = os.path.dirname(output_svg_path)
            if output_dir:
                os.makedirs(output_dir, exist_ok=True)
            
            # Open the PDF document
            doc = fitz.open(pdf_path)
            
            # Validate page number
            if page_number < 1 or page_number > len(doc):
                self.logger.error(f"Invalid page number {page_number}. PDF has {len(doc)} pages.")
                doc.close()
                return False
            
            # Get the specific page (convert to 0-based indexing)
            page = doc[page_number - 1]
            
            # Convert page to SVG
            # Use high DPI for better quality
            svg_text = page.get_svg_image(matrix=fitz.Identity)
            
            # Write SVG to file
            with open(output_svg_path, 'w', encoding='utf-8') as svg_file:
                svg_file.write(svg_text)
            
            doc.close()
            
            self.logger.info(f"Successfully converted page {page_number} of {os.path.basename(pdf_path)} to SVG")
            self.logger.debug(f"SVG saved to: {output_svg_path}")
            
            return True
            
        except Exception as e:
            self.logger.error(f"Failed to convert PDF page to SVG: {e}", exc_info=True)
            return False
    
    def validate_pdf(self, pdf_path: str) -> dict:
        """
        Validate a PDF file and return information about it.
        
        Args:
            pdf_path: Path to the PDF file to validate.
            
        Returns:
            dict: Validation result with 'valid', 'page_count', and 'error' keys.
        """
        result = {
            'valid': False,
            'page_count': 0,
            'error': None
        }
        
        try:
            if not os.path.exists(pdf_path):
                result['error'] = f"File not found: {pdf_path}"
                return result
            
            if not pdf_path.lower().endswith('.pdf'):
                result['error'] = f"File is not a PDF: {pdf_path}"
                return result
            
            # Try to open and get page count
            doc = fitz.open(pdf_path)
            result['page_count'] = len(doc)
            doc.close()
            
            if result['page_count'] == 0:
                result['error'] = "PDF file contains no pages"
                return result
            
            result['valid'] = True
            return result
            
        except Exception as e:
            result['error'] = f"Error validating PDF: {e}"
            return result
