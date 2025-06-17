"""
Validation utilities for file paths and PDF documents.
"""

import os
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import fitz  # PyMuPDF
from ..core.config import Config


class Validators:
    """Utility class for validating files and paths."""
    
    @staticmethod
    def validate_pdf_path(pdf_path: str, base_directory: str) -> Dict[str, any]:
        """
        Validate and resolve a PDF file path.
        
        Args:
            pdf_path: Relative or absolute path to PDF
            base_directory: Base directory for resolving relative paths
            
        Returns:
            Dict with validation results including resolved path and page count
        """
        result = {
            'valid': False,
            'resolved_path': None,
            'page_count': 0,
            'error_message': None,
            'file_size_mb': 0.0
        }
        
        try:
            # Try to resolve the path
            if os.path.isabs(pdf_path):
                resolved_path = pdf_path
            else:
                resolved_path = os.path.join(base_directory, pdf_path)
            
            resolved_path = os.path.abspath(resolved_path)
            
            # Check if file exists
            if not os.path.exists(resolved_path):
                result['error_message'] = f"File not found: {resolved_path}"
                return result
            
            # Check if it's a file (not directory)
            if not os.path.isfile(resolved_path):
                result['error_message'] = f"Path is not a file: {resolved_path}"
                return result
            
            # Check file extension
            if not any(resolved_path.lower().endswith(ext) for ext in Config.SUPPORTED_PDF_EXTENSIONS):
                result['error_message'] = f"Not a PDF file: {resolved_path}"
                return result
            
            # Try to open as PDF and get page count
            try:
                with fitz.open(resolved_path) as pdf_doc:
                    page_count = len(pdf_doc)
                    if page_count == 0:
                        result['error_message'] = f"PDF has no pages: {resolved_path}"
                        return result
                    
                    result['page_count'] = page_count
            except Exception as e:
                result['error_message'] = f"Invalid PDF file: {e}"
                return result
            
            # Get file size
            try:
                result['file_size_mb'] = os.path.getsize(resolved_path) / (1024 * 1024)
            except Exception:
                result['file_size_mb'] = 0.0
            
            result['valid'] = True
            result['resolved_path'] = resolved_path
            
        except Exception as e:
            result['error_message'] = f"Path validation error: {e}"
        
        return result
    
    @staticmethod
    def validate_docx_path(docx_path: str) -> Dict[str, any]:
        """
        Validate a DOCX file path.
        
        Args:
            docx_path: Path to DOCX file
            
        Returns:
            Dict with validation results
        """
        result = {
            'valid': False,
            'resolved_path': None,
            'error_message': None,
            'file_size_mb': 0.0
        }
        
        try:
            resolved_path = os.path.abspath(docx_path)
            
            # Check if file exists
            if not os.path.exists(resolved_path):
                result['error_message'] = f"File not found: {resolved_path}"
                return result
            
            # Check if it's a file
            if not os.path.isfile(resolved_path):
                result['error_message'] = f"Path is not a file: {resolved_path}"
                return result
            
            # Check file extension
            if not any(resolved_path.lower().endswith(ext) for ext in Config.SUPPORTED_DOCX_EXTENSIONS):
                result['error_message'] = f"Not a DOCX file: {resolved_path}"
                return result
            
            # Get file size
            try:
                result['file_size_mb'] = os.path.getsize(resolved_path) / (1024 * 1024)
            except Exception:
                result['file_size_mb'] = 0.0
            
            result['valid'] = True
            result['resolved_path'] = resolved_path
            
        except Exception as e:
            result['error_message'] = f"Path validation error: {e}"
        
        return result
    
    @staticmethod
    def validate_output_path(output_path: str) -> Dict[str, any]:
        """
        Validate an output file path.
        
        Args:
            output_path: Desired output file path
            
        Returns:
            Dict with validation results
        """
        result = {
            'valid': False,
            'resolved_path': None,
            'error_message': None,
            'directory_exists': False,
            'file_exists': False
        }
        
        try:
            resolved_path = os.path.abspath(output_path)
            directory = os.path.dirname(resolved_path)
            
            # Check if directory exists or can be created
            if not os.path.exists(directory):
                try:
                    os.makedirs(directory, exist_ok=True)
                    result['directory_exists'] = True
                except Exception as e:
                    result['error_message'] = f"Cannot create output directory: {e}"
                    return result
            else:
                result['directory_exists'] = True
            
            # Check if file already exists
            result['file_exists'] = os.path.exists(resolved_path)
            
            # Check if we can write to the location
            try:
                # Try to create/touch the file
                with open(resolved_path, 'a'):
                    pass
                # If it didn't exist before, remove it
                if not result['file_exists']:
                    os.remove(resolved_path)
            except Exception as e:
                result['error_message'] = f"Cannot write to output location: {e}"
                return result
            
            result['valid'] = True
            result['resolved_path'] = resolved_path
            
        except Exception as e:
            result['error_message'] = f"Output path validation error: {e}"
        
        return result
    
    @staticmethod
    def validate_placeholders(placeholders: List[Dict]) -> Dict[str, any]:
        """
        Validate a list of placeholders for consistency and conflicts.
        
        Args:
            placeholders: List of placeholder dictionaries
            
        Returns:
            Dict with validation results
        """
        result = {
            'valid': True,
            'errors': [],
            'warnings': [],
            'total_placeholders': len(placeholders),
            'overlay_count': 0,
            'merge_count': 0
        }
        
        overlay_paths = []
        merge_paths = []
        
        for placeholder in placeholders:
            if placeholder.get('type') == 'overlay':
                result['overlay_count'] += 1
                overlay_paths.append(placeholder.get('path', ''))
            elif placeholder.get('type') == 'merge':
                result['merge_count'] += 1
                merge_paths.append(placeholder.get('path', ''))
        
        # Check for duplicate paths
        all_paths = overlay_paths + merge_paths
        seen_paths = set()
        for path in all_paths:
            if path in seen_paths:
                result['warnings'].append(f"Duplicate PDF path: {path}")
            seen_paths.add(path)
        
        # Check for mixed usage of same PDF (both overlay and merge)
        overlay_set = set(overlay_paths)
        merge_set = set(merge_paths)
        overlapping = overlay_set.intersection(merge_set)
        
        if overlapping:
            for path in overlapping:
                result['errors'].append(f"PDF used for both overlay and merge: {path}")
            result['valid'] = False
        
        return result
