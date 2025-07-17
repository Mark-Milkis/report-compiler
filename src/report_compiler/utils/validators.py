"""
Validation utilities for file paths and PDF documents.
"""

import os
import pathlib
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
            # Normalize path separators for cross-platform compatibility
            pdf_path = pdf_path.replace("\\", os.sep).replace("/", os.sep)
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
    
    def validate_placeholders(self, placeholders: List[Dict], base_directory: str) -> Dict[str, any]:
        """
        Validates each placeholder, resolves its path, and checks for consistency.
        
        Args:
            placeholders: List of placeholder dictionaries.
            base_directory: The base directory for resolving relative paths.
            
        Returns:
            Dict with validation results.
        """
        result = {
            'valid': True,
            'errors': [],
            'warnings': [],
        }

        for placeholder in placeholders:
            file_path_raw = placeholder.get('file_path')
            if not file_path_raw:
                msg = f"Placeholder is missing 'file_path': {placeholder}"
                result['errors'].append(msg)
                result['valid'] = False
                continue

            if placeholder.get('is_recursive_docx'):
                # This should not happen if resolution is done before validation.
                msg = f"Unresolved recursive DOCX placeholder found during validation: {file_path_raw}"
                result['errors'].append(msg)
                result['valid'] = False
                continue

            path_validation = self.validate_pdf_path(file_path_raw, base_directory)
            if not path_validation['valid']:
                msg = f"Invalid PDF in placeholder '{file_path_raw}': {path_validation['error_message']}"
                result['errors'].append(msg)
                result['valid'] = False
                continue
            
            # Mutate the placeholder dictionary to include resolved path and page count
            placeholder['resolved_path'] = path_validation['resolved_path']
            placeholder['page_count'] = path_validation['page_count']

        if not result['valid']:
            return result # Stop if there are path errors

        # Now check for consistency
        consistency_result = self._validate_consistency(placeholders)
        result['errors'].extend(consistency_result['errors'])
        result['warnings'].extend(consistency_result['warnings'])
        if not consistency_result['valid']:
            result['valid'] = False
            
        return result

    @staticmethod
    def _validate_consistency(placeholders: List[Dict]) -> Dict[str, any]:
        """
        Validate a list of placeholders for consistency and conflicts.
        """
        result = {
            'valid': True,
            'errors': [],
            'warnings': [],
        }
        
        overlay_paths = [p.get('file_path', '') for p in placeholders if p.get('type') == 'table']
        merge_paths = [p.get('file_path', '') for p in placeholders if p.get('type') == 'paragraph']
        
        # Check for duplicate paths
        all_paths = overlay_paths + merge_paths
        seen_paths = set()
        duplicates = set()
        for path in all_paths:
            if path in seen_paths:
                duplicates.add(path)
            seen_paths.add(path)
        
        if duplicates:
            for path in duplicates:
                result['warnings'].append(f"Duplicate PDF path used: {path}")

        # Check for mixed usage of same PDF (both overlay and merge)
        overlay_set = set(overlay_paths)
        merge_set = set(merge_paths)
        overlapping = overlay_set.intersection(merge_set)
        
        if overlapping:
            for path in overlapping:
                result['errors'].append(f"PDF used for both overlay and merge, which is not allowed: {path}")
            result['valid'] = False
        
        return result
