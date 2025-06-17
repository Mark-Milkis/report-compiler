"""
Advanced marker removal utilities for PDF processing.
"""

import fitz  # PyMuPDF
from typing import Dict, Optional, Tuple


class MarkerRemover:
    """Handles clean removal of marker text from PDF pages."""
    
    def __init__(self):
        pass
    
    def remove_marker_text(self, page: fitz.Page, marker_text: str) -> bool:
        """
        Remove marker text from a PDF page using advanced method.
        
        Args:
            page: PyMuPDF page object
            marker_text: Text to remove
            
        Returns:
            bool: True if marker was found and removed, False otherwise
        """
        try:
            # Find the marker text
            marker_rect = self._find_marker_rect(page, marker_text)
            if not marker_rect:
                return False
            
            print(f"        🎯 Advanced removal of marker text at ({marker_rect.x0:.1f}, {marker_rect.y0:.1f})")
            
            # Remove the text using advanced method
            return self._remove_text_advanced(page, marker_rect)
            
        except Exception as e:
            print(f"        ⚠️ Error removing marker: {e}")
            return False
    
    def _find_marker_rect(self, page: fitz.Page, marker_text: str) -> Optional[fitz.Rect]:
        """
        Find the rectangle containing the marker text.
        
        Args:
            page: PyMuPDF page object
            marker_text: Text to find
            
        Returns:
            fitz.Rect or None: Rectangle containing the text, or None if not found
        """
        try:
            # Get all text instances
            text_instances = page.search_for(marker_text)
            
            if text_instances:
                # Return the first instance
                return text_instances[0]
            
            return None
            
        except Exception:
            return None
    
    def _remove_text_advanced(self, page: fitz.Page, text_rect: fitz.Rect) -> bool:
        """
        Remove text using advanced content stream method.
        
        Args:
            page: PyMuPDF page object
            text_rect: Rectangle containing text to remove
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Get background color for the area
            bg_color = self._get_background_color(page, text_rect)
            
            # Create a rectangle to cover the text
            # Expand slightly to ensure complete coverage
            cover_rect = fitz.Rect(
                text_rect.x0 - 1,
                text_rect.y0 - 1, 
                text_rect.x1 + 1,
                text_rect.y1 + 1
            )
            
            # Draw a rectangle with background color to cover the text
            page.draw_rect(cover_rect, color=bg_color, fill=bg_color, width=0)
            
            return True
            
        except Exception as e:
            print(f"        ⚠️ Advanced removal failed: {e}")
            # Fallback to simple redaction
            return self._remove_text_simple(page, text_rect)
    
    def _remove_text_simple(self, page: fitz.Page, text_rect: fitz.Rect) -> bool:
        """
        Simple text removal using redaction.
        
        Args:
            page: PyMuPDF page object
            text_rect: Rectangle containing text to remove
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            # Add redaction annotation
            redact_annot = page.add_redact_annot(text_rect)
            
            # Apply redaction (removes the text)
            page.apply_redactions()
            
            return True
            
        except Exception:
            return False
    
    def _get_background_color(self, page: fitz.Page, rect: fitz.Rect) -> Tuple[float, float, float]:
        """
        Attempt to determine background color for the given area.
        
        Args:
            page: PyMuPDF page object
            rect: Area to analyze
            
        Returns:
            RGB color tuple (default: white)
        """
        try:
            # For most documents, white background is safe
            # In a more sophisticated implementation, we could:
            # 1. Sample pixels around the text
            # 2. Look for background patterns
            # 3. Check document color scheme
            
            return (1.0, 1.0, 1.0)  # White
            
        except Exception:
            return (1.0, 1.0, 1.0)  # Default to white
    
    def find_marker_position(self, page: fitz.Page, marker_text: str) -> Optional[Dict[str, any]]:
        """
        Find marker position and return detailed information.
        
        Args:
            page: PyMuPDF page object
            marker_text: Marker text to find
            
        Returns:
            Dict with position information, or None if not found
        """
        try:
            marker_rect = self._find_marker_rect(page, marker_text)
            if not marker_rect:
                return None
            
            # Convert to inches for display
            marker_x_in = marker_rect.x0 / 72
            marker_y_in = marker_rect.y0 / 72
            marker_width_in = marker_rect.width / 72
            marker_height_in = marker_rect.height / 72
            
            return {
                'rect': marker_rect,
                'position_points': (marker_rect.x0, marker_rect.y0),
                'position_inches': (marker_x_in, marker_y_in),
                'size_points': (marker_rect.width, marker_rect.height),
                'size_inches': (marker_width_in, marker_height_in)
            }
            
        except Exception:
            return None
