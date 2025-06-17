"""
Marker removal utilities for PDF processing using redaction.
"""

import fitz  # PyMuPDF
from typing import Dict, Optional


class MarkerRemover:
    """Handles clean removal of marker text from PDF pages using redaction."""

    def __init__(self):
        pass

    def remove_marker_text(self, page: fitz.Page, marker_text: str) -> bool:
        """
        Remove marker text from a PDF page using redaction.

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
                # print(f"        ⚠️ Marker '{marker_text}' not found on page.")
                return False

            print(f"        🎯 Applying redaction for marker text '{marker_text}' at ({marker_rect.x0:.1f}, {marker_rect.y0:.1f})")

            # Add redaction annotation
            page.add_redact_annot(marker_rect)
            
            # Apply redaction (removes the text)
            page.apply_redactions()
            
            return True
            
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

            # Basic position information
            position_info = {
                "rect": marker_rect,  # Add the fitz.Rect object itself
                "x0": marker_rect.x0,
                "y0": marker_rect.y0,
                "x1": marker_rect.x1,
                "y1": marker_rect.y1,
                "width": marker_rect.width,
                "height": marker_rect.height,
                "center_x": (marker_rect.x0 + marker_rect.x1) / 2,
                "center_y": (marker_rect.y0 + marker_rect.y1) / 2,
                "page_width": page.rect.width,
                "page_height": page.rect.height,
                "position_inches": (marker_rect.x0 / 72, marker_rect.y0 / 72), # Added
                "size_inches": (marker_rect.width / 72, marker_rect.height / 72) # Added
            }
            
            return position_info
            
        except Exception as e: # It's good practice to log or print the exception
            print(f"      ⚠️ Error in find_marker_position: {e}")
            return None
