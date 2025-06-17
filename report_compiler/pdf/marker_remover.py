"""
Marker removal utilities for PDF processing using redaction.
"""

import fitz  # PyMuPDF
from typing import Dict, Optional
from ..utils.logging_config import get_module_logger


class MarkerRemover:
    """Handles clean removal of marker text from PDF pages using redaction."""

    def __init__(self):
        self.logger = get_module_logger(__name__)

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
                self.logger.debug("        ⚠️ Marker '%s' not found on page.", marker_text)
                return False

            self.logger.debug("        🎯 Applying redaction for marker text '%s' at (%.1f, %.1f)", 
                            marker_text, marker_rect.x0, marker_rect.y0)

            # Add redaction annotation
            page.add_redact_annot(marker_rect)
            
            # Apply redaction (removes the text)
            page.apply_redactions()
            
            return True
            
        except Exception as e:
            self.logger.warning("        ⚠️ Error removing marker: %s", e)
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
        Used by OverlayProcessor.

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

            position_info = {
                "rect": marker_rect,
                "position_inches": (marker_rect.x0 / 72, marker_rect.y0 / 72),
                "size_inches": (marker_rect.width / 72, marker_rect.height / 72)
            }
            
            return position_info
        except Exception as e: # It's good practice to log or print the exception
            self.logger.warning("      ⚠️ Error in find_marker_position: %s", e)
            return None
