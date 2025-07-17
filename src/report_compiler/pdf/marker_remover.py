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

    def remove_markers(self, input_pdf_path: str, markers: list[str], output_pdf_path: str) -> bool:
        """
        Removes all specified markers from the PDF by iterating through pages
        and redacting each marker text.

        Args:
            input_pdf_path: Path to the input PDF file.
            markers: A list of marker strings to remove.
            output_pdf_path: Path to save the cleaned PDF file.

        Returns:
            True if successful, False otherwise.
        """
        self.logger.debug("      Removing %d markers from '%s'...", len(markers), input_pdf_path)
        try:
            pdf_document = fitz.open(input_pdf_path)
            for page in pdf_document:
                for marker in markers:
                    self._remove_text_from_page(page, marker)
            
            pdf_document.save(output_pdf_path, garbage=4, deflate=True, clean=True)
            pdf_document.close()
            self.logger.debug("      Markers removed. Cleaned PDF saved to '%s'", output_pdf_path)
            return True
        except Exception as e:
            self.logger.error("      ❌ Error during marker removal: %s", e, exc_info=True)
            return False

    def _remove_text_from_page(self, page: fitz.Page, text_to_remove: str):
        """
        Remove specified text from a PDF page by applying redactions.

        Args:
            page: PyMuPDF page object.
            text_to_remove: Text to remove.
        """
        try:
            text_instances = page.search_for(text_to_remove)
            for inst in text_instances:
                page.add_redact_annot(inst)
            
            if text_instances:
                page.apply_redactions()
                self.logger.debug("        - Redacted marker '%s' on page %d.", text_to_remove, page.number + 1)

        except Exception as e:
            self.logger.warning("        ⚠️ Error removing text '%s': %s", text_to_remove, e)
