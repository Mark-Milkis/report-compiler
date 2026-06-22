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

    def remove_markers(self, input_pdf_path: str, markers: list[str], output_pdf_path: str,
                       marker_pages: Optional[Dict[str, int]] = None) -> bool:
        """
        Removes all specified markers from the PDF by redacting each marker text.

        Args:
            input_pdf_path: Path to the input PDF file.
            markers: A list of marker strings to remove.
            output_pdf_path: Path to save the cleaned PDF file.
            marker_pages: Optional mapping of marker -> known 0-based page index in
                the input PDF. When supplied (and complete), only those pages are
                searched instead of scanning every page for every marker, which is
                far cheaper on large documents. Falls back to a full scan if the
                map is missing or incomplete.

        Returns:
            True if successful, False otherwise.
        """
        self.logger.debug("      Removing %d markers from '%s'...", len(markers), input_pdf_path)
        try:
            pdf_document = fitz.open(input_pdf_path)
            # Redaction options that leave images and vector line-art untouched.
            # The markers are tiny text strings; without these flags every
            # apply_redactions() call reprocesses all images/graphics on the page,
            # which is extremely slow on pages carrying overlaid CAD content and can
            # also corrupt that content.
            redact_kwargs = self._redaction_kwargs()

            pages_to_markers = self._group_markers_by_page(markers, marker_pages, len(pdf_document))

            if pages_to_markers is not None:
                # Targeted: only touch the specific pages we know hold markers.
                for page_idx in sorted(pages_to_markers):
                    self._redact_markers_on_page(
                        pdf_document[page_idx], pages_to_markers[page_idx], redact_kwargs
                    )
            else:
                # Fallback: scan every page for every marker.
                for page in pdf_document:
                    self._redact_markers_on_page(page, markers, redact_kwargs)

            pdf_document.save(output_pdf_path, garbage=4, deflate=True, clean=True)
            pdf_document.close()
            self.logger.debug("      Markers removed. Cleaned PDF saved to '%s'", output_pdf_path)
            return True
        except Exception as e:
            self.logger.error("      ❌ Error during marker removal: %s", e, exc_info=True)
            return False

    @staticmethod
    def _group_markers_by_page(markers: list[str], marker_pages: Optional[Dict[str, int]],
                               page_count: int) -> Optional[Dict[int, list[str]]]:
        """Group markers by their known page index, or return None to force a full scan.

        Returns None if no page map was provided or any marker's page is unknown or
        out of range, so the caller falls back to the safe full-document scan.
        """
        if not marker_pages:
            return None

        grouped: Dict[int, list[str]] = {}
        for marker in markers:
            page_idx = marker_pages.get(marker)
            if page_idx is None or not (0 <= page_idx < page_count):
                return None
            grouped.setdefault(page_idx, []).append(marker)
        return grouped

    def _redact_markers_on_page(self, page: fitz.Page, markers: list[str], redact_kwargs: dict):
        """Search for each marker on a single page and redact them in one pass.

        apply_redactions() rewrites the page content stream, so it is called at
        most once per page rather than once per marker.
        """
        found_any = False
        for marker in markers:
            rects = page.search_for(marker)
            for inst in rects:
                page.add_redact_annot(inst)
            if rects:
                found_any = True
                self.logger.debug("        - Redacted marker '%s' on page %d.", marker, page.number + 1)

        if found_any:
            page.apply_redactions(**redact_kwargs)

    @staticmethod
    def _redaction_kwargs() -> dict:
        """Build apply_redactions kwargs that skip image/line-art reprocessing.

        These flags exist on modern PyMuPDF (>= 1.24). Fall back gracefully if
        a constant is unavailable so older builds still work.
        """
        kwargs = {}
        image_none = getattr(fitz, "PDF_REDACT_IMAGE_NONE", None)
        if image_none is not None:
            kwargs["images"] = image_none
        graphics_none = getattr(fitz, "PDF_REDACT_LINE_ART_NONE", None)
        if graphics_none is not None:
            kwargs["graphics"] = graphics_none
        return kwargs
