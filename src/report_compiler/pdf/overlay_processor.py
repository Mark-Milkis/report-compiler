"""
PDF overlay processing for table-based insertions.
"""

import fitz  # PyMuPDF
from typing import Dict, List, Any
from ..utils.conversions import points_to_inches
from ..utils.page_selector import PageSelector
from ..utils.logging_config import get_overlay_logger
from .content_analyzer import ContentAnalyzer


class OverlayProcessor:
    """Handles table-based PDF overlay operations."""

    def __init__(self):
        self.page_selector = PageSelector()
        self.content_analyzer = ContentAnalyzer()
        self.logger = get_overlay_logger()
        # Opened source documents are kept alive here until close_sources() is
        # called. show_pdf_page() can keep referencing the source document until
        # the target is saved, so the caller must close these only after the
        # final save of the base document.
        self._source_doc_cache: Dict[str, fitz.Document] = {}
        self.logger.debug("PyMuPDF (fitz) version: %s, path: %s", fitz.__version__, fitz.__file__)

    def process_overlays(self, base_doc: fitz.Document, content_map: Dict[str, Any]) -> bool:
        """
        Process all overlay placeholders directly on the open base document.

        The document is modified in place and is neither opened nor saved here;
        the caller owns its lifecycle so that overlay, merge, and marker removal
        can share a single open document and a single final save.

        Args:
            base_doc: The open base PDF document to overlay content onto.
            content_map: Dictionary mapping markers to their location and metadata.

        Returns:
            bool: True if successful, False otherwise.
        """
        overlay_markers = {
            marker: data for marker, data in content_map.items()
            if data['type'] == 'table'
        }

        if not overlay_markers:
            self.logger.info("No overlay placeholders to process.")
            return True

        # Cache of opened+baked source documents keyed by resolved path. The same
        # source PDF is typically referenced by many markers (one per page), so
        # opening and baking it once and reusing it avoids O(markers x pages) work.
        # Kept open until close_sources() so the base document can be saved first.
        source_doc_cache = self._source_doc_cache
        # Cache of computed content-crop rectangles keyed by (path, source_page_idx).
        crop_rect_cache: Dict[Any, fitz.Rect] = {}
        # Cache of resolved source-page selections keyed by (path, page_spec). The
        # spec is identical for every page-marker of the same overlay table.
        selection_cache: Dict[Any, List[int]] = {}

        try:
            for idx, (marker, data) in enumerate(overlay_markers.items(), 1):
                if not self._process_single_overlay(
                    base_doc, marker, data, idx,
                    source_doc_cache, crop_rect_cache, selection_cache
                ):
                    return False

            self.logger.info("✓ Overlays applied successfully.")
            return True

        except Exception as e:
            self.logger.error("❌ Error during overlay processing: %s", e, exc_info=True)
            return False

    def close_sources(self) -> None:
        """Close all cached overlay source documents.

        Must be called only after the base document has been saved, because
        show_pdf_page() may reference these sources until that save completes.
        """
        for src in self._source_doc_cache.values():
            try:
                src.close()
            except Exception:
                pass
        self._source_doc_cache.clear()

    def _get_source_doc(self, pdf_path: str,
                        source_doc_cache: Dict[str, fitz.Document]) -> fitz.Document:
        """Return an opened, annotation-baked source document, caching by path."""
        source_doc = source_doc_cache.get(pdf_path)
        if source_doc is None:
            self.logger.debug("    > Opening source PDF: %s", pdf_path)
            source_doc = fitz.open(pdf_path)
            self.content_analyzer.bake_annotations(source_doc)
            source_doc_cache[pdf_path] = source_doc
        return source_doc

    def _process_single_overlay(self, base_doc: fitz.Document, marker: str,
                               data: Dict[str, Any], idx: int,
                               source_doc_cache: Dict[str, fitz.Document],
                               crop_rect_cache: Dict[Any, fitz.Rect],
                               selection_cache: Dict[Any, List[int]]) -> bool:
        """
        Process a single overlay placeholder.
        """
        try:
            placeholder = data['placeholder']
            # Use the resolved absolute path, which is guaranteed by the compiler.
            pdf_path = placeholder['resolved_path']
            crop_enabled = placeholder.get('crop_enabled', True)
            # Log the original path for user-facing messages.
            self.logger.info("  Processing overlay %d: %s", idx, placeholder['file_path'])

            # The marker has already been found, its location is in `data`
            page_index = data['page_index']
            marker_rect = fitz.Rect(data['rect'])
            
            self.logger.debug("    > Marker found on page %d at (%.2f, %.2f) inches.",
                           page_index + 1,
                           points_to_inches(marker_rect.x0),
                           points_to_inches(marker_rect.y0))

            # Calculate overlay rectangle based on table dimensions from DOCX
            table_dims = data.get('table_dims', {})
            table_width_pts = table_dims.get('width_pts', 540)  # Default 7.5 inches
            table_height_pts = table_dims.get('height_pts', 288) # Default 4 inches
            
            overlay_rect = fitz.Rect(
                marker_rect.x0,
                marker_rect.y0,
                marker_rect.x0 + table_width_pts,
                marker_rect.y0 + table_height_pts
            )
            
            self.logger.debug("    > Calculated overlay area: %.2f\" x %.2f\"",
                            points_to_inches(overlay_rect.width),
                            points_to_inches(overlay_rect.height))

            # Open source PDF (cached + baked once per unique source file).
            source_doc = self._get_source_doc(pdf_path, source_doc_cache)

            # Determine which pages from the source PDF are requested. The spec is
            # the same for every page-marker of a table, so resolve it once.
            page_spec = placeholder.get('page_spec')
            selection_key = (pdf_path, page_spec)
            selected_source_pages = selection_cache.get(selection_key)
            if selected_source_pages is None:
                page_selection = self.page_selector.parse_specification(page_spec)
                selected_source_pages = self.page_selector.apply_selection(source_doc, page_selection)
                if not selected_source_pages:
                    # If no spec, assume all pages
                    selected_source_pages = list(range(len(source_doc)))
                selection_cache[selection_key] = selected_source_pages

            self.logger.debug("    > Source page selection spec '%s' resolved to %d pages.", page_spec, len(selected_source_pages))

            # Get which page of the overlay this marker represents (1-based)
            overlay_page_num = data.get('overlay_page_num', 1)

            # Check if the requested overlay page is valid
            if overlay_page_num > len(selected_source_pages):
                self.logger.error("  ❌ Marker %s requests overlay page %d, but source selection only has %d pages.",
                                  marker, overlay_page_num, len(selected_source_pages))
                return False

            # Get the specific source page index to overlay
            source_page_idx = selected_source_pages[overlay_page_num - 1]
            source_page = source_doc[source_page_idx]
            target_page = base_doc[page_index]

            self.logger.debug("      - Overlaying source page %d -> Base page %d", source_page_idx + 1, page_index + 1)

            # Content-cropping inspects every drawing/text/image on the page; cache
            # the result so repeated overlays of the same source page are free.
            crop_key = (pdf_path, source_page_idx, crop_enabled)
            crop_rect = crop_rect_cache.get(crop_key)
            if crop_rect is None:
                crop_rect = self.content_analyzer.apply_content_cropping(source_page, crop_enabled)
                crop_rect_cache[crop_key] = crop_rect

            self._overlay_page_content(target_page, source_page, overlay_rect, crop_rect)

            self.logger.info("    ✓ Overlay for %s complete.", placeholder['file_path'])
            return True

        except Exception as e:
            self.logger.error("  ❌ Error processing overlay %d: %s", idx, e, exc_info=True)
            return False

    def _overlay_page_content(self, base_page: fitz.Page, source_page: fitz.Page,
                             overlay_rect: fitz.Rect, crop_rect: fitz.Rect):
        """
        Overlay source page content onto base page, fitting it correctly.
        """
        try:
            self.logger.debug("      - Applying overlay to rect: (%.2f, %.2f) to (%.2f, %.2f) inches",
                             points_to_inches(overlay_rect.x0), points_to_inches(overlay_rect.y0),
                             points_to_inches(overlay_rect.x1), points_to_inches(overlay_rect.y1))
            self.logger.debug("      - Using source content from clip rect: (%.2f, %.2f) to (%.2f, %.2f) inches",
                             points_to_inches(crop_rect.x0), points_to_inches(crop_rect.y0),
                             points_to_inches(crop_rect.x1), points_to_inches(crop_rect.y1))

            # Use the built-in method to overlay the page, keeping proportions
            base_page.show_pdf_page(
                overlay_rect,           # The area on the base page to draw on
                source_page.parent,
                source_page.number,
                clip=crop_rect,         # The area of the source page to use
                keep_proportion=True,   # Maintain aspect ratio
                overlay=True            # Draw on top of existing content
            )
        except Exception as e:
            self.logger.error("    ❌ Overlay failed: %s", e, exc_info=True)
