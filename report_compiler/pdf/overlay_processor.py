"""
PDF overlay processing for table-based insertions.
"""

import fitz  # PyMuPDF
from typing import Dict, List, Any, Optional
from ..core.config import Config
from ..utils.conversions import points_to_inches
from ..utils.page_selector import PageSelector
from ..utils.logging_config import get_overlay_logger
from .content_analyzer import ContentAnalyzer
from .marker_remover import MarkerRemover


class OverlayProcessor:
    """Handles table-based PDF overlay operations."""

    def __init__(self):
        self.page_selector = PageSelector()
        self.content_analyzer = ContentAnalyzer()
        self.marker_remover = MarkerRemover()
        self.logger = get_overlay_logger()
        self.logger.debug("PyMuPDF (fitz) version: %s, path: %s", fitz.__version__, fitz.__file__)

    def process_overlays(self, base_pdf_path: str, content_map: Dict[str, Any], output_path: str) -> bool:
        """
        Process all overlay placeholders in the base PDF.

        Args:
            base_pdf_path: Path to base PDF document.
            content_map: Dictionary mapping markers to their location and metadata.
            output_path: Path for output PDF.

        Returns:
            bool: True if successful, False otherwise.
        """
        overlay_markers = {
            marker: data for marker, data in content_map.items()
            if data['type'] == 'table'
        }

        if not overlay_markers:
            self.logger.info("No overlay placeholders to process.")
            # If no overlays, the output is just a copy of the input
            # This is handled by the compiler's file management.
            return True

        try:
            self.logger.debug("Opening base PDF: %s", base_pdf_path)
            with fitz.open(base_pdf_path) as base_doc:
                for idx, (marker, data) in enumerate(overlay_markers.items(), 1):
                    if not self._process_single_overlay(base_doc, marker, data, idx):
                        return False

                self.logger.debug("Saving overlaid PDF to: %s", output_path)
                base_doc.save(output_path)
                self.logger.info("✓ Overlaid PDF saved successfully.")

            return True

        except Exception as e:
            self.logger.error("❌ Error during overlay processing: %s", e, exc_info=True)
            return False

    def _process_single_overlay(self, base_doc: fitz.Document, marker: str,
                               data: Dict[str, Any], idx: int) -> bool:
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

            # Open source PDF
            self.logger.debug("    > Opening source PDF: %s", pdf_path)
            with fitz.open(pdf_path) as source_doc:
                self.content_analyzer.bake_annotations(source_doc)

                # Determine which pages from the source PDF are requested
                page_spec = placeholder.get('page_spec')
                page_selection = self.page_selector.parse_specification(page_spec)
                selected_source_pages = self.page_selector.apply_selection(source_doc, page_selection)
                if not selected_source_pages:
                    # If no spec, assume all pages
                    selected_source_pages = list(range(len(source_doc)))
                
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

                crop_rect = self.content_analyzer.apply_content_cropping(source_page, crop_enabled)
                
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
