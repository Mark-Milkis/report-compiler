"""
PDF content analysis and cropping utilities.
"""

from typing import Optional, Dict, Any
import fitz  # PyMuPDF
from ..core.config import Config
from ..utils.conversions import points_to_inches
from ..utils.logging_config import get_module_logger


class ContentAnalyzer:
    """Handles PDF content detection and analysis."""

    def __init__(self):
        self.logger = get_module_logger(__name__)

    def _expected_markers(
        self, placeholders: dict[str, Any], table_metadata: Optional[Dict[int, Any]]
    ) -> dict[str, dict[str, Any]]:
        """Build the set of markers we expect to find, keyed by marker string.

        Each value carries everything needed to construct the content-map entry
        once the marker's location is known. Table placeholders contribute one
        marker per source page (matching the row replication done in the DOCX).
        """
        expected: dict[str, dict[str, Any]] = {}

        for placeholder in placeholders.get('paragraph', []):
            marker = Config.get_merge_marker(placeholder['paragraph_index'])
            expected[marker] = {'placeholder': placeholder, 'is_table': False}

        for placeholder in placeholders.get('table', []):
            num_pages = placeholder.get('page_count', 1)
            for page_num in range(1, num_pages + 1):
                marker = Config.get_overlay_marker(placeholder['table_index'], page_num)
                entry = {'placeholder': placeholder, 'is_table': True, 'overlay_page_num': page_num}
                if table_metadata:
                    entry['table_dims'] = table_metadata.get(placeholder['table_index'], {})
                expected[marker] = entry

        return expected

    def get_content_bbox(self, pdf_page: fitz.Page) -> Optional[fitz.Rect]:
        """
        Get the bounding box of actual content (excluding margins) by detecting text, images, and drawings.
        
        Args:
            pdf_page: PyMuPDF page object
            
        Returns:
            fitz.Rect: Bounding box of content, or None if no content found
        """
        content_bbox = None
        try:
            # Combine text, drawings, and images to find the total content area
            paths = pdf_page.get_drawings()
            text_blocks = pdf_page.get_text("dict")["blocks"]
            image_blocks = pdf_page.get_images(full=True)

            if not paths and not text_blocks and not image_blocks:
                self.logger.debug("      - Page has no content to analyze.")
                return None

            for path in paths:
                if content_bbox is None:
                    content_bbox = path["rect"]
                else:
                    content_bbox.include_rect(path["rect"])

            for block in text_blocks:
                if "bbox" in block:
                    if content_bbox is None:
                        content_bbox = fitz.Rect(block["bbox"])
                    else:
                        content_bbox.include_rect(fitz.Rect(block["bbox"]))

            for img in image_blocks:
                img_rect = pdf_page.get_image_bbox(img)
                if img_rect:
                    if content_bbox is None:
                        content_bbox = img_rect
                    else:
                        content_bbox.include_rect(img_rect)

        except Exception as e:
            self.logger.warning("    ⚠️ Error detecting content bbox: %s", e)
            return None
        return content_bbox

    def apply_content_cropping(
        self, pdf_page: fitz.Page, crop_enabled: bool = True, padding: Optional[int] = None
    ) -> fitz.Rect:
        """
        Crop PDF page to content boundaries with border-preserving padding, or return full page.
        
        Args:
            pdf_page: PyMuPDF page object
            crop_enabled: Whether to enable content cropping (default: True)
            padding: Padding around content in points (default: from Config.DEFAULT_PADDING)
            
        Returns:
            fitz.Rect: Content rectangle to use for clipping
        """
        if padding is None:
            padding = Config.DEFAULT_PADDING

        if not crop_enabled:
            self.logger.debug("      - Content cropping disabled, using full page.")
            return pdf_page.rect

        content_bbox = self.get_content_bbox(pdf_page)

        if content_bbox is None or content_bbox.is_empty or content_bbox.is_infinite:
            self.logger.debug("      - No valid content found for cropping, using full page.")
            return pdf_page.rect

        # Apply padding
        padded_rect = fitz.Rect(
            content_bbox.x0 - padding,
            content_bbox.y0 - padding,
            content_bbox.x1 + padding,
            content_bbox.y1 + padding,
        )

        # Ensure the padded rectangle does not exceed the page boundaries
        final_rect = padded_rect & pdf_page.rect
        self.logger.debug("      - Original content box: (%.2f, %.2f) to (%.2f, %.2f) inches",
                         points_to_inches(content_bbox.x0), points_to_inches(content_bbox.y0),
                         points_to_inches(content_bbox.x1), points_to_inches(content_bbox.y1))
        self.logger.debug("      - Final cropped area with padding: %.2f\" x %.2f\"",
                         points_to_inches(final_rect.width), points_to_inches(final_rect.height))

        return final_rect

    def bake_annotations(self, pdf_doc: fitz.Document):
        """Applies all annotations (comments, highlights) permanently to the pages.

        ``Document.bake`` is a whole-document operation, so it must be called
        exactly once. Calling it inside a per-page loop bakes the entire
        document N times (O(N^2)) and was the dominant cost for large reports.
        """
        self.logger.debug("  > Baking annotations for %d pages...", len(pdf_doc))
        pdf_doc.bake(annots=True)  # Apply all annotations across the whole document

    def analyze(self, pdf_doc: fitz.Document, placeholders: dict[str, Any], table_metadata: dict[int, Any]) -> Optional[dict[str, Any]]:
        """
        Locate every placeholder marker in the (already open) base PDF.

        This is a single sweep over the document: each page is visited once and
        searched for any markers not yet found, short-circuiting as soon as all
        expected markers are located. This replaces the previous approach which
        opened the PDF twice and scanned the whole document once per marker
        (O(markers x pages)), plus a separate full-text sweep for a Table of
        Contents that nothing downstream consumed.

        Args:
            pdf_doc: An open PyMuPDF document for the base PDF.
            placeholders: Dictionary of placeholders.
            table_metadata: Dictionary of table metadata.

        Returns:
            A content_map dictionary mapping each found marker to its location
            data, or None on failure.
        """
        self.logger.info("  > Starting PDF analysis...")

        try:
            pending = self._expected_markers(placeholders, table_metadata)
            content_map: dict[str, Any] = {}

            for page_index, page in enumerate(pdf_doc):
                if not pending:
                    break  # Every expected marker has been located.
                for marker in list(pending.keys()):
                    rects = page.search_for(marker)
                    if not rects:
                        continue
                    rect = rects[0]
                    info = pending.pop(marker)
                    self.logger.debug("    - Found marker '%s' on page %d at (%.2f, %.2f) inches.",
                                     marker, page_index + 1,
                                     points_to_inches(rect.x0), points_to_inches(rect.y0))
                    map_entry = {
                        'placeholder': info['placeholder'],
                        'page_index': page_index,
                        'rect': [rect.x0, rect.y0, rect.x1, rect.y1],
                        'type': info['placeholder']['type'],
                    }
                    if info['is_table']:
                        if 'table_dims' in info:
                            map_entry['table_dims'] = info['table_dims']
                        map_entry['overlay_page_num'] = info['overlay_page_num']
                    content_map[marker] = map_entry

            for marker in pending:
                self.logger.warning("    - ⚠️ Marker '%s' not found in the PDF.", marker)

            if not content_map and placeholders.get('total', 0) > 0:
                # This can happen if the DOCX modification failed to insert markers, or if they were removed during PDF conversion.
                self.logger.warning("  > ⚠️ No markers were found in the PDF, but placeholders were expected. Downstream processing may fail.")

            self.logger.info("  > PDF analysis complete (%d markers located).", len(content_map))
            return content_map
        except Exception as e:
            self.logger.error("❌ Top-level error during PDF analysis: %s", e, exc_info=True)
            return None
