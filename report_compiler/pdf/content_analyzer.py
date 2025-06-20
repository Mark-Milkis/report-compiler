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

    def find_toc_pages(self, pdf_path: str) -> list[int]:
        """Identifies pages that appear to be a Table of Contents."""
        toc_pages = []
        try:
            self.logger.debug("  > Scanning for Table of Contents...")
            doc = fitz.open(pdf_path)
            # A simple heuristic: look for pages with the title "Table of Contents"
            for i, page in enumerate(doc):
                if "table of contents" in page.get_text().lower():
                    toc_pages.append(i)
            return toc_pages
        except Exception as e:
            self.logger.error("  > ❌ Error analyzing for ToC pages: %s", e, exc_info=True)
            return []

    def find_placeholder_markers(
        self, pdf_path: str, placeholders: dict[str, Any], table_metadata: dict[int, Any]
    ) -> dict[str, Any]:
        """
        Scans the PDF for all placeholder markers and records their locations and metadata.
        Handles multi-page overlays by searching for each page's unique marker.

        Returns:
            A content_map dictionary mapping a marker string to its location data.
        """
        content_map = {}
        self.logger.debug("  > Scanning PDF for placeholder markers...")
        try:
            doc = fitz.open(pdf_path)
            
            # Process paragraph placeholders
            for placeholder in placeholders.get('paragraph', []):
                self._find_and_map_marker(doc, placeholder, content_map)

            # Process table placeholders, handling multi-page overlays
            for placeholder in placeholders.get('table', []):
                num_pages = placeholder.get('page_count', 1)
                
                for page_num in range(1, num_pages + 1):
                    # For each page of the overlay, find its corresponding marker
                    self._find_and_map_marker(doc, placeholder, content_map, table_metadata, page_num)

            self.logger.debug("  > Found %d markers in total.", len(content_map))
            return content_map
        except Exception as e:
            self.logger.error("❌ Error finding placeholder markers in PDF: %s", e, exc_info=True)
            return {}

    def _find_and_map_marker(self, doc: fitz.Document, placeholder: dict[str, Any], content_map: dict[str, Any], table_metadata: Optional[Dict[int, Any]] = None, page_num: int = 1):
        """Helper to find a single marker and add it to the content map."""
        marker = self._get_marker_for_placeholder(placeholder, page_num)
        if not marker:
            return

        found_marker = self._search_for_marker_in_doc(doc, marker)
        if found_marker:
            page_index, rect = found_marker
            self.logger.debug("    - Found marker '%s' on page %d at (%.2f, %.2f) inches.",
                             marker, page_index + 1, points_to_inches(rect.x0), points_to_inches(rect.y0))
            
            map_entry = {
                'placeholder': placeholder,
                'page_index': page_index,
                'rect': [rect.x0, rect.y0, rect.x1, rect.y1],
                'type': placeholder['type']
            }
            if placeholder['type'] == 'table':
                if table_metadata:
                    map_entry['table_dims'] = table_metadata.get(placeholder['table_index'], {})
                map_entry['overlay_page_num'] = page_num # Add the overlay page number
            content_map[marker] = map_entry
        else:
            self.logger.warning("    - ⚠️ Marker '%s' not found in the PDF.", marker)

    def _get_marker_for_placeholder(self, placeholder: dict[str, Any], page_num: int = 1) -> Optional[str]:
        if placeholder.get('type') == 'table':
            return Config.get_overlay_marker(placeholder['table_index'], page_num)
        elif placeholder.get('type') == 'paragraph':
            return Config.get_merge_marker(placeholder['paragraph_index'])
        return None

    def _search_for_marker_in_doc(self, doc: fitz.Document, marker: str) -> Optional[tuple[int, fitz.Rect]]:
        for page_index, page in enumerate(doc):
            rects = page.search_for(marker)
            if rects:
                return page_index, rects[0]  # Return page and rect of first find
        return None

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
        """Applies all annotations (comments, highlights) permanently to the pages."""
        self.logger.debug("  > Baking annotations for %d pages...", len(pdf_doc))
        for page in pdf_doc:
            pdf_doc.bake(annots=True)  # Apply all annotations to the page
            self.logger.debug("    - Baked annotations for page %d", page.number + 1)


            # for annot in page.annots():
            #     # Applying the annotation renders it onto the page content
            #     annot.update(flags=fitz.ANNOT_FLAG_PRINT)
            #     # Deleting the annotation removes the interactive element
            #     page.delete_annot(annot)

    def analyze(self, pdf_path: str, placeholders: dict[str, Any], table_metadata: dict[int, Any]) -> Optional[tuple[dict[str, Any], list[int]]]:
        """
        Analyzes the PDF to find all markers and the table of contents.

        Args:
            pdf_path: Path to the PDF file.
            placeholders: Dictionary of placeholders.
            table_metadata: Dictionary of table metadata.

        Returns:
            A tuple containing the content_map and a list of TOC page indices, or None on failure.
        """
        self.logger.info("  > Starting PDF analysis...")
        
        try:
            toc_pages = self.find_toc_pages(pdf_path)
            self.logger.info("  > Found %d Table of Contents pages.", len(toc_pages))

            content_map = self.find_placeholder_markers(pdf_path, placeholders, table_metadata)
            if not content_map and placeholders.get('total', 0) > 0:
                # This can happen if the DOCX modification failed to insert markers, or if they were removed during PDF conversion.
                self.logger.warning("  > ⚠️ No markers were found in the PDF, but placeholders were expected. Downstream processing may fail.")

            self.logger.info("  > PDF analysis complete.")
            return content_map, toc_pages
        except Exception as e:
            self.logger.error("❌ Top-level error during PDF analysis: %s", e, exc_info=True)
            return None
