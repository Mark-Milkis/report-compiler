"""
PDF overlay processing for table-based insertions.
"""

import fitz  # PyMuPDF
from typing import Dict, List, Any, Optional
from ..core.config import Config
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
    
    def process_overlays(self, base_pdf_path: str, overlay_placeholders: List[Dict], 
                        table_metadata: Dict[int, Dict[str, float]], 
                        output_path: str) -> bool:
        """
        Process all overlay placeholders in the base PDF.
        
        Args:
            base_pdf_path: Path to base PDF document
            overlay_placeholders: List of overlay placeholder dictionaries
            table_metadata: Table dimension metadata from DocxProcessor
            output_path: Path for output PDF
            
        Returns:
            bool: True if successful, False otherwise        """
        if not overlay_placeholders:
            return True
        
        try:
            self.logger.info("   • Processing %d table-based overlay(s)...", len(overlay_placeholders))
            self.logger.info("    Opening base PDF: %s", base_pdf_path)
            
            with fitz.open(base_pdf_path) as base_doc:
                for idx, placeholder in enumerate(overlay_placeholders, 1):
                    if not self._process_single_overlay(base_doc, placeholder, table_metadata, idx):
                        return False
                
                # Save the final PDF
                self.logger.info("    Saving final PDF: %s", output_path)
                base_doc.save(output_path)
                self.logger.info("    ✓ Final PDF saved successfully")
            
            self.logger.info("   ✓ Overlay processing complete")
            return True
            
        except Exception as e:
            self.logger.error("   ❌ Error processing overlays: %s", e, exc_info=True)
            return False
    
    def _process_single_overlay(self, base_doc: fitz.Document, placeholder: Dict[str, Any], 
                               table_metadata: Dict[int, Dict[str, float]], idx: int) -> bool:
        """
        Process a single overlay placeholder.
        
        Args:
            base_doc: Base PDF document
            placeholder: Overlay placeholder dictionary
            table_metadata: Table dimension metadata
            idx: Placeholder index for naming
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            pdf_path = placeholder['resolved_path']
            table_idx = placeholder['table_index']
            page_count = placeholder['page_count']
            crop_enabled = placeholder.get('crop_enabled', True)            
            self.logger.info("    Processing appendix %d: %s", idx, placeholder['pdf_path_raw'])
            
            # Find main marker position
            main_marker = Config.get_overlay_marker(table_idx)
            overlay_rect, start_page_index = self._find_and_remove_marker(
                base_doc, main_marker, table_metadata.get(table_idx, {}))
            
            if overlay_rect is None:
                self.logger.error("      ❌ Could not find positioning marker")
                return False
            
            self.logger.info("      ✓ Found positioning markers on page %d", start_page_index + 1)
            self.logger.info("      📋 Overlay rectangle: (%.1f, %.1f) to (%.1f, %.1f)", 
                           overlay_rect.x0, overlay_rect.y0, overlay_rect.x1, overlay_rect.y1)
            self.logger.info("      📏 Overlay size: %.1f x %.1f points", overlay_rect.width, overlay_rect.height)
            
            # Open source PDF
            self.logger.info("      Opening appendix PDF: %s", pdf_path)
            with fitz.open(pdf_path) as source_doc:
                # Bake annotations
                self.content_analyzer.bake_annotations(source_doc)
                  # Determine pages to use
                page_selection = self.page_selector.parse_specification(placeholder.get('page_spec'))
                pages_to_overlay = self.page_selector.apply_selection(source_doc, page_selection)
                if not pages_to_overlay:
                    pages_to_overlay = list(range(len(source_doc)))
                
                self.logger.info("        📄 Using %d pages", len(pages_to_overlay))
                  # Overlay each page
                current_page_index = start_page_index
                
                for i, source_page_idx in enumerate(pages_to_overlay, 1):
                    source_page = source_doc[source_page_idx]
                    
                    # Apply content cropping
                    crop_rect = self.content_analyzer.apply_content_cropping(
                        source_page, crop_enabled)
                    
                    if i == 1:
                        # First page goes to the main marker position
                        self.logger.info("        Overlaying source page %d -> base page %d (position %d/%d)", 
                                        source_page_idx + 1, current_page_index + 1, i, len(pages_to_overlay))
                        self._overlay_page_content(base_doc[current_page_index], source_page, 
                                                 overlay_rect, crop_rect)
                        self.logger.info("        📌 Precise overlay within detected cell boundaries")
                    else:
                        # Additional pages need to find their specific markers
                        page_marker = Config.get_overlay_marker(table_idx, i)
                        self.logger.info("        Overlaying source page %d -> searching for marker %s (position %d/%d)", 
                                        source_page_idx + 1, page_marker, i, len(pages_to_overlay))
                        
                        marker_rect, marker_page_idx = self._find_and_remove_marker(
                            base_doc, page_marker, table_metadata.get(table_idx, {}))
                        
                        if marker_rect:
                            self._overlay_page_content(base_doc[marker_page_idx], source_page, 
                                                     marker_rect, crop_rect)
                            self.logger.info("        📌 Precise overlay in replicated table on page %d", marker_page_idx + 1)
                        else:
                            self.logger.warning("        ⚠️ Could not find marker for page %d", i)
            
            self.logger.info("      ✓ Appendix %d overlay complete", idx)
            return True
            
        except Exception as e:
            self.logger.error("      ❌ Error processing overlay %d: %s", idx, e, exc_info=True)
            return False
    
    def _find_and_remove_marker(self, pdf_doc: fitz.Document, marker: str, 
                               table_dims: Dict[str, float]) -> tuple:
        """
        Find marker position and remove it, returning overlay rectangle.
        
        Args:
            pdf_doc: PDF document to search
            marker: Marker text to find
            table_dims: Table dimensions metadata
            
        Returns:
            Tuple of (overlay_rect, page_index) or (None, None) if not found        """
        self.logger.info("      🔍 Searching for marker: %s", marker)
        
        for page_index in range(len(pdf_doc)):
            page = pdf_doc[page_index]
            
            # Find marker position
            marker_info = self.marker_remover.find_marker_position(page, marker)
            if marker_info:
                marker_rect = marker_info['rect']
                
                self.logger.info("      📍 Marker found at: (%.1f, %.1f) points = (%.2f, %.2f) inches", 
                                marker_rect.x0, marker_rect.y0, 
                                marker_info['position_inches'][0], marker_info['position_inches'][1])
                self.logger.info("      📍 Marker size: %.1f x %.1f points = %.2f x %.2f inches", 
                                marker_rect.width, marker_rect.height, 
                                marker_info['size_inches'][0], marker_info['size_inches'][1])
                
                # Calculate overlay rectangle based on table dimensions
                table_width_pts = table_dims.get('width_pts', 540)  # Default 7.5 inches
                table_height_pts = table_dims.get('height_pts', 288)  # Default 4 inches
                
                self.logger.info("      📏 Table dimensions: %.1f x %.1f points = %.2f x %.2f inches", 
                                table_width_pts, table_height_pts, 
                                table_dims.get('width_inches', 7.5), table_dims.get('height_inches', 4.0))
                
                overlay_rect = fitz.Rect(
                    marker_rect.x0,  # left (marker x position)
                    marker_rect.y0,  # top (marker y position)  
                    marker_rect.x0 + table_width_pts,  # right (left + table width)
                    marker_rect.y0 + table_height_pts  # bottom (top + table height)
                )
                  # Display calculated overlay rectangle
                overlay_x_in = overlay_rect.x0 / 72
                overlay_y_in = overlay_rect.y0 / 72
                overlay_width_in = overlay_rect.width / 72
                overlay_height_in = overlay_rect.height / 72
                
                self.logger.info("      📐 Calculated overlay rectangle:")
                self.logger.info("         • Points: (%.1f, %.1f) to (%.1f, %.1f)", 
                                overlay_rect.x0, overlay_rect.y0, overlay_rect.x1, overlay_rect.y1)
                self.logger.info("         • Inches: (%.2f, %.2f) to (%.2f, %.2f)", 
                                overlay_x_in, overlay_y_in, overlay_x_in + overlay_width_in, overlay_y_in + overlay_height_in)
                self.logger.info("         • Size: %.1f x %.1f points = %.2f x %.2f inches", 
                                overlay_rect.width, overlay_rect.height, overlay_width_in, overlay_height_in)
                
                # Remove the marker text
                if self.marker_remover.remove_marker_text(page, marker):
                    self.logger.info("      ✓ Removed marker text from page %d", page_index + 1)
                else:
                    self.logger.warning("      ⚠️ Could not remove marker text from page %d", page_index + 1)
                
                return overlay_rect, page_index
        
        return None, None
    
    def _overlay_page_content(self, base_page: fitz.Page, source_page: fitz.Page, 
                             overlay_rect: fitz.Rect, crop_rect: fitz.Rect) -> None:
        """
        Overlay source page content onto base page at specified position.
        
        Args:
            base_page: Target page in base document
            source_page: Source page to overlay
            overlay_rect: Rectangle where content should be placed
            crop_rect: Rectangle of content to use from source page
        """
        try:
            # Create transformation matrix to fit source content into overlay rectangle
            # Scale and translate source content to fit the overlay area
            
            # Calculate scale factors
            scale_x = overlay_rect.width / crop_rect.width
            scale_y = overlay_rect.height / crop_rect.height
            
            # Use the smaller scale to maintain aspect ratio
            scale = min(scale_x, scale_y)
            
            # Calculate translation to center content in overlay rectangle
            scaled_width = crop_rect.width * scale
            scaled_height = crop_rect.height * scale
            
            translate_x = overlay_rect.x0 + (overlay_rect.width - scaled_width) / 2 - crop_rect.x0 * scale
            translate_y = overlay_rect.y0 + (overlay_rect.height - scaled_height) / 2 - crop_rect.y0 * scale
              # Create transformation matrix
            transform = fitz.Matrix(scale, scale).pretranslate(translate_x, translate_y)
            
            # Show the source page on the base page with transformation
            try:
                # Try with matrix parameter (newer PyMuPDF versions)
                base_page.show_pdf_page(overlay_rect, source_page.parent, source_page.number, 
                                       clip=crop_rect, matrix=transform)
            except TypeError:
                # Fallback for older PyMuPDF versions without matrix parameter
                base_page.show_pdf_page(overlay_rect, source_page.parent, source_page.number, 
                                       clip=crop_rect)
        except Exception as e:
            self.logger.warning("        ⚠️ Error overlaying content: %s", e)
            # Fallback: simple overlay without transformation
            try:
                base_page.show_pdf_page(overlay_rect, source_page.parent, source_page.number)
            except Exception as fallback_error:
                self.logger.warning("        ⚠️ Fallback also failed: %s", fallback_error)
