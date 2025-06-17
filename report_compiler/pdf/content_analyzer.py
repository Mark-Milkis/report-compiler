"""
PDF content analysis and cropping utilities.
"""

from typing import Optional, Dict, Any
import fitz  # PyMuPDF
from ..core.config import Config


class ContentAnalyzer:
    """Handles PDF content detection and cropping operations."""
    
    def __init__(self):
        pass
    
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
            # Get all text blocks
            text_blocks = pdf_page.get_text("dict")
            
            # Include text boundaries
            for block in text_blocks.get("blocks", []):
                if "lines" in block:
                    for line in block["lines"]:
                        for span in line.get("spans", []):
                            bbox = fitz.Rect(span["bbox"])
                            if content_bbox is None:
                                content_bbox = bbox
                            else:
                                content_bbox.include_rect(bbox)
            
            # Get all drawing objects (including borders, lines, shapes)
            try:
                drawings = pdf_page.get_drawings()
                for drawing in drawings:
                    bbox = fitz.Rect(drawing["rect"])
                    if content_bbox is None:
                        content_bbox = bbox
                    else:
                        content_bbox.include_rect(bbox)
            except:
                # get_drawings() might not be available in all PyMuPDF versions
                pass
            
            # Get all images
            try:
                images = pdf_page.get_images()
                for img in images:
                    # Get image bbox - format: (xref, smask, width, height, bpc, colorspace, alt, name, filter)
                    img_rect = pdf_page.get_image_bbox(img)
                    if img_rect:
                        if content_bbox is None:
                            content_bbox = img_rect
                        else:
                            content_bbox.include_rect(img_rect)
            except:
                # Fallback: get images without bbox if method not available
                pass
        
        except Exception as e:
            print(f"        ⚠️ Error detecting content bbox: {e}")
            return None
        
        return content_bbox
    
    def apply_content_cropping(self, pdf_page: fitz.Page, crop_enabled: bool = True, 
                              padding: int = None) -> fitz.Rect:
        """
        Crop PDF page to content boundaries with border-preserving padding, or return full page.
        
        Args:
            pdf_page: PyMuPDF page object
            crop_enabled: Whether to enable content cropping (default: True)
            padding: Padding around content in points (default: from Config)
            
        Returns:
            fitz.Rect: Content rectangle to use for clipping
        """
        if padding is None:
            padding = Config.DEFAULT_PADDING
            
        if not crop_enabled:
            print(f"        📐 Content cropping disabled, using full page ({pdf_page.rect.width / 72:.2f} x {pdf_page.rect.height / 72:.2f} inches)")
            return pdf_page.rect
        
        content_bbox = self.get_content_bbox(pdf_page)
        
        if content_bbox and not content_bbox.is_empty:
            # Add border-preserving padding around content
            # Use larger padding to ensure borders and thin lines are preserved
            border_padding = max(padding, Config.MIN_PADDING)  # Minimum padding to preserve borders
            
            content_bbox.x0 = max(0, content_bbox.x0 - border_padding)
            content_bbox.y0 = max(0, content_bbox.y0 - border_padding)
            content_bbox.x1 = min(pdf_page.rect.width, content_bbox.x1 + border_padding)
            content_bbox.y1 = min(pdf_page.rect.height, content_bbox.y1 + border_padding)
            
            # Convert to inches for display
            content_bbox_inches = (
                content_bbox.x0 / 72, content_bbox.y0 / 72,
                content_bbox.x1 / 72, content_bbox.y1 / 72
            )
            page_size_inches = (pdf_page.rect.width / 72, pdf_page.rect.height / 72)
            
            print(f"        📐 Content area: ({content_bbox_inches[0]:.2f}, {content_bbox_inches[1]:.2f}) to ({content_bbox_inches[2]:.2f}, {content_bbox_inches[3]:.2f}) inches")
            print(f"        📐 Original page: {page_size_inches[0]:.2f} x {page_size_inches[1]:.2f} inches")
            print(f"        📐 Using content-aware cropping (saves {((pdf_page.rect.width * pdf_page.rect.height) - (content_bbox.width * content_bbox.height)) / (pdf_page.rect.width * pdf_page.rect.height) * 100:.1f}% space)")
            
            return content_bbox
        else:
            print(f"        📐 No content detected or empty bbox, using full page")
            return pdf_page.rect
    
    def detect_annotations(self, pdf_doc: fitz.Document) -> Dict[str, any]:
        """
        Detect and analyze annotations in a PDF document.
        
        Args:
            pdf_doc: PyMuPDF document object
            
        Returns:
            Dict with annotation analysis results
        """
        total_annotations = 0
        pages_with_annotations = 0
        annotation_types = set()
        
        for page_num in range(len(pdf_doc)):
            page = pdf_doc[page_num]
            annotations = page.annots()
            
            if annotations:
                page_annot_count = len(list(annotations))
                if page_annot_count > 0:
                    total_annotations += page_annot_count
                    pages_with_annotations += 1
                    
                    # Get annotation types
                    for annot in page.annots():
                        annotation_types.add(annot.type[1])  # Get annotation type name
        
        return {
            'total_annotations': total_annotations,
            'pages_with_annotations': pages_with_annotations,
            'total_pages': len(pdf_doc),
            'annotation_types': list(annotation_types),
            'has_annotations': total_annotations > 0
        }
    
    def bake_annotations(self, pdf_doc: fitz.Document) -> bool:
        """
        Bake annotations into PDF content to preserve them during processing.
        
        Args:
            pdf_doc: PyMuPDF document object
            
        Returns:
            bool: True if annotations were found and baked, False otherwise
        """
        annotation_info = self.detect_annotations(pdf_doc)
        
        if annotation_info['has_annotations']:
            print(f"        📝 Found {annotation_info['total_annotations']} annotation(s), baking into content...")
            
            # Bake annotations into the content
            for page_num in range(len(pdf_doc)):
                page = pdf_doc[page_num]
                
                # This applies all annotations to the page content
                page.apply_redactions()
            
            print(f"        ✓ Annotations baked into PDF content")
            return True
        else:
            print(f"        📝 No annotations found in PDF")
            return False
