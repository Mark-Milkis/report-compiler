"""
PDF merge processing for paragraph-based insertions.
"""

import fitz  # PyMuPDF
from typing import Dict, List, Any, Optional
from ..core.config import Config
from ..utils.page_selector import PageSelector
from .content_analyzer import ContentAnalyzer
from .marker_remover import MarkerRemover


class MergeProcessor:
    """Handles paragraph-based PDF merge operations."""
    
    def __init__(self):
        self.page_selector = PageSelector()
        self.content_analyzer = ContentAnalyzer()
        self.marker_remover = MarkerRemover()
    
    def process_merges(self, base_pdf_path: str, merge_placeholders: List[Dict], 
                      output_path: str) -> bool:
        """
        Process all merge placeholders in the base PDF.
        
        Args:
            base_pdf_path: Path to base PDF document
            merge_placeholders: List of merge placeholder dictionaries
            output_path: Path for output PDF
            
        Returns:
            bool: True if successful, False otherwise
        """
        if not merge_placeholders:
            return True
        
        try:
            print(f"   • Processing {len(merge_placeholders)} paragraph-based merge(s)...")
            print(f"    Opening base PDF for merge insertions: {base_pdf_path}")
            
            with fitz.open(base_pdf_path) as base_doc:
                # Process merges in reverse order to maintain page indices
                for idx, placeholder in enumerate(reversed(merge_placeholders), 1):
                    merge_idx = len(merge_placeholders) - idx + 1
                    if not self._process_single_merge(base_doc, placeholder, merge_idx):
                        return False
                  # Save the PDF with merges
                print(f"    Saving PDF with merges: {output_path}")
                if base_pdf_path == output_path:
                    # If saving to same file, use incremental save
                    base_doc.saveIncr()
                else:
                    base_doc.save(output_path)
                print(f"    ✓ PDF with merges saved successfully")
            
            print("   ✓ Merge processing complete")
            return True
            
        except Exception as e:
            print(f"   ❌ Error processing merges: {e}")
            return False
    
    def _process_single_merge(self, base_doc: fitz.Document, placeholder: Dict[str, Any], 
                             idx: int) -> bool:
        """
        Process a single merge placeholder.
        
        Args:
            base_doc: Base PDF document
            placeholder: Merge placeholder dictionary
            idx: Placeholder index for naming
            
        Returns:
            bool: True if successful, False otherwise
        """
        try:
            pdf_path = placeholder['resolved_path']
            page_count = placeholder['page_count']
            
            print(f"    Processing merge appendix {idx}: {placeholder['pdf_path_raw']}")
            
            # Find and remove the merge marker
            marker = Config.get_merge_marker(idx)
            marker_page_idx = self._find_and_remove_merge_marker(base_doc, marker)
            
            if marker_page_idx is None:
                print(f"      ❌ Could not find merge marker")
                return False
            
            print(f"      📄 Will insert PDF page(s) after page {marker_page_idx + 1}")
            
            # Open source PDF
            print(f"      Opening appendix PDF: {pdf_path}")
            with fitz.open(pdf_path) as source_doc:
                # Bake annotations
                self.content_analyzer.bake_annotations(source_doc)
                
                # Determine pages to use
                page_selection = self.page_selector.parse_specification(placeholder.get('page_spec'))
                pages_to_insert = self.page_selector.apply_selection(source_doc, page_selection)
                
                if not pages_to_insert:
                    pages_to_insert = list(range(len(source_doc)))
                
                print(f"        📄 Using all {len(pages_to_insert)} pages")
                
                # Insert entire PDF at once (more efficient)
                insert_position = marker_page_idx + 1
                print(f"        📥 Inserting entire PDF ({len(pages_to_insert)} pages) at position {insert_position}")
                
                base_doc.insert_pdf(source_doc, from_page=0, to_page=len(source_doc)-1, 
                                   start_at=insert_position)
                
                print(f"        ✓ Inserted all {len(pages_to_insert)} pages in single operation")
            
            print(f"      ✓ Merge appendix {idx} insertion complete")
            return True
            
        except Exception as e:
            print(f"      ❌ Error processing merge {idx}: {e}")
            return False
    
    def _find_and_remove_merge_marker(self, base_doc: fitz.Document, marker: str) -> Optional[int]:
        """
        Find and remove merge marker, returning the page index.
        
        Args:
            base_doc: PDF document to search
            marker: Marker text to find and remove
            
        Returns:
            int or None: Page index where marker was found, or None if not found
        """
        print(f"      🔍 Searching for marker: {marker}")
        
        for page_index in range(len(base_doc)):
            page = base_doc[page_index]
            
            # Check if marker exists on this page
            marker_info = self.marker_remover.find_marker_position(page, marker)
            if marker_info:
                print(f"      ✓ Found marker on page {page_index + 1}")
                print(f"      🧹 Removing marker from page {page_index + 1}")
                
                # Remove the marker
                if self.marker_remover.remove_marker_text(page, marker):
                    print(f"      ✓ Marker removed from page {page_index + 1}")
                else:
                    print(f"      ⚠️ Could not remove marker from page {page_index + 1}")
                
                return page_index
        
        return None
