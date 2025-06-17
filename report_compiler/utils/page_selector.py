"""
Page selection and specification parsing utilities.
"""

import re
from typing import List, Dict, Any, Optional
import fitz  # PyMuPDF


class PageSelector:
    """Handles page specification parsing and selection for PDF processing."""
    
    def __init__(self):
        # Regex for page specification parsing
        self.page_spec_regex = re.compile(r'(\d+)(?:-(\d+)?)?\s*(?:,|$)')
    
    def parse_specification(self, page_spec: str) -> Dict[str, Any]:
        """
        Parse page specification string into structured data.
        
        Args:
            page_spec (str): Page specification (e.g., "1-3,7,9-")
            
        Returns:
            Dict containing parsed page information
        """
        if not page_spec or page_spec.strip() == "":
            return {
                'pages': [],
                'use_all': True,
                'open_range_start': None,
                'total_specified': 0
            }
        
        pages = []
        open_range_start = None
        
        # Find all matches in the specification
        matches = self.page_spec_regex.findall(page_spec.strip())
        
        for match in matches:
            start_str, end_str = match
            start = int(start_str)
            
            if end_str == '':  # Open range (e.g., "9-")
                open_range_start = start - 1  # Convert to 0-based
                break
            elif end_str:  # Closed range (e.g., "1-3")
                end = int(end_str)
                for page_num in range(start, end + 1):
                    pages.append(page_num - 1)  # Convert to 0-based
            else:  # Single page (e.g., "7")
                pages.append(start - 1)  # Convert to 0-based
        
        return {
            'pages': sorted(list(set(pages))),  # Remove duplicates and sort
            'use_all': False,
            'open_range_start': open_range_start,
            'total_specified': len(pages)
        }
    
    def apply_selection(self, pdf_doc: fitz.Document, selection: Dict[str, Any], 
                       max_pages: Optional[int] = None) -> List[int]:
        """
        Apply page selection to a PDF document.
        
        Args:
            pdf_doc: PyMuPDF document object
            selection: Parsed page selection from parse_specification
            max_pages: Optional maximum number of pages to include
            
        Returns:
            List of 0-based page indices to process
        """
        total_pages = len(pdf_doc)
        
        if selection['use_all']:
            pages = list(range(total_pages))
        else:
            pages = selection['pages'].copy()
            
            # Handle open range
            if selection['open_range_start'] is not None:
                start_page = selection['open_range_start']
                pages.extend(range(start_page, total_pages))
        
        # Remove any pages that don't exist in the document
        valid_pages = [p for p in pages if 0 <= p < total_pages]
        
        # Apply max_pages limit if specified
        if max_pages and len(valid_pages) > max_pages:
            valid_pages = valid_pages[:max_pages]
        
        return sorted(list(set(valid_pages)))  # Remove duplicates and sort
    
    def validate_pages(self, selection: Dict[str, Any], total_pages: int) -> Dict[str, Any]:
        """
        Validate page selection against document size.
        
        Args:
            selection: Parsed page selection
            total_pages: Total pages in the document
            
        Returns:
            Dict with validation results
        """
        if selection['use_all']:
            return {
                'valid': True,
                'message': f"Using all {total_pages} pages",
                'page_count': total_pages
            }
        
        valid_pages = [p for p in selection['pages'] if 0 <= p < total_pages]
        invalid_pages = [p + 1 for p in selection['pages'] if p >= total_pages]  # Convert back to 1-based for display
        
        # Handle open range
        additional_from_range = 0
        if selection['open_range_start'] is not None:
            if selection['open_range_start'] < total_pages:
                additional_from_range = total_pages - selection['open_range_start']
            else:
                invalid_pages.append(selection['open_range_start'] + 1)  # Convert to 1-based
        
        total_valid = len(valid_pages) + additional_from_range
        
        if invalid_pages:
            return {
                'valid': False,
                'message': f"Invalid page numbers: {invalid_pages} (document has {total_pages} pages)",
                'page_count': total_valid,
                'invalid_pages': invalid_pages
            }
        
        return {
            'valid': True,
            'message': f"Using {total_valid} of {total_pages} pages",
            'page_count': total_valid
        }
    
    def format_page_list(self, pages: List[int], one_based: bool = True) -> str:
        """
        Format a list of page numbers into a readable string.
        
        Args:
            pages: List of page numbers (0-based or 1-based)
            one_based: Whether to display as 1-based page numbers
            
        Returns:
            Formatted string (e.g., "1-3, 7, 9-11")
        """
        if not pages:
            return "none"
        
        # Convert to 1-based if needed
        display_pages = [p + 1 for p in pages] if not one_based else pages
        display_pages = sorted(display_pages)
        
        if not display_pages:
            return "none"
        
        ranges = []
        start = display_pages[0]
        end = start
        
        for page in display_pages[1:]:
            if page == end + 1:
                end = page
            else:
                if start == end:
                    ranges.append(str(start))
                else:
                    ranges.append(f"{start}-{end}")
                start = end = page
        
        # Add the last range
        if start == end:
            ranges.append(str(start))
        else:
            ranges.append(f"{start}-{end}")
        
        return ", ".join(ranges)
