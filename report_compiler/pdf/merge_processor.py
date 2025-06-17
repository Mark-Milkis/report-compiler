"""
PDF merge processing for paragraph-based insertions.
"""

import fitz  # PyMuPDF
import shutil
import os
from typing import Dict, List, Any, Optional
from ..core.config import Config
from ..utils.page_selector import PageSelector
from ..utils.logging_config import get_merge_logger
from .content_analyzer import ContentAnalyzer
from .marker_remover import MarkerRemover


class MergeProcessor:
    """Handles paragraph-based PDF merge operations with hierarchical TOC generation."""
    
    def __init__(self):
        self.page_selector = PageSelector()
        self.content_analyzer = ContentAnalyzer()
        self.marker_remover = MarkerRemover()
        self.logger = get_merge_logger()

    def _adjust_appendix_toc(self, appendix_toc_entries: list,
                             page_offset_in_final_doc: int, 
                             base_nest_level: int = 2) -> list:
        """
        Adjusts page numbers and levels for an appendix's TOC entries.
        Page numbers in PyMuPDF TOC are 1-indexed.
        page_offset_in_final_doc is the 0-indexed page where the appendix content starts in the output document.
        base_nest_level is the level in the master_toc under which these entries should be nested.
        """
        self.logger.debug(f"Adjusting {len(appendix_toc_entries)} TOC entries with page offset {page_offset_in_final_doc}, nest level {base_nest_level}")
        
        adjusted_entries = []
        for level, title, page_num, opts in appendix_toc_entries:
            # Original page_num is 1-indexed relative to appendix_doc
            # New page_num is 1-indexed relative to output_doc
            new_page_num_in_output = (page_num - 1) + page_offset_in_final_doc + 1
            new_level = base_nest_level + level  # Nest under the existing appendix heading
            adjusted_entries.append([new_level, title, new_page_num_in_output, opts])
            self.logger.debug(f"  TOC entry '{title}': level {level}->{new_level}, page {page_num}->{new_page_num_in_output}")
        
        return adjusted_entries

    def _find_appendix_heading_in_toc(self, toc_entries: list, marker_page_num: int) -> Optional[int]:
        """
        Find the appendix heading in the TOC that corresponds to the page where the marker was found.
        Returns the index of the TOC entry, or None if not found.
        """
        if not toc_entries:
            return None
            
        # Look for TOC entries that point to the same page as the marker or nearby pages
        for idx, entry in enumerate(toc_entries):
            if len(entry) >= 3:
                level, title, page_num = entry[0], entry[1], entry[2]
                # Check if this TOC entry is on the same page or the page before the marker                # (since the marker might be after the heading)
                if page_num == marker_page_num or page_num == marker_page_num + 1:
                    # Additional check: look for "APPENDIX" in the title
                    if "APPENDIX" in title.upper():
                        return idx
        return None

    def process_merges(self, base_pdf_path: str, merge_placeholders: List[Dict],
                       output_path: str) -> bool:
        """
        Process all merge placeholders in the base PDF, creating a merged TOC.
        """
        if not merge_placeholders:
            if base_pdf_path != output_path:
                try:
                    shutil.copy(base_pdf_path, output_path)
                    self.logger.info(f"No merges required, copied base PDF to: {os.path.basename(output_path)}")
                except Exception as e:
                    self.logger.error(f"Failed to copy base PDF: {e}", exc_info=True)
                    return False
            else:
                self.logger.info("No merges required, base PDF is already at output path")
            return True

        self.logger.info(f"Processing {len(merge_placeholders)} merge(s) with hierarchical TOC")

        master_toc = []
        output_doc = None
        
        try:
            # Open and validate base PDF
            self.logger.debug(f"Opening base PDF: {base_pdf_path}")
            self.logger.debug(f"File exists: {os.path.exists(base_pdf_path)}")
            if os.path.exists(base_pdf_path):
                file_size = os.path.getsize(base_pdf_path) / (1024 * 1024)  # MB
                self.logger.debug(f"File size: {file_size:.2f} MB")
            
            try:
                output_doc = fitz.open(base_pdf_path)
                self.logger.debug(f"Successfully opened PDF document, type: {type(output_doc)}")
                self.logger.debug(f"Document has {len(output_doc)} pages")
            except Exception as open_error:
                self.logger.error(f"Failed to open base PDF with fitz.open(): {open_error}", exc_info=True)
                self.logger.debug(f"fitz.open type: {type(fitz.open)}")
                return False
            
            if output_doc.needs_pass: 
                auth_result = output_doc.authenticate("")
                self.logger.debug(f"Document authentication result: {auth_result}")
            
            # Extract base TOC
            base_toc_entries = []
            try:
                self.logger.debug("Extracting Table of Contents from base PDF")
                raw_toc = output_doc.get_toc(simple=False)
                if raw_toc is not None:
                    base_toc_entries = raw_toc
                    master_toc.extend(base_toc_entries)
                    self.logger.info(f"Extracted {len(base_toc_entries)} TOC entries from base PDF")
                    for i, entry in enumerate(base_toc_entries[:5]):  # Log first 5 entries
                        if len(entry) >= 3:
                            level, title, page = entry[0], entry[1], entry[2]
                            self.logger.debug(f"  TOC[{i}]: Level {level}, '{title}', Page {page}")
                        else:
                            self.logger.debug(f"  TOC[{i}]: Unexpected format: {entry}")
                else:
                    self.logger.info("No Table of Contents found in base PDF")
            except TypeError as te:
                self.logger.error(f"TypeError extracting TOC from base PDF: {te}", exc_info=True)
                self.logger.debug(f"output_doc type: {type(output_doc)}")
                self.logger.debug(f"get_toc method type: {type(output_doc.get_toc) if hasattr(output_doc, 'get_toc') else 'Missing'}")
                base_toc_entries = []
            except Exception as e_toc:
                self.logger.error(f"Error extracting TOC from base PDF: {e_toc}", exc_info=True)
                base_toc_entries = []

        except Exception as e:
            self.logger.error(f"Error during base PDF processing: {e}", exc_info=True)
            if output_doc and not output_doc.is_closed:
                output_doc.close()
            return False

        try:
            # Find markers and their positions in the original base PDF
            initial_marker_scan_doc = fitz.open(base_pdf_path)
            if initial_marker_scan_doc.needs_pass: 
                initial_marker_scan_doc.authenticate("")

            sorted_placeholder_infos = []
            for idx, placeholder in enumerate(merge_placeholders, 1):
                marker_text = Config.get_merge_marker(idx)
                found_on_page_idx = -1
                for page_num, page in enumerate(initial_marker_scan_doc):
                    if page.search_for(marker_text, quads=False):
                        found_on_page_idx = page_num
                        break
                if found_on_page_idx != -1:
                    sorted_placeholder_infos.append({
                        'placeholder': placeholder,
                        'original_marker_page_idx': found_on_page_idx, 
                        'marker_text': marker_text
                    })
                else:
                    self.logger.warning(f"Marker '{marker_text}' for '{placeholder.get('pdf_path_raw', 'N/A')}' not found in base PDF. Skipping.")
            
            initial_marker_scan_doc.close()
            sorted_placeholder_infos.sort(key=lambda x: x['original_marker_page_idx'])

            page_offset_from_prior_insertions = 0 

            for info in sorted_placeholder_infos:
                placeholder = info['placeholder']
                original_marker_page_idx = info['original_marker_page_idx']
                marker_text = info['marker_text']

                current_marker_page_in_output_idx = original_marker_page_idx + page_offset_from_prior_insertions
                
                if not (0 <= current_marker_page_in_output_idx < len(output_doc)):
                    self.logger.warning(f"Calculated marker page index {current_marker_page_in_output_idx + 1} is out of bounds. Skipping appendix.")
                    continue
                  # Remove marker from the output document
                page_to_clean = output_doc[current_marker_page_in_output_idx]
                if not self.marker_remover.remove_marker_text(page_to_clean, marker_text):
                    self.logger.warning(f"Failed to remove marker '{marker_text}' from page {current_marker_page_in_output_idx + 1}.")
                
                insertion_start_page_idx_in_output = current_marker_page_in_output_idx + 1

                pdf_path = placeholder['resolved_path']
                appendix_name = os.path.basename(pdf_path)
                
                self.logger.info("    Processing merge for appendix: %s", appendix_name)

                with fitz.open(pdf_path) as appendix_doc:
                    if appendix_doc.needs_pass: 
                        appendix_doc.authenticate("")
                    self.content_analyzer.bake_annotations(appendix_doc) 
                    
                    appendix_original_toc = []
                    try:
                        raw_appendix_toc = appendix_doc.get_toc(simple=False)
                        if raw_appendix_toc is not None:
                            appendix_original_toc = raw_appendix_toc
                    except Exception as e_app_toc:
                        self.logger.warning("      ⚠️ Error getting TOC from appendix PDF: %s", e_app_toc)
                    
                    # Find the corresponding appendix heading in the base TOC
                    appendix_heading_idx = self._find_appendix_heading_in_toc(
                        base_toc_entries, current_marker_page_in_output_idx + 1
                    )
                    
                    if appendix_heading_idx is not None and appendix_original_toc:
                        # Adjust appendix TOC entries to nest under the existing heading
                        existing_heading = base_toc_entries[appendix_heading_idx]
                        existing_level = existing_heading[0]
                        
                        adjusted_appendix_toc = self._adjust_appendix_toc(
                            appendix_original_toc, 
                            insertion_start_page_idx_in_output, 
                            base_nest_level=existing_level
                        )
                          # Insert the adjusted TOC entries after the appendix heading
                        toc_insert_position = appendix_heading_idx + 1
                        for entry in reversed(adjusted_appendix_toc):
                            master_toc.insert(toc_insert_position, entry)
                        
                        self.logger.info("      ✓ Nested %d TOC entries under existing appendix heading", len(adjusted_appendix_toc))
                    elif appendix_original_toc:
                        # Fallback: add as top-level entries if no corresponding heading found
                        adjusted_appendix_toc = self._adjust_appendix_toc(
                            appendix_original_toc, 
                            insertion_start_page_idx_in_output, 
                            base_nest_level=1
                        )
                        master_toc.extend(adjusted_appendix_toc)
                        self.logger.info("      ✓ Added %d TOC entries as top-level (no corresponding heading found)", len(adjusted_appendix_toc))

                    selected_pages_spec = placeholder.get('page_spec')
                    page_selection = self.page_selector.parse_specification(selected_pages_spec)
                    page_indices_to_insert = self.page_selector.apply_selection(appendix_doc, page_selection)
                    
                    num_pages_to_insert = len(page_indices_to_insert)
                    if num_pages_to_insert > 0:
                        # Convert page indices to a format suitable for insert_pdf
                        if len(page_indices_to_insert) == len(appendix_doc):
                            # All pages - no need to specify page selection
                            output_doc.insert_pdf(appendix_doc, start_at=insertion_start_page_idx_in_output)
                        else:
                            # Specific pages - use the pages parameter
                            output_doc.insert_pdf(appendix_doc, 
                                                  from_page=page_indices_to_insert[0], 
                                                  to_page=page_indices_to_insert[-1] if len(page_indices_to_insert) > 1 else page_indices_to_insert[0],
                                                  start_at=insertion_start_page_idx_in_output)
                        page_offset_from_prior_insertions += num_pages_to_insert
                        self.logger.info("      ✓ Merged '%s', %d pages inserted.", appendix_name, num_pages_to_insert)
                    else:
                        self.logger.info("      • No pages selected from '%s'. Skipping insertion.", appendix_name)
            
            if master_toc:
                output_doc.set_toc(master_toc)
                self.logger.info("   ✓ Hierarchical Table of Contents constructed and applied.")
              # If we're saving to the same path as the base file, use incremental save
            if os.path.abspath(base_pdf_path) == os.path.abspath(output_path):
                output_doc.saveIncr()
                self.logger.info("   ✓ Merged PDF with hierarchical TOC saved incrementally to: %s", output_path)
            else:
                output_doc.save(output_path, garbage=3, deflate=True, clean=True)
                self.logger.info("   ✓ Merged PDF with hierarchical TOC saved to: %s", output_path)
            return True

        except Exception as e:
            self.logger.error("   ❌ Error processing merges with TOC: %s", e, exc_info=True)
            return False
        finally:
            if output_doc and not output_doc.is_closed:
                output_doc.close()

    def _process_single_merge(self, base_doc: fitz.Document, placeholder: Dict[str, Any], 
                             idx: int) -> bool:
        """
        Process a single merge placeholder (legacy method - not used in hierarchical TOC mode).
        
        Args:
            base_doc: Base PDF document
            placeholder: Merge placeholder dictionary
            idx: Placeholder index for naming
            
        Returns:
            bool: True if successful, False otherwise        """
        try:
            pdf_path = placeholder['resolved_path']
            page_count = placeholder['page_count']
            
            self.logger.info("    Processing merge appendix %d: %s", idx, placeholder['pdf_path_raw'])
            
            # Find and remove the merge marker
            marker = Config.get_merge_marker(idx)
            marker_page_idx = self._find_and_remove_merge_marker(base_doc, marker)
            
            if marker_page_idx is None:
                self.logger.error("      ❌ Could not find merge marker")
                return False
            
            self.logger.info("      📄 Will insert PDF page(s) after page %d", marker_page_idx + 1)
            
            # Open source PDF
            self.logger.info("      Opening appendix PDF: %s", pdf_path)
            with fitz.open(pdf_path) as source_doc:
                # Bake annotations
                self.content_analyzer.bake_annotations(source_doc)
                
                # Determine pages to use
                page_selection = self.page_selector.parse_specification(placeholder.get('page_spec'))
                pages_to_insert = self.page_selector.apply_selection(source_doc, page_selection)
                
                if not pages_to_insert:
                    pages_to_insert = list(range(len(source_doc)))
                
                self.logger.info("        📄 Using all %d pages", len(pages_to_insert))
                
                # Insert entire PDF at once (more efficient)
                insert_position = marker_page_idx + 1
                self.logger.info("        📥 Inserting entire PDF (%d pages) at position %d", len(pages_to_insert), insert_position)
                
                base_doc.insert_pdf(source_doc, from_page=0, to_page=len(source_doc)-1, 
                                   start_at=insert_position)
                
                self.logger.info("        ✓ Inserted all %d pages in single operation", len(pages_to_insert))
            
            self.logger.info("      ✓ Merge appendix %d insertion complete", idx)
            return True
            
        except Exception as e:
            self.logger.error("      ❌ Error processing merge %d: %s", idx, e, exc_info=True)
            return False
    
    def _find_and_remove_merge_marker(self, base_doc: fitz.Document, marker: str) -> Optional[int]:
        """
        Find and remove merge marker, returning the page index.
        
        Args:
            base_doc: PDF document to search
            marker: Marker text to find and remove
            
        Returns:
            int or None: Page index where marker was found, or None if not found        """
        self.logger.info("      🔍 Searching for marker: %s", marker)
        
        for page_index in range(len(base_doc)):
            page = base_doc[page_index]
            
            # Check if marker exists on this page
            marker_info = self.marker_remover.find_marker_position(page, marker)
            if marker_info:
                self.logger.info("      ✓ Found marker on page %d", page_index + 1)
                self.logger.info("      🧹 Removing marker from page %d", page_index + 1)
                
                # Remove the marker
                if self.marker_remover.remove_marker_text(page, marker):
                    self.logger.info("      ✓ Marker removed from page %d", page_index + 1)
                else:
                    self.logger.warning("      ⚠️ Could not remove marker from page %d", page_index + 1)
                
                return page_index
        
        return None
