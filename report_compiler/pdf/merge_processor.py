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

    def process_merges(self, base_pdf_path: str, content_map: Dict[str, Any],
                       toc_pages: List[int], output_path: str) -> bool:
        """
        Process all merge placeholders, inserting PDFs and creating a hierarchical TOC.

        Args:
            base_pdf_path: Path to the PDF to insert content into.
            content_map: Dictionary mapping markers to their location and metadata.
            toc_pages: A list of page numbers that contain the Table of Contents.
            output_path: Path for the final output PDF.

        Returns:
            True if successful, False otherwise.
        """
        merge_markers = sorted([
            (marker, data) for marker, data in content_map.items()
            if data['type'] == 'paragraph'
        ], key=lambda item: item[1]['page_index'])

        if not merge_markers:
            self.logger.info("No merge placeholders to process.")
            # File is copied by the compiler if no merges are needed.
            return True

        try:
            self.logger.debug("Opening base PDF for merging: %s", base_pdf_path)
            output_doc = fitz.open(base_pdf_path)
            master_toc = output_doc.get_toc(simple=False)
            self.logger.debug("  > Extracted %d root TOC entries from base document.", len(master_toc))

            page_offset = 0
            for idx, (marker, data) in enumerate(merge_markers, 1):
                placeholder = data['placeholder']
                pdf_path = placeholder['resolved_path']
                self.logger.info("  Processing merge %d: %s", idx, placeholder['pdf_path_raw'])

                # The page where the marker was found in the *original* document
                original_marker_page = data['page_index']
                # The page to insert *at* in the *current* (potentially modified) document
                insert_at_page = original_marker_page + page_offset + 1

                self.logger.debug("    > Marker found on original page %d. Inserting at page %d in output.",
                                 original_marker_page + 1, insert_at_page)

                with fitz.open(pdf_path) as appendix_doc:
                    self.content_analyzer.bake_annotations(appendix_doc)

                    # Select pages from the appendix PDF as specified in the placeholder
                    page_spec = placeholder.get('page_spec')
                    page_selection = self.page_selector.parse_specification(page_spec)
                    pages_to_insert = self.page_selector.apply_selection(appendix_doc, page_selection)
                    if not pages_to_insert:
                        pages_to_insert = list(range(len(appendix_doc)))
                    
                    num_pages_to_insert = len(pages_to_insert)
                    if num_pages_to_insert == 0:
                        self.logger.warning("    > No pages selected from %s. Skipping.", pdf_path)
                        continue

                    self.logger.info("    > Merging %d page(s) from appendix.", num_pages_to_insert)

                    # Handle TOC merging
                    appendix_toc = appendix_doc.get_toc(simple=False)
                    if appendix_toc:
                        self._merge_toc_entries(master_toc, appendix_toc, original_marker_page + 1, insert_at_page, placeholder)

                    # Insert the selected pages
                    output_doc.insert_pdf(
                        appendix_doc,
                        from_page=pages_to_insert[0],
                        to_page=pages_to_insert[-1],
                        start_at=insert_at_page
                    )
                    
                    page_offset += num_pages_to_insert

            self.logger.info("  > Applying final hierarchical Table of Contents.")
            output_doc.set_toc(master_toc)
            
            self.logger.debug("Saving final merged PDF to: %s", output_path)
            output_doc.save(output_path, garbage=3, deflate=True, clean=True)
            self.logger.info("✓ Merge processing complete.")
            return True

        except Exception as e:
            self.logger.error("❌ Error during merge processing: %s", e, exc_info=True)
            return False
        finally:
            if 'output_doc' in locals() and output_doc and not output_doc.is_closed:
                output_doc.close()

    def _merge_toc_entries(self, master_toc, appendix_toc, marker_page, insert_page, placeholder):
        """Finds the correct position in the master TOC and inserts the appendix TOC."""
        self.logger.debug("    > Merging %d TOC entries from appendix.", len(appendix_toc))
        
        # Find the heading in the master_toc that corresponds to this appendix
        heading_idx = self._find_appendix_heading_in_toc(master_toc, marker_page)
        
        base_level = 1
        insert_pos = len(master_toc)

        if heading_idx is not None:
            base_level = master_toc[heading_idx][0]
            insert_pos = heading_idx + 1
            self.logger.debug("    > Found corresponding TOC heading '%s' at level %d.", master_toc[heading_idx][1], base_level)
        else:
            self.logger.warning("    > Could not find a matching heading in the main TOC for this appendix.")
            self.logger.warning("    > Appending TOC entries at the root level.")

        # Adjust and insert the appendix TOC entries
        adjusted_toc = self._adjust_appendix_toc(appendix_toc, insert_page, base_level)
        
        for entry in reversed(adjusted_toc):
            master_toc.insert(insert_pos, entry)
        self.logger.debug("    > Inserted %d adjusted TOC entries.", len(adjusted_toc))

    def _adjust_appendix_toc(self, appendix_toc, page_offset, base_nest_level):
        """Adjusts page numbers and levels for an appendix's TOC entries."""
        adjusted_entries = []
        for level, title, page_num, opts in appendix_toc:
            new_page_num = (page_num - 1) + page_offset
            new_level = base_nest_level + level
            adjusted_entries.append([new_level, title, new_page_num, opts])
        return adjusted_entries

    def _find_appendix_heading_in_toc(self, toc_entries, marker_page_num):
        """
        Find the TOC entry that corresponds to the page where the marker was found.
        """
        # Search for a TOC entry pointing to the marker's page.
        # This assumes the heading for the appendix is on the same page as the marker.
        for idx, entry in enumerate(toc_entries):
            if len(entry) >= 3 and entry[2] == marker_page_num:
                # A simple heuristic: if the title contains "Appendix", it's probably the one.
                if "APPENDIX" in entry[1].upper():
                    return idx
        # Fallback: if no exact match, maybe it's the last entry on the previous page.
        for idx, entry in reversed(list(enumerate(toc_entries))):
             if len(entry) >= 3 and entry[2] < marker_page_num:
                 return idx
        return None
