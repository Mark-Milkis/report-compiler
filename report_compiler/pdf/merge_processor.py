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
            if os.path.exists(base_pdf_path) and not os.path.samefile(base_pdf_path, output_path):
                shutil.copy(base_pdf_path, output_path)
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
                self.logger.info("  Processing merge %d: %s", idx, os.path.basename(pdf_path))

                original_marker_page_idx = data['page_index']  # 0-indexed

                # The page with the marker, in the context of the evolving output document
                current_marker_page_idx = original_marker_page_idx + page_offset

                # We insert the appendix *after* the page with the marker.
                # PyMuPDF's `start_at` is the 0-indexed page number to insert *before*.
                insertion_point_idx = current_marker_page_idx + 1

                self.logger.debug(
                    "    > Marker on original page %d (current: %d). Inserting before page %d (0-indexed).",
                    original_marker_page_idx + 1, current_marker_page_idx + 1, insertion_point_idx
                )

                with fitz.open(pdf_path) as appendix_doc:
                    self.content_analyzer.bake_annotations(appendix_doc)

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

                    # Adjust the page numbers of the master TOC before merging the new TOC.
                    # This ensures that links in the original document's TOC are updated.
                    self.logger.debug("    > Adjusting master TOC page numbers for %d inserted pages.", num_pages_to_insert)
                    for entry in master_toc:
                        # entry[2] is 1-based page number. insertion_point_idx is 0-based.
                        # Any entry pointing to a page at or after the insertion point needs to be shifted.
                        if entry[2] >= insertion_point_idx + 1:
                            entry[2] += num_pages_to_insert

                    appendix_toc = appendix_doc.get_toc(simple=False)
                    if appendix_toc:
                        # The new content will start at page `insertion_point_idx + 1` (1-based)
                        new_content_start_page_num = insertion_point_idx + 1
                        # The page where the marker is now located (1-based)
                        current_marker_page_num = original_marker_page_idx + page_offset + 1
                        marker_rect = data.get('rect')
                        self._merge_toc_entries(
                            master_toc,
                            appendix_toc,
                            current_marker_page_num,
                            new_content_start_page_num,
                            placeholder,
                            marker_rect
                        )

                    output_doc.insert_pdf(
                        appendix_doc,
                        from_page=pages_to_insert[0],
                        to_page=pages_to_insert[-1],
                        start_at=insertion_point_idx
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

    def _merge_toc_entries(self, master_toc, appendix_toc, marker_page_num, new_content_start_page_num, placeholder, marker_rect: Optional[List[float]]):
        """Finds the correct position in the master TOC and inserts the appendix TOC."""
        self.logger.debug("    > Merging %d TOC entries from appendix.", len(appendix_toc))
        
        heading_idx = self._find_appendix_heading_in_toc(master_toc, marker_page_num, marker_rect)
        
        base_level = 1
        insert_pos = len(master_toc)

        if heading_idx is not None:
            base_level = master_toc[heading_idx][0]
            insert_pos = heading_idx + 1
            self.logger.debug("    > Found corresponding TOC heading '%s' at level %d.", master_toc[heading_idx][1], base_level)
        else:
            self.logger.warning("    > Could not find a matching heading in the main TOC for this appendix.")
            self.logger.warning("    > Appending TOC entries at the root level.")

        adjusted_toc = self._adjust_appendix_toc(appendix_toc, new_content_start_page_num, base_level)
        
        for entry in reversed(adjusted_toc):
            master_toc.insert(insert_pos, entry)
        self.logger.debug("    > Inserted %d adjusted TOC entries.", len(adjusted_toc))

    def _adjust_appendix_toc(self, appendix_toc, new_content_start_page_num, base_nest_level):
        """Adjusts page numbers and levels for an appendix's TOC entries."""
        adjusted_entries = []
        for level, title, page_num, opts in appendix_toc:
            new_page_num = (new_content_start_page_num - 1) + page_num
            new_level = base_nest_level + level
            
            # Create a completely fresh destination dictionary to avoid any lingering
            # invalid references (like xrefs) from the source PDF's opts dictionary.
            new_opts = {'kind': fitz.LINK_GOTO, 'zoom': 0}

            # Extract and recreate the destination point to remove any hidden state.
            original_to = opts.get('to')
            if isinstance(original_to, fitz.Point):
                new_opts['to'] = fitz.Point(original_to.x, original_to.y)
            else:
                new_opts['to'] = fitz.Point(0, 0) # Default to top of page
            
            original_zoom = opts.get('zoom')
            if original_zoom:
                new_opts['zoom'] = original_zoom

            adjusted_entries.append([new_level, title, new_page_num, new_opts])
        return adjusted_entries

    def _find_appendix_heading_in_toc(self, toc_entries: List[Any], marker_page_num: int, marker_rect_coords: Optional[List[float]]) -> Optional[int]:
        """
        Finds the TOC entry that most likely corresponds to the section where content is being inserted.
        It prioritizes finding the heading immediately preceding the insertion marker on the same page.
        """
        if not marker_rect_coords:
            self.logger.debug("    > Marker coordinates not available. Using page-based fallback for TOC heading search.")
            # Fallback to old logic if rect is not available for some reason
            for idx, entry in reversed(list(enumerate(toc_entries))):
                if len(entry) >= 3 and entry[2] <= marker_page_num:
                    return idx
            return None

        marker_y = marker_rect_coords[1]  # y0 of the marker's rectangle

        # Find all headings on the same page as the marker
        headings_on_page = []
        for idx, entry in enumerate(toc_entries):
            if len(entry) >= 3 and entry[2] == marker_page_num:
                headings_on_page.append((idx, entry))

        # Find the heading immediately preceding the marker on the same page
        best_match_idx = None
        min_distance = float('inf')

        if headings_on_page:
            self.logger.debug("    > Found %d headings on page %d. Analyzing position.", len(headings_on_page), marker_page_num)
            for idx, entry in headings_on_page:
                dest = entry[3].get('to') if len(entry) > 3 and isinstance(entry[3], dict) else None
                if dest and isinstance(dest, fitz.Point):
                    heading_y = dest.y
                    # We are looking for a heading *above* the marker
                    if heading_y < marker_y:
                        distance = marker_y - heading_y
                        if distance < min_distance:
                            min_distance = distance
                            best_match_idx = idx
            
            if best_match_idx is not None:
                self.logger.debug("    > Found best matching heading on same page: '%s'", toc_entries[best_match_idx][1])
                return best_match_idx

        # Fallback: If no heading is found above the marker on the same page,
        # or if there are no headings on the marker's page at all,
        # find the last heading on a previous page.
        self.logger.debug("    > No heading found above marker on page %d. Searching previous pages.", marker_page_num)
        last_heading_idx = None
        # Iterate up to the marker page to find the last entry
        for idx, entry in enumerate(toc_entries):
            if len(entry) >= 3 and entry[2] < marker_page_num:
                last_heading_idx = idx
            elif len(entry) >= 3 and entry[2] >= marker_page_num:
                # We have passed the target page, no need to search further
                break
        
        if last_heading_idx is not None:
            self.logger.debug("    > Found last heading on a previous page: '%s'", toc_entries[last_heading_idx][1])
        
        return last_heading_idx
