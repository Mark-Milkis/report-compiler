"""Qt-free logic for the overlay dialog: page expansion, tag building, relative paths.

Kept separate from the PySide6 dialog so it can be unit-tested without a GUI or Word.
"""

from __future__ import annotations

import os
from typing import List, Set

from report_compiler.utils.page_selector import PageSelector

_selector = PageSelector()


def expand_selection(spec: str, total_pages: int) -> Set[int]:
    """Resolve a page-spec string to a concrete set of 0-based page indices.

    Reuses ``PageSelector.parse_specification`` and clamps to the document size.
    Empty/blank spec means "all pages".
    """
    selection = _selector.parse_specification(spec)
    if selection["use_all"]:
        return set(range(total_pages))

    pages: Set[int] = set(selection["pages"])
    open_start = selection["open_range_start"]
    if open_start is not None:
        pages |= set(range(open_start, total_pages))
    return {p for p in pages if 0 <= p < total_pages}


def format_spec(pages_zero_based: Set[int]) -> str:
    """Serialize a set of 0-based page indices to a compact 1-based spec ("1-3, 5")."""
    if not pages_zero_based:
        return ""
    # PageSelector.format_page_list collapses runs into ranges; feed it 1-based pages.
    one_based = sorted(p + 1 for p in pages_zero_based)
    return _selector.format_page_list(one_based, one_based=True)


def relative_pdf_path(pdf_path: str, doc_path: str) -> str:
    """Path of ``pdf_path`` relative to the document's folder, using forward slashes.

    Falls back to the absolute path when the two are on different drives (Windows),
    where a relative path is impossible.
    """
    doc_dir = os.path.dirname(doc_path)
    try:
        rel = os.path.relpath(pdf_path, doc_dir)
    except ValueError:
        rel = os.path.abspath(pdf_path)
    return rel.replace(os.sep, "/")


def build_overlay_tag(
    rel_path: str,
    selected_pages_zero_based: Set[int],
    total_pages: int,
    crop: bool,
) -> str:
    """Build an ``[[OVERLAY: ...]]`` tag.

    Omits ``page=`` when every page is selected (the parser treats absent page as all).
    Cropping defaults to off, so only the on case is written (``crop=true``); when off we
    omit ``crop=`` entirely to keep the tag clean.
    """
    tag = f"[[OVERLAY: {rel_path}"
    if selected_pages_zero_based and len(selected_pages_zero_based) < total_pages:
        tag += f", page={format_spec(selected_pages_zero_based)}"
    if crop:
        tag += ", crop=true"
    tag += "]]"
    return tag
