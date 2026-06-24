"""Write back to the live Word document via COM (used by the GUI dialogs).

The dialog runs in its own process; it attaches to the already-running Word instance
with ``GetActiveObject`` and inserts at a bookmark that the VBA launcher dropped at the
user's selection. This mirrors the table insertion the VBA used to do directly.
"""

from __future__ import annotations

import os

try:
    import win32com.client
except ImportError:  # pragma: no cover - non-Windows
    win32com = None


def _find_document(word, doc_path: str):
    """Return the open Document matching doc_path, else the active document."""
    target = os.path.normcase(os.path.abspath(doc_path))
    for doc in word.Documents:
        try:
            if os.path.normcase(os.path.abspath(doc.FullName)) == target:
                return doc
        except Exception:
            continue
    return word.ActiveDocument


def insert_overlay_table(doc_path: str, anchor_bookmark: str, tag_text: str) -> None:
    """Insert a 1x1 borderless table containing ``tag_text`` at the anchor bookmark.

    Raises on failure so the caller can surface a message.
    """
    if win32com is None:
        raise RuntimeError("pywin32 is not installed; cannot reach Word.")

    word = win32com.client.GetActiveObject("Word.Application")
    doc = _find_document(word, doc_path)

    if anchor_bookmark and doc.Bookmarks.Exists(anchor_bookmark):
        rng = doc.Bookmarks(anchor_bookmark).Range
    else:
        rng = word.Selection.Range

    table = doc.Tables.Add(rng, 1, 1)
    table.Borders.Enable = False
    table.Cell(1, 1).Range.Text = tag_text

    if anchor_bookmark and doc.Bookmarks.Exists(anchor_bookmark):
        doc.Bookmarks(anchor_bookmark).Delete
