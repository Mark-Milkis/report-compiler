"""Raster rendering of PDF pages for GUI previews (PyMuPDF).

Shared by all GUI features. ``render_page_png`` is pure PyMuPDF (testable without Qt);
``render_page_pixmap`` wraps it into a QPixmap and is cached for the thumbnail grid.
"""

from __future__ import annotations

import fitz  # PyMuPDF


def page_count(pdf_path: str) -> int:
    with fitz.open(pdf_path) as doc:
        return len(doc)


def render_page_png(pdf_path: str, page_index: int, target_width_px: int) -> bytes:
    """Render one page to PNG bytes, scaled so its width is ~``target_width_px``."""
    with fitz.open(pdf_path) as doc:
        page = doc[page_index]
        zoom = target_width_px / page.rect.width if page.rect.width else 1.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
        return pix.tobytes("png")


_pixmap_cache: dict = {}


def render_page_pixmap(pdf_path: str, page_index: int, target_width_px: int):
    """Render a page to a cached QPixmap. Imported lazily so this module stays
    importable (for the PNG path) without PySide6 installed."""
    from PySide6.QtGui import QPixmap

    key = (pdf_path, page_index, target_width_px)
    cached = _pixmap_cache.get(key)
    if cached is not None:
        return cached

    png = render_page_png(pdf_path, page_index, target_width_px)
    pixmap = QPixmap()
    pixmap.loadFromData(png, "PNG")
    _pixmap_cache[key] = pixmap
    return pixmap


def clear_cache() -> None:
    _pixmap_cache.clear()
