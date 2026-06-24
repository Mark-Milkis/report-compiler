"""GUI-side PDF rendering: re-exports the pure renderer and adds a cached QPixmap wrapper.

The pure functions live in ``report_compiler.utils.pdf_render`` so non-GUI code (the
in-document overlay preview) can render without importing PySide6.
"""

from __future__ import annotations

from report_compiler.utils.pdf_render import page_count, render_page_png  # noqa: F401 (re-export)

_pixmap_cache: dict = {}


def render_page_pixmap(pdf_path: str, page_index: int, target_width_px: int):
    """Render a page to a cached QPixmap. PySide6 is imported lazily so this module
    stays importable (for the re-exported PNG path) without PySide6 installed."""
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
