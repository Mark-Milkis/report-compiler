"""Pure PyMuPDF raster rendering of PDF pages (no GUI dependency).

Used by both the GUI thumbnail grid and the in-document overlay preview. Returning PNG
bytes keeps this importable without PySide6; the QPixmap wrapper lives in
``report_compiler.gui.pdf_render``.
"""

from __future__ import annotations

from typing import Optional, Tuple

import fitz  # PyMuPDF


def page_count(pdf_path: str) -> int:
    with fitz.open(pdf_path) as doc:
        return len(doc)


def render_page_png(
    pdf_path: str,
    page_index: int,
    target_width_px: int,
    clip: Optional[Tuple[float, float, float, float]] = None,
) -> bytes:
    """Render one page to PNG bytes, scaled so the rendered width is ~``target_width_px``.

    ``clip`` is an optional (x0, y0, x1, y1) rectangle in PDF points; when given, only
    that region is rendered (used for crop-accurate previews). The zoom is derived from
    the clipped width so the output still lands near ``target_width_px``.
    """
    with fitz.open(pdf_path) as doc:
        page = doc[page_index]
        clip_rect = fitz.Rect(clip) if clip is not None else None
        width_pts = clip_rect.width if clip_rect is not None else page.rect.width
        zoom = target_width_px / width_pts if width_pts else 1.0
        pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip_rect, alpha=False)
        return pix.tobytes("png")
