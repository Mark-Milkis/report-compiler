"""In-document overlay preview: render OVERLAY pages into the live Word doc via COM.

Driven from the COM server (``SetOverlayPreview``). Three modes:
  - ``tags``  : restore every overlay to its canonical 1x1 tag table.
  - ``quick`` : first selected page in the cell + a "+N more" caption.
  - ``full``  : expand into one row per selected page so surrounding content reflows.

The ``[[OVERLAY: …]]`` tag is always preserved as hidden text in the cell (and redundantly
in each preview image's AltText), so the document still compiles at full resolution. The
compile pipeline independently normalizes any leftover previews (see
``DocxProcessor.normalize_overlay_previews``), so recovery never depends on this toggle.

This module drives Word and can only be exercised on a real interactive session.
"""

from __future__ import annotations

import os
import shutil
import tempfile

import fitz  # PyMuPDF

try:
    import win32com.client
except ImportError:  # pragma: no cover - non-Windows
    win32com = None

from ..core.config import Config
from ..gui.overlay_logic import expand_selection, format_spec, parse_overlay_tag
from ..pdf.content_analyzer import ContentAnalyzer
from ..utils.logging_config import get_logger
from ..utils.pdf_render import page_count

_MARKER = Config.OVERLAY_PREVIEW_MARKER
_WD_COLLAPSE_END = 0
_WD_COLOR_RED = 255
_QUICK_WIDTH_PX = 600
_FULL_WIDTH_PX = 800


def set_overlay_view(doc_path: str, mode: str) -> str:
    """Apply ``mode`` ('tags' | 'quick' | 'full') to every overlay in the document."""
    if win32com is None:
        raise RuntimeError("pywin32 is not installed; cannot reach Word.")
    if mode not in ("tags", "quick", "full"):
        raise ValueError(f"Unknown overlay view mode: {mode}")

    logger = get_logger()
    word = win32com.client.GetActiveObject("Word.Application")
    doc = _find_document(word, doc_path)

    overlays = [t for t in doc.Tables if _is_overlay_table(t)]
    # Restore-then-apply: every switch starts from canonical tags (idempotent).
    for table in overlays:
        _restore_table(table)

    if mode == "tags":
        return f"Overlay view: tags ({len(overlays)} overlay(s))."

    rendered, errors = 0, 0
    tmp_dir = tempfile.mkdtemp(prefix="rc_ovlprev_")
    try:
        for table in overlays:
            try:
                _apply_table(table, mode, doc_path, tmp_dir)
                rendered += 1
            except Exception as exc:  # noqa: BLE001 - one bad overlay shouldn't abort the rest
                errors += 1
                logger.warning("Overlay preview failed: %s", exc)
                _mark_error(table, str(exc))
    finally:
        shutil.rmtree(tmp_dir, ignore_errors=True)

    msg = f"Overlay view: {mode} ({rendered} overlay(s)"
    msg += f", {errors} error(s))" if errors else ")"
    return msg


# --- document / table helpers -------------------------------------------------
def _find_document(word, doc_path: str):
    target = os.path.normcase(os.path.abspath(doc_path))
    for doc in word.Documents:
        try:
            if os.path.normcase(os.path.abspath(doc.FullName)) == target:
                return doc
        except Exception:
            continue
    return word.ActiveDocument


def _is_overlay_table(table) -> bool:
    try:
        if table.Columns.Count != 1:
            return False
    except Exception:
        return False
    return _tag_of_table(table) is not None


def _tag_of_table(table):
    """Tag string for an overlay table, from cell text (hidden ok) or image AltText."""
    try:
        for r in range(1, table.Rows.Count + 1):
            match = Config.OVERLAY_REGEX.search(table.Cell(r, 1).Range.Text)
            if match:
                return match.group(0)
    except Exception:
        pass
    try:
        for shape in table.Range.InlineShapes:
            descr = shape.AlternativeText or ""
            if descr.startswith(_MARKER):
                match = Config.OVERLAY_REGEX.search(descr)
                if match:
                    return match.group(0)
    except Exception:
        pass
    return None


def _restore_table(table) -> None:
    """Collapse to a single row and set the cell to the plain, visible tag."""
    tag = _tag_of_table(table) or ""
    while table.Rows.Count > 1:
        table.Rows(table.Rows.Count).Delete()
    cell = table.Cell(1, 1)
    cell.Range.Text = tag  # replaces text and removes any inline images
    cell.Range.Font.Hidden = False


# --- applying previews --------------------------------------------------------
def _apply_table(table, mode: str, doc_path: str, tmp_dir: str) -> None:
    tag = _tag_of_table(table)
    parsed = parse_overlay_tag(tag)
    if parsed is None:
        raise ValueError("not an overlay tag")

    pdf_path = os.path.abspath(os.path.join(os.path.dirname(doc_path), parsed["file"]))
    if not os.path.isfile(pdf_path):
        raise FileNotFoundError(f"missing: {parsed['file']}")

    total = page_count(pdf_path)
    indices = sorted(expand_selection(parsed["page"] or "", total)) or [0]
    col_width = _column_width(table)

    if mode == "quick":
        pngs = _render_pages(pdf_path, indices[:1], parsed["crop"], tmp_dir, _QUICK_WIDTH_PX)
        cell = table.Cell(1, 1)
        _reset_cell(cell)
        _insert_image(cell, pngs[0], tag, col_width)
        _append_hidden_tag(cell, tag)
        caption = f"{os.path.basename(parsed['file'])} · p.{format_spec(set(indices))}"
        if len(indices) > 1:
            caption += f"  (+{len(indices) - 1} more)"
        _append_caption(cell, caption)
    else:  # full
        pngs = _render_pages(pdf_path, indices, parsed["crop"], tmp_dir, _FULL_WIDTH_PX)
        cell = table.Cell(1, 1)
        _reset_cell(cell)
        _insert_image(cell, pngs[0], tag, col_width)
        _append_hidden_tag(cell, tag)
        for png in pngs[1:]:
            table.Rows.Add()
            new_cell = table.Cell(table.Rows.Count, 1)
            _reset_cell(new_cell)
            _insert_image(new_cell, png, tag, col_width)


def _render_pages(pdf_path, indices, crop, tmp_dir, width_px):
    """Render selected pages to temp PNGs, applying the same crop the compiler would."""
    analyzer = ContentAnalyzer()
    out = []
    with fitz.open(pdf_path) as doc:
        for i in indices:
            page = doc[i]
            clip = analyzer.apply_content_cropping(page, crop)
            zoom = width_px / clip.width if clip.width else 1.0
            pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), clip=clip, alpha=False)
            path = os.path.join(tmp_dir, f"ovl_{i}.png")
            pix.save(path)
            out.append(path)
    return out


def _reset_cell(cell) -> None:
    """Empty a cell and clear hidden formatting so inserted images stay visible."""
    cell.Range.Text = ""
    cell.Range.Font.Hidden = False


def _append_hidden_tag(cell, tag: str) -> None:
    """Append the tag at the end of the cell and hide only that run (kept for compile)."""
    rng = cell.Range
    rng.Collapse(_WD_COLLAPSE_END)
    rng.Text = tag
    rng.Font.Hidden = True


def _insert_image(cell, png_path: str, tag: str, col_width) -> None:
    rng = cell.Range
    rng.Collapse(_WD_COLLAPSE_END)
    rng.Font.Hidden = False  # ensure the picture's run is not hidden-formatted
    shape = rng.InlineShapes.AddPicture(FileName=png_path, LinkToFile=False, SaveWithDocument=True)
    try:
        shape.LockAspectRatio = -1  # msoTrue
        if col_width:
            shape.Width = col_width
    except Exception:
        pass
    try:
        shape.AlternativeText = f"{_MARKER}:{tag}"
    except Exception:
        pass


def _append_caption(cell, text: str) -> None:
    rng = cell.Range
    rng.Collapse(_WD_COLLAPSE_END)
    rng.Text = "\r" + text
    rng.Font.Hidden = False


def _mark_error(table, message: str) -> None:
    """Show a red warning in the cell while keeping the tag (hidden) for recovery."""
    try:
        tag = _tag_of_table(table) or ""
        cell = table.Cell(1, 1)
        _reset_cell(cell)
        rng = cell.Range
        rng.Collapse(_WD_COLLAPSE_END)
        rng.Text = "⚠ " + message
        rng.Font.Hidden = False
        rng.Font.Color = _WD_COLOR_RED
        _append_hidden_tag(cell, tag)
    except Exception:
        pass


def _column_width(table):
    try:
        return table.Columns(1).Width
    except Exception:
        return None
