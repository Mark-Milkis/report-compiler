"""Enumerate and classify the links (placeholders) in the live Word document.

Powers the link manager. Scanning and navigation/rewrite touch Word via COM, but the
classification and tag-rewrite logic are pure functions so they can be unit-tested
without Word.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field
from typing import Optional

try:
    import win32com.client  # noqa: F401  (presence check; COM used via passed objects)
except ImportError:  # pragma: no cover - non-Windows
    win32com = None

from ..core.config import Config
from ..gui.overlay_logic import parse_overlay_tag
from ..utils.page_selector import PageSelector
from ..utils.validators import Validators

# Statuses
OK = "ok"
MISSING = "missing"
WRONG_TYPE = "wrong_type"
PAGE_RANGE = "page_out_of_range"
UNKNOWN = "unknown"

# Kinds
OVERLAY = "overlay"
IMAGE = "image"
APPENDIX = "appendix"
DOCX = "docx"

_REGEX_FOR_KIND = {
    OVERLAY: Config.OVERLAY_REGEX,
    IMAGE: Config.IMAGE_REGEX,
    APPENDIX: Config.INSERT_REGEX,
    DOCX: Config.INSERT_REGEX,
}

_selector = PageSelector()


# --- pure logic (testable) ----------------------------------------------------
def resolve_forms(stored_path: str, doc_dir: str):
    """Return (absolute_form, relative_form). relative_form is None when no relative
    path exists (no doc dir, or a different drive)."""
    if os.path.isabs(stored_path):
        absolute = os.path.normpath(stored_path)
    elif doc_dir:
        absolute = os.path.normpath(os.path.join(doc_dir, stored_path))
    else:
        absolute = stored_path
    relative = None
    if doc_dir:
        try:
            relative = os.path.relpath(absolute, doc_dir).replace(os.sep, "/")
        except ValueError:
            relative = None  # different drive
    return absolute, relative


def _max_requested_page(page_spec: Optional[str]) -> Optional[int]:
    """Highest explicit 1-based page in a spec, or None (open ranges/all are unbounded)."""
    if not page_spec:
        return None
    sel = _selector.parse_specification(page_spec)
    if sel["use_all"] or not sel["pages"]:
        return None
    return max(sel["pages"]) + 1


def classify(kind: str, stored_path: str, page_spec: Optional[str], doc_dir: str) -> dict:
    """Resolve and validate a link. Pure: filesystem + Validators only, no Word."""
    absolute, relative = resolve_forms(stored_path, doc_dir)
    out = {
        "resolved_path": absolute,
        "absolute_form": absolute,
        "relative_form": relative,
        "page_count": 0,
        "status": UNKNOWN,
        "message": "",
    }
    if kind in (OVERLAY, APPENDIX):
        result = Validators.validate_pdf_path(absolute, "")
    elif kind == IMAGE:
        result = Validators.validate_image_path(absolute, "")
    elif kind == DOCX:
        result = Validators.validate_docx_path(absolute)
    else:
        out["message"] = "Unrecognized link"
        return out

    out["page_count"] = result.get("page_count", 0)
    if not result["valid"]:
        error = result.get("error_message") or "Invalid file"
        out["status"] = MISSING if "not found" in error.lower() else WRONG_TYPE
        out["message"] = error
        return out

    if kind in (OVERLAY, APPENDIX):
        max_page = _max_requested_page(page_spec)
        if max_page and out["page_count"] and max_page > out["page_count"]:
            out["status"] = PAGE_RANGE
            out["message"] = f"Requested page {max_page} exceeds {out['page_count']} pages"
            return out

    out["status"] = OK
    return out


def rewrite_tag_file(raw_tag: str, kind: str, new_path: str) -> str:
    """Return raw_tag with only its file path (regex group 1) replaced — params kept."""
    regex = _REGEX_FOR_KIND.get(kind)
    match = regex.search(raw_tag) if regex else None
    if not match:
        return raw_tag
    start, end = match.start(1), match.end(1)
    return raw_tag[:start] + new_path + raw_tag[end:]


# --- link record --------------------------------------------------------------
@dataclass
class LinkRecord:
    kind: str
    raw_tag: str
    stored_path: str
    page_spec: Optional[str]
    status: str
    message: str
    resolved_path: str
    relative_form: Optional[str]
    absolute_form: str
    page_count: int
    is_absolute: bool
    locator: object = field(default=None, repr=False)


# --- COM locators (navigate / rewrite the live doc) ---------------------------
class _TableLocator:
    """Overlay/image link living in a 1x1 table cell."""

    def __init__(self, table):
        self.table = table

    def select(self):
        self.table.Range.Select()

    def set_tag(self, new_tag, old_tag=None):
        table = self.table
        for shape in list(table.Range.InlineShapes):  # drop any preview images
            try:
                shape.Delete()
            except Exception:
                pass
        while table.Rows.Count > 1:  # collapse any full-preview expansion
            table.Rows(table.Rows.Count).Delete()
        cell = table.Cell(1, 1)
        cell.Range.Text = new_tag
        cell.Range.Font.Hidden = False


class _ParagraphLocator:
    """INSERT/appendix/docx link living in a paragraph."""

    def __init__(self, rng):
        self.range = rng

    def select(self):
        self.range.Select()

    def set_tag(self, new_tag, old_tag=None):
        find = self.range.Find
        find.ClearFormatting()
        find.Replacement.ClearFormatting()
        find.MatchWildcards = False
        find.Text = old_tag
        find.Replacement.Text = new_tag
        find.Execute(Replace=1)  # wdReplaceOne


# --- scanning the live document ----------------------------------------------
def find_document(word, doc_path: str):
    target = os.path.normcase(os.path.abspath(doc_path))
    for doc in word.Documents:
        try:
            if os.path.normcase(os.path.abspath(doc.FullName)) == target:
                return doc
        except Exception:
            continue
    return word.ActiveDocument


def _make_record(kind, raw_tag, stored_path, page_spec, doc_dir, locator) -> LinkRecord:
    info = classify(kind, stored_path, page_spec, doc_dir)
    return LinkRecord(
        kind=kind,
        raw_tag=raw_tag,
        stored_path=stored_path,
        page_spec=page_spec,
        status=info["status"],
        message=info["message"],
        resolved_path=info["resolved_path"],
        relative_form=info["relative_form"],
        absolute_form=info["absolute_form"],
        page_count=info["page_count"],
        is_absolute=os.path.isabs(stored_path),
        locator=locator,
    )


def scan_links(doc, doc_path: str) -> list:
    """Scan the live document for every link. Returns a list of LinkRecord."""
    doc_dir = os.path.dirname(doc_path)
    records: list = []

    for table in doc.Tables:
        try:
            cell_text = table.Cell(1, 1).Range.Text
        except Exception:
            continue
        overlay = Config.OVERLAY_REGEX.search(cell_text)
        if overlay:
            parsed = parse_overlay_tag(overlay.group(0)) or {}
            records.append(_make_record(
                OVERLAY, overlay.group(0), parsed.get("file", ""), parsed.get("page"),
                doc_dir, _TableLocator(table)))
            continue
        image = Config.IMAGE_REGEX.search(cell_text)
        if image:
            records.append(_make_record(
                IMAGE, image.group(0), image.group(1).strip(), None,
                doc_dir, _TableLocator(table)))

    for paragraph in doc.Paragraphs:
        try:
            text = paragraph.Range.Text
        except Exception:
            continue
        insert = Config.INSERT_REGEX.search(text)
        if insert:
            path = insert.group(1).strip()
            page_spec = (insert.group(2) or "").strip() or None
            kind = DOCX if path.lower().endswith(".docx") else APPENDIX
            records.append(_make_record(
                kind, insert.group(0), path, page_spec, doc_dir,
                _ParagraphLocator(paragraph.Range)))

    return records


# --- actions ------------------------------------------------------------------
def go_to(word, record: LinkRecord) -> None:
    if record.locator is not None:
        record.locator.select()
    try:
        word.Activate()
    except Exception:
        pass


def open_source(record: LinkRecord) -> bool:
    if record.resolved_path and os.path.isfile(record.resolved_path):
        os.startfile(record.resolved_path)  # noqa: S606 - opening the user's own linked file
        return True
    return False


def set_link_path(record: LinkRecord, new_stored_path: str, doc_dir: str) -> LinkRecord:
    """Rewrite the tag's file to new_stored_path (params preserved) and re-classify."""
    new_tag = rewrite_tag_file(record.raw_tag, record.kind, new_stored_path)
    if record.locator is not None:
        record.locator.set_tag(new_tag, record.raw_tag)
    return _make_record(record.kind, new_tag, new_stored_path, record.page_spec, doc_dir, record.locator)
