# Link Manager — Design

## Context

**Problem:** A report references external files through placeholders scattered across the document — overlays, appendices (`INSERT` PDFs), images, and recursive DOCX inserts. Today there's no way to see them all at once or know which paths are still valid; a broken link only surfaces as a compile error. Engineers expect an xref-style window (AutoCAD) / link manager (Revit): one list of every link, its validity, and quick actions to fix it.

**Goal:** A PySide6 window, launched from the ribbon, that scans the **live** Word document via COM, lists every link with its validity, and offers per-link **Go to**, **Open file**, **Relink**, and a **Relative ⇄ Absolute** path toggle.

This is the final GUI-roadmap feature. It reuses the foundation already built: the COM server + launch pattern, the PySide6 shell, the live-doc COM iteration in `overlay_preview.py`, the placeholder regexes in `Config`, and `Validators`.

**Decisions locked in:**
- Scan the **live Word doc via COM** (reflects unsaved edits; enables navigation), not the saved file.
- Actions: **Go to**, **Open file**, **Relink** (fix/move path). No bulk operations.
- **Per-link Relative ⇄ Absolute toggle** (rewrites the tag's path form; the common direction is →absolute). No bulk convert.
- **Skip preview-freshness** — v1 shows path validity only.

## Architecture / flow

```
Ribbon "Link manager"  (VBA thin launcher)
  └ localDocPath = GetLocalPath(ActiveDocument.FullName)
  └ CreateObject("ReportCompiler.Application").LaunchLinkManager(localDocPath)

COM server  LaunchLinkManager(doc_path)
  └ subprocess.Popen([sys.executable, "-m", "report_compiler.gui", "link-manager", "--doc", doc_path])

GUI process (PySide6)
  ├ GetActiveObject("Word.Application") → find document
  ├ scan_links(doc, doc_path) → [LinkRecord]   (validity via Validators)
  ├ table view + detail panel (both path forms)
  └ actions operate on the live doc via COM:
       Go to    → record.locator.Select() + ActiveWindow.ScrollIntoView + Word.Activate
       Open     → os.startfile(resolved_path)
       Relink   → file dialog → rewrite tag's file (preserve params) → re-validate
       Rel/Abs  → rewrite tag's file to the other form → re-validate
```

Same model as the overlay dialog: the GUI runs in its own process and reads/writes the live document through COM.

## Components

- **`document/link_index.py`** — the scanner (win32com):
  - `scan_links(doc, doc_path) -> list[LinkRecord]`: iterate `doc.Tables` (cells matching `OVERLAY_REGEX` / `IMAGE_REGEX`) and `doc.Paragraphs` (text matching `INSERT_REGEX`; `.docx` target ⇒ recursive-docx kind). Mirrors the live iteration already in [overlay_preview.py](src/report_compiler/document/overlay_preview.py).
  - `LinkRecord`: `kind` (overlay/image/appendix/docx), `raw_tag`, `stored_path`, `is_absolute`, `relative_form`, `absolute_form`, `pages`, `page_count`, `status`, `message`, `locator` (the Word `Range`/table for navigation).
  - `classify(kind, stored_path, doc_dir) -> (status, resolved_path, page_count, message)` — **pure, testable**, delegating to `Validators.validate_pdf_path` / `validate_image_path` / `validate_docx_path` ([utils/validators.py](src/report_compiler/utils/validators.py)). Statuses: `ok`, `missing`, `wrong_type`, `page_out_of_range` (warning).
  - `rewrite_link_path(record, new_path)` — replace only the file portion (regex group 1) inside `raw_tag`, preserving all params, and write it back into the cell/paragraph via COM. Used by both Relink and the Rel/Abs toggle. For an overlay currently showing a preview, restore it to tags first (reuse `overlay_preview` restore).
- **`gui/link_manager_dialog.py`** — `QTableView` (Type · Link · Pages · Status badge) above a **detail panel** showing the selected link's relative and absolute paths, page info, the Rel/Abs segmented toggle, and the action buttons. Double-click a row = Go to. A **Refresh** re-scans.
- **`com_server.LaunchLinkManager(doc_path)`**; ribbon button + thin VBA launcher (mirrors `InsertOverlayPlaceholder`).

## Path forms (Relative ⇄ Absolute)
Both forms are always computed for display: `absolute_form = abspath(join(doc_dir, stored))` (or `stored` if already absolute); `relative_form = relpath(absolute_form, doc_dir)` with forward slashes. The toggle rewrites the stored path to the chosen form. **Absolute→Relative is disabled when the file is on a different drive** than the document (no relative path exists). "Absolute" is the resolved **local** path (the launcher passes the OneDrive/SharePoint-resolved doc path).

## Error handling
- COM server not registered → the standard "COM Server Not Registered" message.
- Document never saved → relative resolution has no base; show a notice and still list links (absolute-only).
- A link whose tag can't be parsed → listed as `kind=unknown`, status `wrong_type`, never silently dropped.
- Relink/convert failures surface per-row without aborting the others; Refresh always re-syncs from the live doc.

## Out of scope
- Bulk convert-all and bulk relink.
- Preview-freshness / refresh-preview (deferred).
- Editing page ranges or crop from the manager (that's the overlay dialog).
- IMAGE/INSERT to non-file targets.

## Verification
1. **Unit (no Word):** `classify` over a temp dir — ok / missing / wrong_type / page_out_of_range for PDF, image, and DOCX targets; `relative_form`/`absolute_form` computation incl. the cross-drive guard; `rewrite_link_path` swaps the file and preserves params (`page=`, `crop=`, `:1-3`) for all tag types.
2. **End-to-end in Word (real machine):** open a doc with overlays/appendices/images/DOCX inserts (some broken) → list shows correct statuses; Go to scrolls to the link; Open launches the file; Relink fixes a broken path; the Rel/Abs toggle rewrites the tag and the row re-validates; Refresh picks up external edits.
3. **Caveat:** the live COM scan / navigation / write-back can't run in the agent sandbox (same registry/COM isolation as the other Word features); verified on a real interactive session. `classify`, path-form computation, and `rewrite_link_path` string logic are fully testable headless.
