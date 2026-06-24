# In-Document Overlay Preview — Design

## Context

**Problem:** An OVERLAY placeholder is an empty 1×1 table cell holding a `[[OVERLAY: file.pdf, page=…]]` tag until compile. Coworkers can't tell *which* pages an overlay will insert or *how* they'll look, and can't see how the rest of the document reflows around multi-page overlays.

**Goal:** Let users flip the whole document between **Tags**, a lightweight **Quick preview**, and a reflow-accurate **Full preview** of every overlay — rendered from the real PDFs, crop-accurate — without ever endangering a full-resolution compile.

This is feature #3 of the GUI roadmap. It builds on the existing foundation: the COM server, `gui/pdf_render` (PyMuPDF raster), `content_analyzer` (cropping), `placeholder_parser`/`Config` (OVERLAY regex), and the Python-drives-Word pattern.

**Decisions locked in:**
- Document-wide **ribbon dropdown**: Tags / Quick / Full.
- **Quick** = first selected page + `+N` badge (cell stays 1×1). **Full** = expand into per-page rows so surrounding content reflows.
- Previews are **crop-accurate** (reuse `content_analyzer`).
- **Recovery is a property of the compiler**, not just the toggle (see below).
- Safeguards: redundant tag stored in each preview image's AltText; a **Repair overlays** command. (No strip-on-save, no save-blocking.)

## Recovery & fidelity (the core guarantee)

Preview images are **low-res throwaways**; the compiler ignores them entirely — it reads the tag, resolves the original PDF, and renders it vector-accurate via `show_pdf_page`. So previews can never reduce output quality; they only bloat the file while shown.

The `[[OVERLAY: …]]` tag is always preserved as **hidden text** in the cell (survives save/reopen; still present in `cell.text`), and **redundantly** in each preview image's AltText. Canonical state of an overlay is a **1×1 single-cell borderless table** whose cell text matches the OVERLAY regex.

**Compile-pipeline normalization (mandatory, covers every compile path — CLI, COM, button):** before parsing, the compiler normalizes overlays on its working copy so a saved-in-preview document always compiles correctly:
- remove inline images marked `RCPREVIEW`,
- collapse any expanded multi-row overlay table back to the single tagged row,
- (hidden tag text is read fine by the parser, so no unhide needed for compile).

Without this, a saved Full-preview overlay (multi-row table) would be **silently skipped** by the parser (it only treats 1×1 tables as overlays) and quick-preview images would leak into output. Normalization makes recovery guaranteed regardless of who compiles or how.

## Architecture / flow

```
Ribbon dropdown "Overlay view"  (Tags | Quick | Full)   [VBA thin callback]
  └ localDocPath = GetLocalPath(ActiveDocument.FullName)
  └ CreateObject("ReportCompiler.Application").SetOverlayPreview(localDocPath, mode)

COM server  SetOverlayPreview(doc_path, mode)
  └ document/overlay_preview.set_overlay_view(doc_path, mode)
       GetActiveObject("Word.Application") → restore-then-apply (idempotent):
         1. restore_to_tags(doc): delete RCPREVIEW shapes, collapse rows to the tagged row, unhide tag
         2. if mode == quick: per overlay → render first page → insert image + "+N" caption, hide tag
            if mode == full:  per overlay → render each selected page → replicate rows + images, hide tag

Compile (any path) → docx_processor normalizes overlays first (python-docx) → parse → render from source PDFs
"Repair overlays" ribbon button → SetOverlayPreview(doc, "tags")  (force canonical)
```

Word stays the source of truth for the live preview; the operation runs synchronously in the COM call (typical docs have a handful of overlays). If it proves slow on large jobs, it can move to the async job+poll pattern later.

## Components

- **`utils/pdf_render.py`** (move from `gui/`): `page_count`, `render_page_png(pdf, page_index, target_width_px, clip=None)` — add optional `clip` rect for crop-accurate rendering. `gui/pdf_render` keeps only the `QPixmap` wrapper, importing from here (so non-GUI code no longer imports the GUI package).
- **`document/overlay_preview.py`** (new): the live-Word engine — `set_overlay_view(doc_path, mode)`, `restore_to_tags(doc)`, `apply_quick(doc, …)`, `apply_full(doc, …)`, `iter_overlay_tables(doc)`. Uses win32com (`GetActiveObject`), the OVERLAY regex + `PageSelector` (page expansion) + `content_analyzer` (crop rect → clip). Marks inserted shapes with AltText = the tag, prefixed by the `RCPREVIEW` marker.
- **`document/docx_processor.py`**: add `_normalize_overlay_previews(doc)` (python-docx) run at the start of processing — removes `RCPREVIEW` `<w:drawing>` runs and collapses multi-row overlay tables to the tagged row. Reuse the existing OVERLAY regex.
- **`core/config.py`**: add `OVERLAY_PREVIEW_MARKER = "RCPREVIEW"` so the live engine and the compile normalizer agree on the marker.
- **`com_server.py`**: `SetOverlayPreview(doc_path, mode)` added to `_public_methods_`.
- **VBA `ReportingTools.bas` + `report_compiler_UI.xml`**: the "Overlay view" dropdown (Tags/Quick/Full) + a "Repair overlays" button; rebuild the `.dotm`. `RunReportCompiler` keeps working since normalization now lives in the pipeline (the button no longer needs its own strip, though it may still set Tags first for a clean visual).

## Crop fidelity
For each page, compute the crop rect with `content_analyzer.apply_content_cropping(page, crop_enabled)` and pass it as `clip` to `render_page_png`, so the preview matches what the compiler would overlay.

## Error handling
- Missing PDF / invalid page spec: the cell shows a red `⚠ missing: file.pdf` note but keeps the tag; the toggle continues and reports `rendered X overlays, Y errors`.
- COM server not registered: the standard "COM Server Not Registered" message (mirrors the other buttons).
- Restore is defensive: an overlay is any single-column table with a cell whose text matches the OVERLAY regex; the row carrying the tag is kept, others deleted.

## Out of scope
- Editing overlay parameters from preview (that's the overlay dialog / future link manager).
- Per-overlay (non-document-wide) toggling.
- Strip-on-save and save-blocking (explicitly declined).
- Previews for IMAGE/INSERT placeholders (overlay-only for now).

## Verification
1. **Unit (no Word):** tag parse → (file, pages, crop); `PageSelector` expansion; `render_page_png` with a `clip` produces a valid cropped PNG; `docx_processor._normalize_overlay_previews` on a hand-built docx (multi-row table with `RCPREVIEW` images) collapses to a 1×1 with the tag intact.
2. **Compile recovery (no live Word):** build a docx in the expanded/preview state (multi-row + marked images + hidden tag), run a normal `compile`, and confirm the overlay renders from the source PDF at full resolution and the preview images don't appear in output.
3. **End-to-end in Word (real machine):** doc with 2 overlays → cycle Tags→Quick→Full→Tags and confirm a clean round-trip (cell text restored, no leftover rows/images); save while in Full preview, reopen, run **Repair overlays**, compile → correct high-res output.
4. **Caveat:** the live-Word apply/restore (COM `GetActiveObject` manipulation) can't run in the agent sandbox — same registry/COM isolation as before; verified on a real interactive Windows session. The pure logic and the compile-time normalization are fully testable headless.
