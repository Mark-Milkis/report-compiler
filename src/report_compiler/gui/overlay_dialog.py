"""The Insert PDF Overlay dialog (PySide6).

Grid of page thumbnails with the selected pages highlighted, two-way synced with a
page-range box. On Insert it builds the [[OVERLAY: ...]] tag and writes it back into the
live Word document via ``word_writer``.
"""

from __future__ import annotations

import os

from PySide6.QtCore import QObject, Qt, QThread, Signal, Slot
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (
    QCheckBox,
    QDialog,
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from report_compiler.gui import pdf_render, word_writer
from report_compiler.gui.overlay_logic import (
    build_overlay_tag,
    expand_selection,
    format_spec,
    relative_pdf_path,
)

_THUMB_W = 120
_COLS = 4
_SEL_STYLE = "border: 2px solid #378ADD; border-radius: 6px; background: #E6F1FB;"
_UNSEL_STYLE = "border: 1px solid #C9C9C4; border-radius: 6px; background: white;"


class _ThumbWorker(QObject):
    """Renders page PNGs off the UI thread; emits bytes per page."""

    rendered = Signal(int, bytes)
    done = Signal()

    def __init__(self, pdf_path: str, count: int, width: int):
        super().__init__()
        self._pdf_path = pdf_path
        self._count = count
        self._width = width

    @Slot()
    def run(self) -> None:
        for i in range(self._count):
            try:
                png = pdf_render.render_page_png(self._pdf_path, i, self._width)
            except Exception:
                png = b""
            self.rendered.emit(i, png)
        self.done.emit()


class _PageThumb(QFrame):
    """One clickable page thumbnail; click toggles selection."""

    clicked = Signal(int)

    def __init__(self, index: int):
        super().__init__()
        self._index = index
        self.setFixedWidth(_THUMB_W + 16)
        self.setStyleSheet(_UNSEL_STYLE)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(6, 6, 6, 6)
        self._image = QLabel("…")
        self._image.setAlignment(Qt.AlignCenter)
        self._image.setMinimumHeight(int(_THUMB_W / 0.77))
        self._number = QLabel(str(index + 1))
        self._number.setAlignment(Qt.AlignCenter)
        self._number.setStyleSheet("border: none; color: #5F5E5A; font-size: 11px;")
        layout.addWidget(self._image)
        layout.addWidget(self._number)

    def set_pixmap(self, pixmap: QPixmap) -> None:
        self._image.setPixmap(pixmap.scaledToWidth(_THUMB_W, Qt.SmoothTransformation))

    def set_selected(self, selected: bool) -> None:
        self.setStyleSheet(_SEL_STYLE if selected else _UNSEL_STYLE)

    def mousePressEvent(self, event) -> None:  # noqa: N802 (Qt override)
        self.clicked.emit(self._index)


class OverlayDialog(QDialog):
    def __init__(self, doc_path: str = "", anchor: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle("Insert PDF overlay")
        self.resize(620, 640)
        self._doc_path = doc_path
        self._anchor = anchor
        self._pdf_path = ""
        self._total = 0
        self._selected: set[int] = set()
        self._thumbs: list[_PageThumb] = []
        self._syncing = False
        self._thread: QThread | None = None

        root = QVBoxLayout(self)

        # PDF file row
        file_row = QHBoxLayout()
        file_row.addWidget(QLabel("PDF file"))
        self._path_edit = QLineEdit()
        self._path_edit.setReadOnly(True)
        file_row.addWidget(self._path_edit, 1)
        browse = QPushButton("Browse…")
        browse.clicked.connect(self._on_browse)
        file_row.addWidget(browse)
        root.addLayout(file_row)

        # Selection toolbar (select/deselect all + live count)
        sel_bar = QHBoxLayout()
        self._select_all_btn = QPushButton("Select all")
        self._select_all_btn.clicked.connect(self._on_select_all)
        self._deselect_all_btn = QPushButton("Deselect all")
        self._deselect_all_btn.clicked.connect(self._on_deselect_all)
        self._select_all_btn.setEnabled(False)
        self._deselect_all_btn.setEnabled(False)
        sel_bar.addWidget(self._select_all_btn)
        sel_bar.addWidget(self._deselect_all_btn)
        sel_bar.addStretch(1)
        self._count_label = QLabel("")
        self._count_label.setStyleSheet("color: #5F5E5A;")
        sel_bar.addWidget(self._count_label)
        root.addLayout(sel_bar)

        # Thumbnail grid in a scroll area
        self._grid_host = QWidget()
        self._grid = QGridLayout(self._grid_host)
        self._grid.setAlignment(Qt.AlignTop)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(self._grid_host)
        root.addWidget(scroll, 1)

        # Page range + crop
        opts = QHBoxLayout()
        opts.addWidget(QLabel("Pages"))
        self._pages_edit = QLineEdit()
        self._pages_edit.setPlaceholderText("all  (e.g. 1-3, 5, 8-)")
        self._pages_edit.textEdited.connect(self._on_pages_edited)
        opts.addWidget(self._pages_edit, 1)
        self._crop = QCheckBox("Auto-crop to content")
        self._crop.setChecked(False)
        opts.addWidget(self._crop)
        root.addLayout(opts)

        # Buttons
        btns = QHBoxLayout()
        btns.addStretch(1)
        cancel = QPushButton("Cancel")
        cancel.clicked.connect(self.reject)
        btns.addWidget(cancel)
        self._insert = QPushButton("Insert overlay")
        self._insert.setDefault(True)
        self._insert.setEnabled(False)
        self._insert.clicked.connect(self._on_insert)
        btns.addWidget(self._insert)
        root.addLayout(btns)

    # --- PDF loading ---------------------------------------------------------
    def _on_browse(self) -> None:
        start_dir = os.path.dirname(self._doc_path) if self._doc_path else ""
        path, _ = QFileDialog.getOpenFileName(self, "Select a PDF", start_dir, "PDF files (*.pdf)")
        if path:
            self._load_pdf(path)

    def _load_pdf(self, path: str) -> None:
        try:
            total = pdf_render.page_count(path)
        except Exception as exc:
            QMessageBox.critical(self, "Cannot open PDF", str(exc))
            return
        self._pdf_path = path
        self._total = total
        self._path_edit.setText(path)
        self._build_grid()
        # Default: all pages selected (empty box == all).
        self._pages_edit.clear()
        self._selected = set(range(total))
        self._refresh_selection_styles()
        self._insert.setEnabled(True)
        self._select_all_btn.setEnabled(True)
        self._deselect_all_btn.setEnabled(True)
        self._start_render()

    def _build_grid(self) -> None:
        while self._grid.count():
            item = self._grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self._thumbs = []
        for i in range(self._total):
            thumb = _PageThumb(i)
            thumb.clicked.connect(self._on_thumb_clicked)
            self._thumbs.append(thumb)
            self._grid.addWidget(thumb, i // _COLS, i % _COLS)

    def _start_render(self) -> None:
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait()
        self._thread = QThread()
        self._worker = _ThumbWorker(self._pdf_path, self._total, _THUMB_W * 2)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.rendered.connect(self._on_thumb_rendered)
        self._worker.done.connect(self._thread.quit)
        self._thread.start()

    @Slot(int, bytes)
    def _on_thumb_rendered(self, index: int, png: bytes) -> None:
        if png and 0 <= index < len(self._thumbs):
            pixmap = QPixmap()
            pixmap.loadFromData(png, "PNG")
            self._thumbs[index].set_pixmap(pixmap)

    # --- selection sync ------------------------------------------------------
    def _on_pages_edited(self, text: str) -> None:
        if self._syncing:
            return
        self._selected = expand_selection(text, self._total)
        self._refresh_selection_styles()

    def _on_thumb_clicked(self, index: int) -> None:
        if index in self._selected:
            self._selected.discard(index)
        else:
            self._selected.add(index)
        self._sync_box_from_selection()
        self._refresh_selection_styles()

    def _on_select_all(self) -> None:
        self._selected = set(range(self._total))
        self._sync_box_from_selection()
        self._refresh_selection_styles()

    def _on_deselect_all(self) -> None:
        self._selected = set()
        self._sync_box_from_selection()
        self._refresh_selection_styles()

    def _sync_box_from_selection(self) -> None:
        """Update the page-range box to match the current selection (no feedback loop)."""
        self._syncing = True
        self._pages_edit.setText(format_spec(self._selected))
        self._syncing = False

    def _refresh_selection_styles(self) -> None:
        for i, thumb in enumerate(self._thumbs):
            thumb.set_selected(i in self._selected)
        self._count_label.setText(f"{len(self._selected)} of {self._total} pages selected")

    # --- insert --------------------------------------------------------------
    def _on_insert(self) -> None:
        if not self._pdf_path:
            return
        if not self._selected:
            QMessageBox.warning(self, "No pages", "Select at least one page to overlay.")
            return
        rel = relative_pdf_path(self._pdf_path, self._doc_path or self._pdf_path)
        tag = build_overlay_tag(rel, self._selected, self._total, self._crop.isChecked())

        if not self._doc_path:
            # Standalone/dev mode (no Word attached): just echo the tag.
            print(tag)
            self.accept()
            return
        try:
            word_writer.insert_overlay_table(self._doc_path, self._anchor, tag)
        except Exception as exc:
            QMessageBox.critical(self, "Insert failed", f"Could not insert into Word:\n{exc}")
            return
        self.accept()
