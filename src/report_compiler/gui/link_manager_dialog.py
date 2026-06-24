"""Link manager window (PySide6).

Lists every link in the live Word document with its validity, and offers per-link
Go to / Open file / Relink and a Relative <-> Absolute path toggle. Reads and writes the
document through COM (``document.link_index``).
"""

from __future__ import annotations

import os

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import (
    QButtonGroup,
    QDialog,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QMessageBox,
    QPushButton,
    QRadioButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
)

from report_compiler.document import link_index as li

_KIND_LABEL = {li.OVERLAY: "Overlay", li.IMAGE: "Image", li.APPENDIX: "Appendix", li.DOCX: "DOCX"}
_STATUS_LABEL = {
    li.OK: "OK",
    li.MISSING: "Missing",
    li.WRONG_TYPE: "Wrong type",
    li.PAGE_RANGE: "Page range",
    li.UNKNOWN: "Unknown",
}
_STATUS_COLOR = {
    li.OK: "#1D9E75",
    li.MISSING: "#E24B4A",
    li.WRONG_TYPE: "#E24B4A",
    li.PAGE_RANGE: "#BA7517",
    li.UNKNOWN: "#888780",
}
_FILE_FILTER = {
    li.OVERLAY: "PDF files (*.pdf)",
    li.APPENDIX: "PDF files (*.pdf)",
    li.IMAGE: "Images (*.png *.jpg *.jpeg *.gif *.bmp *.tif *.tiff *.svg *.emf *.wmf)",
    li.DOCX: "Word documents (*.docx)",
}


class LinkManagerDialog(QDialog):
    def __init__(self, doc_path: str = "", parent=None):
        super().__init__(parent)
        self.setWindowTitle(f"Link manager — {os.path.basename(doc_path) or 'document'}")
        self.resize(720, 560)
        self._doc_path = doc_path
        self._doc_dir = os.path.dirname(doc_path)
        self._word = None
        self._doc = None
        self._records: list = []
        self._updating = False

        root = QVBoxLayout(self)

        # Toolbar
        bar = QHBoxLayout()
        bar.addStretch(1)
        refresh = QPushButton("Refresh")
        refresh.clicked.connect(self._refresh)
        bar.addWidget(refresh)
        root.addLayout(bar)

        # Table
        self._table = QTableWidget(0, 4)
        self._table.setHorizontalHeaderLabels(["Type", "Link", "Pages", "Status"])
        self._table.verticalHeader().setVisible(False)
        self._table.setSelectionBehavior(QTableWidget.SelectRows)
        self._table.setSelectionMode(QTableWidget.SingleSelection)
        self._table.setEditTriggers(QTableWidget.NoEditTriggers)
        header = self._table.horizontalHeader()
        header.setSectionResizeMode(0, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(1, QHeaderView.Stretch)
        header.setSectionResizeMode(2, QHeaderView.ResizeToContents)
        header.setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self._table.itemSelectionChanged.connect(self._on_selection)
        self._table.itemDoubleClicked.connect(lambda *_: self._go_to())
        root.addWidget(self._table, 1)

        root.addWidget(self._build_detail_panel())

        self._summary = QLabel("")
        self._summary.setStyleSheet("color: #5F5E5A;")
        root.addWidget(self._summary)

        self._connect_word()
        self._refresh()

    # --- detail panel --------------------------------------------------------
    def _build_detail_panel(self) -> QFrame:
        panel = QFrame()
        panel.setFrameShape(QFrame.StyledPanel)
        layout = QVBoxLayout(panel)

        self._title_label = QLabel("Select a link")
        self._title_label.setStyleSheet("font-weight: 500;")
        layout.addWidget(self._title_label)

        self._rel_label = QLabel("")
        self._abs_label = QLabel("")
        for lbl in (self._rel_label, self._abs_label):
            lbl.setTextInteractionFlags(Qt.TextSelectableByMouse)
            layout.addWidget(lbl)

        store = QHBoxLayout()
        store.addWidget(QLabel("Store path as"))
        self._rel_radio = QRadioButton("Relative")
        self._abs_radio = QRadioButton("Absolute")
        self._form_group = QButtonGroup(self)
        self._form_group.addButton(self._rel_radio)
        self._form_group.addButton(self._abs_radio)
        self._rel_radio.toggled.connect(self._on_form_toggled)
        store.addWidget(self._rel_radio)
        store.addWidget(self._abs_radio)
        store.addStretch(1)
        layout.addLayout(store)

        actions = QHBoxLayout()
        self._goto_btn = QPushButton("Go to")
        self._goto_btn.clicked.connect(self._go_to)
        self._open_btn = QPushButton("Open file")
        self._open_btn.clicked.connect(self._open_file)
        self._relink_btn = QPushButton("Relink…")
        self._relink_btn.clicked.connect(self._relink)
        actions.addWidget(self._goto_btn)
        actions.addWidget(self._open_btn)
        actions.addWidget(self._relink_btn)
        actions.addStretch(1)
        layout.addLayout(actions)

        self._set_detail_enabled(False)
        return panel

    # --- Word + scan ---------------------------------------------------------
    def _connect_word(self) -> None:
        try:
            import win32com.client

            self._word = win32com.client.GetActiveObject("Word.Application")
            self._doc = li.find_document(self._word, self._doc_path)
        except Exception as exc:
            QMessageBox.warning(self, "Word not found", f"Could not attach to Word:\n{exc}")

    def _refresh(self) -> None:
        if self._doc is None:
            return
        try:
            self._records = li.scan_links(self._doc, self._doc_path)
        except Exception as exc:
            QMessageBox.critical(self, "Scan failed", str(exc))
            return
        self._populate()

    def _populate(self) -> None:
        self._table.setRowCount(len(self._records))
        for row, rec in enumerate(self._records):
            self._set_row(row, rec)
        missing = sum(1 for r in self._records if r.status in (li.MISSING, li.WRONG_TYPE))
        warns = sum(1 for r in self._records if r.status == li.PAGE_RANGE)
        parts = [f"{len(self._records)} links"]
        if missing:
            parts.append(f"{missing} broken")
        if warns:
            parts.append(f"{warns} warning")
        self._summary.setText("  ·  ".join(parts))
        self._set_detail_enabled(False)
        self._title_label.setText("Select a link")

    def _set_row(self, row: int, rec) -> None:
        pages = rec.page_spec or ("all" if rec.kind in (li.OVERLAY, li.APPENDIX) else "—")
        cells = [_KIND_LABEL.get(rec.kind, rec.kind), rec.stored_path, pages, _STATUS_LABEL.get(rec.status, rec.status)]
        for col, text in enumerate(cells):
            item = QTableWidgetItem(text)
            if col == 3:
                item.setForeground(QColor(_STATUS_COLOR.get(rec.status, "#888780")))
                if rec.message:
                    item.setToolTip(rec.message)
            self._table.setItem(row, col, item)

    # --- selection + detail --------------------------------------------------
    def _current_row(self) -> int:
        rows = self._table.selectionModel().selectedRows()
        return rows[0].row() if rows else -1

    def _on_selection(self) -> None:
        row = self._current_row()
        if row < 0 or row >= len(self._records):
            self._set_detail_enabled(False)
            return
        rec = self._records[row]
        self._title_label.setText(f"{_KIND_LABEL.get(rec.kind, rec.kind)} · {os.path.basename(rec.stored_path)}")
        self._rel_label.setText(f"Relative:  {rec.relative_form or '(unavailable — different drive)'}")
        self._abs_label.setText(f"Absolute:  {rec.absolute_form}")
        self._set_detail_enabled(True)
        self._updating = True
        self._abs_radio.setChecked(rec.is_absolute)
        self._rel_radio.setChecked(not rec.is_absolute)
        self._rel_radio.setEnabled(rec.relative_form is not None)
        self._updating = False
        self._open_btn.setEnabled(rec.status in (li.OK, li.PAGE_RANGE))

    def _set_detail_enabled(self, enabled: bool) -> None:
        for w in (self._goto_btn, self._open_btn, self._relink_btn, self._rel_radio, self._abs_radio):
            w.setEnabled(enabled)
        if not enabled:
            self._rel_label.setText("")
            self._abs_label.setText("")

    # --- actions -------------------------------------------------------------
    def _go_to(self) -> None:
        rec = self._selected()
        if rec:
            try:
                li.go_to(self._word, rec)
            except Exception as exc:
                QMessageBox.warning(self, "Go to failed", str(exc))

    def _open_file(self) -> None:
        rec = self._selected()
        if rec and not li.open_source(rec):
            QMessageBox.information(self, "File not found", f"Cannot open:\n{rec.absolute_form}")

    def _relink(self) -> None:
        rec = self._selected()
        if not rec:
            return
        start = os.path.dirname(rec.absolute_form) if rec.absolute_form else self._doc_dir
        path, _ = QFileDialog.getOpenFileName(self, "Relink to file", start, _FILE_FILTER.get(rec.kind, "All files (*.*)"))
        if not path:
            return
        # Preserve the link's current path form (relative unless it was absolute / cross-drive).
        new_stored = path
        if not rec.is_absolute:
            try:
                new_stored = os.path.relpath(path, self._doc_dir).replace(os.sep, "/")
            except ValueError:
                new_stored = path
        self._apply_new_path(rec, new_stored)

    def _on_form_toggled(self) -> None:
        if self._updating:
            return
        rec = self._selected()
        if not rec:
            return
        want_absolute = self._abs_radio.isChecked()
        if want_absolute == rec.is_absolute:
            return
        new_stored = rec.absolute_form if want_absolute else rec.relative_form
        if not new_stored:
            return
        self._apply_new_path(rec, new_stored)

    def _apply_new_path(self, rec, new_stored: str) -> None:
        row = self._current_row()
        try:
            new_rec = li.set_link_path(rec, new_stored, self._doc_dir)
        except Exception as exc:
            QMessageBox.critical(self, "Update failed", str(exc))
            return
        self._records[row] = new_rec
        self._set_row(row, new_rec)
        self._on_selection()

    def _selected(self):
        row = self._current_row()
        return self._records[row] if 0 <= row < len(self._records) else None
