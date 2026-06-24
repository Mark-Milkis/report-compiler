"""PySide6 GUI dialogs for the Word integration (driven from the COM server).

The dialogs run as their own process (launched by the COM server) and attach to the
live Word instance via COM to write back. Pure, Qt-free logic lives in
``overlay_logic`` so it can be unit-tested without PySide6 or Word.
"""
