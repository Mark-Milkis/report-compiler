"""Entry point for GUI dialogs: ``python -m report_compiler.gui <command> [opts]``.

Launched by the COM server (e.g. ``LaunchOverlayDialog``) as its own process, or run
manually for development.
"""

from __future__ import annotations

import argparse
import sys


def main(argv: list[str] | None = None) -> int:
    parser = argparse.ArgumentParser(prog="report_compiler.gui")
    sub = parser.add_subparsers(dest="command", required=True)

    overlay = sub.add_parser("overlay", help="Insert PDF overlay dialog")
    overlay.add_argument("--doc", default="", help="Local path of the active Word document")
    overlay.add_argument("--anchor", default="", help="Bookmark name to insert at")

    args = parser.parse_args(argv)

    if args.command == "overlay":
        from PySide6.QtWidgets import QApplication
        from report_compiler.gui.overlay_dialog import OverlayDialog

        app = QApplication(sys.argv)
        dialog = OverlayDialog(doc_path=args.doc, anchor=args.anchor)
        dialog.show()
        return app.exec()

    parser.error(f"unknown command: {args.command}")
    return 2


if __name__ == "__main__":
    raise SystemExit(main())
