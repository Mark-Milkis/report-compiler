"""Out-of-process COM server exposing Report Compiler to Word (and other COM clients).

The Word ribbon calls ``CreateObject("ReportCompiler.Application")`` and drives
compilation through this server instead of shelling out to ``uvx``. Compilation runs
on a background thread so the COM call returns immediately (async job + status),
keeping Word responsive while the report builds.

Registration is per-user (``HKCU\\Software\\Classes``) so no administrator rights are
needed. Because COM needs a stable command line to launch the server, registration is
captured against a persistent ``uv tool install`` of the package rather than an
ephemeral ``uvx`` environment -- see :func:`bootstrap_register`.
"""

from __future__ import annotations

import os
import subprocess
import sys
import threading
import uuid
from dataclasses import dataclass
from pathlib import Path

# COM machinery is Windows-only. Guard the import so the module (and the CLI that
# imports it) still loads on other platforms / without pywin32 -- the register and
# serve paths simply raise a clear error when actually used.
try:
    import pythoncom  # noqa: F401  (imported for side effects + worker CoInitialize)
    import winreg

    _COM_AVAILABLE = True
except ImportError:  # pragma: no cover - non-Windows
    pythoncom = None
    winreg = None
    _COM_AVAILABLE = False


# --- Stable identity. NEVER change these once registered against real installs. ---
CLSID = "{6B43D8F8-E8E7-4F05-89F8-E4607F9E5702}"
PROGID = "ReportCompiler.Application"
DESC = "Report Compiler COM Server"
CLASS_SPEC = "report_compiler.com_server.ReportCompilerCOMServer"

# Job states (plain strings so any COM client can compare them).
PENDING = "pending"
RUNNING = "running"
SUCCEEDED = "succeeded"
FAILED = "failed"
UNKNOWN = "unknown"


# ---------------------------------------------------------------------------
# Async job tracking
# ---------------------------------------------------------------------------
@dataclass
class _Job:
    status: str = PENDING
    message: str = ""
    output_path: str = ""


class _JobRegistry:
    """Thread-safe table of compile jobs keyed by an opaque job id."""

    def __init__(self) -> None:
        self._jobs: dict[str, _Job] = {}
        self._lock = threading.Lock()

    def create(self, output_path: str) -> str:
        job_id = uuid.uuid4().hex
        with self._lock:
            self._jobs[job_id] = _Job(output_path=output_path)
        return job_id

    def update(self, job_id: str, *, status: str | None = None, message: str | None = None) -> None:
        with self._lock:
            job = self._jobs.get(job_id)
            if job is None:
                return
            if status is not None:
                job.status = status
            if message is not None:
                job.message = message

    def get(self, job_id: str) -> _Job | None:
        with self._lock:
            return self._jobs.get(job_id)


_JOBS = _JobRegistry()


def _run_compile(job_id: str, input_path: str, output_path: str) -> None:
    """Worker-thread body: run the compile pipeline and record the outcome."""
    # Any thread that touches COM (we drive Word via win32com) must init COM itself.
    if pythoncom is not None:
        pythoncom.CoInitialize()
    try:
        _JOBS.update(job_id, status=RUNNING)

        from report_compiler.cli import handle_compilation
        from report_compiler.utils.logging_config import get_logger, setup_logging

        # A windowless server has no console, so log each job to a file beside the
        # output. Users (and the future GUI) can open it to see what happened.
        log_file = os.path.splitext(output_path)[0] + ".compile.log"
        setup_logging(log_file=log_file, verbose=True)
        logger = get_logger()

        rc = handle_compilation(input_path, output_path, False, logger)
        if rc == 0:
            _JOBS.update(job_id, status=SUCCEEDED, message=f"Compiled to {output_path}")
        else:
            _JOBS.update(
                job_id,
                status=FAILED,
                message=f"Compilation failed. See log: {log_file}",
            )
    except Exception as exc:  # noqa: BLE001 - surface any failure to the COM client
        _JOBS.update(job_id, status=FAILED, message=str(exc))
    finally:
        if pythoncom is not None:
            pythoncom.CoUninitialize()


def _run_svg_import(job_id: str, pdf_path: str, output_path: str, page: str) -> None:
    """Worker-thread body: convert PDF page(s) to SVG and record the outcome.

    PDF->SVG uses PyMuPDF (no Word), so no COM init is needed on this thread.
    """
    try:
        _JOBS.update(job_id, status=RUNNING)

        from report_compiler.cli import handle_svg_import
        from report_compiler.utils.logging_config import get_logger, setup_logging

        log_file = os.path.splitext(output_path)[0] + ".svg.log"
        setup_logging(log_file=log_file, verbose=True)
        logger = get_logger()

        rc = handle_svg_import(pdf_path, output_path, page, logger)
        if rc == 0:
            _JOBS.update(job_id, status=SUCCEEDED, message=f"Converted to {output_path}")
        else:
            _JOBS.update(
                job_id,
                status=FAILED,
                message=f"SVG conversion failed. See log: {log_file}",
            )
    except Exception as exc:  # noqa: BLE001 - surface any failure to the COM client
        _JOBS.update(job_id, status=FAILED, message=str(exc))


# ---------------------------------------------------------------------------
# COM-exposed object
# ---------------------------------------------------------------------------
class ReportCompilerCOMServer:
    """COM entry point for the Report Compiler.

    Methods take and return plain strings so any COM client (VBA, PowerShell, a
    future GUI) can call them without type-library marshalling. Compilation is
    asynchronous: :meth:`CompileAsync` returns a job id immediately and the client
    polls :meth:`GetJobStatus` / :meth:`GetJobMessage`.
    """

    _public_methods_ = [
        "CompileAsync",
        "SvgImportAsync",
        "LaunchOverlayDialog",
        "LaunchLinkManager",
        "SetOverlayPreview",
        "GetJobStatus",
        "GetJobMessage",
    ]
    _reg_clsid_ = CLSID
    _reg_progid_ = PROGID
    _reg_desc_ = DESC
    _reg_class_spec_ = CLASS_SPEC

    def CompileAsync(self, input_path, output_path):
        """Start a compile on a background thread; return a job id immediately."""
        job_id = _JOBS.create(str(output_path))
        thread = threading.Thread(
            target=_run_compile,
            args=(job_id, str(input_path), str(output_path)),
            daemon=True,
        )
        thread.start()
        return job_id

    def SvgImportAsync(self, pdf_path, output_path, page):
        """Start a PDF->SVG conversion on a background thread; return a job id.

        ``page`` is a spec string: a single page, range ("1-3"), list ("1,3,5"), or
        "all". Multiple pages produce ``<stem>_page_<n>.svg`` files next to output_path.
        """
        job_id = _JOBS.create(str(output_path))
        thread = threading.Thread(
            target=_run_svg_import,
            args=(job_id, str(pdf_path), str(output_path), str(page)),
            daemon=True,
        )
        thread.start()
        return job_id

    def LaunchOverlayDialog(self, doc_path, anchor):
        """Spawn the Insert PDF Overlay GUI as its own process and return immediately.

        Not part of the async-job model — this just launches a UI process, which then
        attaches to Word itself to write back. Uses this interpreter (the stable tool
        env) so PySide6 and the package are available.
        """
        subprocess.Popen(
            [
                sys.executable,
                "-m",
                "report_compiler.gui",
                "overlay",
                "--doc",
                str(doc_path),
                "--anchor",
                str(anchor),
            ],
            close_fds=True,
        )
        return "launched"

    def LaunchLinkManager(self, doc_path):
        """Spawn the Link Manager GUI as its own process and return immediately."""
        subprocess.Popen(
            [sys.executable, "-m", "report_compiler.gui", "link-manager", "--doc", str(doc_path)],
            close_fds=True,
        )
        return "launched"

    def SetOverlayPreview(self, doc_path, mode):
        """Switch the document's overlay view: 'tags', 'quick', or 'full'.

        Runs synchronously (manipulates the live document) and returns a short summary.
        Errors are raised to the COM client.
        """
        from report_compiler.document import overlay_preview

        return overlay_preview.set_overlay_view(str(doc_path), str(mode))

    def GetJobStatus(self, job_id):
        job = _JOBS.get(str(job_id))
        return job.status if job else UNKNOWN

    def GetJobMessage(self, job_id):
        job = _JOBS.get(str(job_id))
        return job.message if job else "Unknown job id"


# ---------------------------------------------------------------------------
# Per-user (HKCU) registration
# ---------------------------------------------------------------------------
_CLASSES = r"Software\Classes"


def _require_com() -> None:
    if not _COM_AVAILABLE:
        raise RuntimeError("COM registration is only available on Windows with pywin32 installed.")


def _clsid_key() -> str:
    return rf"{_CLASSES}\CLSID\{CLSID}"


def _progid_key() -> str:
    return rf"{_CLASSES}\{PROGID}"


def _set_default(subkey: str, value: str) -> None:
    with winreg.CreateKey(winreg.HKEY_CURRENT_USER, subkey) as key:
        winreg.SetValueEx(key, "", 0, winreg.REG_SZ, value)


def _read_default(subkey: str) -> str | None:
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, subkey) as key:
            return winreg.QueryValueEx(key, "")[0]
    except FileNotFoundError:
        return None


def _delete_tree(subkey: str) -> None:
    """Recursively delete an HKCU subkey (DeleteKey only removes leaf keys)."""
    try:
        with winreg.OpenKey(winreg.HKEY_CURRENT_USER, subkey, 0, winreg.KEY_ALL_ACCESS) as key:
            while True:
                try:
                    child = winreg.EnumKey(key, 0)
                except OSError:
                    break
                _delete_tree(subkey + "\\" + child)
        winreg.DeleteKey(winreg.HKEY_CURRENT_USER, subkey)
    except FileNotFoundError:
        pass


def _localserver_command() -> str:
    """Build the LocalServer32 command for the *currently running* interpreter.

    Captures this interpreter's ``pythonw.exe`` plus pywin32's ``localserver.py``,
    which is exactly what we want when this runs inside the stable ``uv tool``
    install (see :func:`bootstrap_register`). COM appends ``-Embedding`` at launch.
    """
    import win32com.server.register as reg

    exe = reg._find_localserver_exe(True)
    pyfile = reg._find_localserver_module()
    return f'"{exe}" "{pyfile}" {CLSID}'


def register_self() -> str:
    """Write the per-user COM registration for *this* interpreter. Returns the command."""
    _require_com()
    command = _localserver_command()

    clsid_key = _clsid_key()
    _set_default(clsid_key, DESC)
    _set_default(clsid_key + r"\LocalServer32", command)
    _set_default(clsid_key + r"\PythonCOM", CLASS_SPEC)

    progid_key = _progid_key()
    _set_default(progid_key, DESC)
    _set_default(progid_key + r"\CLSID", CLSID)
    return command


def unregister_self() -> None:
    """Remove the per-user COM registration."""
    _require_com()
    _delete_tree(_clsid_key())
    _delete_tree(_progid_key())


def status() -> dict:
    """Return current registration state for display."""
    _require_com()
    local_server = _read_default(_clsid_key() + r"\LocalServer32")
    progid_clsid = _read_default(_progid_key() + r"\CLSID")
    return {
        "progid": PROGID,
        "clsid": CLSID,
        "registered": bool(local_server) and progid_clsid == CLSID,
        "local_server": local_server,
        "class_spec": _read_default(_clsid_key() + r"\PythonCOM"),
    }


# ---------------------------------------------------------------------------
# Bootstrap: install a stable interpreter, then register against it
# ---------------------------------------------------------------------------
def _source_checkout_root() -> Path | None:
    """Return the repo root if we're running from a source checkout, else None."""
    # src/report_compiler/com_server.py -> repo root is parents[2]
    root = Path(__file__).resolve().parents[2]
    return root if (root / "pyproject.toml").exists() else None


def _installed_tool_python() -> Path:
    """Locate the python.exe inside the persistent ``uv tool install`` of this package."""
    tool_dir = subprocess.run(
        ["uv", "tool", "dir"], capture_output=True, text=True, check=True
    ).stdout.strip()
    py = Path(tool_dir) / "report-compiler" / "Scripts" / "python.exe"
    if not py.exists():
        raise RuntimeError(f"Could not find the installed tool interpreter at: {py}")
    return py


def bootstrap_register() -> str:
    """Public registration entry point used by ``com-server register``.

    1. ``uv tool install`` the package so a stable interpreter exists.
    2. Re-run ``com-server _register-self`` *with that interpreter* so the registry
       captures a path that survives uvx's ephemeral environments.

    Returns the resolved LocalServer32 command for reporting.
    """
    _require_com()

    source = _source_checkout_root()
    target = str(source) if source else "report-compiler"
    subprocess.run(["uv", "tool", "install", "--force", target], check=True)

    py = _installed_tool_python()
    subprocess.run(
        [str(py), "-m", "report_compiler.cli", "com-server", "_register-self"],
        check=True,
    )
    # Report what the installed interpreter registered.
    return _read_default(_clsid_key() + r"\LocalServer32") or "(registered)"
