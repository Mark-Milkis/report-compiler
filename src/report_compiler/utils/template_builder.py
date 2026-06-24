"""Compile the VBA modules (.bas) into a ``vbaProject.bin`` using Word.

``vbaProject.bin`` is an OLE2 / MS-OVBA binary (compressed source + compiled p-code);
Word's VBA engine is the only reliable way to author it. This is the single
Word-dependent step in building the ribbon template — the rest is assembled by
:mod:`report_compiler.utils.template_packager` with no Word.

We work on a *copy* of an existing macro-enabled template (the carrier): wipe its
standard/class modules, import the ``.bas`` files, then do a plain in-place ``Save()``
(no ``SaveAs`` — that path is finicky and raised Word error 4198 here). Finally we lift
``word/vbaProject.bin`` out of the saved copy. Wiping every std/class module first means
the project ends up containing exactly our sources, with no stale leftovers (this is
what fixes the duplicate-module problem).

Requires Word and "Trust access to the VBA project object model" (we enable that
HKCU setting best-effort).
"""

from __future__ import annotations

import shutil
import tempfile
import zipfile
from pathlib import Path
from typing import Tuple

try:
    import win32com.client
except ImportError:  # pragma: no cover - non-Windows
    win32com = None

try:
    import winreg
except ImportError:  # pragma: no cover - non-Windows
    winreg = None

# VBComponent types we replace (standard + class modules); ThisDocument is left alone.
_VBEXT_CT_STDMODULE = 1
_VBEXT_CT_CLASSMODULE = 2


def _enable_access_vbom(word_version: str) -> bool:
    """Best-effort enable of 'Trust access to the VBA project object model' (HKCU)."""
    if winreg is None:
        return False
    key_path = rf"Software\Microsoft\Office\{word_version}\Word\Security"
    try:
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER, key_path) as key:
            winreg.SetValueEx(key, "AccessVBOM", 0, winreg.REG_DWORD, 1)
        return True
    except OSError:
        return False


def build_vba_bin(bas_dir: Path, carrier_dotm: Path, output_bin: Path, logger) -> Tuple[bool, str]:
    """Compile every ``*.bas`` in ``bas_dir`` into ``output_bin`` (vbaProject.bin).

    ``carrier_dotm`` is an existing macro-enabled template used only as a host for the
    VBA project; we copy it, refresh its modules, save in place, and extract the bin.

    Returns (success, message).
    """
    if win32com is None:
        return False, "pywin32 is not installed; cannot automate Word."

    bas_dir = Path(bas_dir)
    carrier_dotm = Path(carrier_dotm)
    output_bin = Path(output_bin)
    bas_files = sorted(bas_dir.glob("*.bas"))
    if not bas_files:
        return False, f"No .bas source files found in {bas_dir}."
    if not carrier_dotm.exists():
        return False, (
            f"Carrier template not found: {carrier_dotm}. "
            "Run 'word-integration package' once to create it (or restore it from git)."
        )

    tmp_dir = Path(tempfile.mkdtemp(prefix="rc_vba_"))
    # Keep the carrier's extension so Save() round-trips the same macro-enabled format.
    tmp_doc = tmp_dir / ("carrier" + carrier_dotm.suffix)
    shutil.copy2(carrier_dotm, tmp_doc)

    word = None
    doc = None
    try:
        word = win32com.client.Dispatch("Word.Application")
        word.Visible = False
        try:
            word.DisplayAlerts = 0
        except Exception:
            pass

        _enable_access_vbom(str(word.Version))

        logger.info(f"Opening carrier template: {tmp_doc.name}")
        doc = word.Documents.Open(str(tmp_doc.resolve()))

        try:
            components = doc.VBProject.VBComponents
        except Exception as e:
            return (
                False,
                "Could not access the VBA project. Enable Word > File > Options > "
                "Trust Center > Macro Settings > 'Trust access to the VBA project "
                f"object model' and try again. ({e})",
            )

        # Wipe every std/class module so the project ends up as exactly our sources.
        for comp in list(components):
            if comp.Type in (_VBEXT_CT_STDMODULE, _VBEXT_CT_CLASSMODULE):
                logger.info(f"  Removing existing module: {comp.Name}")
                components.Remove(comp)

        for bas in bas_files:
            logger.info(f"  Importing {bas.name}")
            components.Import(str(bas.resolve()))

        logger.info("Saving template (in place)...")
        doc.Save()
        doc.Close(SaveChanges=False)
        doc = None

        # Lift word/vbaProject.bin out of the saved copy.
        with zipfile.ZipFile(tmp_doc) as z:
            if "word/vbaProject.bin" not in z.namelist():
                return False, "Saved document contains no vbaProject.bin (no VBA compiled?)."
            data = z.read("word/vbaProject.bin")

        output_bin.parent.mkdir(parents=True, exist_ok=True)
        output_bin.write_bytes(data)

        names = ", ".join(b.stem for b in bas_files)
        return True, f"Built {output_bin.name} from {len(bas_files)} module(s): {names}"

    except Exception as e:  # noqa: BLE001
        return False, f"Failed to build vbaProject.bin: {e}"
    finally:
        if doc is not None:
            try:
                doc.Close(SaveChanges=False)
            except Exception:
                pass
        if word is not None:
            try:
                word.Quit()
            except Exception:
                pass
        shutil.rmtree(tmp_dir, ignore_errors=True)
