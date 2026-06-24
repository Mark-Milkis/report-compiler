"""Assemble the Word ribbon template (.dotm) from tracked plain-text sources.

A .dotm is an OPC (Open Packaging Conventions) ZIP. Every part is plain text or PNG
except ``word/vbaProject.bin`` (the compiled VBA, produced separately by
``template_builder``). This module builds the package with the standard library only —
no Word, no Office RibbonX Editor:

    skeleton/**            static WordprocessingML + content-types + rels (tracked)
    report_compiler_UI.xml -> customUI/customUI14.xml   (the ribbon)
    icons/<name>.png       -> customUI/images/<name>.png (only those the ribbon uses)
    vbaProject.bin         -> word/vbaProject.bin        (the compiled VBA blob)
"""

from __future__ import annotations

import xml.etree.ElementTree as ET
import zipfile
from pathlib import Path
from typing import List, Tuple

_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
_IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
_CONTENT_TYPES = "[Content_Types].xml"


def _image_targets(customui_rels: Path) -> List[str]:
    """Return the image filenames the ribbon references, from customUI14.xml.rels.

    Driving the image list off the relationships file keeps the package in sync with
    what the ribbon actually declares, so we never ship orphan icons or miss one.
    """
    tree = ET.parse(customui_rels)
    names: List[str] = []
    for rel in tree.getroot().findall(f"{{{_REL_NS}}}Relationship"):
        if rel.get("Type") == _IMAGE_REL_TYPE:
            target = rel.get("Target", "")  # e.g. "images/compile-report.png"
            names.append(target.rsplit("/", 1)[-1])
    return names


def _iter_skeleton(skeleton_dir: Path):
    """Yield (arcname, abspath) for every file under the skeleton, content-types first."""
    files = [p for p in skeleton_dir.rglob("*") if p.is_file()]

    def arcname(p: Path) -> str:
        return p.relative_to(skeleton_dir).as_posix()

    # [Content_Types].xml conventionally comes first in an OPC package.
    files.sort(key=lambda p: (arcname(p) != _CONTENT_TYPES, arcname(p)))
    for p in files:
        yield arcname(p), p


def package_template(
    skeleton_dir: Path,
    customui_xml: Path,
    icons_dir: Path,
    vbaproject_bin: Path,
    output_dotm: Path,
    logger,
) -> Tuple[bool, str]:
    """Build ``output_dotm`` from the tracked sources. Pure Python; no Word required."""
    skeleton_dir = Path(skeleton_dir)
    customui_xml = Path(customui_xml)
    icons_dir = Path(icons_dir)
    vbaproject_bin = Path(vbaproject_bin)
    output_dotm = Path(output_dotm)

    # Validate inputs up front with actionable messages.
    if not skeleton_dir.is_dir():
        return False, f"Skeleton directory not found: {skeleton_dir}"
    customui_rels = skeleton_dir / "customUI" / "_rels" / "customUI14.xml.rels"
    if not customui_rels.exists():
        return False, f"Missing ribbon relationships: {customui_rels}"
    if not customui_xml.exists():
        return False, f"Ribbon XML not found: {customui_xml}"
    if not vbaproject_bin.exists():
        return False, (
            f"vbaProject.bin not found: {vbaproject_bin}. "
            "Run 'word-integration build-vba' first to generate it."
        )

    images = _image_targets(customui_rels)
    missing = [n for n in images if not (icons_dir / n).exists()]
    if missing:
        return False, f"Missing icon(s) in {icons_dir}: {', '.join(missing)}"

    output_dotm.parent.mkdir(parents=True, exist_ok=True)
    logger.info(f"Packaging template -> {output_dotm}")
    with zipfile.ZipFile(output_dotm, "w", zipfile.ZIP_DEFLATED) as z:
        # 1. Static skeleton parts.
        for arcname, path in _iter_skeleton(skeleton_dir):
            z.write(path, arcname)
            logger.debug(f"  + {arcname}")
        # 2. Ribbon.
        z.write(customui_xml, "customUI/customUI14.xml")
        logger.debug("  + customUI/customUI14.xml")
        # 3. Icons the ribbon references.
        for name in images:
            z.write(icons_dir / name, f"customUI/images/{name}")
            logger.debug(f"  + customUI/images/{name}")
        # 4. Compiled VBA.
        z.write(vbaproject_bin, "word/vbaProject.bin")
        logger.debug("  + word/vbaProject.bin")

    return True, (
        f"Packaged {output_dotm.name} "
        f"({len(images)} icon(s), ribbon + VBA from sources)"
    )
