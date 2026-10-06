#!/usr/bin/env python3
"""Read a ``make_face_source_probe.py`` export: which copy of each face PowerPoint used.

For every probe it reads the glyph origins PowerPoint's PDF puts down (PyMuPDF's
``rawdict``) and the face the PDF names for them, then predicts the same origins from
every copy of that family and style installed on this machine -- macOS's fonts,
PowerPoint's bundle, Office's cloud-font cache -- and prints each copy's largest error.
The copy that agrees to a hundredth of a point is the one PowerPoint laid the line out
with; the name the PDF embeds is the one it drew.

Usage::

    python3 tools/read_face_source_probe.py deck.pdf

Needs PyMuPDF and fontTools.  The fonts are read where they are installed; nothing is
copied, and only numbers are printed.
"""

from __future__ import annotations

import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "src"))
sys.path.insert(0, str(HERE))

import pymupdf  # noqa: E402
from fontTools.ttLib import TTCollection, TTFont  # noqa: E402

from make_face_source_probe import BOX_H, LEFT, PROBES, ROW, SIZE, TOP  # noqa: E402

from pptx2svg.fonts import office  # noqa: E402

PT = SIZE / 100


def copies(family: str, bold: bool, italic: bool) -> list[tuple[str, str, object]]:
    """``(location, path, TTFont)`` of every installed face of ``family`` in this style."""
    out = []
    for location, directory in office.search_dirs():
        for path in sorted(Path(directory).iterdir()):
            if path.suffix.lower() not in (".ttf", ".otf", ".ttc"):
                continue
            try:
                fonts = (TTCollection(str(path), lazy=True).fonts if path.suffix.lower() == ".ttc"
                         else [TTFont(str(path), lazy=True)])
            except Exception:
                continue
            for font in fonts:
                try:
                    name = font["name"]
                    selection = font["OS/2"].fsSelection
                except Exception:
                    continue
                if (name.getDebugName(1) or "").lower() != family.lower():
                    continue
                if (bool(selection & 0x20), bool(selection & 0x01)) != (bold, italic):
                    continue
                out.append((location, f"{path.name} ({name.getDebugName(5)})", font))
    return out


def predicted(font, text: str) -> list[float]:
    """Origins of ``text``'s glyphs from the first, in points, by this face's advances."""
    cmap = font.getBestCmap() or {}
    metrics = font["hmtx"].metrics
    upem = font["head"].unitsPerEm
    out, x = [], 0.0
    for char in text:
        out.append(x)
        glyph = cmap.get(ord(char)) or cmap.get(0xF000 + ord(char))
        x += (metrics[glyph][0] if glyph in metrics else 0) * PT / upem
    return out


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    with pymupdf.open(sys.argv[1]) as document:
        page = document[0].get_text("rawdict")
    for index, (key, family, bold, italic, text) in enumerate(PROBES):
        y = TOP + index * ROW
        glyphs = sorted(
            (char["origin"][0], span["font"])
            for block in page["blocks"] for line in block.get("lines", [])
            for span in line["spans"] for char in span["chars"]
            if y <= char["origin"][1] <= y + BOX_H and LEFT - 5 <= char["origin"][0]
        )
        if len(glyphs) != len(text):
            print(f"{key:20} {len(glyphs)} glyphs for {len(text)} characters; skipped")
            continue
        drawn = sorted({name for _, name in glyphs})
        observed = [x - glyphs[0][0] for x, _ in glyphs]
        print(f"{key:20} PDF draws {', '.join(drawn)}")
        found = copies(family, bold, italic)
        if not found:
            print(f"{'':20}   no copy of {family} installed in this style")
        for location, label, font in found:
            error = max(abs(a - b) for a, b in zip(observed, predicted(font, text)))
            print(f"{'':20}   {location:8} {label:52} max error {error:7.3f} pt")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
