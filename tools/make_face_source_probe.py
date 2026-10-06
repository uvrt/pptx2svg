#!/usr/bin/env python3
"""Build a deck that measures *which copy* of a face PowerPoint lays out and draws with.

A family can be installed in more than one place on a Mac with Office: macOS's own fonts,
PowerPoint's bundle (``Microsoft PowerPoint.app/Contents/Resources/DFonts``) and Office's
cloud-font cache.  Where two copies have different advance widths, the positions of the
glyphs in PowerPoint's PDF export say which one it used, and so does the face the PDF
embeds.  ``tools/read_face_source_probe.py`` reads both and names the matching copy.

Every probe is one line in its own text box -- no wrap, no insets, left aligned, 24 pt --
in one family and style:

``rockwell-*``
    Rockwell, in all four styles.  macOS ships Rockwell 13.0 and PowerPoint's bundle
    Rockwell 1.65, and the two disagree on every printable ASCII advance: this is the
    family that settles the precedence between the system and the bundle.
``symbol``
    ``a b g`` in Symbol.  macOS's Symbol maps Unicode Greek only; the bundle's is
    symbol-encoded, so these three letters are alpha, beta and gamma only in the bundle.
``arial`` / ``verdana`` / ``tahoma`` / ``times``
    Families in both places with equal advances: controls, and the face the PDF names.
``aptos`` / ``aptos-display`` / ``aptos-narrow`` / ``calibri`` / ``segoe`` / ``lato``
    Families in exactly one place -- PowerPoint's bundle or the cloud cache -- so the
    reader can confirm they are found there at all.

Usage::

    python3 tools/make_face_source_probe.py ~/face-source-probe.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/face-source-probe.pptx ~/face-source-probe.pdf
    python3 tools/read_face_source_probe.py ~/face-source-probe.pptx ~/face-source-probe.pdf

A probe deck and its export are throwaway and must be deleted again.
"""

from __future__ import annotations

import sys
from pathlib import Path
from xml.sax.saxutils import escape

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(HERE))

from make_feature_sweep import NAMESPACES, PT  # noqa: E402
from make_style_probe import write_deck  # noqa: E402

SOURCE = ROOT / "tests/fixtures/real-financial-report.pptx"

LEFT, TOP, ROW, BOX_W, BOX_H = 30, 20, 34, 900, 30
SIZE = 2400
TEXT = "Hamburgefonstiv 0123"

#: ``(key, typeface, bold, italic, text)``; one box per row, top to bottom.
PROBES: list[tuple[str, str, bool, bool, str]] = [
    ("rockwell-regular", "Rockwell", False, False, TEXT),
    ("rockwell-bold", "Rockwell", True, False, TEXT),
    ("rockwell-italic", "Rockwell", False, True, TEXT),
    ("rockwell-bolditalic", "Rockwell", True, True, TEXT),
    ("symbol", "Symbol", False, False, "abgdpqw"),
    ("arial", "Arial", False, False, TEXT),
    ("verdana", "Verdana", False, False, TEXT),
    ("tahoma", "Tahoma", False, False, TEXT),
    ("times", "Times New Roman", False, False, TEXT),
    ("aptos", "Aptos", False, False, TEXT),
    ("aptos-display", "Aptos Display", False, False, TEXT),
    ("aptos-narrow", "Aptos Narrow", False, False, TEXT),
    ("calibri", "Calibri", False, False, TEXT),
    ("segoe", "Segoe UI", False, False, TEXT),
    ("lato", "Lato", False, False, TEXT),
]


def box(index: int, key: str, typeface: str, bold: bool, italic: bool, text: str) -> str:
    attrs = f"lang='en-US' sz='{SIZE}'" + (" b='1'" if bold else "") + (" i='1'" if italic else "")
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='{100 + index}' name='{key}'/>"
        "<p:cNvSpPr txBox='1'/><p:nvPr/></p:nvSpPr><p:spPr>"
        f"<a:xfrm><a:off x='{LEFT * PT}' y='{(TOP + index * ROW) * PT}'/>"
        f"<a:ext cx='{BOX_W * PT}' cy='{BOX_H * PT}'/></a:xfrm>"
        "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
        "<p:txBody><a:bodyPr wrap='none' lIns='0' tIns='0' rIns='0' bIns='0'/><a:lstStyle/>"
        f"<a:p><a:r><a:rPr {attrs} dirty='0'><a:latin typeface='{escape(typeface)}'/>"
        f"<a:cs typeface='{escape(typeface)}'/></a:rPr><a:t>{escape(text)}</a:t></a:r>"
        "</a:p></p:txBody></p:sp>"
    )


def slides() -> list[str]:
    return [
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
        '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + "".join(box(index, *probe) for index, probe in enumerate(PROBES))
        + "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
    ]


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    target = Path(sys.argv[1]).expanduser()
    count = write_deck(target, SOURCE, slides())
    print(f"{len(PROBES)} probes on {count} slide(s) -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
