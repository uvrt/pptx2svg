#!/usr/bin/env python3
"""Build a deck that measures where PowerPoint puts a run that follows another.

Every probe is one line in its own text box -- no wrap, no insets, left aligned -- made of
two runs: the first ends in a space (``Revenue grew `` then ``12%``), and the second is in
another face, or the same face in another style, or the same face again as the control.
``tools/read_run_probe.py`` reads the origin of the second run's first glyph off
PowerPoint's PDF export and compares it with where pptx2svg's SVG puts it.

``pair-*``
    The faces E6 drafted: Calibri then Calibri Italic, Bold, Consolas, Cambria; Arial and
    Aptos then Consolas; Consolas then Calibri.
``nospace-*`` / ``lead-*``
    The same pairs with no space at the join, and with the space opening the second run
    instead of closing the first: which face the space is measured in.
``size-*``
    The first pair at 12, 18 and 36 pt, and with the second run at another size.
``bullet-*`` / ``ctr-*``
    A bulleted paragraph and a centred one, where the line does not start at the inset.

Usage::

    python3 tools/make_run_probe.py ~/Documents/run-probe.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/Documents/run-probe.pptx ~/Documents/run-probe.pdf
    python3 tools/read_run_probe.py ~/Documents/run-probe.pptx ~/Documents/run-probe.pdf

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

#: The 2013 Office theme: Calibri Light and Calibri.
SOURCE = ROOT / "tests/fixtures/real-financial-report.pptx"

#: Boxes down the slide, two columns; 960 x 540 pt.
LEFT, TOP, ROW, COLUMN_GAP, BOX_W, BOX_H = 30, 20, 40, 470, 440, 36
ROWS = (540 - TOP) // ROW

BULLET = (
    "<a:pPr marL='342900' indent='-342900'><a:buFont typeface='Arial'/>"
    "<a:buChar char='&#8226;'/></a:pPr>"
)


def face(typeface: str | None = None, *, italic: bool = False, bold: bool = False,
         size: int = 2400) -> dict:
    return {"typeface": typeface, "italic": italic, "bold": bold, "size": size}


#: Every probe: its key, the two runs' text and faces, and the paragraph's ``a:pPr``.
PROBES: list[dict] = []


def probe(key: str, first: str, second: str, a: dict, b: dict, ppr: str = "") -> None:
    PROBES.append({"key": key, "first": first, "second": second, "a": a, "b": b, "ppr": ppr})


PAIRS = {
    "calibri-italic": (face(), face(italic=True)),
    "calibri-bold": (face(), face(bold=True)),
    "calibri-consolas": (face(), face("Consolas")),
    "calibri-cambria": (face(), face("Cambria")),
    "arial-consolas": (face("Arial"), face("Consolas")),
    "aptos-consolas": (face("Aptos"), face("Consolas")),
    "aptos-italic": (face("Aptos"), face("Aptos", italic=True)),
    "consolas-calibri": (face("Consolas"), face()),
    "calibri-calibri": (face(), face()),
}
for name, (a, b) in PAIRS.items():
    probe(f"pair-{name}", "Revenue grew ", "12%", a, b)
    probe(f"nospace-{name}", "Revenue grew", "12%", a, b)
    probe(f"lead-{name}", "Revenue grew", " 12%", a, b)
    probe(f"spaces-{name}", "Operating margin   ", "11.9%", a, b)
for size in (1200, 1800, 3600):
    probe(f"size-{size}", "Revenue grew ", "12%", face(size=size), face("Consolas", size=size))
probe("size-mixed", "Revenue grew ", "12%", face(size=1800), face("Consolas", size=3600))
probe("bullet-consolas", "Operating margin ", "11.9%", face(), face("Consolas"), BULLET)
probe("bullet-italic", "Operating margin ", "11.9%", face(), face(italic=True), BULLET)
probe("ctr-consolas", "Operating margin ", "11.9%", face(), face("Consolas"), "<a:pPr algn='ctr'/>")


def place(index: int) -> tuple[int, float, float]:
    """Slide, x and y (pt) of probe ``index``'s box."""
    slide, within = divmod(index, 2 * ROWS)
    column, row = divmod(within, ROWS)
    return slide, LEFT + column * COLUMN_GAP, TOP + row * ROW


def run_xml(text: str, spec: dict) -> str:
    attrs = f"lang='en-US' sz='{spec['size']}'"
    if spec["italic"]:
        attrs += " i='1'"
    if spec["bold"]:
        attrs += " b='1'"
    latin = f"<a:latin typeface='{spec['typeface']}'/>" if spec["typeface"] else ""
    return f"<a:r><a:rPr {attrs} dirty='0'>{latin}</a:rPr><a:t>{escape(text)}</a:t></a:r>"


def box(shape_id: int, spec: dict, x: float, y: float) -> str:
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='{shape_id}' name='{spec['key']}'/>"
        "<p:cNvSpPr txBox='1'/><p:nvPr/></p:nvSpPr><p:spPr>"
        f"<a:xfrm><a:off x='{int(x * PT)}' y='{int(y * PT)}'/>"
        f"<a:ext cx='{BOX_W * PT}' cy='{BOX_H * PT}'/></a:xfrm>"
        "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom><a:noFill/></p:spPr>"
        "<p:txBody><a:bodyPr wrap='none' lIns='0' tIns='0' rIns='0' bIns='0'/><a:lstStyle/>"
        f"<a:p>{spec['ppr']}{run_xml(spec['first'], spec['a'])}"
        f"{run_xml(spec['second'], spec['b'])}</a:p></p:txBody></p:sp>"
    )


def slides() -> list[str]:
    trees: dict[int, list[str]] = {}
    for index, spec in enumerate(PROBES):
        slide, x, y = place(index)
        trees.setdefault(slide, []).append(box(100 + index, spec, x, y))
    return [
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
        '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + "".join(trees[slide])
        + "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
        for slide in sorted(trees)
    ]


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    target = Path(sys.argv[1]).expanduser()
    count = write_deck(target, SOURCE, slides())
    print(f"{len(PROBES)} probes on {count} slides -> {target}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
