#!/usr/bin/env python3
"""Build a deck that measures what PowerPoint draws for each autofit mode on open.

One probe per slide: a grey 400 x 150 pt text box at (60, 40) holding short Calibri
paragraphs that do not wrap, so how many lines there are is known.  The box is filled so
its drawn height can be read off the export too, which is how ``spAutoFit`` shows whether
PowerPoint grew the shape.  ``tools/read_autofit_probe.py`` reads each run's size and each
line's baseline off PowerPoint's PDF export and compares them with pptx2svg's SVG.

``norm-*``
    ``a:normAutofit`` with nothing stored: text that fits, text that overflows at 12, 18
    and 28 pt, a paragraph that wraps past the bottom, and a centred overflow.  (One
    that does not wrap, ``wrap='none'``, overflowing, is left out: PowerPoint 16.106 drops
    the AppleEvent connection exporting it, -609, and refuses every open after that with
    -9074 until it is killed.)
``fs*-*`` / ``ls*-*``
    A stored ``fontScale`` (50, 62.5, 92.5 %) on text that fits and on text that still
    overflows at that scale; a stored ``lnSpcReduction`` alone and with a ``fontScale``,
    and against ``spcBef``, a percentage ``lnSpc`` and an exact ``spcPts`` -- each with a
    ``ctl-*`` twin under ``noAutofit`` that shows the spacing unreduced.
``sp-*`` / ``no-*`` / ``none-*``
    ``a:spAutoFit`` on text taller and shorter than the box; ``a:noAutofit``; no autofit
    element at all.
``round-*``
    A stored ``fontScale`` on sizes it does not take to a whole point (25 pt at 50 %,
    18 pt at 70 %, ...), on a stated 10.5 pt, and on a paragraph with a second, larger
    run; ``ctl-1050`` the 10.5 pt under ``noAutofit``, ``ctl-mixed`` the two runs' scaled
    sizes stated.

Usage::

    python3 tools/make_autofit_probe.py ~/autofit-probe.pptx
    osascript tools/powerpoint_export_pdf.applescript \\
        ~/autofit-probe.pptx ~/autofit-probe.pdf
    python3 tools/read_autofit_probe.py ~/autofit-probe.pptx ~/autofit-probe.pdf

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

#: The 2013 Office theme: Calibri Light and Calibri; a 960 x 540 pt slide.
SOURCE = ROOT / "tests/fixtures/real-financial-report.pptx"

#: The box every probe draws, pt.
BOX_X, BOX_Y, BOX_W, BOX_H = 60, 40, 400, 150

WORDS = "alpha bravo charlie delta echo foxtrot golf hotel india juliet kilo lima".split()
LONG = " ".join(WORDS * 4)

#: Every probe: key, autofit element, run size (hundredths), paragraphs, ``a:pPr`` body,
#: ``a:bodyPr`` attributes beyond the wrap.
PROBES: list[dict] = []


def probe(key: str, fit: str, size: int, lines: int, ppr: str = "", body: str = "",
          text: str | None = None, second: int | None = None) -> None:
    PROBES.append({"key": key, "fit": fit, "size": size, "lines": lines, "ppr": ppr,
                   "body": body, "text": text, "second": second})


def scale(percent: float) -> str:
    return f"<a:normAutofit fontScale='{round(percent * 1000)}'/>"


NORM = "<a:normAutofit/>"
probe("norm-fit-18", NORM, 1800, 4)
probe("norm-over-18", NORM, 1800, 12)
probe("norm-over-12", NORM, 1200, 20)
probe("norm-over-28", NORM, 2800, 8)
probe("norm-over-wrap", NORM, 1800, 1, text=LONG)
probe("norm-over-ctr", NORM, 1800, 12, body=" anchor='ctr'")
probe("fs50-fit", "<a:normAutofit fontScale='50000'/>", 1800, 4)
probe("fs50-over", "<a:normAutofit fontScale='50000'/>", 1800, 30)
probe("fs625-fit", "<a:normAutofit fontScale='62500'/>", 1800, 6)
probe("fs925-fit", "<a:normAutofit fontScale='92500'/>", 1800, 4)
probe("fs925-over", "<a:normAutofit fontScale='92500'/>", 1800, 12)
LS20 = "<a:normAutofit lnSpcReduction='20000'/>"
probe("ls20", LS20, 1800, 6)
probe("ls20-over", LS20, 1800, 12)
probe("fs75-ls20", "<a:normAutofit fontScale='75000' lnSpcReduction='20000'/>", 1800, 10)
SPCBEF = "<a:spcBef><a:spcPts val='1200'/></a:spcBef>"
LNSPC150 = "<a:lnSpc><a:spcPct val='150000'/></a:lnSpc>"
SPCPTS = "<a:lnSpc><a:spcPts val='3000'/></a:lnSpc>"
SPCAFT = "<a:spcAft><a:spcPct val='50000'/></a:spcAft>"
for name, spacing in (("spcbef", SPCBEF), ("lnspc150", LNSPC150), ("spcpts", SPCPTS),
                      ("spcaftpct", SPCAFT)):
    probe(f"ls20-{name}", LS20, 1800, 5, ppr=spacing)
    probe(f"ctl-{name}", "<a:noAutofit/>", 1800, 5, ppr=spacing)
probe("sp-over-18", "<a:spAutoFit/>", 1800, 12)
probe("sp-under-18", "<a:spAutoFit/>", 1800, 2)
probe("sp-under-b", "<a:spAutoFit/>", 1800, 2, body=" anchor='b'")
probe("no-over-18", "<a:noAutofit/>", 1800, 12)
probe("none-over-18", "", 1800, 12)
# What a stored fontScale does to a size it does not divide into whole points, to a run
# of a second size, and to spacing stated in points; a stated half point as the control.
for size, percent in ((2500, 50), (2100, 50), (1300, 50), (2700, 50), (1100, 50),
                      (2000, 62.5), (1400, 90), (1800, 70), (1000, 85), (1800, 55),
                      (1050, 50), (1050, 100)):
    probe(f"round-{size}-{percent}", scale(percent), size, 3)
probe("ctl-1050", "<a:noAutofit/>", 1050, 3)
probe("round-mixed-62.5", scale(62.5), 1800, 3, second=2400)
probe("ctl-mixed", "<a:noAutofit/>", 1100, 3, second=1500)
probe("fs50-spcbef", scale(50), 1800, 5, ppr=SPCBEF)
probe("fs50-spcpts", scale(50), 1800, 5, ppr=SPCPTS)
probe("ls20-spcbefpct", LS20, 1800, 5, ppr="<a:spcBef><a:spcPct val='50000'/></a:spcBef>")
probe("ls10-lnspc90", "<a:normAutofit lnSpcReduction='10000'/>", 1800, 5,
      ppr="<a:lnSpc><a:spcPct val='90000'/></a:lnSpc>")
probe("ctl-lnspc80", "<a:noAutofit/>", 1800, 5, ppr="<a:lnSpc><a:spcPct val='80000'/></a:lnSpc>")


def run(text: str, size: int) -> str:
    return (
        f"<a:r><a:rPr lang='en-US' sz='{size}' dirty='0'>"
        "<a:solidFill><a:srgbClr val='000000'/></a:solidFill><a:latin typeface='Calibri'/>"
        f"</a:rPr><a:t>{escape(text)}</a:t></a:r>"
    )


def paragraph(spec: dict, number: int) -> str:
    text = spec["text"] or f"Line {number:02d} {WORDS[number % len(WORDS)]}"
    ppr = f"<a:pPr>{spec['ppr']}</a:pPr>" if spec["ppr"] else ""
    second = run(" and more", spec["second"]) if spec["second"] else ""
    return f"<a:p>{ppr}{run(text, spec['size'])}{second}</a:p>"


def box(spec: dict) -> str:
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='2' name='{spec['key']}'/>"
        "<p:cNvSpPr txBox='1'/><p:nvPr/></p:nvSpPr><p:spPr>"
        f"<a:xfrm><a:off x='{BOX_X * PT}' y='{BOX_Y * PT}'/>"
        f"<a:ext cx='{BOX_W * PT}' cy='{BOX_H * PT}'/></a:xfrm>"
        "<a:prstGeom prst='rect'><a:avLst/></a:prstGeom>"
        "<a:solidFill><a:srgbClr val='D9D9D9'/></a:solidFill></p:spPr>"
        f"<p:txBody><a:bodyPr wrap='square' rtlCol='0'{spec['body']}>{spec['fit']}</a:bodyPr>"
        "<a:lstStyle/>"
        + "".join(paragraph(spec, n + 1) for n in range(spec["lines"]))
        + "</p:txBody></p:sp>"
    )


def slides() -> list[str]:
    return [
        "<?xml version='1.0' encoding='UTF-8' standalone='yes'?>"
        f"<p:sld {NAMESPACES}><p:cSld>"
        '<p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/>'
        '</p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/>'
        '<a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr>'
        + box(spec)
        + "</p:spTree></p:cSld>"
        "<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:sld>"
        for spec in PROBES
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
