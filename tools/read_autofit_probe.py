#!/usr/bin/env python3
"""Read a ``make_autofit_probe.py`` deck's PDF export next to pptx2svg's SVG of it.

For every probe it reads, from PowerPoint's export, the size of each run (PyMuPDF's
``rawdict`` span size), the baseline of each line (its glyphs' origins), and the height
of the grey box -- and the same from pptx2svg's SVG: each ``tspan``'s ``font-size`` and
the running sum of the ``dy`` advances, from the box's ``rect``.  A row is ``ok`` when the
sizes agree to 0.05 pt (or, for a single size, the size the first line's advance implies
agrees to ``DRAWN``), the line count agrees, the first and last baselines agree to
``TOLERANCE`` and the box is as tall.

Usage::

    python3 tools/read_autofit_probe.py deck.pptx deck.pdf

Needs PyMuPDF; pptx2svg itself is imported from the checkout.
"""

from __future__ import annotations

import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "src"))
sys.path.insert(0, str(HERE))

import pymupdf  # noqa: E402

from make_autofit_probe import BOX_H, BOX_X, BOX_Y, PROBES  # noqa: E402

#: Points a baseline may differ by and still agree.
TOLERANCE = 1.0
#: How far, as a fraction, the size a run's advance implies may be from ours: hinting
#: puts 11 pt Calibri's advance 1 % wide of its outline's.
DRAWN = 0.015
#: SVG pixels per point.
PX = 4 / 3


def read_export(pdf: Path) -> list[dict]:
    found = []
    with pymupdf.open(pdf) as document:
        for page in document:
            lines: dict[float, None] = {}
            sizes = set()
            drawn = None
            for block in page.get_text("rawdict")["blocks"]:
                for line in block.get("lines", []):
                    if drawn is None:
                        drawn = _drawn_size(line["spans"][0]["chars"])
                    for span in line["spans"]:
                        for char in span["chars"]:
                            if char["c"].isspace():
                                continue
                            lines[round(char["origin"][1], 2)] = None
                            sizes.add(round(span["size"], 2))
            boxes = [
                drawing["rect"] for drawing in page.get_drawings()
                if drawing.get("fill") and abs(drawing["rect"].x0 - BOX_X) < 0.5
            ]
            found.append({
                "sizes": sorted(sizes),
                "drawn": drawn,
                "baselines": sorted(lines),
                "box": round(boxes[0].height, 2) if boxes else None,
            })
    return found


def _drawn_size(chars: list[dict]) -> float:
    """The size a run's glyphs were drawn at, from their advance, pt.

    The PDF's own font size is not it for a size off the whole point: PowerPoint writes a
    stated 10.5 pt run as an 11 pt font whose advances are 10.5 pt's.
    """
    from pptx2svg import DefaultTextMeasurer

    text = "".join(char["c"] for char in chars[:-1])
    advance = chars[-1]["origin"][0] - chars[0]["origin"][0]
    at_one = DefaultTextMeasurer().measure_text_width(text, 1.0, False, "Calibri", None) / PX
    return advance / at_one


def read_ours(pptx: Path) -> list[dict]:
    from pptx2svg import convert_pptx_to_svg

    found = []
    for svg in convert_pptx_to_svg(str(pptx)):
        top = float(re.search(r'transform="translate\([^,]+, ([^)]+)\)"', svg).group(1))
        box = float(re.search(r'<rect fill="#d9d9d9"[^>]* height="([^"]+)"', svg).group(1))
        text = re.search(r'<text x="[^"]*" y="([^"]+)"', svg)
        y = top + float(text.group(1))
        baselines, sizes = [], set()
        for attrs in re.findall(r"<tspan([^>]*)>", svg):
            dy = re.search(r' dy="([^"]+)"', attrs)
            if dy and (float(dy.group(1)) or not baselines):
                y += float(dy.group(1))
                baselines.append(round(y / PX, 2))
            sizes.add(round(float(re.search(r'font-size="([^"]+)"', attrs).group(1)) / PX, 2))
        found.append({"sizes": sorted(sizes), "baselines": baselines, "box": round(box / PX, 2)})
    return found


def describe(entry: dict) -> str:
    lines = entry["baselines"]
    span = f"{lines[0]:7.2f} {lines[-1]:7.2f}" if lines else " " * 15
    sizes = "/".join(f"{size:g}" for size in entry["sizes"])
    return f"{sizes:>11} {len(lines):3d} {span} {entry['box']!s:>6}"


def agrees(theirs: dict, ours: dict) -> bool:
    sizes = (
        len(theirs["sizes"]) == len(ours["sizes"])
        and all(abs(a - b) <= 0.05 for a, b in zip(theirs["sizes"], ours["sizes"]))
    ) or (
        len(ours["sizes"]) == 1 and abs(theirs["drawn"] / ours["sizes"][0] - 1) <= DRAWN
    )
    return (
        sizes
        and len(theirs["baselines"]) == len(ours["baselines"])
        and abs(theirs["baselines"][0] - ours["baselines"][0]) <= TOLERANCE
        and abs(theirs["baselines"][-1] - ours["baselines"][-1]) <= TOLERANCE
        and theirs["box"] is not None
        and abs(theirs["box"] - ours["box"]) <= 0.5
    )


def main() -> int:
    if len(sys.argv) != 3:
        print(__doc__)
        return 2
    theirs = read_export(Path(sys.argv[2]))
    ours = read_ours(Path(sys.argv[1]))
    header = f"{'sizes':>11} {'n':>3} {'first':>7} {'last':>7} {'box':>6}"
    print(f"{'probe':<18} PowerPoint {header[10:]}   pptx2svg {header[9:]}")
    agreed = 0
    for spec, a, b in zip(PROBES, theirs, ours):
        ok = agrees(a, b)
        agreed += ok
        print(f"{spec['key']:<18} {describe(a)}   {describe(b)}  {'ok' if ok else 'DIFF'}")
    print(f"{agreed} of {len(PROBES)} agree; the box is {BOX_H} pt tall at y = {BOX_Y}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
