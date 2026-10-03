#!/usr/bin/env python3
"""Read a ``make_run_probe.py`` deck's PDF export: where the second run starts, and ours.

For every probe it reads, from PowerPoint's export, the origin of the line's first glyph
and of the second run's first glyph (PyMuPDF's ``rawdict``: each glyph's own origin, as
the content stream places it; a bullet is skipped), and their difference -- the advance
PowerPoint gave the first run, its trailing space included, and a space opening the
second run.  Ours is the same advance as pptx2svg's measurer gives it, which is what its
layout wraps and aligns with and where it puts a run that opens a chunk.  A row is ``ok``
within ``TOLERANCE`` points.

Usage::

    python3 tools/read_run_probe.py deck.pdf

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

from make_run_probe import BOX_H, BOX_W, PROBES, place  # noqa: E402

#: Points a position may differ by and still agree.
TOLERANCE = 0.5
#: SVG pixels per point.
PX = 4 / 3


def read_export(pdf: Path) -> dict[str, tuple[float, float]]:
    """Probe key -> (line start, second run start), points from the box's left edge."""
    found = {}
    with pymupdf.open(pdf) as document:
        pages = [page.get_text("rawdict") for page in document]
    for index, spec in enumerate(PROBES):
        slide, x, y = place(index)
        glyphs = sorted(
            (char["origin"][0], char["c"])
            for block in pages[slide]["blocks"]
            for line in block.get("lines", [])
            for span in line["spans"]
            for char in span["chars"]
            if x - 40 <= char["origin"][0] <= x + BOX_W and y <= char["origin"][1] <= y + BOX_H
            and not char["c"].isspace()
        )
        if "buChar" in spec["ppr"]:
            glyphs = glyphs[1:]
        inked = len(re.sub(r"\s", "", spec["first"]))
        if len(glyphs) <= inked:
            continue
        found[spec["key"]] = (glyphs[0][0] - x, glyphs[inked][0] - x)
    return found


def our_advance(spec: dict) -> float:
    """The first run's advance, and a space opening the second run, as measured, pt."""
    advance = _measure(spec["first"], spec["a"])
    if spec["second"].startswith(" "):
        advance += _measure(" ", spec["b"])
    return advance / PX


def _measure(text: str, spec: dict) -> float:
    from pptx2svg import DefaultTextMeasurer

    return DefaultTextMeasurer().measure_text_width(
        text, spec["size"] / 100, spec["bold"], spec["typeface"] or "Calibri", None
    )


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    power = read_export(Path(sys.argv[1]).expanduser())
    agree = 0
    for spec in PROBES:
        key = spec["key"]
        if key not in power:
            print(f"{key:28s} not read")
            continue
        start, second = power[key]
        drawn, ours = second - start, our_advance(spec)
        ok = abs(ours - drawn) <= TOLERANCE
        agree += ok
        print(
            f"{key:28s} PowerPoint {drawn:7.2f}   ours {ours:7.2f}"
            f"   {'ok' if ok else 'DIFFERS ' + format(ours - drawn, '+.2f')}"
        )
    print(f"{agree} of {len(PROBES)} agree within {TOLERANCE} pt")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
