#!/usr/bin/env python3
"""Read a ``make_kern_source_probe.py`` export: which kern pairs PowerPoint applied.

For every probe line it reads the glyph origins PowerPoint's PDF puts down (PyMuPDF's
``rawdict``), takes off each glyph's own advance -- from the same installed face the
probe chose its pairs from -- and prints what is left at every join, in font units,
beside what the legacy ``kern`` table and the ``GPOS`` ``kern`` feature say it should
be.  A join counts for a table when it is within 3/1000 em of its value (PowerPoint
places a glyph to about 0.05 pt; 3/1000 em is 0.072 pt at 24 pt), for both when the two
agree.

Measured on PowerPoint 16 for Mac (ROADMAP.md, 5.10): every static face -- from
PowerPoint's bundle (Aptos, Calibri, Arial, Times New Roman, Cambria, Verdana, Tahoma,
Meiryo, Yu Gothic), Office's cloud cache (Aptos Display, Segoe UI, Lato, Raleway) or
macOS's folders (Minion Pro, PT Sans, Avenir Next, Hiragino Sans, its kana too) -- was
kerned with its legacy ``kern`` table and never with a pair only ``GPOS`` holds, the
legacy value winning where the two disagree; the two variable faces, Noto Sans JP (Latin
and kana) and STIX Two Text, with their ``GPOS`` pairs.  28 of 29 lines read (PT Sans's
quotes come back out of order from the PDF's text).

Usage::

    python3 tools/read_kern_source_probe.py ~/kern-source-probe.pdf

Needs PyMuPDF and fontTools.  The fonts are read where they are installed; nothing is
copied, and only numbers are printed.
"""

from __future__ import annotations

import collections
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "src"))
sys.path.insert(0, str(HERE))

import pymupdf  # noqa: E402

import make_kern_source_probe as probe  # noqa: E402

PT = probe.SIZE / 100
TOLERANCE_EM = 0.003


def main() -> int:
    if len(sys.argv) != 2:
        print(__doc__)
        return 2
    found = probe.probes()
    with pymupdf.open(sys.argv[1]) as document:
        pages = [page.get_text("rawdict") for page in document]
    totals: dict[str, collections.Counter] = collections.defaultdict(collections.Counter)
    for index, (key, family, script, text) in enumerate(found):
        page = pages[index // probe.ROWS_PER_SLIDE]
        y = probe.TOP + (index % probe.ROWS_PER_SLIDE) * probe.ROW
        glyphs = sorted(
            # A PDF may map a quote's glyph back to U+02BC (Yu Gothic's does).
            (char["origin"][0], char["c"].replace("\u02bc", "\u2019"), span["font"])
            for block in page["blocks"] for line in block.get("lines", [])
            for span in line["spans"] for char in span["chars"]
            if y <= char["origin"][1] <= y + probe.BOX_H and char["c"].strip()
        )
        if [c for _, c, _ in glyphs] != list(text):
            print(f"{key:30} read {''.join(c for _, c, _ in glyphs)!r} for {text!r}; skipped")
            continue
        font = probe.open_face(family)
        cmap = font.getBestCmap()
        upem = font["head"].unitsPerEm
        advances = font["hmtx"].metrics
        legacy, gpos = probe.kern_tables(font)
        tolerance = TOLERANCE_EM * upem
        joins = []
        for (x0, a, _), (x1, b, _) in zip(glyphs, glyphs[1:]):
            first, second = cmap[ord(a)], cmap[ord(b)]
            observed = (x1 - x0) * upem / PT - advances[first][0]
            old, new = legacy.get((first, second), 0), gpos(first, second)
            hit_old, hit_new = abs(observed - old) <= tolerance, abs(observed - new) <= tolerance
            verdict = ("both" if old == new else "legacy+gpos?") if hit_old and hit_new else \
                "legacy" if hit_old else "gpos" if hit_new else "neither"
            if old or new:
                totals[key][verdict] += 1
                joins.append(f"{a}{b} {observed:+.0f} (legacy {old:+d}, gpos {new:+d}) {verdict}")
            elif not hit_old:
                totals[key]["unkerned join moved"] += 1
                joins.append(f"{a}{b} {observed:+.0f} (no pair)")
        drawn = sorted({name for _, _, name in glyphs})
        print(f"{key:30} {', '.join(drawn)}: {dict(totals[key])}")
        for join in joins:
            print(f"{'':32}{join}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
