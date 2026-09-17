#!/usr/bin/env python3
"""Read a ``make_cjk_wrap_probe.py`` export and say where PowerPoint broke each line.

The filler character is 東, whose advance is exactly 1 em, so the first line is reported
as a **count** rather than a width and the boundary case is exact.  Character origins are
read with ``FPDFText_GetCharOrigin`` rather than ink boxes: a box is a glyph's extent and
an origin is an advance, and the two differ by the side bearings.

``--check`` renders the same deck through this library and prints our own first-line
count beside PowerPoint's, so the residual is a signed number of characters rather than
an impression.  It needs the deck the PDF was exported from, found next to it by name.

``--baselines`` reports the distance between consecutive baselines instead of the break,
for the ``spacing`` family, whose subject is vertical.  With ``--check`` it puts our own
advance beside PowerPoint's; ours is read off the ``dy`` chain in the SVG, which is the
same quantity by construction.

Usage::

    python3 tools/read_cjk_wrap_probe.py ~/pptx2svg-oracle/cjk-wrap-box.pdf
    python3 tools/read_cjk_wrap_probe.py ~/pptx2svg-oracle/cjk-wrap-box.pdf --check
"""

from __future__ import annotations

import ctypes
import re
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path[:0] = [str(ROOT / "src"), str(ROOT / "packages/pptx2svg-fonts/src"), str(ROOT / "tools")]

#: Two origins on the same line never differ by more than this many points vertically.
SAME_LINE_PT = 2.0


def pdf_lines(path: Path) -> list[list[str]]:
    """The text of each page's lines, in reading order, one list per page."""
    import pypdfium2 as pdfium
    import pypdfium2.raw as raw

    document = pdfium.PdfDocument(str(path))
    pages: list[list[str]] = []
    for page in document:
        text_page = page.get_textpage()
        rows: list[tuple[float, list[str]]] = []
        for index in range(raw.FPDFText_CountChars(text_page)):
            char = chr(raw.FPDFText_GetUnicode(text_page, index))
            x, y = ctypes.c_double(), ctypes.c_double()
            raw.FPDFText_GetCharOrigin(text_page, index, x, y)
            if not rows or abs(y.value - rows[-1][0]) > SAME_LINE_PT:
                rows.append((y.value, []))
            rows[-1][1].append(char)
        lines = ["".join(chars).strip() for _y, chars in rows]
        pages.append([line for line in lines if line])
    return pages


def our_lines(deck: Path) -> list[list[str]]:
    """The same reading off our own SVG: a new ``dy`` starts a line.

    Rendered with the bundled metric tables rather than the fidelity harness's local
    font profile, which for this deck is the same measurement: Noto Sans JP's table is
    extracted from the file PowerPoint is drawing with.
    """
    import pptx2svg

    pages: list[list[str]] = []
    for svg in pptx2svg.convert_pptx_to_svg(str(deck)):
        body = re.search(r"<text [^>]*>(.*?)</text>", svg, re.S)
        if body is None:
            pages.append([])
            continue
        lines: list[str] = []
        for attrs, text in re.findall(r"<tspan([^>]*)>([^<]*)</tspan>", body.group(1)):
            if 'dy="' in attrs or not lines:
                lines.append("")
            lines[-1] += text
        pages.append([line.strip() for line in lines if line.strip()])
    return pages


def pdf_baselines(path: Path) -> list[list[float]]:
    """Each page's text baselines, top to bottom, in points."""
    import pypdfium2 as pdfium
    import pypdfium2.raw as raw

    document = pdfium.PdfDocument(str(path))
    pages: list[list[float]] = []
    for page in document:
        text_page = page.get_textpage()
        rows: list[tuple[float, list[str]]] = []
        for index in range(raw.FPDFText_CountChars(text_page)):
            char = chr(raw.FPDFText_GetUnicode(text_page, index))
            x, y = ctypes.c_double(), ctypes.c_double()
            raw.FPDFText_GetCharOrigin(text_page, index, x, y)
            if not rows or abs(y.value - rows[-1][0]) > SAME_LINE_PT:
                rows.append((y.value, []))
            rows[-1][1].append(char)
        pages.append([y for y, chars in rows if "".join(chars).strip()])
    return pages


def our_baselines(deck: Path) -> list[list[float]]:
    """The same, recovered from the ``dy`` chain: ``<text y>`` plus a running sum."""
    import pptx2svg
    from pptx2svg.units import PX_PER_PT

    pages: list[list[float]] = []
    for svg in pptx2svg.convert_pptx_to_svg(str(deck)):
        body = re.search(r'<text [^>]*y="([-\d.]+)"[^>]*>(.*?)</text>', svg, re.S)
        if body is None:
            pages.append([])
            continue
        position = float(body.group(1))
        out = []
        for attrs, _text in re.findall(r"<tspan([^>]*)>([^<]*)</tspan>", body.group(2)):
            match = re.search(r'dy="([-\d.]+)"', attrs)
            if match is None:
                continue
            position += float(match.group(1))
            out.append(position / PX_PER_PT)
        pages.append(out)
    return pages


def report_baselines(pdf: Path, probes: list[dict], check: bool) -> int:
    theirs = pdf_baselines(pdf)
    ours = our_baselines(pdf.with_suffix(".pptx")) if check else None
    header = f"{'probe':18s} {'sizes':>12s} {'PPT advance':>12s}"
    if check:
        header += f" {'ours':>9s} {'d':>7s}"
    print(header)
    print("-" * len(header))
    for index, probe in enumerate(probes):
        sizes = [size for size, _percent, _text in probe.get("paragraphs", [])]
        rows = theirs[index] if index < len(theirs) else []
        gap = rows[0] - rows[1] if len(rows) > 1 else float("nan")
        row = (
            f"{probe['key']:18s} {'/'.join(f'{s:.0f}' for s in sizes):>12s} {gap:12.3f}"
        )
        if check:
            mine_rows = ours[index] if index < len(ours) else []
            mine = mine_rows[1] - mine_rows[0] if len(mine_rows) > 1 else float("nan")
            row += f" {mine:9.3f} {mine - gap:+7.3f}"
        print(row)
    return 0


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    pdf = Path(sys.argv[1]).expanduser()
    check = "--check" in sys.argv[2:]

    import make_cjk_wrap_probe as maker

    probes = maker.probes_for(pdf)
    # A spacing probe has no break to report, so it reads as baselines whether or not
    # the flag is given.
    if "--baselines" in sys.argv[2:] or any("paragraphs" in probe for probe in probes):
        return report_baselines(pdf, probes, check)
    pages = pdf_lines(pdf)
    ours = our_lines(pdf.with_suffix(".pptx")) if check else None

    header = f"{'probe':18s} {'box pt':>8s} {'size':>5s} {'k':>4s} {'PPT':>4s}"
    if check:
        header += f" {'ours':>5s} {'d':>3s}"
    print(header + "  line 1")
    print("-" * (len(header) + 30))

    disagreements = 0
    for index, probe in enumerate(probes):
        lines = pages[index] if index < len(pages) else []
        first = lines[0] if lines else ""
        size = probe["runs"][0][1]
        budget = probe["width_pt"] - probe.get("l_ins", 0.0) - probe.get("r_ins", 0.0)
        budget -= probe.get("mar_l_pt", 0.0)
        row = (
            f"{probe['key']:18s} {probe['width_pt']:8.2f} {size:5.1f} "
            f"{budget / size:4.2f} {len(first):4d}"
        )
        if check:
            mine = ours[index][0] if index < len(ours) and ours[index] else ""
            delta = len(mine) - len(first)
            disagreements += delta != 0
            row += f" {len(mine):5d} {delta:+3d}"
        print(row + f"  {first[:24]}")

    if check:
        print(f"\n{disagreements} of {len(probes)} first lines disagree")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
