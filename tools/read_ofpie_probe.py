#!/usr/bin/env python3
"""Read an ``ofPieChart`` probe deck's PDF export back as packing geometry.

Every slide yields the same four numbers -- the first plot's radius and centre, and the
second plot's extent and centre -- and they come off the page as **exact vertices**, not
as a bounding box:

* the **bar** form's second plot is a stack of rectangles, so its width is ``s*r`` and its
  height ``2*s*r`` with no fitting at all, and one slide reads ``r`` twice over;
* a **slice**'s two straight edges run from the centre to the circle, so their far
  endpoints are the only path points lying *on* it.  An arc's bezier control points sit
  outside the circle -- by 9% at a quarter turn -- which is why the radius here is the
  **minimum** distance from the centre and never the bounding box.  The centre itself is
  the vertex every slice of one plot shares.

What the columns say:

``r``        the first plot's radius, from its slices.
``r2``       the second plot's, from the bar's width and height (they must agree) or from
             the second pie's own slices.
``s_read``   ``r2 / r``, which must come back as the stated ``secondPieSize``.
``divisor``  ``W / r`` for the region width ``W`` this library computes -- the quantity
             :data:`~pptx2svg.resolve.chart.OF_PIE_BAR_GAP_DIVISOR` is a term of.
``span``     the whole drawing's width, first plot's left edge to second plot's right.
``centre``   that span's midpoint, against the region's own midpoint.

Usage::

    python3 tools/read_ofpie_probe.py ~/pptx2svg-oracle/ofpie-pack.pdf [substring]
    python3 tools/read_ofpie_probe.py ~/pptx2svg-oracle/ofpie-pack.pdf --check [substring]

``--check`` renders the same deck through this library and prints the residual against
PowerPoint's radius and centres, which is the assertion SSIM cannot make.
"""

from __future__ import annotations

import ctypes
import math
import sys
from collections import Counter
from pathlib import Path

import pypdfium2 as pdfium
import pypdfium2.raw as raw

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from make_ofpie_probe import EMU, FRAME_OFF, probes_for  # noqa: E402
from read_axis_probe import chartmod, path_points, walk  # noqa: E402

#: The page is in points with y up; the frame's own top is measured down from the slide's.
PAGE_HEIGHT = 405.0
FRAME_LEFT = FRAME_OFF[0] / EMU
FRAME_TOP = FRAME_OFF[1] / EMU

#: A legend swatch is 5.49 pt square and a slice or a bar is tens of points across.
MIN_DRAWN_PT = 12.0


def filled_paths(page):
    """Every filled path's vertices, ignoring the page backdrop and the legend keys."""
    out = []
    for obj, kind, parent in walk(page):
        if kind != raw.FPDF_PAGEOBJ_PATH:
            continue
        fill = ctypes.c_int()
        stroke = ctypes.c_int()
        raw.FPDFPath_GetDrawMode(obj, ctypes.byref(fill), ctypes.byref(stroke))
        if not fill.value:
            continue
        points = path_points(obj, parent)
        if not points:
            continue
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        width, height = max(xs) - min(xs), max(ys) - min(ys)
        if width < MIN_DRAWN_PT or height < MIN_DRAWN_PT:
            continue
        if width > 700.0 and height > 390.0:  # the white page behind everything
            continue
        out.append(points)
    return out


def _rounded(point) -> tuple[float, float]:
    return (round(point[0], 2), round(point[1], 2))


def _is_rectangle(points) -> bool:
    corners = {_rounded(p) for p in points}
    if len(corners) != 4:
        return False
    xs = {c[0] for c in corners}
    ys = {c[1] for c in corners}
    return len(xs) == 2 and len(ys) == 2


def read_page(page) -> dict | None:
    """The two plots' geometry, however the second one is drawn."""
    paths = filled_paths(page)
    if not paths:
        return None
    rectangles = [p for p in paths if _is_rectangle(p)]
    wedges = [p for p in paths if not _is_rectangle(p)]

    # A pie's centre is the vertex its slices share, and with three slices or more it is
    # the **only** vertex three distinct paths have in common -- two adjacent slices share
    # their boundary point, but no third one does.  Counting per path rather than per
    # vertex matters: a closed subpath repeats its start, so a two-slice junction would
    # otherwise reach three occurrences on its own.
    shared = Counter(point for wedge in wedges for point in {_rounded(p) for p in wedge})
    centres = sorted(
        (point for point, count in shared.items() if count >= 3), key=lambda p: p[0]
    )
    if not centres:
        return None

    def radius_of(centre) -> float:
        mine = [w for w in wedges if any(_rounded(p) == centre for p in w)]
        distances = [
            math.hypot(p[0] - centre[0], p[1] - centre[1])
            for wedge in mine
            for p in wedge
            if _rounded(p) != centre
        ]
        return min(distances) if distances else 0.0

    first = centres[0]
    out = {
        "first_x": first[0],
        "first_y": first[1],
        "r": radius_of(first),
        "left": first[0] - radius_of(first),
    }
    if rectangles:
        xs = [p[0] for rect in rectangles for p in rect]
        ys = [p[1] for rect in rectangles for p in rect]
        out["r2"] = (max(xs) - min(xs))
        out["r2_tall"] = (max(ys) - min(ys)) / 2.0
        out["second_x"] = (min(xs) + max(xs)) / 2.0
        out["second_y"] = (min(ys) + max(ys)) / 2.0
        out["right"] = max(xs)
    elif len(centres) > 1:
        second = centres[-1]
        out["r2"] = radius_of(second)
        out["r2_tall"] = radius_of(second)
        out["second_x"] = second[0]
        out["second_y"] = second[1]
        out["right"] = second[0] + radius_of(second)
    else:
        return None
    return out


def ours(deck: Path) -> list[dict]:
    """The region and geometry this library computes for the same deck."""
    rows: list[dict] = []
    original = chartmod.ChartBuilder._of_pie_geometry

    def record(self, region):
        answer = original(self, region)
        radius, first_x, fraction, second_x, is_bar = answer
        rows.append(
            {
                "region_left": region.left + FRAME_LEFT,
                "region_right": region.right + FRAME_LEFT,
                "region_top": region.top,
                "region_bottom": region.bottom,
                "width": region.width,
                "height": region.height,
                "r": radius,
                "first_x": first_x + FRAME_LEFT,
                "fraction": fraction,
                "second_x": second_x + FRAME_LEFT,
                "is_bar": is_bar,
            }
        )
        return answer

    chartmod.ChartBuilder._of_pie_geometry = record
    try:
        import pptx2svg

        pptx2svg.convert_pptx_to_svg(deck.read_bytes())
    finally:
        chartmod.ChartBuilder._of_pie_geometry = original
    return rows


def report(path: Path, only: str | None) -> int:
    probes = probes_for(path)
    doc = pdfium.PdfDocument(path)
    mine_all = ours(path.with_suffix(".pptx"))
    print(
        f"{'key':16s} {'form':4s} {'g':>4s} {'s':>4s} {'W':>7s} {'H':>7s} {'r':>8s} "
        f"{'r2/r':>6s} {'W/r':>7s} {'span':>8s} {'dctr':>7s} {'r-ours':>8s}"
    )
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        read = read_page(doc[index].raw)
        mine = mine_all[index] if index < len(mine_all) else None
        if read is None or mine is None:
            print(f"{probe['key']:16s}  -- not read --")
            continue
        width, height = mine["width"], mine["height"]
        span = read["right"] - read["left"]
        centre = (read["right"] + read["left"]) / 2
        region_centre = (mine["region_left"] + mine["region_right"]) / 2
        print(
            f"{probe['key']:16s} {probe.get('form', 'bar'):4s} {probe.get('gap', 100):4d} "
            f"{probe.get('second', 75):4d} {width:7.3f} {height:7.3f} {read['r']:8.4f} "
            f"{read['r2'] / read['r']:6.4f} {width / read['r']:7.4f} {span:8.3f} "
            f"{centre - region_centre:7.3f} {mine['r'] - read['r']:8.3f}"
        )
    return 0


def check(path: Path, only: str | None) -> int:
    """Our radius and plot centres against PowerPoint's, slide by slide."""
    probes = probes_for(path)
    doc = pdfium.PdfDocument(path)
    mine_all = ours(path.with_suffix(".pptx"))
    worst = 0.0
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        read = read_page(doc[index].raw)
        mine = mine_all[index] if index < len(mine_all) else None
        if read is None or mine is None:
            print(f"{probe['key']:16s}  -- not read --")
            continue
        deltas = (
            mine["r"] - read["r"],
            mine["first_x"] - read["first_x"],
            mine["second_x"] - read["second_x"],
        )
        worst = max(worst, max(abs(d) for d in deltas))
        flag = "" if max(abs(d) for d in deltas) < 0.5 else "   <<<"
        print(
            f"{probe['key']:16s} dr {deltas[0]:+7.3f}  dx1 {deltas[1]:+7.3f}  "
            f"dx2 {deltas[2]:+7.3f}{flag}"
        )
    print(f"\nworst ofPie residual {worst:.3f} pt")
    return 0


def main() -> int:
    path = Path(sys.argv[1]).expanduser()
    arguments = sys.argv[2:]
    if "--check" in arguments:
        arguments.remove("--check")
        return check(path, arguments[0] if arguments else None)
    return report(path, arguments[0] if arguments else None)


if __name__ == "__main__":
    raise SystemExit(main())
