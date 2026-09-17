#!/usr/bin/env python3
"""Read a bubble probe deck's PDF export back as an axis domain and a set of radii.

Each slide yields four things, all of them exact:

* the **drawn domain** of each axis, from its tick labels -- text, so there is no
  geometry in it at all;
* the **plot rectangle**, from the gridlines' own extent;
* every **bubble**'s centre and diameter, which come off a circle's bounding box because
  PowerPoint draws them axis-aligned and square (:data:`BUBBLE_REGION_INSET_PT`);
* what this library computes for the same slide, by instrumenting the resolver.

The quantity under test is the **headroom**: how much of the axis' span sits above the
largest value, expressed as the fraction of the plot the extreme bubble's radius takes.
If PowerPoint pads for the drawn ink then

    span >= max(value(i) / (1 - radius(i) / length))

over the points, where ``length`` is the plot's own extent along that axis -- and the
column ``need`` below is the right-hand side, to be read against the ``span`` PowerPoint
actually drew.  If it pads for the numbers instead, ``need`` moves with ``c:bubbleScale``
and the drawn span does not.

Usage::

    python3 tools/read_bubble_probe.py ~/pptx2svg-oracle/bubble-axis.pdf [substring]
    python3 tools/read_bubble_probe.py ~/pptx2svg-oracle/bubble-axis.pdf --check [substring]
"""

from __future__ import annotations

import ctypes
import sys
from pathlib import Path

import pypdfium2 as pdfium
import pypdfium2.raw as raw

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from make_bubble_probe import EMU, FRAME_OFF, probes_for  # noqa: E402
from read_axis_probe import (  # noqa: E402
    _number,
    chartmod,
    horizontal_strokes,
    labels,
    path_points,
    vertical_strokes,
    walk,
)

FRAME_LEFT = FRAME_OFF[0] / EMU


def circles(page) -> list[tuple[float, float, float]]:
    """Every drawn bubble as ``(centre x, centre y, diameter)``.

    A circle is filled, square to well inside a rounding, and bigger than the 5.49 pt
    legend swatch -- which is square too, and is what the size bound is here for.
    """
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
        if len(points) < 9:
            continue
        xs = [p[0] for p in points]
        ys = [p[1] for p in points]
        width, height = max(xs) - min(xs), max(ys) - min(ys)
        if width < 7.0 or abs(width - height) > 0.05:
            continue
        out.append(((min(xs) + max(xs)) / 2, (min(ys) + max(ys)) / 2, width))
    return sorted(out)


def plot_rect(page) -> tuple[float, float, float, float] | None:
    """``(left, bottom, right, top)`` in page points, from the gridlines and axis lines."""
    horizontal = horizontal_strokes(page)
    vertical = vertical_strokes(page)
    if not horizontal or not vertical:
        return None
    left = min(x for _, x, _ in horizontal)
    right = max(x for _, _, x in horizontal)
    bottom = min(y for _, y, _ in vertical)
    top = max(y for _, _, y in vertical)
    return left, bottom, right, top


def _cluster(rows, key, tolerance: float):
    """The largest group of *rows* sharing *key* to within *tolerance*."""
    best: list = []
    for row in rows:
        group = [other for other in rows if abs(key(other) - key(row)) <= tolerance]
        if len(group) > len(best):
            best = group
    return best


def domains(page, rect) -> tuple[list[float], list[float]]:
    """The two axes' tick values, read off their labels.

    **Not** by asking where the plot is: once a domain goes negative PowerPoint moves the
    value labels in to the *zero line* rather than leaving them outside the plot, and half
    this deck does exactly that.  What is stable either way is that the x axis' labels are
    the one row of numbers on the page, so it is found by clustering their **ink tops** --
    not their centres, which a comma in ``0,2`` drops by two thirds of a point -- and
    everything numeric that is not on that row belongs to the y axis.
    """
    numbered = []
    for text, x0, y_centre, height, width, x_centre in labels(page):
        top = y_centre + height / 2
        try:
            value = _number(text)
        except ValueError:
            continue
        numbered.append((value, x0 + width, y_centre, x_centre, top))
    # The x axis is the one row of numbers; everything numeric that is not on it belongs to
    # the y axis, which is a safer rule than clustering its right edges -- ``1`` carries a
    # side bearing an em wider than ``0``'s and falls out of any tolerance tight enough to
    # be worth having.
    across = _cluster(numbered, lambda row: row[4], 0.5)
    down = [row for row in numbered if row not in across]
    return (
        [row[0] for row in sorted(across, key=lambda row: row[3])],
        [row[0] for row in sorted(down, key=lambda row: row[2])],
    )


def ours(deck: Path) -> list[dict]:
    """The domain, plot rectangle and largest diameter this library computes.

    ``_draw_scatter_series`` is where the two scales and the plot rectangle are all in
    scope at once, so one hook answers everything; it runs once per series, and the first
    series of each chart is the one recorded.
    """
    rows: list[dict] = []
    original = chartmod.ChartBuilder._draw_scatter_series

    def record(self, rect, index, item, xs, x_scale, y_scale):
        if index == 0:
            region = self._bubble_region()
            rows.append(
                {
                    "x": x_scale,
                    "y": y_scale,
                    "left": rect.left + FRAME_LEFT,
                    "width": rect.width,
                    "height": rect.height,
                    "largest": self._largest_bubble() if self._is_bubble else 0.0,
                    "region_w": region.width,
                    "region_h": region.height,
                }
            )
        return original(self, rect, index, item, xs, x_scale, y_scale)

    chartmod.ChartBuilder._draw_scatter_series = record
    try:
        import pptx2svg

        pptx2svg.convert_pptx_to_svg(deck.read_bytes())
    finally:
        chartmod.ChartBuilder._draw_scatter_series = original
    return rows


def _headroom(points, length: float) -> float:
    """``max(value / (1 - radius / length))`` -- the span that clears every drawn bubble."""
    need = 0.0
    for value, radius in points:
        room = 1.0 - radius / length
        need = max(need, value / room if room > 0 else float("inf"))
    return need


def report(path: Path, only: str | None) -> int:
    probes = probes_for(path)
    doc = pdfium.PdfDocument(path)
    mine_all = ours(path.with_suffix(".pptx"))
    print(
        f"{'key':14s} {'scale':>5s} {'xmax':>7s} {'ymax':>7s} {'plotW':>7s} {'plotH':>7s} "
        f"{'M':>7s} {'x dom':>12s} {'y dom':>12s} {'xneed':>7s} {'yneed':>7s} "
        f"{'ours x':>10s} {'ours y':>10s}"
    )
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        page_handle = doc[index]
        page = page_handle.raw
        rect = plot_rect(page)
        mine = mine_all[index] if index < len(mine_all) else None
        if rect is None or mine is None:
            print(f"{probe['key']:14s}  -- not read --")
            continue
        left, bottom, right, top = rect
        x_ticks, y_ticks = domains(page, rect)
        if len(x_ticks) < 2 or len(y_ticks) < 2:
            print(f"{probe['key']:14s}  -- ticks not read -- {x_ticks} {y_ticks}")
            continue
        drawn = circles(page)
        biggest = max((d for _, _, d in drawn), default=0.0)
        width, height = right - left, top - bottom
        x_span = x_ticks[-1] - x_ticks[0]
        y_span = y_ticks[-1] - y_ticks[0]
        # Map each drawn bubble back onto its data value through the drawn axes.
        points_x = [
            ((cx - left) / width * x_span + x_ticks[0], d / 2)
            for cx, _, d in drawn
            if left - 1 <= cx <= right + 1
        ]
        points_y = [
            ((cy - bottom) / height * y_span + y_ticks[0], d / 2)
            for _, cy, d in drawn
            if bottom - 1 <= cy <= top + 1
        ]
        mine_x = mine.get("x")
        mine_y = mine.get("y")
        print(
            f"{probe['key']:14s} {probe.get('scale', 100):5d} "
            f"{max(v for v, _ in points_x):7.3f} {max(v for v, _ in points_y):7.3f} "
            f"{width:7.3f} {height:7.3f} {biggest:7.3f} "
            f"{x_ticks[0]:5.4g}..{x_ticks[-1]:<6.4g} {y_ticks[0]:5.4g}..{y_ticks[-1]:<6.4g} "
            f"{_headroom(points_x, width):7.3f} {_headroom(points_y, height):7.3f} "
            f"{(mine_x[0] if mine_x else 0):4.4g}..{(mine_x[1] if mine_x else 0):<5.4g} "
            f"{(mine_y[0] if mine_y else 0):4.4g}..{(mine_y[1] if mine_y else 0):<5.4g}"
        )
    return 0


def check(path: Path, only: str | None) -> int:
    """Our drawn domain against PowerPoint's, slide by slide."""
    probes = probes_for(path)
    doc = pdfium.PdfDocument(path)
    mine_all = ours(path.with_suffix(".pptx"))
    bad = 0
    total = 0
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        page_handle = doc[index]
        page = page_handle.raw
        rect = plot_rect(page)
        mine = mine_all[index] if index < len(mine_all) else None
        if rect is None or mine is None:
            print(f"{probe['key']:14s}  -- not read --")
            continue
        x_ticks, y_ticks = domains(page, rect)
        if len(x_ticks) < 2 or len(y_ticks) < 2:
            print(f"{probe['key']:14s}  -- ticks not read --")
            continue
        total += 1
        want = (x_ticks[0], x_ticks[-1], y_ticks[0], y_ticks[-1])
        got = (mine["x"][0], mine["x"][1], mine["y"][0], mine["y"][1])
        ok = all(abs(a - b) < 1e-6 * max(1.0, abs(a)) for a, b in zip(want, got))
        bad += not ok
        steps = (x_ticks[1] - x_ticks[0], y_ticks[1] - y_ticks[0])
        unit_ok = (
            abs(steps[0] - mine["x"][2]) < 1e-6 and abs(steps[1] - mine["y"][2]) < 1e-6
        )
        print(
            f"{probe['key']:14s} x {want[0]:g}..{want[1]:g} by {steps[0]:g}  "
            f"y {want[2]:g}..{want[3]:g} by {steps[1]:g}   "
            f"ours x {got[0]:g}..{got[1]:g} by {mine['x'][2]:g}  "
            f"y {got[2]:g}..{got[3]:g} by {mine['y'][2]:g}"
            f"{'' if ok and unit_ok else '   <<<'}"
        )
    print(f"\n{total - bad}/{total} domains agree")
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
