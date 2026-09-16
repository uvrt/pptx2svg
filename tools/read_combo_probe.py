#!/usr/bin/env python3
"""Read a combo probe deck's PDF export back as geometry.

The questions ROADMAP.md 3.3 asks are about *several* groups at once, so unlike
``read_axis_probe.py`` -- which reads one axis -- this reports the whole page: every page
object in **paint order** with what it is, and the text runs grouped into the four rows
and columns an axis probe cares about (the left ticks, the right ticks, the categories
along the bottom, and the legend).

Paint order is the point of the object listing.  A PDF content stream is drawn in order,
so "which group is on top" is read directly off the index at which each group's marks
appear -- no inference from what covers what.

Usage::

    python3 tools/read_combo_probe.py ~/pptx2svg-oracle/combo-order.pdf [substring]
    python3 tools/read_combo_probe.py ~/pptx2svg-oracle/combo-order.pdf --check [substring]

``--check`` renders the same deck through this library and prints, per slide, the
residual between the two plot rectangles and whether the two axes carry the same ticks.
That is the assertion the SSIM in ``tools/fidelity.py`` cannot make: a slide can score
0.95 with its plot 3 pt narrow, and 3 pt is the difference between a measured rule and a
plausible one.
"""

from __future__ import annotations

import ctypes
import sys
from pathlib import Path

import pypdfium2 as pdfium
import pypdfium2.raw as raw

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from make_combo_probe import FRAME_OFF, METER, probes_for, values_for  # noqa: E402
from read_axis_probe import _mul, _number, labels, path_points  # noqa: E402

from pptx2svg.resolve import chart as chartmod  # noqa: E402


def _matrix(obj):
    m = raw.FS_MATRIX()
    raw.FPDFPageObj_GetMatrix(obj, ctypes.byref(m))
    return (m.a, m.b, m.c, m.d, m.e, m.f)


def _draw_mode(obj) -> tuple[bool, bool]:
    fill = ctypes.c_int()
    stroke = ctypes.c_int()
    if not raw.FPDFPath_GetDrawMode(obj, ctypes.byref(fill), ctypes.byref(stroke)):
        return (False, True)
    return (bool(fill.value), bool(stroke.value))


def _colour(obj, stroke: bool) -> tuple[int, int, int]:
    parts = [ctypes.c_uint() for _ in range(4)]
    getter = raw.FPDFPageObj_GetStrokeColor if stroke else raw.FPDFPageObj_GetFillColor
    if not getter(obj, *[ctypes.byref(p) for p in parts]):
        return (-1, -1, -1)
    return tuple(p.value for p in parts[:3])


def walk_ordered(page_or_form, parent=(1, 0, 0, 1, 0, 0), form=False):
    """Every leaf object in paint order, with its accumulated matrix."""
    if form:
        count = raw.FPDFFormObj_CountObjects(page_or_form)

        def get(i):
            return raw.FPDFFormObj_GetObject(page_or_form, i)
    else:
        count = raw.FPDFPage_CountObjects(page_or_form)

        def get(i):
            return raw.FPDFPage_GetObject(page_or_form, i)

    for i in range(count):
        obj = get(i)
        kind = raw.FPDFPageObj_GetType(obj)
        if kind == raw.FPDF_PAGEOBJ_FORM:
            yield from walk_ordered(obj, _mul(parent, _matrix(obj)), form=True)
        else:
            yield obj, kind, parent


def classify(points, filled: bool, stroked: bool) -> str:
    """What a path is, from its own geometry.

    ``grid`` is a stroke every one of whose segments is horizontal -- gridlines arrive as
    one path holding all of them -- and ``axis`` is the vertical equivalent.  A stroke with
    any diagonal segment is a **series line**; a filled path is a ``bar`` if its points
    make an axis-aligned rectangle and an ``area`` otherwise.
    """
    if len(points) < 2:
        return "dot"
    pairs = list(zip(points[::2], points[1::2]))
    if stroked and not filled:
        horizontal = all(abs(a[1] - b[1]) < 0.05 for a, b in pairs)
        vertical = all(abs(a[0] - b[0]) < 0.05 for a, b in pairs)
        if horizontal and not vertical:
            return "grid"
        if vertical and not horizontal:
            return "axis"
        return "series-line"
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    width, height = max(xs) - min(xs), max(ys) - min(ys)
    if width < 9.0 and height < 9.0:
        return "marker"
    # **A whole bar group arrives as one filled path**, four points per bar, so the count
    # of axis-aligned quads is the count of bars -- and a path that is not made of quads
    # is the area polygon.  Classifying on the point count alone reads five bars as an
    # area, which is the one mistake that would answer the draw-order question backwards.
    if len(points) % 5 == 0 and all(
        len({round(p[0], 2) for p in quad}) == 2 and len({round(p[1], 2) for p in quad}) == 2
        for quad in (points[i:i + 5] for i in range(0, len(points), 5))
    ):
        return f"bars[{len(points) // 5}]"
    return "area"


def bbox(points) -> tuple[float, float, float, float]:
    xs = [p[0] for p in points]
    ys = [p[1] for p in points]
    return (round(min(xs), 2), round(min(ys), 2), round(max(xs), 2), round(max(ys), 2))


def page_objects(page) -> list[dict]:
    out = []
    for index, (obj, kind, parent) in enumerate(walk_ordered(page)):
        if kind != raw.FPDF_PAGEOBJ_PATH:
            continue
        points = path_points(obj, parent)
        if not points:
            continue
        filled, stroked = _draw_mode(obj)
        what = classify(points, filled, stroked)
        out.append(
            {
                "order": index,
                "what": what,
                "bbox": bbox(points),
                "points": len(points),
                "colour": _colour(obj, stroke=stroked and not filled),
                "pairs": [
                    (round(a[0], 2), round(a[1], 2), round(b[0], 2), round(b[1], 2))
                    for a, b in zip(points[::2], points[1::2])
                ],
            }
        )
    return out


def numeric_rows(page) -> dict:
    """The text runs, split into the left ticks, the right ticks and everything else."""
    text = labels(page)
    numbers, words = [], []
    for label, left, y, height, width, x in text:
        try:
            numbers.append({"value": _number(label), "text": label, "x": x, "y": y,
                            "left": left, "width": width})
        except ValueError:
            words.append({"text": label, "x": x, "y": y, "left": left, "width": width})
    if not numbers:
        return {"left": [], "right": [], "words": words}
    # Two label columns at most: the left axis' and the right axis'.  Split on the widest
    # gap in the x centres rather than on the frame's midpoint, so an axis whose labels
    # happen to sit near the middle is still one column.
    xs = sorted({round(n["x"], 1) for n in numbers})
    split = None
    if len(xs) > 1:
        gaps = [(b - a, (a + b) / 2) for a, b in zip(xs, xs[1:])]
        widest, middle = max(gaps)
        if widest > 30.0:
            split = middle
    left = [n for n in numbers if split is None or n["x"] < split]
    right = [] if split is None else [n for n in numbers if n["x"] >= split]
    left.sort(key=lambda n: n["y"])
    right.sort(key=lambda n: n["y"])
    return {"left": left, "right": right, "words": words}


def unit_of(values) -> float | list:
    steps = [round(b - a, 9) for a, b in zip(values, values[1:])]
    return steps[0] if steps and len(set(steps)) == 1 else steps


def meter_count(high: float, unit: float) -> int | None:
    """The interval count a drawn *unit* names, for one of the five meter datasets."""
    for intervals in range(1, 11):
        scale = chartmod.nice_axis_scale(0.0, high, intervals=intervals, anchor_zero=True)
        if abs(scale[2] - unit) < unit * 1e-9:
            return intervals
    return None


def describe(probe: dict, page) -> dict:
    rows = numeric_rows(page)
    objects = page_objects(page)
    # A ``majorTickMark="out"`` tick is a horizontal stroke too, and a short one outside
    # the plot; taking the plot's edges from every horizontal stroke reads it as part of
    # the plot and puts the left edge 3 pt too far out.
    grids = [o for o in objects
             if o["what"] == "grid" and o["bbox"][2] - o["bbox"][0] > 20.0]
    grid_xs = [(o["bbox"][0], o["bbox"][2]) for o in grids]
    plot_left = min((x0 for x0, _ in grid_xs), default=0.0)
    plot_right = max((x1 for _, x1 in grid_xs), default=0.0)
    grid_ys = sorted({round(p[1], 2) for o in grids for p in
                      [(o["pairs"][i][0], o["pairs"][i][1]) for i in range(len(o["pairs"]))]})
    return {
        "left": [n["value"] for n in rows["left"]],
        "left_y": [round(n["y"], 2) for n in rows["left"]],
        "right": [n["value"] for n in rows["right"]],
        "right_y": [round(n["y"], 2) for n in rows["right"]],
        "words": [(w["text"], round(w["left"], 1), round(w["y"], 1)) for w in rows["words"]],
        "plot": (round(plot_left, 2), round(plot_right, 2)),
        "grid_y": grid_ys,
        "objects": objects,
    }


def ours(deck: Path) -> list[dict]:
    """The same deck rendered by this library, as one dict of geometry per slide.

    Coordinates come back in the chart frame's own space, so the frame offset is added
    to put them beside the PDF's slide points.  The PDF's y grows upwards and the SVG's
    downwards, so only x is compared; the tick *values* carry the vertical answer.
    """
    import re

    import pptx2svg

    scale = 96.0 / 72.0
    offset = FRAME_OFF[0] / 12700
    out = []
    for svg in pptx2svg.convert_pptx_to_svg(deck.read_bytes()):
        if not isinstance(svg, str):
            svg = svg.decode()
        lines = []
        for match in re.finditer(
            r'translate\(([-\d.]+), ([-\d.]+)\)"><line [^>]*x1="([-\d.]+)" y1="([-\d.]+)" '
            r'x2="([-\d.]+)" y2="([-\d.]+)"',
            svg,
        ):
            tx, ty = float(match.group(1)) / scale, float(match.group(2)) / scale
            x1, y1, x2, y2 = (float(match.group(i)) / scale for i in (3, 4, 5, 6))
            lines.append((tx + x1 + offset, ty + y1, tx + x2 + offset, ty + y2))
        # Long horizontal strokes only: the legend's own key is a horizontal rule too, and
        # it sits outside the plot, so taking the extreme of every one of them reads the
        # legend as part of the plot rectangle.
        horizontal = [
            line
            for line in lines
            if abs(line[1] - line[3]) < 0.01 and abs(line[2] - line[0]) > 20.0
        ]
        texts = []
        for match in re.finditer(r"<tspan[^>]*>([^<]*)</tspan>", svg):
            texts.append(match.group(1))
        out.append(
            {
                "left": min((min(line[0], line[2]) for line in horizontal), default=0.0),
                "right": max((max(line[0], line[2]) for line in horizontal), default=0.0),
                "texts": texts,
            }
        )
    return out


def check(path: Path, only: str | None) -> int:
    """Compare every slide's plot rectangle and tick labels against the export."""
    probes = probes_for(path)
    doc = pdfium.PdfDocument(path)
    rendered = ours(path.with_suffix(".pptx"))
    worst = 0.0
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        pdf_page = doc[index]
        read = describe(probe, pdf_page.raw)
        mine = rendered[index]
        dl = mine["left"] - read["plot"][0]
        dr = mine["right"] - read["plot"][1]
        worst = max(worst, abs(dl), abs(dr))
        # The label as *this library* would print it, so a difference is a difference in
        # the axis and not in how two formatters spell a million.
        drawn = [
            chartmod.format_number(value, "General")
            for value in read["left"] + read["right"]
        ]
        def bare(text: str) -> str:
            # The export writes the machine's locale separators and this library writes
            # none, so both sides are compared on their digits alone.
            return text.replace(".", "").replace(",", "").replace("-", "")

        drawn_bare = [bare(text) for text in drawn]
        mine_bare = {bare(text) for text in mine["texts"]}
        missing = [
            text for text, key in zip(drawn, drawn_bare) if key not in mine_bare
        ]
        flag = "" if abs(dl) < 0.5 and abs(dr) < 0.5 and not missing else "   <<<"
        print(f"{probe['key']:18s} plot dx {dl:+7.2f} / {dr:+7.2f}"
              f"   ticks drawn {len(read['left'])}+{len(read['right'])}"
              f"   missing {missing}{flag}")
    print(f"\nworst plot edge residual {worst:.2f} pt")
    return 0


def main() -> int:
    path = Path(sys.argv[1]).expanduser()
    arguments = sys.argv[2:]
    if "--check" in arguments:
        arguments.remove("--check")
        return check(path, arguments[0] if arguments else None)
    only = arguments[0] if arguments else None
    doc = pdfium.PdfDocument(path)
    rows = []
    for index, probe in enumerate(probes_for(path)):
        if only and only not in probe["key"]:
            continue
        # Hold the page: pdfium frees it with the wrapper, and reading objects out of
        # a freed page segfaults rather than raising.
        pdf_page = doc[index]
        read = describe(probe, pdf_page.raw)
        width_pt = probe["frame"][0] / 12700
        height_pt = probe["frame"][1] / 12700
        print(f"=== {index + 1:3d} {probe['key']:18s} frame={width_pt:.0f}x{height_pt:.0f}pt")
        groups = " + ".join(
            f"{g['kind']}{'(2nd)' if g.get('secondary') else ''}"
            f"[{','.join(f'{h:g}' for h in g['highs'])}]"
            for g in probe["groups"]
        )
        print(f"    groups   {groups}")
        print(f"    left     {read['left']}  unit={unit_of(read['left'])}")
        if read["right"]:
            print(f"    right    {read['right']}  unit={unit_of(read['right'])}")
            print(f"    align    left_y={read['left_y']}")
            print(f"             right_y={read['right_y']}")
        print(f"    plot     x {read['plot'][0]} .. {read['plot'][1]} "
              f"({read['plot'][1] - read['plot'][0]:.2f} pt wide)")
        print(f"    words    {read['words']}")
        order = [(o["order"], o["what"], o["bbox"], o["colour"]) for o in read["objects"]
                 if o["what"].startswith("bars") or o["what"] in ("area", "series-line")]
        for entry in order:
            print(f"    paint    {entry[0]:4d} {entry[1]:10s} {entry[2]} {entry[3]}")
        rows.append((probe, read))

    # The N-meter table.  A drawn unit names an interval count for one of the five meter
    # datasets, so each row reads the count off *both* axes and puts the frame rule's
    # prediction beside them.
    print("\n# key\tframe_pt\tprimary_high\tleft_unit\tN_left\t"
          "second_high\tright_unit\tN_right\tN_rule")
    for probe, read in rows:
        height_pt = probe["frame"][1] / 12700
        box = chartmod.font_box("Aptos", probe.get("size", 1000) / 100)
        rule = chartmod.side_axis_intervals(height_pt, box.pitch)
        primary = next((g for g in probe["groups"] if not g.get("secondary")), None)
        second = next((g for g in probe["groups"] if g.get("secondary")), None)
        left_unit = unit_of(read["left"])
        right_unit = unit_of(read["right"])
        left_n = (meter_count(max(primary["highs"]), left_unit)
                  if primary and isinstance(left_unit, float) else None)
        right_n = (meter_count(max(second["highs"]), right_unit)
                   if second and isinstance(right_unit, float) else None)
        print(f"{probe['key']}\t{height_pt:.0f}\t"
              f"{max(primary['highs']) if primary else '-'}\t{left_unit}\t{left_n}\t"
              f"{max(second['highs']) if second else '-'}\t{right_unit}\t{right_n}\t{rule}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
