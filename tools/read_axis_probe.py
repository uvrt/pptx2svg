#!/usr/bin/env python3
"""Read the drawn tick unit back out of an axis probe deck's PDF export.

For each probe slide this reports what PowerPoint actually drew: the y positions of the
major gridlines (path objects that are horizontal strokes) and the axis tick labels with
their vertical centres, so the tick unit is *read* rather than inferred.  The labels are
the primary reading; the gridlines confirm the count and give the axis length in points,
which is what the coarsening stage turns out to depend on.

Two traps, both of which quietly produce a plausible wrong number:

* **The gridlines are one path, not one path each.**  A page object holding sixteen points
  is eight lines; requiring a two-point path finds none of them.  The page backdrop is a
  horizontal-edged rectangle too, and is excluded by where it starts rather than by size.
* **The labels are in the machine's locale.**  This one writes ``1,02`` for a tenth-scale
  tick and would write ``1.020`` for a thousand and twenty, so a naive ``float()`` reads
  0.98 as 98 -- which looks exactly like a chart with a hundredfold axis.

Usage -- the table is chosen by the PDF's file name, and the optional substring filters
to the probes whose key contains it::

    python3 tools/read_axis_probe.py ~/pptx2svg-oracle/axis-decade.pdf [substring]
"""

from __future__ import annotations

import ctypes
import math
import re
import sys
from pathlib import Path

import pypdfium2 as pdfium
import pypdfium2.raw as raw

sys.path.insert(0, str(Path(__file__).resolve().parent))
sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
from make_axis_probe import probes_for  # noqa: E402

from pptx2svg.resolve import chart as chartmod  # noqa: E402


def _matrix(obj):
    m = raw.FS_MATRIX()
    raw.FPDFPageObj_GetMatrix(obj, ctypes.byref(m))
    return (m.a, m.b, m.c, m.d, m.e, m.f)


def _mul(outer, inner):
    a1, b1, c1, d1, e1, f1 = inner
    a2, b2, c2, d2, e2, f2 = outer
    return (
        a1 * a2 + b1 * c2,
        a1 * b2 + b1 * d2,
        c1 * a2 + d1 * c2,
        c1 * b2 + d1 * d2,
        e1 * a2 + f1 * c2 + e2,
        e1 * b2 + f1 * d2 + f2,
    )


def _apply(m, x, y):
    a, b, c, d, e, f = m
    return (a * x + c * y + e, b * x + d * y + f)


def walk(page_or_form, parent=(1, 0, 0, 1, 0, 0), form=False):
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
            yield from walk(obj, _mul(parent, _matrix(obj)), form=True)
        else:
            yield obj, kind, parent


def path_points(obj, parent):
    m = _mul(parent, _matrix(obj))
    points = []
    for i in range(raw.FPDFPath_CountSegments(obj)):
        seg = raw.FPDFPath_GetPathSegment(obj, i)
        x, y = ctypes.c_float(), ctypes.c_float()
        raw.FPDFPathSegment_GetPoint(seg, ctypes.byref(x), ctypes.byref(y))
        points.append(_apply(m, x.value, y.value))
    return points


def horizontal_strokes(page):
    """Every horizontal path stroke, as ``(y, x0, x1)`` in page points."""
    out = []
    for obj, kind, parent in walk(page):
        if kind != raw.FPDF_PAGEOBJ_PATH:
            continue
        # Gridlines arrive as one path holding every line, so pair the points up rather
        # than requiring a two-point path.
        pts = path_points(obj, parent)
        for (x0, y0), (x1, y1) in zip(pts[::2], pts[1::2]):
            # The page backdrop is a horizontal-edged rectangle too; the gridlines start
            # where the axis does, never at the page edge.
            if abs(y0 - y1) < 0.05 and abs(x1 - x0) > 20.0 and min(x0, x1) > 5.0:
                out.append((round(y0, 3), round(min(x0, x1), 2), round(max(x0, x1), 2)))
    return sorted(out)


def labels(page):
    """Text runs as ``(text, x_left, y_centre)``, one per rectangle pdfium reports.

    Character boxes are not enough to reassemble a number: the decimal separator's box
    abuts its neighbours differently from the digits' and any gap threshold either splits
    ``1,02`` into three runs or joins two labels into one.  ``FPDFText_GetRect`` already
    groups a drawn run, which for an axis label is exactly the label.

    The text comes back exactly as PowerPoint drew it, separators and all; :func:`_number`
    is where the locale is undone.
    """
    textpage = raw.FPDFText_LoadPage(page)
    out = []
    for i in range(raw.FPDFText_CountRects(textpage, 0, -1)):
        vals = [ctypes.c_double() for _ in range(4)]
        raw.FPDFText_GetRect(textpage, i, *[ctypes.byref(v) for v in vals])
        left, top, right, bottom = (v.value for v in vals)
        buf = ctypes.create_string_buffer(512)
        count = raw.FPDFText_GetBoundedText(
            textpage,
            left,
            top,
            right,
            bottom,
            ctypes.cast(buf, ctypes.POINTER(ctypes.c_ushort)),
            255,
        )
        text = bytes(buf)[: count * 2].decode("utf-16-le", "replace").rstrip("\x00").strip()
        if text:
            out.append(
                (text, round(left, 2), round((top + bottom) / 2, 3), round(top - bottom, 2))
            )
    raw.FPDFText_ClosePage(textpage)
    return out


def _number(label: str) -> float:
    """A tick label as a number, in whichever separator convention drew it.

    PowerPoint writes the *locale's* separators, and this machine's are Dutch: ``0,098``
    is a decimal and ``1.020`` is a thousand and twenty.  A comma settles it outright; a
    lone point is read as a group separator only where it splits digits into a group of
    exactly three, which no tick label's fraction does.
    """
    if "," in label:
        return float(label.replace(".", "").replace(",", "."))
    if re.fullmatch(r"-?\d{1,3}(\.\d{3})+", label):
        return float(label.replace(".", ""))
    return float(label)


def prediction(low, high, slack):
    saved = chartmod._DECADE_SLACK
    chartmod._DECADE_SLACK = slack
    try:
        return chartmod.nice_axis_scale(low, high, horizontal=False, anchor_zero=False)
    finally:
        chartmod._DECADE_SLACK = saved


def unit_from(values):
    steps = [round(b - a, 12) for a, b in zip(values, values[1:])]
    return steps[0] if steps and len(set(steps)) == 1 else steps


def _ceil_125(value: float) -> float:
    """The smallest 1-2-5 step at or above *value*."""
    decade = float(f"1e{math.floor(math.log10(value))}")
    for step in (1, 2, 5, 10):
        if value <= step * decade * (1 + 1e-12):
            return step * decade
    raise AssertionError(value)


def measured_rule(low: float, high: float) -> tuple[float, float, float]:
    """PowerPoint's axis as these probes measured it: ``(minimum, maximum, unit)``.

    The **base** unit only.  On a short axis PowerPoint then coarsens up the 1-2-5 ladder
    by a rule that is still unknown, so a reading coarser than this is not a failure of
    this function -- see ROADMAP.md, "The unit rule is not the power of ten below the
    span".  Written here rather than in ``src`` because it is half a rule, and shipping
    half of it would draw eleven gridlines where PowerPoint draws four.
    """
    anchored = not (
        (low > 0 and high > 0 and low > chartmod.AXIS_ZERO_ANCHOR_RATIO * high)
        or (low < 0 and high < 0 and high < chartmod.AXIS_ZERO_ANCHOR_RATIO * low)
    )
    if anchored:
        low, high = min(0.0, low), max(0.0, high)
    pad = 0.05 * (high - low)
    padded_low = 0.0 if anchored and low >= 0 else low - pad
    padded_high = 0.0 if anchored and high <= 0 else high + pad
    unit = _ceil_125((padded_high - padded_low) / 10)
    return (
        math.floor(padded_low / unit + 1e-9) * unit,
        math.ceil(padded_high / unit - 1e-9) * unit,
        unit,
    )


def main() -> int:
    path = Path(sys.argv[1]).expanduser()
    only = sys.argv[2] if len(sys.argv) > 2 else None
    doc = pdfium.PdfDocument(path)
    for index, probe in enumerate(probes_for(path)):
        if only and only not in probe["key"]:
            continue
        pdf_page = doc[index]
        page = pdf_page.raw
        strokes = horizontal_strokes(page)
        text = labels(page)
        numbers = []
        for label, x0, y, _height in text:
            try:
                numbers.append((_number(label), y, x0))
            except ValueError:
                pass
        glyph = max((row[3] for row in text), default=0.0)
        numbers.sort(key=lambda n: n[1])
        values = [n[0] for n in numbers]
        with_slack = prediction(probe["low"], probe["high"], 1e-12)
        without = prediction(probe["low"], probe["high"], 0.0)
        span = probe["high"] - probe["low"]
        unit = unit_from(values)
        ys = [s[0] for s in strokes]
        height = round(ys[-1] - ys[0], 2) if len(ys) > 1 else 0.0
        step = round(height / (len(ys) - 1), 2) if len(ys) > 1 else 0.0
        print(f"=== {index + 1:2d} {probe['key']:12s} {probe['low']!r}..{probe['high']!r} "
              f"span={span!r} short={1 - span / probe['decade']:.2e} "
              f"frame={probe['frame'][1] / 12700:.0f}pt")
        print(f"    drawn    {values[0] if values else '?'}..{values[-1] if values else '?'} "
              f"by {unit}   ticks={len(values)} intervals={max(len(ys) - 1, 0)} "
              f"axis={height}pt step={step}pt glyph={glyph}pt")
        print(f"    labels   {values}")
        print(f"    ours     slack={with_slack}  noslack={without}")
        base = measured_rule(probe["low"], probe["high"])
        verdict = "base"
        if values and not (
            isinstance(unit, float)
            and abs(unit - base[2]) < base[2] * 1e-9
            and abs(values[0] - base[0]) < base[2] * 1e-6
            and abs(values[-1] - base[1]) < base[2] * 1e-6
        ):
            verdict = (
                f"coarsened x{unit / base[2]:g}"
                if isinstance(unit, float) and unit > base[2]
                else "*** NEITHER ***"
            )
        print(f"    measured {base}  {verdict}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
