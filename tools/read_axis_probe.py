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

``--insets`` reads something else off the same exports: the **plot rectangle**, as four
insets from the chart frame, ours beside PowerPoint's.  It takes them from the gridlines
rather than from the tick labels' centres, which is the difference between measuring the
rectangle and measuring a text rect's idea of where a label's middle is -- see
:func:`plot_insets` and ROADMAP.md 3.6.  Our side comes from the ``.pptx`` beside the
export, so the deck has to be there too.

Usage -- the table is chosen by the PDF's file name, and the optional substring filters
to the probes whose key contains it::

    python3 tools/read_axis_probe.py ~/pptx2svg-oracle/axis-decade.pdf [substring]
    python3 tools/read_axis_probe.py ~/pptx2svg-oracle/axis-inset.pdf --insets [substring]
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
sys.path.insert(
    0, str(Path(__file__).resolve().parent.parent / "packages/pptx2svg-fonts/src")
)
from make_axis_probe import FRAME_OFF, probes_for  # noqa: E402

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


def _is_stroked(obj) -> bool:
    """Whether this path is drawn as a stroke rather than only filled.

    A **bar** and an **area** are filled rectangles and polygons whose edges include
    horizontal segments, and pairing a filled path's points produces lines that look
    exactly like gridlines.  The scatter decks never hit this because a scatter draws
    nothing but markers and a line; the type sweep does, on every slide.
    """
    fill = ctypes.c_int()
    stroke = ctypes.c_int()
    if not raw.FPDFPath_GetDrawMode(obj, ctypes.byref(fill), ctypes.byref(stroke)):
        return True
    return bool(stroke.value)


def horizontal_strokes(page):
    """Every horizontal path stroke, as ``(y, x0, x1)`` in page points."""
    out = []
    for obj, kind, parent in walk(page):
        if kind != raw.FPDF_PAGEOBJ_PATH or not _is_stroked(obj):
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


def vertical_strokes(page):
    """Every vertical path stroke, as ``(x, y0, y1)`` -- the gridlines of a bottom axis."""
    out = []
    for obj, kind, parent in walk(page):
        if kind != raw.FPDF_PAGEOBJ_PATH or not _is_stroked(obj):
            continue
        pts = path_points(obj, parent)
        for (x0, y0), (x1, y1) in zip(pts[::2], pts[1::2]):
            if abs(x0 - x1) < 0.05 and abs(y1 - y0) > 20.0 and min(y0, y1) > 5.0:
                out.append((round(x0, 3), round(min(y0, y1), 2), round(max(y0, y1), 2)))
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
                (
                    text,
                    round(left, 2),
                    round((top + bottom) / 2, 3),
                    round(top - bottom, 2),
                    round(right - left, 2),
                    round((left + right) / 2, 3),
                )
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


def unit_from(values):
    """The step between consecutive labels, or the list of steps when they differ."""
    steps = [round(b - a, 12) for a, b in zip(values, values[1:])]
    return steps[0] if steps and len(set(steps)) == 1 else steps


def prediction(probe: dict, axis_pt: float = 0.0) -> tuple[float, float, float]:
    """The axis the shipped rule draws for this probe, frame and all.

    ``axis_pt`` is the length PowerPoint drew, and it is used for one thing only: a
    **radar**'s interval count is a function of its radius, which is a layout output
    rather than a frame dimension, so the reader feeds back the radius it read instead of
    reimplementing ``_radar_geometry`` here.

    The probes carry no title and no legend except the ``g`` family, whose key says which
    it has, so the band the rule divides is the frame itself less that furniture.  The face
    is the deck's theme minor font, Aptos, because none of the probes names another.
    """
    size = probe.get("size", 1000) / 100
    width_pt, height_pt = (dimension / 12700 for dimension in probe["frame"])
    face = probe.get("face", "Aptos")
    box = chartmod.font_box(face, size)
    radar = probe.get("kind") == "radar"
    horizontal = probe.get("kind") == "bar" or probe.get("read") == "x"
    if horizontal:
        provisional = chartmod.nice_axis_scale(
            probe["low"], probe["high"], anchor_zero=_anchor_zero(probe)
        )
        widest = max(
            (
                chartmod.text_width(
                    chartmod.format_number(value, "General"), face, size
                )
                for value in _tick_values(provisional)
            ),
            default=0.0,
        )
        intervals = chartmod.bottom_axis_intervals(width_pt, size, widest)
    elif radar:
        intervals = chartmod.radial_axis_intervals(axis_pt, box.line_height)
    else:
        if probe.get("legend") in ("b", "t", "tr"):
            height_pt -= chartmod.legend_row_pitch(box) + chartmod.LEGEND_BAND_PAD_PT
        if probe.get("title"):
            # An unstyled chart title is Arial 18 pt, which is the fallback any unstyled
            # text box gets and what `_title_size` returns when nothing names a size.
            title = chartmod.font_box("Arial", 18.0)
            height_pt -= chartmod.TITLE_BAND_LINES * title.line_height
        intervals = chartmod.side_axis_intervals(height_pt, box.line_height)
    return chartmod.nice_axis_scale(
        probe["low"],
        probe["high"],
        intervals=intervals,
        # A radar pads nothing and stops at the data, which is what `strict=False` is.
        strict=not radar,
        anchor_zero=_anchor_zero(probe),
    )


def _anchor_zero(probe: dict) -> bool:
    """Whether this probe's axis is held at zero, which every type but a scatter is.

    The category-chart probes carry a ``kind``; the scatter ones do not.
    """
    return bool(probe.get("kind"))


def _tick_values(scale) -> list[float]:
    minimum, maximum, unit = scale
    values, value = [], minimum
    while value <= maximum + unit * 1e-9 and len(values) < 200:
        values.append(value)
        value += unit
    return values


def read_page(page, probe: dict) -> dict:
    """Everything one probe slide says: the drawn axis, its length and its labels.

    The axis length is taken from the **extreme tick labels' own centres**, not from the
    gridlines: a bar or an area fills paths whose edges survive every stroke filter on
    some exports, and the first and last tick sit exactly at the ends of the plot, so the
    labels measure the same length more robustly.  The gridline count is still read and
    reported, as the cross-check that the labels were all found.
    """
    horizontal = probe.get("kind") == "bar" or probe.get("read") == "x"
    text = labels(page)
    if probe.get("read") == "x":
        # Both axes carry numbers on this probe, so the bottom row -- the labels sharing
        # the lowest centre -- is the x axis and everything above it is the y axis.
        floor_y = min((row[2] for row in text), default=0.0)
        text = [row for row in text if abs(row[2] - floor_y) < 3.0]
    numbers = []
    for label, left, y, _height, width, x in text:
        try:
            numbers.append((_number(label), x if horizontal else y, width))
        except ValueError:
            pass
    # A bottom axis runs left to right and a side axis bottom to top, and PDF y grows
    # upwards, so sorting on the position ascending puts both in increasing value order.
    numbers.sort(key=lambda n: n[1])
    values = [n[0] for n in numbers]
    positions = [n[1] for n in numbers]
    strokes = vertical_strokes(page) if horizontal else horizontal_strokes(page)
    axis = round(abs(positions[-1] - positions[0]), 3) if len(positions) > 1 else 0.0
    return {
        "values": values,
        "axis": axis,
        "intervals": max(len(values) - 1, 0),
        "step": round(axis / max(len(values) - 1, 1), 3),
        "glyph": max((row[3] for row in text), default=0.0),
        "widest": max((n[2] for n in numbers), default=0.0),
        "gridlines": len(strokes),
        "unit": unit_from(values),
    }


#: Where the chart frame sits on the page, in points, as ``tools/make_axis_probe.py``
#: places every probe: one frame per slide at :data:`make_axis_probe.FRAME_OFF`.
def _frame(probe: dict, page_height: float) -> tuple[float, float, float, float]:
    """The probe's chart frame as ``(left, top, right, bottom)`` in PDF points."""
    left = FRAME_OFF[0] / 12700
    top = page_height - FRAME_OFF[1] / 12700
    return (left, top, left + probe["frame"][0] / 12700, top - probe["frame"][1] / 12700)


def plot_insets(page, probe: dict, page_height: float) -> dict | None:
    """The drawn plot rectangle as four insets from the frame, read off the gridlines.

    The **gridlines are the rectangle**: the topmost major gridline is the plot's top edge
    and the category axis line its bottom, and both run the plot's full width.  That is
    what separates the top inset from the bottom band, which no reading of the tick
    labels' own centres can do -- a label centre carries whatever a PDF text rect's centre
    is against the tick it marks, and the same unknown then enters both ends with opposite
    signs.
    """
    strokes = [row for row in horizontal_strokes(page) if row[2] - row[1] > 100.0]
    if len(strokes) < 2:
        return None
    left_edge, top_edge, right_edge, bottom_edge = _frame(probe, page_height)
    # The widest stroke is a full-width gridline; a tick mark is not.
    width = max(row[2] - row[1] for row in strokes)
    lines = [row for row in strokes if row[2] - row[1] > width - 0.5]
    return {
        "top": round(top_edge - max(row[0] for row in lines), 3),
        "bottom": round(min(row[0] for row in lines) - bottom_edge, 3),
        "left": round(min(row[1] for row in lines) - left_edge, 3),
        "right": round(right_edge - max(row[2] for row in lines), 3),
        "count": len(lines),
    }


def our_insets(deck: Path, probes: list[dict]) -> list[dict | None]:
    """The same four insets, off this library's own model, for the same deck.

    Our gridlines are :class:`~pptx2svg.model.ConnectorElement` lines inside the chart
    frame, so the rectangle comes back the same way PowerPoint's does rather than through
    a second reimplementation of the layout.
    """
    from pptx2svg import ConvertOptions, convert_pptx_to_model
    from pptx2svg import model as m

    presentation = convert_pptx_to_model(deck, ConvertOptions())
    page_height = presentation.slide_size.height / 12700
    out: list[dict | None] = []
    for probe, slide in zip(probes, presentation.slides):
        lines: list[tuple[float, float, float]] = []

        def walk(elements, ox=0.0, oy=0.0):
            for element in elements:
                transform = getattr(element, "transform", None)
                if isinstance(element, m.ChartElement):
                    walk(
                        element.children,
                        ox + transform.offset_x / 12700,
                        oy + transform.offset_y / 12700,
                    )
                elif isinstance(element, m.GroupElement):
                    walk(element.children, ox, oy)
                elif isinstance(element, m.ConnectorElement) and transform is not None:
                    if transform.extent_height < 0.5 and transform.extent_width > 100.0:
                        x = ox + transform.offset_x / 12700
                        lines.append(
                            (
                                page_height - (oy + transform.offset_y / 12700),
                                x,
                                x + transform.extent_width / 12700,
                            )
                        )

        walk(slide.elements)
        if len(lines) < 2:
            out.append(None)
            continue
        left_edge, top_edge, right_edge, bottom_edge = _frame(probe, page_height)
        width = max(row[2] - row[1] for row in lines)
        lines = [row for row in lines if row[2] - row[1] > width - 0.5]
        out.append(
            {
                "top": round(top_edge - max(row[0] for row in lines), 3),
                "bottom": round(min(row[0] for row in lines) - bottom_edge, 3),
                "left": round(min(row[1] for row in lines) - left_edge, 3),
                "right": round(right_edge - max(row[2] for row in lines), 3),
                "count": len(lines),
            }
        )
    return out


def insets_main(path: Path, only: str | None) -> int:
    """``--insets``: the plot rectangle PowerPoint drew, against the one we draw.

    The deck is the ``.pptx`` beside the export, which is where our side comes from.
    """
    probes = probes_for(path)
    doc = pdfium.PdfDocument(path)
    mine = our_insets(path.with_suffix(".pptx"), probes)
    print(
        "# key\tsize\tface\tframe\tpp_top\tpp_bot\tpp_h\tour_top\tour_bot\tour_h\t"
        "d_top\td_bot\tpp_left\tour_left\td_left\tpp_right\tour_right"
    )
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        page = doc[index]
        theirs = plot_insets(page.raw, probe, page.get_height())
        ours = mine[index]
        if theirs is None or ours is None:
            print(f"{probe['key']}\t(no gridlines)")
            continue
        height = probe["frame"][1] / 12700
        print(
            f"{probe['key']}\t{probe.get('size', 1000) / 100:g}\t"
            f"{probe.get('face', 'Aptos')}\t{height:.0f}\t"
            f"{theirs['top']:.3f}\t{theirs['bottom']:.3f}\t"
            f"{height - theirs['top'] - theirs['bottom']:.3f}\t"
            f"{ours['top']:.3f}\t{ours['bottom']:.3f}\t"
            f"{height - ours['top'] - ours['bottom']:.3f}\t"
            f"{ours['top'] - theirs['top']:+.3f}\t{ours['bottom'] - theirs['bottom']:+.3f}\t"
            f"{theirs['left']:.3f}\t{ours['left']:.3f}\t{ours['left'] - theirs['left']:+.3f}\t"
            f"{theirs['right']:.3f}\t{ours['right']:.3f}"
        )
    return 0


def main() -> int:
    path = Path(sys.argv[1]).expanduser()
    args = sys.argv[2:]
    if "--insets" in args:
        rest = [arg for arg in args if arg != "--insets"]
        return insets_main(path, rest[0] if rest else None)
    only = args[0] if args else None
    doc = pdfium.PdfDocument(path)
    rows = []
    for index, probe in enumerate(probes_for(path)):
        if only and only not in probe["key"]:
            continue
        # Hold the page: pdfium frees it with the wrapper, and reading objects out of a
        # freed page segfaults rather than raising.
        pdf_page = doc[index]
        read = read_page(pdf_page.raw, probe)
        values, unit = read["values"], read["unit"]
        ours = prediction(probe, read['axis'])
        span = probe["high"] - probe["low"]
        print(f"=== {index + 1:2d} {probe['key']:16s} {probe['low']!r}..{probe['high']!r} "
              f"span={span!r} short={1 - span / probe['decade']:.2e} "
              f"frame={probe['frame'][1] / 12700:.0f}pt size={probe.get('size', 1000) / 100:g}pt")
        print(f"    drawn    {values[0] if values else '?'}..{values[-1] if values else '?'} "
              f"by {unit}   ticks={len(values)} gridlines={read['gridlines']} "
              f"axis={read['axis']}pt step={read['step']}pt glyph={read['glyph']}pt "
              f"widest={read['widest']}pt")
        print(f"    labels   {values}")
        verdict = "MATCH" if _agrees(values, unit, ours) else "*** DIFFERS ***"
        print(f"    ours     {ours}  {verdict}")
        rows.append((probe, read, ours, verdict))
    print("\n# key size frame_pt axis_pt intervals step_pt drawn_unit ours_unit verdict")
    for probe, read, ours, verdict in rows:
        print(f"{probe['key']}\t{probe.get('size', 1000) / 100:g}\t"
              f"{probe['frame'][1] / 12700:.1f}x{probe['frame'][0] / 12700:.1f}\t"
              f"{read['axis']}\t{read['intervals']}\t{read['step']}\t{read['unit']}\t"
              f"{ours[2]:g}\t{read['widest']}\t{verdict}")
    agreed = sum(1 for row in rows if row[3] == "MATCH")
    print(f"\n{agreed} of {len(rows)} probes reproduced")
    return 0


def _agrees(values, unit, ours) -> bool:
    """Whether the drawn axis and ours are the same axis."""
    if not values or not isinstance(unit, float):
        return False
    minimum, maximum, our_unit = ours
    return (
        abs(unit - our_unit) < our_unit * 1e-6
        and abs(values[0] - minimum) < our_unit * 1e-6
        and abs(values[-1] - maximum) < our_unit * 1e-6
    )


if __name__ == "__main__":
    raise SystemExit(main())
