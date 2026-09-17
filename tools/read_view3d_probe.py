#!/usr/bin/env python3
"""Read a 3-D probe deck's PDF export: the axis PowerPoint chose and where it drew it.

PowerPoint rasterises 3-D chart geometry (ROADMAP.md 3.4) but leaves every string vector,
so this reads nothing but text.  The value axis' tick labels give

* the **axis** -- its ends and its unit, hence the interval count that produced it, and
* the **plot rectangle** -- the extreme ticks sit on the plot's own top and bottom edges,
  and the category labels' centres bracket its left and right ones.

That is the whole of the depth reservation: the difference between those insets and the
ones a flat chart of the same data on the same frame gets.  ``--check`` renders the deck
through this library and prints that difference per probe.

Usage::

    python3 tools/read_view3d_probe.py ~/pptx2svg-oracle/view3d-meter.pdf [substring]
    python3 tools/read_view3d_probe.py ~/pptx2svg-oracle/view3d-meter.pdf --check
"""

from __future__ import annotations

import sys
from pathlib import Path

import pypdfium2 as pdfium

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))
sys.path.insert(0, str(HERE.parent / "src"))
sys.path.insert(0, str(HERE.parent / "packages/pptx2svg-fonts/src"))

from make_axis_probe import FRAME_OFF  # noqa: E402
from make_view3d_probe import probes_for  # noqa: E402
from read_axis_probe import _number, labels, unit_from  # noqa: E402

EMU = 12700.0


def read_page(page, probe: dict) -> dict:
    """One probe slide's axis and plot rectangle, in page points."""
    text = labels(page)
    ticks, categories = [], []
    for label, left, y, _height, width, x in text:
        try:
            ticks.append((_number(label), y, width))
        except ValueError:
            categories.append((label, x, y))
    ticks.sort(key=lambda row: row[1])
    categories.sort(key=lambda row: row[1])
    values = [row[0] for row in ticks]
    positions = [row[1] for row in ticks]
    xs = [row[1] for row in categories]
    band = (xs[1] - xs[0]) if len(xs) > 1 else 0.0
    return {
        "values": values,
        "unit": unit_from(values),
        "bottom": positions[0] if positions else 0.0,
        "top": positions[-1] if positions else 0.0,
        "height": round(positions[-1] - positions[0], 3) if len(positions) > 1 else 0.0,
        "left": round(xs[0] - band / 2, 3) if xs else 0.0,
        "right": round(xs[-1] + band / 2, 3) if xs else 0.0,
        "cat_y": round(categories[0][2], 3) if categories else 0.0,
    }


def insets(read: dict, probe: dict, page_height: float) -> dict:
    """The plot rectangle as four insets from the chart frame's own edges, in points."""
    frame_top = page_height - FRAME_OFF[1] / EMU
    frame_bottom = frame_top - probe["frame"][1] / EMU
    frame_left = FRAME_OFF[0] / EMU
    frame_right = frame_left + probe["frame"][0] / EMU
    return {
        "top": round(frame_top - read["top"], 3),
        "bottom": round(read["bottom"] - frame_bottom, 3),
        "left": round(read["left"] - frame_left, 3),
        "right": round(frame_right - read["right"], 3),
        "height": read["height"],
        "width": round(read["right"] - read["left"], 3),
    }


def ours(deck: Path, probes: list[dict]) -> list[dict]:
    """What this library draws for the same deck: the same readings, same units.

    The library's own model is read rather than its SVG, so the numbers are the layout's
    and not a rasteriser's: a chart's tick labels are ordinary text elements inside the
    chart frame, and their boxes' centres are what PowerPoint's rect centres are compared
    against.
    """
    from pptx2svg import ConvertOptions, convert_pptx_to_model
    from pptx2svg import model as m

    presentation = convert_pptx_to_model(deck, ConvertOptions())
    page_height = presentation.slide_size.height / EMU
    rows = []
    for probe, slide in zip(probes, presentation.slides):
        ticks, categories = [], []

        def walk(elements, ox=0.0, oy=0.0):
            for element in elements:
                if isinstance(element, m.ChartElement):
                    walk(
                        element.children,
                        ox + element.transform.offset_x / EMU,
                        oy + element.transform.offset_y / EMU,
                    )
                elif isinstance(element, m.GroupElement):
                    walk(element.children, ox, oy)
                elif isinstance(element, m.ShapeElement) and element.text_body is not None:
                    text = "".join(
                        run.text
                        for paragraph in element.text_body.paragraphs
                        for run in paragraph.runs
                    ).strip()
                    if not text:
                        continue
                    transform = element.transform
                    x = ox + transform.offset_x / EMU + transform.extent_width / EMU / 2
                    y = page_height - (
                        oy + transform.offset_y / EMU + transform.extent_height / EMU / 2
                    )
                    try:
                        ticks.append((_number(text), y))
                    except ValueError:
                        categories.append((text, x, y))

        walk(slide.elements)
        ticks.sort(key=lambda row: row[1])
        categories.sort(key=lambda row: row[1])
        xs = [row[1] for row in categories]
        band = (xs[1] - xs[0]) if len(xs) > 1 else 0.0
        positions = [row[1] for row in ticks]
        rows.append(
            {
                "values": [row[0] for row in ticks],
                "unit": unit_from([row[0] for row in ticks]),
                "bottom": positions[0] if positions else 0.0,
                "top": positions[-1] if positions else 0.0,
                "height": round(positions[-1] - positions[0], 3) if len(positions) > 1 else 0.0,
                "left": round(xs[0] - band / 2, 3) if xs else 0.0,
                "right": round(xs[-1] + band / 2, 3) if xs else 0.0,
            }
        )
    return rows


def solved_intervals(read: dict, high: float, strict: bool) -> set[int]:
    """Every interval count that would draw the axis this probe drew.

    The axis is a product of two rules -- how many intervals there is room for, and what
    range they divide -- and only the drawn unit and ends are observable.  So *strict*
    picks the range rule and this reports the counts consistent with it: an empty set says
    the drawn axis is not reachable at any count, which is the rule being refuted.
    """
    from pptx2svg.resolve import chart as chartmod

    values, unit = read["values"], read["unit"]
    if not values or not isinstance(unit, float):
        return set()
    out = set()
    for count in range(1, chartmod.AXIS_MAX_INTERVALS + 1):
        minimum, maximum, step = chartmod.nice_axis_scale(
            0.0, high, intervals=count, strict=strict
        )
        if (
            abs(step - unit) < unit * 1e-6
            and abs(minimum - values[0]) < unit * 1e-6
            and abs(maximum - values[-1]) < unit * 1e-6
        ):
            out.add(count)
    return out


def solve(path: Path, probes: list[dict]) -> None:
    """Per probe, the counts each range rule allows; per cell, what they agree on.

    A "cell" is the probes sharing a frame and a view -- the N-meter, whose datasets'
    drawn units name one count between them.  The rule that survives is the one whose
    cells intersect non-empty.
    """
    doc = pdfium.PdfDocument(path)
    cells: dict[tuple, dict[str, set[int]]] = {}
    for index, probe in enumerate(probes):
        read = read_page(doc[index].raw, probe)
        view = probe["view"]
        cell = (
            probe["kind"],
            probe["frame"][1],
            "absent" if view is None else tuple(sorted(view.items())),
        )
        strictly = solved_intervals(read, probe["high"], strict=True)
        barely = solved_intervals(read, probe["high"], strict=False)
        row = cells.setdefault(cell, {"padded": set(range(1, 11)), "bare": set(range(1, 11)),
                                      "height": read["height"], "keys": []})
        row["padded"] &= strictly
        row["bare"] &= barely
        row["keys"].append(probe["key"])
        print(
            f"{probe['key']:20s} {_axis(read):22s} h={read['height']:8.3f}  "
            f"padded={sorted(strictly)}  bare={sorted(barely)}"
        )
    print("\n# cell\tplot_h\tpadded_N\tbare_N\tprobes")
    for cell, row in cells.items():
        kind, frame, view = cell
        shown = view if isinstance(view, str) else ",".join(f"{k}={v}" for k, v in view)
        print(
            f"{kind}\t{frame / EMU:.0f}\t{shown}\t{row['height']:.3f}\t"
            f"{sorted(row['padded'])}\t{sorted(row['bare'])}\t{len(row['keys'])}"
        )


def _axis(read: dict) -> str:
    values, unit = read["values"], read["unit"]
    if not values:
        return "(no ticks)"
    return f"{values[0]:g}..{values[-1]:g} by {unit}"


def main() -> int:
    path = Path(sys.argv[1]).expanduser()
    rest = sys.argv[2:]
    check = "--check" in rest
    only = next((arg for arg in rest if not arg.startswith("--")), None)
    probes = probes_for(path)
    if "--solve" in rest:
        solve(path, probes)
        return 0
    doc = pdfium.PdfDocument(path)
    mine = ours(path.with_suffix(".pptx"), probes) if check else None

    print(
        "# key\tkind\tframe\tview\tdrawn\tplot_h\ttop\tbottom\tleft\tright"
        + ("\tours\tour_h\tour_top\tour_bottom\td_top\td_bottom\td_left" if check else "")
    )
    for index, probe in enumerate(probes):
        if only and only not in probe["key"]:
            continue
        page = doc[index]
        read = read_page(page.raw, probe)
        box = insets(read, probe, page.get_height())
        view = probe["view"]
        shown = (
            "absent"
            if view is None
            else ",".join(f"{k}={v}" for k, v in view.items()) or "empty"
        )
        row = [
            probe["key"],
            probe["kind"],
            f"{probe['frame'][1] / EMU:.0f}",
            shown,
            _axis(read),
            f"{box['height']:.3f}",
            f"{box['top']:.3f}",
            f"{box['bottom']:.3f}",
            f"{box['left']:.3f}",
            f"{box['right']:.3f}",
        ]
        if check:
            our = mine[index]
            our_box = insets(our, probe, page.get_height())
            row += [
                _axis(our),
                f"{our_box['height']:.3f}",
                f"{our_box['top']:.3f}",
                f"{our_box['bottom']:.3f}",
                f"{box['top'] - our_box['top']:+.3f}",
                f"{box['bottom'] - our_box['bottom']:+.3f}",
                f"{box['left'] - our_box['left']:+.3f}",
            ]
        print("\t".join(row))
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
