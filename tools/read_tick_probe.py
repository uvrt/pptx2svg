#!/usr/bin/env python3
"""Read the axis lines and tick marks back out of a tick probe deck's PDF export.

For each probe slide this lists every **stroked** segment on the page -- with no
gridlines and filled, unstroked series those are exactly the axis lines and their ticks
-- in frame coordinates (points from the chart frame's top-left corner), with the
stroke's colour, width, dash and cap.  A long segment is an axis; the short ones beside it
are its ticks, and their offsets either side of the axis are the reading.

``--compare`` scores this library's own strokes for the same deck against them instead.
``--summary`` folds each page into one line per axis: the axis line, then each distinct
tick shape as ``side-extents x count`` along with the positions the ticks stand at.

Needs PyMuPDF (dev only, like ``tools/fidelity.py``).

Usage::

    python3 tools/read_tick_probe.py ~/zz-tick-marks.pdf [--summary | --compare] [substring]
"""

from __future__ import annotations

import math
import sys
from collections import defaultdict
from pathlib import Path

import pymupdf as fitz

sys.path.insert(0, str(Path(__file__).resolve().parent))
from make_tick_probe import probes_for  # noqa: E402

#: Where ``make_axis_probe.slide_xml`` puts every frame, in points.
FRAME_OFF = (18.0, 12.0)


def segments(page) -> list[dict]:
    """Every stroked straight segment on *page*, in frame points."""
    out = []
    for path in page.get_drawings():
        if path.get("type") not in ("s", "fs") or path.get("color") is None:
            continue
        for item in path["items"]:
            if item[0] == "l":
                a, b = item[1], item[2]
            elif item[0] == "re":
                continue
            else:
                continue
            out.append(
                {
                    "x0": round(a.x - FRAME_OFF[0], 3),
                    "y0": round(a.y - FRAME_OFF[1], 3),
                    "x1": round(b.x - FRAME_OFF[0], 3),
                    "y1": round(b.y - FRAME_OFF[1], 3),
                    "color": "#%02X%02X%02X" % tuple(round(c * 255) for c in path["color"]),
                    "width": round(path.get("width") or 0.0, 3),
                    "dash": path.get("dashes"),
                    "cap": path.get("lineCap"),
                }
            )
    return out


def summary(segs: list[dict]) -> list[str]:
    """Group short segments against the long ones they stand on."""
    lines = []
    long = [s for s in segs if max(abs(s["x1"] - s["x0"]), abs(s["y1"] - s["y0"])) > 20]
    short = [s for s in segs if s not in long]
    for axis in long:
        vertical = abs(axis["x1"] - axis["x0"]) < 0.01
        at = axis["x0"] if vertical else axis["y0"]
        lines.append(
            f"  axis {'V' if vertical else 'H'} at {at:.3f} "
            f"span {min(axis['y0'], axis['y1']) if vertical else min(axis['x0'], axis['x1']):.3f}.."
            f"{max(axis['y0'], axis['y1']) if vertical else max(axis['x0'], axis['x1']):.3f} "
            f"{axis['color']} w{axis['width']} dash={axis['dash']} cap={axis['cap']}"
        )
        shapes = defaultdict(list)
        for tick in short:
            tick_vertical = abs(tick["x1"] - tick["x0"]) < 0.01
            if tick_vertical == vertical:
                continue
            if vertical:
                lo, hi = sorted((tick["x0"] - at, tick["x1"] - at))
                where = tick["y0"]
            else:
                lo, hi = sorted((tick["y0"] - at, tick["y1"] - at))
                where = tick["x0"]
            if lo > 0.01 or hi < -0.01:
                # Not touching this axis.
                if not (lo <= 0.01 and hi >= -0.01):
                    continue
            key = (round(lo, 3), round(hi, 3), tick["color"], tick["width"], str(tick["dash"]), tick["cap"])
            shapes[key].append(round(where, 3))
        for key, where in sorted(shapes.items()):
            lo, hi, color, width, dash, cap = key
            lines.append(
                f"    tick {lo:+.3f}..{hi:+.3f} x{len(where)} {color} w{width} dash={dash} "
                f"cap={cap} at {sorted(where)}"
            )
    return lines


def our_segments(deck: Path) -> list[list[dict]]:
    """The same list, per slide, off this library's own model for the same deck.

    Our axis lines and ticks are :class:`~pptx2svg.model.ConnectorElement` lines inside
    the chart frame, so they come back the way PowerPoint's strokes do.
    """
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "src"))
    from pptx2svg import ConvertOptions, convert_pptx_to_model
    from pptx2svg import model as m

    presentation = convert_pptx_to_model(deck, ConvertOptions())
    out = []
    for slide in presentation.slides:
        found: list[dict] = []

        def walk(elements, ox=0.0, oy=0.0):
            for element in elements:
                transform = getattr(element, "transform", None)
                if isinstance(element, m.ChartElement):
                    walk(
                        element.children,
                        ox + transform.offset_x / 12700,
                        oy + transform.offset_y / 12700,
                    )
                elif isinstance(element, m.ConnectorElement) and transform is not None:
                    x0 = ox + transform.offset_x / 12700
                    y0 = oy + transform.offset_y / 12700
                    x1 = x0 + transform.extent_width / 12700
                    y1 = y0 + transform.extent_height / 12700
                    if transform.flip_v:
                        y0, y1 = y1, y0
                    outline = element.outline
                    fill = getattr(outline, "fill", None)
                    color = getattr(getattr(fill, "color", None), "hex", None)
                    found.append(
                        {
                            "x0": round(x0 - FRAME_OFF[0], 3),
                            "y0": round(y0 - FRAME_OFF[1], 3),
                            "x1": round(x1 - FRAME_OFF[0], 3),
                            "y1": round(y1 - FRAME_OFF[1], 3),
                            "color": (color or "?").upper(),
                            "width": round((outline.width or 0) / 12700, 3) if outline else 0,
                            "dash": None,
                            "cap": None,
                        }
                    )

        walk(slide.elements)
        out.append(found)
    return out


def compare(theirs: list[dict], ours: list[dict]) -> str:
    """How far our short strokes are from PowerPoint's, matched nearest first."""

    def short(segs):
        return [s for s in segs if max(abs(s["x1"] - s["x0"]), abs(s["y1"] - s["y0"])) <= 20]

    a, b = short(theirs), short(ours)
    worst = 0.0
    for tick in a:
        best = min(
            (
                max(
                    abs(min(tick["x0"], tick["x1"]) - min(o["x0"], o["x1"])),
                    abs(max(tick["x0"], tick["x1"]) - max(o["x0"], o["x1"])),
                    abs(min(tick["y0"], tick["y1"]) - min(o["y0"], o["y1"])),
                    abs(max(tick["y0"], tick["y1"]) - max(o["y0"], o["y1"])),
                )
                for o in b
            ),
            default=math.inf,
        )
        worst = max(worst, best)
    return f"PowerPoint {len(a)} ticks, ours {len(b)}; worst nearest-match error {worst:.3f} pt"


def main() -> int:
    args = [a for a in sys.argv[1:] if not a.startswith("--")]
    path = Path(args[0]).expanduser()
    only = args[1] if len(args) > 1 else None
    probes = probes_for(path)
    document = fitz.open(path)
    ours = our_segments(path.with_suffix(".pptx")) if "--compare" in sys.argv else None
    for index, page in enumerate(document):
        key = probes[index]["key"] if index < len(probes) else f"page{index + 1}"
        if only and only not in key:
            continue
        segs = segments(page)
        if ours is not None:
            print(f"{index + 1:3d} {key:16s} {compare(segs, ours[index])}")
            continue
        print(f"{index + 1:3d} {key}: {len(segs)} stroked segments")
        if "--summary" in sys.argv:
            for line in summary(segs):
                print(line)
        else:
            for seg in segs:
                print(
                    f"    ({seg['x0']:8.3f},{seg['y0']:8.3f})-({seg['x1']:8.3f},{seg['y1']:8.3f}) "
                    f"{seg['color']} w{seg['width']} dash={seg['dash']} cap={seg['cap']}"
                )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
