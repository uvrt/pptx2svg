#!/usr/bin/env python3
"""Read a ``make_style_probe.py`` deck's PDF export, and say where pptx2svg disagrees.

For every probe shape it reads, from PowerPoint's export:

* **the fill** -- the colour at three points down the shape's left side, away from the
  text and inside any outline, so a gradient shows as three different colours;
* **the outline** -- a scan across the shape's left edge at mid-height: the colour of its
  darkest pixel and how many points of it are not white;
* **a shadow** -- whether the page under the shape's bottom edge is darker than white;
* **the text** -- the colour and the font of the span that says "Styled".

and the same four things from pptx2svg's model of the same deck, with the model's fill
composited over the white page the way the export's is.  A row is marked ``ok`` when
every channel is within ``TOLERANCE`` levels, every width within half a point, and the
shadows, colours and faces agree.

Usage::

    python3 tools/read_style_probe.py deck.pptx deck.pdf

Needs PyMuPDF to read the PDF; pptx2svg itself is imported from the checkout.
"""

from __future__ import annotations

import math
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "src"))
sys.path.insert(0, str(HERE))

import pymupdf  # noqa: E402

from make_style_probe import BOX_H, BOX_W, PROBES, place  # noqa: E402
from pptx2svg import convert_pptx_to_model  # noqa: E402
from pptx2svg import model as m  # noqa: E402

#: Pixels per point the export is read at.
SCALE = 8
#: Levels a channel may differ by and still agree: PDF colour, antialiasing and a
#: shading's interpolation all cost a level or two.
TOLERANCE = 4
#: Where down the shape the fill is sampled, as fractions of its height.
FILL_SAMPLES = (0.12, 0.5, 0.88)
#: How far in from the left edge, pt: past a 6 pt outline's inner half.
FILL_INSET = 8


def _hex(rgb) -> str:
    return "#" + "".join(f"{int(round(c)):02x}" for c in rgb[:3])


def _rgb(hex_value: str) -> tuple[int, int, int]:
    value = hex_value.lstrip("#")
    return tuple(int(value[i : i + 2], 16) for i in (0, 2, 4))


def _over_white(color: m.ResolvedColor) -> tuple[float, float, float]:
    return tuple(c * color.alpha + 255 * (1 - color.alpha) for c in _rgb(color.hex))


def _close(a, b) -> bool:
    return all(abs(x - y) <= TOLERANCE for x, y in zip(a, b))


# ----------------------------------------------------------------------------- export


def read_export(pdf: Path) -> dict[str, dict]:
    document = pymupdf.open(pdf)
    pages = {}
    found = {}
    for index, spec in enumerate(PROBES):
        slide, x, y = place(index)
        if slide not in pages:
            page = document[slide]
            pix = page.get_pixmap(matrix=pymupdf.Matrix(SCALE, SCALE), alpha=False)
            spans = [
                span
                for block in page.get_text("dict")["blocks"]
                for line in block.get("lines", [])
                for span in line["spans"]
            ]
            pages[slide] = (pix, spans)
        pix, spans = pages[slide]

        def pixel(px: float, py: float):
            return pix.pixel(int(px * SCALE), int(py * SCALE))

        fill = [pixel(x + FILL_INSET, y + BOX_H * f) for f in FILL_SAMPLES]

        # Ink that is neither the page, nor the fill beside it, nor the light grey a
        # shadow's blur spreads past the edge, is the outline.
        inside = pixel(x + FILL_INSET, y + BOX_H / 2)
        scan = [pixel(x + dx / SCALE, y + BOX_H / 2) for dx in range(-6 * SCALE, 6 * SCALE)]
        inked = [
            p for p in scan
            if min(p) < 235
            and not all(abs(a - b) <= 2 * TOLERANCE for a, b in zip(p, inside))
            and not (max(p) - min(p) < 8 and min(p) > 200)
        ]
        line = None
        if inked:
            darkest = min(inked, key=sum)
            line = (_hex(darkest), len(inked) / SCALE)

        # Below the outline's outer half, where a shadow offset downwards shows.
        clear = 2 + (line[1] / 2 if line else 0)
        below = pixel(x + BOX_W / 2, y + BOX_H + clear)
        shadow = min(below) < 245

        text = next(
            (
                span
                for span in spans
                if span["text"].strip() == "Styled"
                and x <= (span["bbox"][0] + span["bbox"][2]) / 2 <= x + BOX_W
                and y <= (span["bbox"][1] + span["bbox"][3]) / 2 <= y + BOX_H
            ),
            None,
        )
        found[spec["key"]] = {
            "fill": [_hex(p) for p in fill],
            "fill_rgb": fill,
            "line": line,
            "shadow": shadow,
            "text": (f"#{text['color']:06x}", text["font"]) if text else None,
        }
    return found


# ------------------------------------------------------------------------------ model


def _gradient_at(fill: m.GradientFill, fx: float, fy: float) -> tuple[float, float, float]:
    """The model gradient's colour at a point given as fractions of the box."""
    angle = math.radians(fill.angle)
    dx, dy = math.cos(angle), math.sin(angle)
    span = abs(BOX_W * dx) + abs(BOX_H * dy)
    t = ((fx - 0.5) * BOX_W * dx + (fy - 0.5) * BOX_H * dy) / span + 0.5
    stops = sorted(fill.stops, key=lambda s: s.position)
    position = t
    if position <= stops[0].position:
        return _over_white(stops[0].color)
    for a, b in zip(stops, stops[1:]):
        if position <= b.position:
            k = (position - a.position) / max(1e-9, b.position - a.position)
            ca, cb = _over_white(a.color), _over_white(b.color)
            return tuple(p + (q - p) * k for p, q in zip(ca, cb))
    return _over_white(stops[-1].color)


def read_model(pptx: Path) -> dict[str, dict]:
    resolved = convert_pptx_to_model(str(pptx))
    shapes = {}
    for slide in resolved.slides:
        for element in slide.elements:
            if isinstance(element, m.ShapeElement) and element.alt_text:
                shapes[element.alt_text] = element
    found = {}
    for spec in PROBES:
        shape = shapes[spec["key"]]
        fx = FILL_INSET / BOX_W
        if isinstance(shape.fill, m.SolidFill):
            fill = [_over_white(shape.fill.color)] * len(FILL_SAMPLES)
        elif isinstance(shape.fill, m.GradientFill) and shape.fill.gradient_type == "linear":
            fill = [_gradient_at(shape.fill, fx, f) for f in FILL_SAMPLES]
        elif isinstance(shape.fill, m.GradientFill):
            # A path gradient's colour at a point depends on its fill-to rectangle, which
            # this does not model: the row says what was drawn and is not held to it.
            fill = None
        else:
            fill = [(255, 255, 255)] * len(FILL_SAMPLES)
        line = None
        outline = shape.outline
        if outline is not None and isinstance(outline.fill, m.SolidFill):
            line = (outline.fill.color.hex, outline.width / 12700)
        effects = shape.effects
        shadow = effects is not None and effects.outer_shadow is not None
        run = shape.text_body.paragraphs[0].runs[0].properties if shape.text_body else None
        text = (run.color.hex if run and run.color else "#000000", run.font_family if run else None)
        found[spec["key"]] = {
            "fill": [_hex(p) for p in fill] if fill else ["(path gradient)"],
            "fill_rgb": fill,
            "line": line,
            "shadow": shadow,
            "text": text,
        }
    return found


def _face(name: str | None) -> str:
    """A face as the PDF names it and as the model does, alike: ``Calibri-Light``."""
    return (name or "").split("+")[-1].replace(" ", "").replace("-", "").lower()


def compare(power: dict, ours: dict) -> list[str]:
    wrong = []
    if ours["fill_rgb"] is not None and not all(
        _close(a, b) for a, b in zip(power["fill_rgb"], ours["fill_rgb"])
    ):
        wrong.append("fill")
    if (power["line"] is None) != (ours["line"] is None):
        wrong.append("line")
    elif power["line"] is not None:
        # A stroke's colour is read off its darkest pixel, so only a stroke wide enough
        # to have an unblended one is held to its colour.
        if power["line"][1] >= 1.5 and not _close(_rgb(power["line"][0]), _rgb(ours["line"][0])):
            wrong.append("line colour")
        if abs(power["line"][1] - ours["line"][1]) > 0.75:
            wrong.append("line width")
    if power["shadow"] != ours["shadow"]:
        wrong.append("shadow")
    if power["text"] is not None and not _close(
        _rgb(power["text"][0]), _rgb(ours["text"][0])
    ):
        wrong.append("text colour")
    if power["text"] is not None and _face(power["text"][1]) != _face(ours["text"][1]):
        wrong.append("text face")
    return wrong


def main() -> int:
    if len(sys.argv) != 3:
        print(__doc__)
        return 2
    pptx, pdf = Path(sys.argv[1]).expanduser(), Path(sys.argv[2]).expanduser()
    power, ours = read_export(pdf), read_model(pptx)
    disagreements = 0
    for spec in PROBES:
        key = spec["key"]
        p, o = power[key], ours[key]
        wrong = compare(p, o)
        disagreements += bool(wrong)
        print(f"{key:26s} {'ok' if not wrong else 'DIFFERS: ' + ', '.join(wrong)}")
        print(f"    PowerPoint fill {' '.join(p['fill'])}  line {p['line']}  "
              f"shadow {p['shadow']}  text {p['text']}")
        print(f"    pptx2svg   fill {' '.join(o['fill'])}  line {o['line']}  "
              f"shadow {o['shadow']}  text {o['text']}")
    print(f"\n{len(PROBES) - disagreements} of {len(PROBES)} probes agree")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
