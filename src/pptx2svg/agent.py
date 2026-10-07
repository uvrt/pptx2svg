"""The agent view: a compact, per-slide SVG for a model to read, not to look at.

:func:`convert_pptx_to_agent_svg` describes each slide as an SVG whose user units are
**points** (``viewBox="0 0 960 540"`` for a 16:9 slide), with every shape addressable:

* each shape is one element, or a ``<g>`` holding its geometry and its text, carrying
  ``data-pptx-id`` exactly as the normal render writes it (``<slide id>.<cNvPr id>``;
  ``lay:``/``mst:`` for the layout's and master's own shapes, which are drawn first and
  marked ``data-layer="layout"``/``"master"``), in document order;
* geometry is the shape's box in slide points, numbers rounded to a tenth of a point; a shape's
  preset is ``data-preset`` (an ellipse is an ``<ellipse>``, a line a ``<line>``, every
  other preset a ``<rect>`` of its box), rotation a ``rotate()`` about the box's centre;
  a group's members are in slide points too, and the group is ``<g data-kind="group">``;
* colours are the hex with the theme name where the deck used one:
  ``fill="#0B6E79" data-fill="accent1"``, ``data-stroke="tx1 lumMod=75%"``; a gradient is
  drawn as its first stop with ``data-fill-kind="gradient"``; dashes and arrowheads are
  named (``data-dash="dash"``, ``data-arrow-end="triangle"``);
* text is a ``<text>`` per shape at its text area's top left (the box less its insets,
  which are written as ``data-insets`` only when they are not PowerPoint's default 7.2 and
  3.6 pt), one ``<tspan>`` per paragraph and one per run, with size, weight, style and colour; the
  vertical anchor is ``data-anchor``.  Lines are not broken: wrapping is the renderer's
  job, and the text is what a reader needs;
* pictures, charts, tables, SmartArt and media are placeholders --
  ``<rect data-kind="picture" .../>`` -- a table with its cell text, a chart with its
  type and title, SmartArt with its nodes' text.

There are no glyph outlines, fonts, ``@font-face``, images, base64, filters or ``<defs>``:
an agent view is typically a tenth of the normal render's size or less.  Text is not
measured, so no font is read either.
"""

from __future__ import annotations

import math
import os
from collections import Counter
from typing import Callable, Iterable
from xml.sax.saxutils import escape, quoteattr

from . import model as m
from .opc import OpcPackage
from .parse.parts import read_presentation
from .resolve import resolve_presentation
from .resolve.naming import name_of

__all__ = ["convert_pptx_to_agent_svg", "render_slide_to_agent_svg"]

EMU_PER_POINT = 12700

#: Presets drawn as their own SVG element; every other preset is a ``<rect>`` of its box.
_ELLIPSES = {"ellipse", "flowChartConnector"}
_LINES = {"line", "straightConnector1"}

#: Maps a box in some child coordinate space (points) to slide points.
Mapper = Callable[[float, float, float, float], "tuple[float, float, float, float]"]


def convert_pptx_to_agent_svg(source, options=None) -> list[str]:
    """Describe a deck as one compact agent-view SVG per slide (see the module docstring).

    ``source`` is a path, bytes, a file object or an :class:`~pptx2svg.opc.OpcPackage`, as
    for :func:`pptx2svg.convert_pptx_to_svg`; of ``options`` (a
    :class:`~pptx2svg.ConvertOptions`) only ``slide_numbers`` and ``warnings`` apply.
    """
    if isinstance(source, OpcPackage):
        package = source
    elif isinstance(source, (str, os.PathLike)):
        package = OpcPackage.open(os.fspath(source))
    else:
        package = OpcPackage.open(source)
    presentation = read_presentation(package)
    names: dict = {}
    slide_numbers = getattr(options, "slide_numbers", None) if options is not None else None
    resolved = resolve_presentation(package, presentation, slide_numbers=slide_numbers,
                                    color_names=names)
    if options is not None and getattr(options, "warnings", None) is not None:
        options.warnings.extend(resolved.warnings)
    return [render_slide_to_agent_svg(slide, resolved.slide_size, names)
            for slide in resolved.slides]


def render_slide_to_agent_svg(slide: m.Slide, slide_size: m.SlideSize,
                              names: dict | None = None) -> str:
    """One resolved slide as an agent-view SVG.  ``names`` is what
    :func:`~pptx2svg.resolve.resolve_presentation` collected with ``color_names=``."""
    writer = _Writer(names or {})
    width, height = slide_size.width / EMU_PER_POINT, slide_size.height / EMU_PER_POINT
    writer.font = _common_family(slide.elements)
    body = [writer.background(slide, width, height)]
    for element in slide.elements:
        body.append(writer.element(element, _identity))
    font = f' font-family={quoteattr(writer.font)}' if writer.font else ""
    return (f'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 {_n(width)} {_n(height)}" '
            f'data-slide="{slide.slide_number}"{font}>'
            + "".join(part for part in body if part) + "</svg>")


# -- numbers and attributes --------------------------------------------------------------------


def _n(value: float) -> str:
    """One decimal at most (a tenth of a point), no trailing zeros, no ``-0``."""
    text = f"{round(float(value), 1):.1f}".rstrip("0").rstrip(".")
    return "0" if text in ("-0", "") else text


def _pt(emu: float) -> float:
    return emu / EMU_PER_POINT


def _attrs(pairs: Iterable[tuple[str, object]]) -> str:
    out = []
    for name, value in pairs:
        if value is None or value is False:
            continue
        if isinstance(value, float) or isinstance(value, int) and not isinstance(value, bool):
            value = _n(value)
        out.append(f" {name}={quoteattr(str(value))}")
    return "".join(out)


def _identity(x: float, y: float, w: float, h: float):
    return x, y, w, h


def _box(transform: m.Transform, mapper: Mapper) -> tuple[float, float, float, float]:
    return mapper(_pt(transform.offset_x), _pt(transform.offset_y),
                  _pt(transform.extent_width), _pt(transform.extent_height))


def _rotation(transform: m.Transform, box) -> str | None:
    angle = transform.rotation % 360
    if not angle:
        return None
    x, y, w, h = box
    return f"rotate({_n(angle)} {_n(x + w / 2)} {_n(y + h / 2)})"


def _flip(transform: m.Transform) -> str | None:
    flips = ("h" if transform.flip_h else "") + ("v" if transform.flip_v else "")
    return flips or None


def _common_family(elements) -> str | None:
    """The typeface most runs use: written once on the root, not on every run."""
    counts: Counter = Counter()

    def walk(items):
        for element in items:
            body = getattr(element, "text_body", None)
            if body is not None:
                for paragraph in body.paragraphs:
                    for run in paragraph.runs:
                        if run.properties.font_family and run.text.strip():
                            counts[run.properties.font_family] += 1
            walk(getattr(element, "children", ()) or ())

    walk(elements)
    return counts.most_common(1)[0][0] if counts else None


# -- the writer --------------------------------------------------------------------------------


class _Writer:
    def __init__(self, names: dict) -> None:
        self.names = names
        self.font: str | None = None

    # colours

    def _color(self, attribute: str, data: str, color) -> list[tuple[str, object]]:
        if color is None:
            return [(attribute, "none")]
        pairs: list[tuple[str, object]] = [(attribute, color.hex.upper())]
        if color.alpha < 1:
            pairs.append((f"{attribute}-opacity", round(color.alpha, 2)))
        name = name_of(self.names, color)
        if name:
            pairs.append((data, name))
        return pairs

    def fill(self, fill) -> list[tuple[str, object]]:
        if fill is None or isinstance(fill, m.NoFill):
            return [("fill", "none")]
        if isinstance(fill, m.SolidFill):
            return self._color("fill", "data-fill", fill.color)
        if isinstance(fill, m.GradientFill):
            first = fill.stops[0].color if fill.stops else None
            return self._color("fill", "data-fill", first) + [("data-fill-kind", "gradient")]
        if isinstance(fill, m.PatternFill):
            return self._color("fill", "data-fill", fill.foreground_color) + [
                ("data-fill-kind", "pattern")]
        if isinstance(fill, m.ImageFill):
            return [("fill", "none"), ("data-fill-kind", "picture")]
        return [("fill", "none")]

    def stroke(self, outline) -> list[tuple[str, object]]:
        if outline is None or outline.fill is None:
            return []
        fill = outline.fill
        color = fill.color if isinstance(fill, m.SolidFill) else (
            fill.stops[0].color if isinstance(fill, m.GradientFill) and fill.stops else None)
        if color is None:
            return []
        pairs = self._color("stroke", "data-stroke", color)
        pairs.append(("stroke-width", _pt(outline.width)))
        if outline.dash_style and outline.dash_style != "solid":
            pairs.append(("data-dash", outline.dash_style))
        for end, arrow in (("start", outline.head_end), ("end", outline.tail_end)):
            if arrow is not None and getattr(arrow, "type", "none") != "none":
                pairs.append((f"data-arrow-{end}", arrow.type))
        return pairs

    # elements

    def background(self, slide: m.Slide, width: float, height: float) -> str:
        fill = slide.background.fill if slide.background else None
        if fill is None:
            return ""
        pairs = self.fill(fill)
        if pairs[0] == ("fill", "#FFFFFF") and not any(k == "fill-opacity" for k, _ in pairs):
            return ""
        return f'<rect data-kind="background"{_attrs([("width", width), ("height", height)])}' \
               f'{_attrs(pairs)}/>'

    def element(self, element, mapper: Mapper) -> str:
        identity = self._identity(element)
        if isinstance(element, m.ChartElement):
            return self.chart(element, mapper, identity)
        if isinstance(element, m.GroupElement):
            return self.group(element, mapper, identity)
        if isinstance(element, m.ShapeElement):
            return self.shape(element, mapper, identity)
        if isinstance(element, m.ConnectorElement):
            return self.connector(element, mapper, identity)
        if isinstance(element, m.ImageElement):
            box = _box(element.transform, mapper)
            return self._placeholder("picture", box, identity, element.transform)
        if isinstance(element, m.TableElement):
            return self.table(element, mapper, identity)
        return ""

    def _identity(self, element) -> list[tuple[str, object]]:
        element_id = getattr(element, "element_id", None)
        pairs: list[tuple[str, object]] = [("data-pptx-id", element_id)]
        if element_id and element_id.startswith("lay:"):
            pairs.append(("data-layer", "layout"))
        elif element_id and element_id.startswith("mst:"):
            pairs.append(("data-layer", "master"))
        return pairs

    def _placeholder(self, kind: str, box, identity, transform=None, extra=()) -> str:
        x, y, w, h = box
        pairs = [*identity, ("data-kind", kind), ("x", x), ("y", y), ("width", w),
                 ("height", h), ("fill", "none"), *extra]
        if transform is not None:
            pairs.append(("transform", _rotation(transform, box)))
        return f"<rect{_attrs(pairs)}/>"

    def shape(self, element: m.ShapeElement, mapper: Mapper, identity) -> str:
        box = _box(element.transform, mapper)
        x, y, w, h = box
        geometry = element.geometry
        preset = geometry.preset if isinstance(geometry, m.PresetGeometry) else "custom"
        paint = self.fill(element.fill) + self.stroke(element.outline)
        rotate = _rotation(element.transform, box)
        flip = _flip(element.transform)
        shape_pairs: list[tuple[str, object]]
        if preset in _ELLIPSES:
            tag = "ellipse"
            shape_pairs = [("cx", x + w / 2), ("cy", y + h / 2), ("rx", w / 2), ("ry", h / 2)]
        elif preset in _LINES:
            tag = "line"
            x1, x2 = (x + w, x) if element.transform.flip_h else (x, x + w)
            y1, y2 = (y + h, y) if element.transform.flip_v else (y, y + h)
            shape_pairs = [("x1", x1), ("y1", y1), ("x2", x2), ("y2", y2)]
            paint = [pair for pair in paint if pair[0] != "fill"]
            flip = None
        else:
            tag = "rect"
            shape_pairs = [("x", x), ("y", y), ("width", w), ("height", h)]
            if preset == "roundRect":
                adjust = min(max(geometry.adjust_values.get("adj", 16667), 0), 50000)
                shape_pairs.append(("rx", min(w, h) * adjust / 100000))
        if preset not in ("rect", "ellipse", "line") and tag != "ellipse":
            shape_pairs.append(("data-preset", preset))
        if element.placeholder_type is not None or element.placeholder_idx is not None:
            shape_pairs.append(("data-placeholder", element.placeholder_type or "body"))
        text_box = _box(element.text_transform, mapper) \
            if element.text_transform is not None else box
        text = self.text(element.text_body, text_box) if element.text_body is not None else ""
        if not text:
            return f"<{tag}{_attrs([*identity, *shape_pairs, *paint, ('data-flip', flip), ('transform', rotate)])}/>"
        inner = f"<{tag}{_attrs([*shape_pairs, *paint, ('data-flip', flip)])}/>"
        return f"<g{_attrs([*identity, ('transform', rotate)])}>{inner}{text}</g>"

    def connector(self, element: m.ConnectorElement, mapper: Mapper, identity) -> str:
        transform = element.transform
        x, y, w, h = _box(transform, mapper)
        ends = [(x, y), (x + w, y + h)]
        if transform.flip_h:
            ends = [(x + w, ends[0][1]), (x, ends[1][1])]
        if transform.flip_v:
            ends = [(ends[0][0], y + h), (ends[1][0], y)]
        angle = transform.rotation % 360
        if angle:
            cx, cy = x + w / 2, y + h / 2
            cos, sin = math.cos(math.radians(angle)), math.sin(math.radians(angle))
            ends = [(cx + (px - cx) * cos - (py - cy) * sin, cy + (px - cx) * sin + (py - cy) * cos)
                    for px, py in ends]
        geometry = element.geometry
        preset = geometry.preset if isinstance(geometry, m.PresetGeometry) else "custom"
        pairs = [*identity, ("x1", ends[0][0]), ("y1", ends[0][1]), ("x2", ends[1][0]),
                 ("y2", ends[1][1]), *self.stroke(element.outline)]
        if preset not in _LINES:
            pairs.append(("data-preset", preset))
        return f"<line{_attrs(pairs)}/>"

    def group(self, element: m.GroupElement, mapper: Mapper, identity) -> str:
        box = _box(element.transform, mapper)
        children = list(element.children)
        is_smartart = any("/" in (getattr(child, "element_id", None) or "") for child in children)
        child_mapper = _group_mapper(element, mapper)
        if is_smartart:
            texts = [self.element(child, child_mapper) for child in _text_leaves(children)]
            texts = [_strip_identity(text) for text in texts if text]
            body = self._placeholder("smartart", box, [], element.transform)
            return f"<g{_attrs([*identity, ('data-kind', 'smartart')])}>{body}{''.join(texts)}</g>"
        inner = "".join(self.element(child, child_mapper) for child in children)
        pairs = [*identity, ("data-kind", "group"), ("transform", _rotation(element.transform, box))]
        return f"<g{_attrs(pairs)}>{inner}</g>"

    def chart(self, element: m.ChartElement, mapper: Mapper, identity) -> str:
        box = _box(element.transform, mapper)
        chart = element.chart
        extra = [("data-chart", chart.kind), ("data-title", chart.title),
                 ("data-series", len(chart.series) or None),
                 ("data-categories", len(chart.categories) or None)]
        return self._placeholder("chart", box, identity, element.transform, extra)

    def table(self, element: m.TableElement, mapper: Mapper, identity) -> str:
        box = _box(element.transform, mapper)
        x0, y0, _, _ = box
        table = element.table
        scale_x = box[2] / max(1e-9, sum(_pt(c.width) for c in table.columns) or box[2]) \
            if table.columns else 1.0
        rect = self._placeholder("table", box, [], None,
                                 [("data-rows", len(table.rows)),
                                  ("data-columns", len(table.columns))])
        cells = []
        y = y0
        for r, row in enumerate(table.rows):
            x = x0
            for c, cell in enumerate(row.cells):
                width = sum(_pt(col.width) for col in table.columns[c:c + max(1, cell.grid_span)]) \
                    * scale_x if c < len(table.columns) else 0
                if cell.text_body is not None and not cell.h_merge and not cell.v_merge:
                    text = self.text(cell.text_body, (x, y, width, _pt(row.height)),
                                     extra=[("data-cell", f"{r},{c}")])
                    if text:
                        cells.append(text)
                if c < len(table.columns):
                    x += _pt(table.columns[c].width) * scale_x
            y += _pt(row.height)
        return f"<g{_attrs([*identity, ('data-kind', 'table')])}>{rect}{''.join(cells)}</g>"

    # text

    def text(self, body: m.TextBody | None, box, *, extra=()) -> str:
        if body is None or not any(run.text.strip() for p in body.paragraphs for run in p.runs):
            return ""
        x, y, w, h = box
        props = body.body_properties
        left, top, right, bottom = x, y, x + w, y + h
        left += _pt(props.margin_left)
        right -= _pt(props.margin_right)
        top += _pt(props.margin_top)
        bottom -= _pt(props.margin_bottom)
        area_w, area_h = max(0.0, right - left), max(0.0, bottom - top)
        scale = props.font_scale or 1.0
        paragraphs = []
        for index, paragraph in enumerate(body.paragraphs):
            runs = [run for run in paragraph.runs if run.text]
            size = next((run.properties.font_size for run in runs if run.properties.font_size),
                        None) or (paragraph.end_para_run_properties.font_size
                                  if paragraph.end_para_run_properties else None) or 18
            align = paragraph.properties.alignment or "l"
            anchor_x = {"ctr": left + area_w / 2, "r": right}.get(align, left)
            pairs: list[tuple[str, object]] = [("x", anchor_x),
                                               ("dy", size * scale * (1 if index == 0 else 1.2))]
            if align in ("ctr", "r"):
                pairs.append(("text-anchor", "middle" if align == "ctr" else "end"))
            bullet = paragraph.properties.bullet
            if isinstance(bullet, m.CharBullet):
                pairs.append(("data-bullet", bullet.char))
            elif isinstance(bullet, m.AutoNumBullet):
                pairs.append(("data-bullet", bullet.scheme))
            if paragraph.properties.level:
                pairs.append(("data-level", paragraph.properties.level))
            if len(runs) == 1:
                # One run: its formatting goes on the paragraph's own tspan.
                paragraphs.append(self.run(runs[0], scale, pairs))
                continue
            spans = "".join(self.run(run, scale) for run in runs)
            paragraphs.append(f"<tspan{_attrs(pairs)}>{spans}</tspan>")
        text_pairs = [*extra, ("x", left), ("y", top)]
        insets = (props.margin_left, props.margin_top, props.margin_right, props.margin_bottom)
        if insets != (91440, 45720, 91440, 45720):
            text_pairs.append(("data-insets", " ".join(_n(_pt(v)) for v in insets)))
        if props.anchor != "t":
            text_pairs.append(("data-anchor", {"ctr": "middle", "b": "bottom"}[props.anchor]))
        if scale < 1:
            text_pairs.append(("data-font-scale", round(scale, 3)))
        if props.wrap == "none":
            text_pairs.append(("data-wrap", "none"))
        if props.vert and props.vert != "horz":
            text_pairs.append(("data-vert", props.vert))
        return f"<text{_attrs(text_pairs)}>{''.join(paragraphs)}</text>"

    def run(self, run: m.TextRun, scale: float, paragraph=None) -> str:
        p = run.properties
        pairs: list[tuple[str, object]] = list(paragraph or [])
        if p.font_size:
            pairs.append(("font-size", p.font_size * scale))
        if p.bold:
            pairs.append(("font-weight", "bold"))
        if p.italic:
            pairs.append(("font-style", "italic"))
        if p.underline:
            pairs.append(("text-decoration", "underline"))
        if p.font_family and p.font_family != self.font:
            pairs.append(("font-family", p.font_family))
        if p.color is not None:
            pairs += self._color("fill", "data-fill", p.color)
        text = escape(run.text.replace("\v", "\n"))
        return f"<tspan{_attrs(pairs)}>{text}</tspan>" if pairs or paragraph is not None else text


def _group_mapper(element: m.GroupElement, outer: Mapper) -> Mapper:
    """Child coordinates (EMU-derived points in the group's child space) to slide points."""
    t, c = element.transform, element.child_transform
    off_x, off_y = _pt(t.offset_x), _pt(t.offset_y)
    ext_w, ext_h = _pt(t.extent_width), _pt(t.extent_height)
    ch_x, ch_y = _pt(c.offset_x), _pt(c.offset_y)
    ch_w, ch_h = _pt(c.extent_width), _pt(c.extent_height)
    sx = ext_w / ch_w if ch_w else 1.0
    sy = ext_h / ch_h if ch_h else 1.0

    def mapper(x: float, y: float, w: float, h: float):
        return outer(off_x + (x - ch_x) * sx, off_y + (y - ch_y) * sy, w * sx, h * sy)

    return mapper


def _text_leaves(children) -> list:
    out = []
    for child in children:
        if isinstance(child, m.GroupElement):
            out += _text_leaves(child.children)
        elif getattr(child, "text_body", None) is not None:
            out.append(child)
    return out


def _strip_identity(fragment: str) -> str:
    """A SmartArt piece's text without its drawn-piece id: the frame is what is addressed."""
    import re

    return re.sub(r' data-pptx-id="[^"]*"', "", fragment, count=1)
