"""Fill, stroke and arrow-marker attributes.

Renderers here return *attribute strings* rather than elements, because the caller
splices them into whichever geometry element the shape produced.  Anything that needs a
``<defs>`` entry (gradients, image patterns, hatch patterns, markers) registers it on the
:class:`RenderContext` and returns a ``url(#id)`` reference.

SVG output uses inline attributes only -- no CSS classes.  librsvg and resvg, the usual
rasterisation backends, do not apply CSS selectors reliably.
"""

from __future__ import annotations

import base64
import math

from .. import model as m
from ..imagemeta import natural_size_pt
from ..units import PX_PER_PT, emu_to_px
from .context import RenderContext, num
from .pattern import PATTERN_CELL_BITS, PATTERN_CELL_PT, cell_rectangles

#: Dash patterns as multiples of the stroke width (ECMA-376 §20.1.10.49).
DASH_PATTERNS: dict[str, list[float]] = {
    "dash": [4, 3],
    "dot": [1, 3],
    "dashDot": [4, 3, 1, 3],
    "lgDash": [8, 3],
    "lgDashDot": [8, 3, 1, 3],
    "lgDashDotDot": [8, 3, 1, 3, 1, 3],
    "sysDash": [3, 1],
    "sysDot": [1, 1],
}

ARROW_SIZE_PX: dict[str, float] = {"sm": 5, "med": 8, "lg": 12}


def render_fill_attrs(
    fill: m.Fill | None,
    context: RenderContext,
    box: tuple[float, float, float, float] | None = None,
) -> str:
    """``fill="..."`` (plus ``fill-opacity``) for a shape.

    ``box`` is the filled rectangle -- ``(x, y, width, height)`` in the user space the
    fill is referenced from, all in pixels.  Only a tiled image fill needs it, and it
    needs it for a reason no other fill does: ``a:tile@algn`` registers the tile grid
    against one of the box's nine corners and edges, so without the box there is no
    right answer, only a guess that it is the top-left one.
    """
    if fill is None or isinstance(fill, m.NoFill):
        return 'fill="none"'

    if isinstance(fill, m.SolidFill):
        opacity = f' fill-opacity="{num(fill.color.alpha)}"' if fill.color.alpha < 1 else ""
        return f'fill="{fill.color.hex}"{opacity}'

    if isinstance(fill, m.GradientFill):
        return f'fill="{_gradient_ref(fill, context)}"'

    if isinstance(fill, m.ImageFill):
        return f'fill="{_image_fill_ref(fill, context, box)}"'

    if isinstance(fill, m.PatternFill):
        return _pattern_fill_attrs(fill, context)

    return 'fill="none"'


def _gradient_ref(fill: m.GradientFill, context: RenderContext) -> str:
    gradient_id = context.new_id("grad")

    stops = "".join(
        f'<stop offset="{num(stop.position * 100)}%" stop-color="{stop.color.hex}"'
        + (f' stop-opacity="{num(stop.color.alpha)}"' if stop.color.alpha < 1 else "")
        + "/>"
        for stop in fill.stops
    )

    if fill.gradient_type == "radial":
        cx = (fill.center_x if fill.center_x is not None else 0.5) * 100
        cy = (fill.center_y if fill.center_y is not None else 0.5) * 100
        # Radius reaches the farthest corner from the focus point.
        dx = max(cx, 100 - cx)
        dy = max(cy, 100 - cy)
        r = math.hypot(dx, dy)
        context.add_def(
            f'<radialGradient id="{gradient_id}" cx="{num(cx)}%" cy="{num(cy)}%" '
            f'r="{num(r)}%">{stops}</radialGradient>'
        )
        return f"url(#{gradient_id})"

    radians = math.radians(fill.angle)
    x1 = 50 - math.cos(radians) * 50
    y1 = 50 - math.sin(radians) * 50
    x2 = 50 + math.cos(radians) * 50
    y2 = 50 + math.sin(radians) * 50
    context.add_def(
        f'<linearGradient id="{gradient_id}" x1="{num(x1)}%" y1="{num(y1)}%" '
        f'x2="{num(x2)}%" y2="{num(y2)}%">{stops}</linearGradient>'
    )
    return f"url(#{gradient_id})"


def _image_fill_ref(
    fill: m.ImageFill,
    context: RenderContext,
    box: tuple[float, float, float, float] | None = None,
) -> str:
    pattern_id = context.new_id("imgfill")
    href = f"data:{fill.mime_type};base64,{fill.image_data}"

    if fill.tile is not None:
        tile = tile_pattern(pattern_id, href, fill.image_data, fill.tile, box)
        if tile is not None:
            context.add_def(tile)
            return f"url(#{pattern_id})"
        # The picture's natural size is what a tile is measured in, and a format
        # `imagemeta` cannot read has none to measure.  Stretching one copy over the
        # shape is wrong, but it is the same wrong as an untiled fill rather than a
        # tiling at an invented pitch.

    context.add_def(
        f'<pattern id="{pattern_id}" patternContentUnits="objectBoundingBox" '
        f'width="1" height="1">'
        f'<image href="{href}" width="1" height="1" preserveAspectRatio="none"/>'
        "</pattern>"
    )
    return f"url(#{pattern_id})"


#: ``@flip`` -> whether the cell carries a copy mirrored across x, and across y.
_TILE_FLIPS: dict[str, tuple[bool, bool]] = {
    "none": (False, False),
    "x": (True, False),
    "y": (False, True),
    "xy": (True, True),
}


def tile_pattern(
    pattern_id: str,
    href: str,
    image_data: str,
    tile: m.ImageFillTile | m.TileInfo,
    box: tuple[float, float, float, float] | None,
    image_attrs: str = "",
) -> str | None:
    """One ``<pattern>`` for ``a:tile``, sized and registered the way PowerPoint does.

    **The tile is the picture's own size scaled by ``sx``/``sy``.**  It has nothing to do
    with the shape.  This file used to say "a tiled fill repeats at sx/sy of the shape's
    bounding box", and the measurement refutes it: deck ``fill-tile`` drew the same 32 px
    picture at ``sx=100%`` on boxes of 68x48, 136x96, 272x192 and 400x96 pt and got a
    16.0000 pt cell on all four.  Scaling the picture's natural size instead (see
    :mod:`pptx2svg.imagemeta`) reproduces every row of that deck: 25, 50, 60, 150 and 200%
    of a 16 pt picture drew 4, 8, 9.6, 24 and 32 pt, and ``sx != sy`` moved the two axes
    independently.  ``feature-sweep`` slide 10 is the same arithmetic -- a 32 px untagged
    PNG at ``sx=60%`` is 16 x 0.6 = 9.6 pt, which is what PowerPoint drew there, against
    the 82.08 pt this drew from the box.

    Shared with the ``p:pic`` path in :mod:`pptx2svg.render.shape`, which carries the
    identical ``a:tile`` and used to size it from the frame for the stated reason that the
    two paths agreeing mattered more than either being right.  They agree here too, on the
    measurement.

    ``@algn`` registers the grid against the box and ``@tx``/``@ty`` then translate it;
    see :func:`_tile_origin`.  ``@flip`` mirrors alternate copies, which an SVG
    ``<pattern>`` cannot do by repeating one tile -- so the cell is doubled and holds the
    mirrored copies itself, which is exactly what PowerPoint's own export does (a
    ``flip="xy"`` tile of a 32 px picture exports as a **64 x 64** image on a doubled
    cell).  Measured: from the registration point the order is original then mirror in
    both axes, and ``flip="x"`` mirrors **horizontally**.
    """
    try:
        data = base64.b64decode(image_data, validate=True)
    except (ValueError, TypeError):
        return None
    natural = natural_size_pt(data)
    if natural is None:
        return None

    width = natural[0] * PX_PER_PT * tile.sx
    height = natural[1] * PX_PER_PT * tile.sy
    if width <= 0 or height <= 0:
        return None

    mirror_x, mirror_y = _TILE_FLIPS.get(tile.flip, (False, False))
    cell_width = width * (2 if mirror_x else 1)
    cell_height = height * (2 if mirror_y else 1)

    x, y = _tile_origin(tile.align, box, width, height)
    x += emu_to_px(tile.tx)
    y += emu_to_px(tile.ty)

    copies = [(0.0, 0.0, 1, 1)]
    if mirror_x:
        copies.append((2 * width, 0.0, -1, 1))
    if mirror_y:
        copies.append((0.0, 2 * height, 1, -1))
    if mirror_x and mirror_y:
        copies.append((2 * width, 2 * height, -1, -1))

    images = "".join(
        f'<image href="{href}" width="{num(width)}" height="{num(height)}" '
        f'preserveAspectRatio="none"'
        + (f" {image_attrs}" if image_attrs else "")
        + (
            ""
            if (scale_x, scale_y) == (1, 1)
            else f' transform="translate({num(offset_x)}, {num(offset_y)}) '
            f'scale({scale_x}, {scale_y})"'
        )
        + "/>"
        for offset_x, offset_y, scale_x, scale_y in copies
    )
    return (
        f'<pattern id="{pattern_id}" patternUnits="userSpaceOnUse" '
        f'x="{num(x)}" y="{num(y)}" '
        f'width="{num(cell_width)}" height="{num(cell_height)}">{images}</pattern>'
    )


def _tile_origin(
    align: str,
    box: tuple[float, float, float, float] | None,
    width: float,
    height: float,
) -> tuple[float, float]:
    """Where ``@algn`` puts the grid, in the same space as ``box``.

    Measured on deck ``fill-algn``, which had to be built twice.  The first sweep put all
    nine alignments on a 136 x 96 pt box with an 8 pt tile, and 136 and 96 are 17 and 12
    whole tiles -- so left-, centre- and right-registration all landed on the same lattice
    and the sweep said nothing at all.  Re-run on a 130 x 90 box with an 8 pt tile and a
    137 x 83 box with a 12 pt one -- indivisible in both axes both times -- the three rules
    separate cleanly:

    * leading (``tl``/``l``/``bl`` in x, ``tl``/``t``/``tr`` in y) puts the tile's leading
      edge on the box's, so the origin is the box's own left or top;
    * trailing puts the tile's trailing edge on the box's, so the origin is the box's end
      minus one tile -- read as -6 on 130 pt / 8 pt and -7 on 137 pt / 12 pt, which is
      that value modulo the cell;
    * centred puts **one tile's centre on the box's centre** -- read as -3 and -9.5 for the
      same two, which is ``(box - tile) / 2`` modulo the cell.

    The two axes are chosen independently by the two halves of the name, so the nine
    values are one product of three rules with three rather than nine separate cases.
    """
    if box is None:
        return 0.0, 0.0
    left, top, box_width, box_height = box
    if align in ("tr", "r", "br"):
        x = left + box_width - width
    elif align in ("t", "ctr", "b"):
        x = left + (box_width - width) / 2
    else:
        x = left
    if align in ("bl", "b", "br"):
        y = top + box_height - height
    elif align in ("l", "ctr", "r"):
        y = top + (box_height - height) / 2
    else:
        y = top
    return x, y


def _pattern_fill_attrs(fill: m.PatternFill, context: RenderContext) -> str:
    """``a:pattFill`` as a ``<pattern>`` of the preset's measured 8 x 8 cell.

    The cell is **8.0 pt**, which is ``PATTERN_CELL_PT * PX_PER_PT`` pixels here, and one
    bit of the preset's bitmap is one point.  The old code used 8 *pixels*, which is 6 pt,
    so every pattern in the library tiled a third too finely; see
    :mod:`pptx2svg.render.pattern` for how the cell and all 54 bitmaps were measured.

    **The lattice's phase is measured but not reproduced.**  PowerPoint registers the grid
    to the slide's own top-left corner: deck ``fill-pitch`` put a shape's left edge at
    36.0, 36.5, 38.0, 41.0, 100.3, 173.75, 260.125 and 411.0 pt and the exported pattern
    origin snapped every one down to the 8 pt lattice -- 32, 32, 32, 40, 96, 168, 256, 408
    -- so two shapes whose left edges differ by 4 pt get patterns half a cell out of step
    with each other.  This registers to each shape's own top-left instead, which is the
    phase PowerPoint gives a shape that happens to sit on an 8 pt boundary.  The cost is
    exactly that offset and never the pitch, and it is visible: rendered at 5760 px,
    ``feature-sweep`` slide 13's ``cross`` box (at 452, 52 pt, both 4 past a lattice point)
    puts its first rule 3.938 pt into the box for PowerPoint and 0.188 pt for us, on an
    identical 8.0000 pt period -- while the ``horz`` box at 244, 168 pt, which *is* on the
    lattice, matches rule for rule.

    It is left alone because reproducing it needs a concept this renderer does not have --
    the shape's position on the *slide* -- and because that concept immediately raises two
    further questions no probe here answers.  A shape inside a group knows only its offset
    within the group, and a group may also rotate, flip and scale; and a *rotated* shape
    would need to know whether PowerPoint turns the hatch with it or leaves it square to
    the page, which is a device-space brush's usual behaviour and would make the phase a
    property of the drawing surface rather than of the document.  Landing one corner of
    that -- the ungrouped, unrotated case -- would be fitting the fixture rather than the
    law.  Measure those two first; the instrument is ``tools/make_fill_probe.py``.
    """
    rectangles = cell_rectangles(fill.preset)
    if rectangles is None:
        # Not a value ST_PresetPatternVal allows at all.  A flat foreground is a poor
        # picture, but every value the schema does allow is measured.
        opacity = (
            f' fill-opacity="{num(fill.foreground_color.alpha)}"'
            if fill.foreground_color.alpha < 1
            else ""
        )
        return f'fill="{fill.foreground_color.hex}"{opacity}'

    cell = PATTERN_CELL_PT * PX_PER_PT
    unit = cell / PATTERN_CELL_BITS
    pattern_id = context.new_id("patt")

    foreground = fill.foreground_color
    fg_opacity = f' fill-opacity="{num(foreground.alpha)}"' if foreground.alpha < 1 else ""
    bg_opacity = (
        f' fill-opacity="{num(fill.background_color.alpha)}"'
        if fill.background_color.alpha < 1
        else ""
    )
    marks = "".join(
        f'<rect x="{num(x * unit)}" y="{num(y * unit)}" '
        f'width="{num(width * unit)}" height="{num(height * unit)}" '
        f'fill="{foreground.hex}"{fg_opacity}/>'
        for x, y, width, height in rectangles
    )
    context.add_def(
        f'<pattern id="{pattern_id}" patternUnits="userSpaceOnUse" '
        f'width="{num(cell)}" height="{num(cell)}">'
        f'<rect width="{num(cell)}" height="{num(cell)}" '
        f'fill="{fill.background_color.hex}"{bg_opacity}/>{marks}</pattern>'
    )
    return f'fill="url(#{pattern_id})"'


def render_outline_attrs(outline: m.Outline | None, context: RenderContext) -> str:
    """``stroke``/``stroke-width``/``stroke-dasharray`` etc. for a shape."""
    if outline is None:
        return 'stroke="none"'

    width_px = emu_to_px(outline.width)
    parts = [f'stroke-width="{num(width_px)}"']

    if outline.fill is None:
        parts.append('stroke="none"')
    elif isinstance(outline.fill, m.SolidFill):
        parts.append(f'stroke="{outline.fill.color.hex}"')
        if outline.fill.color.alpha < 1:
            parts.append(f'stroke-opacity="{num(outline.fill.color.alpha)}"')
    elif isinstance(outline.fill, m.GradientFill):
        parts.append(f'stroke="{_gradient_ref(outline.fill, context)}"')

    # Dash lengths are multiples of the stroke width in OOXML, absolute in SVG.
    if outline.custom_dash:
        parts.append(
            'stroke-dasharray="'
            + " ".join(num(value * width_px) for value in outline.custom_dash)
            + '"'
        )
    elif outline.dash_style != "solid":
        pattern = DASH_PATTERNS.get(outline.dash_style)
        if pattern:
            parts.append(
                'stroke-dasharray="' + " ".join(num(v * width_px) for v in pattern) + '"'
            )

    if outline.line_cap:
        parts.append(f'stroke-linecap="{outline.line_cap}"')
    if outline.line_join:
        parts.append(f'stroke-linejoin="{outline.line_join}"')

    return " ".join(parts)


def render_markers(outline: m.Outline | None, context: RenderContext) -> str:
    """``marker-start``/``marker-end`` attributes for a connector's arrowheads."""
    if outline is None or (outline.head_end is None and outline.tail_end is None):
        return ""

    color, alpha = "#000000", 1.0
    if isinstance(outline.fill, m.SolidFill):
        color, alpha = outline.fill.color.hex, outline.fill.color.alpha
    elif isinstance(outline.fill, m.GradientFill) and outline.fill.stops:
        color, alpha = outline.fill.stops[0].color.hex, outline.fill.stops[0].color.alpha

    attrs: list[str] = []
    if outline.head_end is not None:
        marker_id = context.new_id("marker")
        definition = _marker_def(marker_id, outline.head_end, color, alpha)
        if definition:
            context.add_def(definition)
            attrs.append(f'marker-start="url(#{marker_id})"')
    if outline.tail_end is not None:
        marker_id = context.new_id("marker")
        definition = _marker_def(marker_id, outline.tail_end, color, alpha)
        if definition:
            context.add_def(definition)
            attrs.append(f'marker-end="url(#{marker_id})"')

    return " ".join(attrs)


def _marker_def(
    marker_id: str, endpoint: m.ArrowEndpoint, color: str, alpha: float
) -> str | None:
    mw = ARROW_SIZE_PX[endpoint.length]
    mh = ARROW_SIZE_PX[endpoint.width]
    opacity = f' opacity="{num(alpha)}"' if alpha < 1 else ""

    if endpoint.type == "none":
        return None

    if endpoint.type == "oval":
        return (
            f'<marker id="{marker_id}" markerWidth="{num(mw)}" markerHeight="{num(mh)}" '
            f'refX="{num(mw/2)}" refY="{num(mh/2)}" orient="auto" markerUnits="userSpaceOnUse">'
            f'<ellipse cx="{num(mw/2)}" cy="{num(mh/2)}" rx="{num(mw/2)}" ry="{num(mh/2)}" '
            f'fill="{color}"{opacity}/></marker>'
        )

    if endpoint.type == "triangle":
        path = f"M 0 0 L {num(mw)} {num(mh/2)} L 0 {num(mh)} Z"
        fill_attr = f'fill="{color}"'
    elif endpoint.type == "stealth":
        path = f"M 0 0 L {num(mw)} {num(mh/2)} L 0 {num(mh)} L {num(mw*0.3)} {num(mh/2)} Z"
        fill_attr = f'fill="{color}"'
    elif endpoint.type == "diamond":
        path = (
            f"M 0 {num(mh/2)} L {num(mw/2)} 0 L {num(mw)} {num(mh/2)} L {num(mw/2)} {num(mh)} Z"
        )
        fill_attr = f'fill="{color}"'
    elif endpoint.type == "arrow":
        path = f"M 0 0 L {num(mw)} {num(mh/2)} L 0 {num(mh)}"
        fill_attr = f'fill="none" stroke="{color}" stroke-width="1"'
    else:
        return None

    return (
        f'<marker id="{marker_id}" markerWidth="{num(mw)}" markerHeight="{num(mh)}" '
        f'refX="{num(mw)}" refY="{num(mh/2)}" orient="auto" markerUnits="userSpaceOnUse">'
        f'<path d="{path}" {fill_attr}{opacity}/></marker>'
    )
