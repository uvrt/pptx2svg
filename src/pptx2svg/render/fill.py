"""Fill, stroke and arrow-marker attributes.

Renderers here return *attribute strings* rather than elements, because the caller
splices them into whichever geometry element the shape produced.  Anything that needs a
``<defs>`` entry (gradients, image patterns, hatch patterns, markers) registers it on the
:class:`RenderContext` and returns a ``url(#id)`` reference.

SVG output uses inline attributes only -- no CSS classes.  librsvg and resvg, the usual
rasterisation backends, do not apply CSS selectors reliably.
"""

from __future__ import annotations

import math

from .. import model as m
from ..units import emu_to_px
from .context import RenderContext, num

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


def render_fill_attrs(fill: m.Fill | None, context: RenderContext) -> str:
    """``fill="..."`` (plus ``fill-opacity``) for a shape."""
    if fill is None or isinstance(fill, m.NoFill):
        return 'fill="none"'

    if isinstance(fill, m.SolidFill):
        opacity = f' fill-opacity="{num(fill.color.alpha)}"' if fill.color.alpha < 1 else ""
        return f'fill="{fill.color.hex}"{opacity}'

    if isinstance(fill, m.GradientFill):
        return f'fill="{_gradient_ref(fill, context)}"'

    if isinstance(fill, m.ImageFill):
        return f'fill="{_image_fill_ref(fill, context)}"'

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


def _image_fill_ref(fill: m.ImageFill, context: RenderContext) -> str:
    pattern_id = context.new_id("imgfill")
    href = f"data:{fill.mime_type};base64,{fill.image_data}"

    if fill.tile is not None:
        # A tiled fill repeats at sx/sy of the shape's bounding box.
        width = num(fill.tile.sx * 100)
        height = num(fill.tile.sy * 100)
        context.add_def(
            f'<pattern id="{pattern_id}" patternUnits="objectBoundingBox" '
            f'width="{width}%" height="{height}%">'
            f'<image href="{href}" width="100%" height="100%" preserveAspectRatio="none"/>'
            "</pattern>"
        )
    else:
        context.add_def(
            f'<pattern id="{pattern_id}" patternContentUnits="objectBoundingBox" '
            f'width="1" height="1">'
            f'<image href="{href}" width="1" height="1" preserveAspectRatio="none"/>'
            "</pattern>"
        )
    return f"url(#{pattern_id})"


def _pattern_fill_attrs(fill: m.PatternFill, context: RenderContext) -> str:
    content = _pattern_content(
        fill.preset, fill.foreground_color.hex, fill.foreground_color.alpha
    )
    if content is None:
        opacity = (
            f' fill-opacity="{num(fill.foreground_color.alpha)}"'
            if fill.foreground_color.alpha < 1
            else ""
        )
        return f'fill="{fill.foreground_color.hex}"{opacity}'

    svg, size = content
    pattern_id = context.new_id("patt")
    bg_opacity = (
        f' fill-opacity="{num(fill.background_color.alpha)}"'
        if fill.background_color.alpha < 1
        else ""
    )
    context.add_def(
        f'<pattern id="{pattern_id}" patternUnits="userSpaceOnUse" '
        f'width="{num(size)}" height="{num(size)}">'
        f'<rect width="{num(size)}" height="{num(size)}" '
        f'fill="{fill.background_color.hex}"{bg_opacity}/>{svg}</pattern>'
    )
    return f'fill="url(#{pattern_id})"'


def _pattern_content(preset: str, fg: str, alpha: float) -> tuple[str, float] | None:
    """Hatch/stipple tiles for ``a:pattFill`` presets."""
    size = 8.0
    opacity = f' opacity="{num(alpha)}"' if alpha < 1 else ""

    def line(x1: float, y1: float, x2: float, y2: float) -> str:
        return (
            f'<line x1="{num(x1)}" y1="{num(y1)}" x2="{num(x2)}" y2="{num(y2)}" '
            f'stroke="{fg}" stroke-width="1"{opacity}/>'
        )

    def dot(x: float, y: float, w: float, h: float) -> str:
        return (
            f'<rect x="{num(x)}" y="{num(y)}" width="{num(w)}" height="{num(h)}" '
            f'fill="{fg}"{opacity}/>'
        )

    if preset in ("ltHorz", "horz"):
        return line(0, 4, 8, 4), size
    if preset in ("ltVert", "vert"):
        return line(4, 0, 4, 8), size
    if preset in ("ltDnDiag", "dnDiag"):
        return line(0, 0, 8, 8), size
    if preset in ("ltUpDiag", "upDiag"):
        return line(0, 8, 8, 0), size
    if preset == "dkHorz":
        return line(0, 2, 8, 2) + line(0, 6, 8, 6), size
    if preset == "dkVert":
        return line(2, 0, 2, 8) + line(6, 0, 6, 8), size
    if preset == "dkDnDiag":
        return line(0, 0, 8, 8) + line(-4, 0, 4, 8), size
    if preset == "dkUpDiag":
        return line(0, 8, 8, 0) + line(4, 8, 12, 0), size
    if preset in ("cross", "smGrid"):
        return line(0, 4, 8, 4) + line(4, 0, 4, 8), size
    if preset == "lgGrid":
        return line(0, 0, 16, 0) + line(0, 0, 0, 16), 16.0
    if preset == "diagCross":
        return line(0, 0, 8, 8) + line(0, 8, 8, 0), size
    if preset == "pct5":
        return dot(0, 0, 1, 1), size
    if preset == "pct10":
        return dot(0, 0, 1, 1) + dot(4, 4, 1, 1), size
    if preset == "pct20":
        return dot(0, 0, 2, 2) + dot(4, 4, 2, 2), size
    if preset == "pct25":
        return dot(0, 0, 2, 2) + dot(4, 0, 2, 2) + dot(2, 4, 2, 2) + dot(6, 4, 2, 2), size
    if preset.startswith("pct"):
        try:
            percent = int(preset[3:])
        except ValueError:
            return None
        return (
            f'<rect width="{num(size)}" height="{num(size)}" fill="{fg}" '
            f'opacity="{num(percent / 100)}"{opacity}/>',
            size,
        )
    return None


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
