"""Slide-to-SVG assembly.

Drawing an element or a group of them moved to :mod:`ooxml_common.drawingml.elements`, with
its history, so that a chart or a SmartArt diagram in a Word document is drawn by the same
code; the slide -- its background, its size and the document around it -- is drawn here.

Output is SVG 1.1 with inline attributes only -- no CSS classes, no stylesheets.  The
usual rasterisation backends (librsvg via cairosvg, and resvg) do not apply CSS selectors
reliably, so every property is written as a presentation attribute.
"""

from __future__ import annotations

from ooxml_common.drawingml.elements import (  # noqa: F401  (moved; re-exported)
    UNIFORM_SCALE_TOLERANCE,
    render_element,
    render_group,
    swaps_group_axes,
)

from .. import model as m
from ..units import emu_to_px
from .context import RenderContext, num
from .fill import render_fill_attrs


def render_slide_to_svg(
    slide: m.Slide,
    slide_size: m.SlideSize,
    context: RenderContext | None = None,
    *,
    width: float | None = None,
    height: float | None = None,
) -> str:
    """Render one resolved slide to a standalone SVG document.

    ``width``/``height`` override the output size in pixels; give just one and the other
    follows the slide's aspect ratio.  The ``viewBox`` always stays at the slide's natural
    pixel size, so the drawing scales rather than reflows.
    """
    context = context or RenderContext()

    view_width = emu_to_px(slide_size.width)
    view_height = emu_to_px(slide_size.height)
    out_width, out_height = _output_size(view_width, view_height, width, height)

    body: list[str] = [_render_background(slide, view_width, view_height, context)]

    for element in slide.elements:
        rendered = render_element(element, context)
        if rendered:
            body.append(rendered)

    defs = f"<defs>{''.join(context.defs)}</defs>" if context.defs else ""

    return (
        '<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
        f'viewBox="0 0 {num(view_width)} {num(view_height)}" '
        f'width="{num(out_width)}" height="{num(out_height)}">'
        f"{defs}{''.join(body)}</svg>"
    )


def _output_size(
    view_width: float, view_height: float, width: float | None, height: float | None
) -> tuple[float, float]:
    """Resolve the output size, keeping the slide's aspect ratio when only one axis is set.

    Letterboxing would otherwise appear: the viewBox keeps the slide's proportions, so a
    width-only override with an unchanged height leaves bars down the sides.
    """
    if width is None and height is None:
        return view_width, view_height
    if width is not None and height is not None:
        return width, height
    if width is not None:
        return width, width * (view_height / view_width) if view_width else view_height
    return height * (view_width / view_height) if view_height else view_width, height


def _render_background(
    slide: m.Slide, width: float, height: float, context: RenderContext
) -> str:
    fill = slide.background.fill if slide.background else None

    if isinstance(fill, m.ImageFill):
        href = f"data:{fill.mime_type};base64,{fill.image_data}"
        return (
            f'<image href="{href}" width="{num(width)}" height="{num(height)}" '
            'preserveAspectRatio="none"/>'
        )

    if fill is not None:
        attrs = render_fill_attrs(fill, context, (0, 0, width, height))
        return f'<rect width="{num(width)}" height="{num(height)}" {attrs}/>'

    # PowerPoint's implicit background is white, not transparent.
    return f'<rect width="{num(width)}" height="{num(height)}" fill="#FFFFFF"/>'
