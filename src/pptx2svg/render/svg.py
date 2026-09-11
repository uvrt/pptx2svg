"""Slide-to-SVG assembly.

Output is SVG 1.1 with inline attributes only -- no CSS classes, no stylesheets.  The
usual rasterisation backends (librsvg via cairosvg, and resvg) do not apply CSS selectors
reliably, so every property is written as a presentation attribute.
"""

from __future__ import annotations

from .. import model as m
from ..units import emu_to_px
from .context import RenderContext, escape_xml_attr, num
from .fill import render_fill_attrs
from .shape import render_connector, render_image, render_shape, render_table


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
        attrs = render_fill_attrs(fill, context)
        return f'<rect width="{num(width)}" height="{num(height)}" {attrs}/>'

    # PowerPoint's implicit background is white, not transparent.
    return f'<rect width="{num(width)}" height="{num(height)}" fill="#FFFFFF"/>'


def render_element(element: m.SlideElement, context: RenderContext) -> str:
    if isinstance(element, m.ShapeElement):
        rendered = render_shape(element, context)
    elif isinstance(element, m.ImageElement):
        rendered = render_image(element, context)
    elif isinstance(element, m.ConnectorElement):
        rendered = render_connector(element, context)
    elif isinstance(element, m.GroupElement):
        rendered = render_group(element, context)
    elif isinstance(element, m.TableElement):
        rendered = render_table(element, context)
    else:
        return ""

    alt_text = getattr(element, "alt_text", None)
    if rendered and alt_text:
        rendered = _add_aria_label(rendered, alt_text)

    hyperlink = getattr(element, "hyperlink", None)
    if rendered and hyperlink is not None:
        rendered = f'<a href="{escape_xml_attr(hyperlink.url)}">{rendered}</a>'

    return rendered


def _add_aria_label(fragment: str, alt_text: str) -> str:
    """Attach the shape's alt text to its outermost element, for screen readers."""
    for tag in ("<g", "<image", "<path"):
        if fragment.startswith(tag):
            label = escape_xml_attr(alt_text)
            return f'{tag} role="img" aria-label="{label}"{fragment[len(tag):]}'
    return fragment


def render_group(group: m.GroupElement, context: RenderContext) -> str:
    """Map the group's child coordinate space onto its on-slide box.

    A group declares both where it sits (``a:off``/``a:ext``) and what coordinate system
    its children were authored in (``a:chOff``/``a:chExt``).  The composite transform is
    ``T(off) . R . F . S(ext/chExt) . T(-chOff)``, and nested groups compose outward-in.
    """
    x = emu_to_px(group.transform.offset_x)
    y = emu_to_px(group.transform.offset_y)
    w = emu_to_px(group.transform.extent_width)
    h = emu_to_px(group.transform.extent_height)
    child_w = emu_to_px(group.child_transform.extent_width)
    child_h = emu_to_px(group.child_transform.extent_height)
    child_x = emu_to_px(group.child_transform.offset_x)
    child_y = emu_to_px(group.child_transform.offset_y)

    # A zero child extent is malformed but does occur; fall back to identity scale
    # rather than dividing by zero or inventing an extent.
    scale_x = w / child_w if child_w else 1.0
    scale_y = h / child_h if child_h else 1.0

    parts = [f"translate({num(x)}, {num(y)})"]
    if group.transform.rotation:
        parts.append(f"rotate({num(group.transform.rotation)}, {num(w / 2)}, {num(h / 2)})")
    if group.transform.flip_h or group.transform.flip_v:
        parts.append(
            f"translate({num(w if group.transform.flip_h else 0)}, "
            f"{num(h if group.transform.flip_v else 0)})"
        )
        parts.append(
            f"scale({-1 if group.transform.flip_h else 1}, "
            f"{-1 if group.transform.flip_v else 1})"
        )
    parts.append(f"scale({num(scale_x)}, {num(scale_y)})")
    parts.append(f"translate({num(-child_x)}, {num(-child_y)})")

    children = "".join(
        rendered
        for rendered in (render_element(child, context) for child in group.children)
        if rendered
    )
    return f'<g transform="{" ".join(parts)}">{children}</g>'
