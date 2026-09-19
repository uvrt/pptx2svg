"""Slide-to-SVG assembly.

Output is SVG 1.1 with inline attributes only -- no CSS classes, no stylesheets.  The
usual rasterisation backends (librsvg via cairosvg, and resvg) do not apply CSS selectors
reliably, so every property is written as a presentation attribute.
"""

from __future__ import annotations

from dataclasses import replace

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
        attrs = render_fill_attrs(fill, context, (0, 0, width, height))
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
    elif isinstance(element, (m.GroupElement, m.ChartElement)):
        # A chart is a group that also carries its data: the resolver has already lowered
        # it to rectangles, lines and text in the frame's own coordinate space, so there
        # is nothing chart-specific left to draw.
        rendered = render_group(element, context)
    elif isinstance(element, m.TableElement):
        rendered = render_table(element, context)
    else:
        return ""

    if rendered:
        attributes: dict[str, str] = {}

        alt_text = getattr(element, "alt_text", None)
        if alt_text:
            attributes["role"] = "img"
            attributes["aria-label"] = alt_text

        # Identity travels with the drawing so a caller can map what it sees back to the
        # shape it came from -- which is what makes the SVG addressable by an editor.
        element_id = getattr(element, "element_id", None)
        if element_id:
            attributes["data-pptx-id"] = element_id
        element_path = getattr(element, "element_path", None)
        if element_path:
            attributes["data-pptx-path"] = element_path

        if attributes:
            rendered = _add_attrs(rendered, attributes)

    hyperlink = getattr(element, "hyperlink", None)
    if rendered and hyperlink is not None:
        rendered = f'<a href="{escape_xml_attr(hyperlink.url)}">{rendered}</a>'

    return rendered


def _add_attrs(fragment: str, attributes: dict[str, str]) -> str:
    """Splice attributes into a rendered fragment's outermost element.

    Elements are emitted as strings rather than built as a tree, so this matches on the
    opening tag.  Anything not in the list is returned untouched rather than corrupted.
    """
    for tag in ("<g", "<image", "<path", "<rect", "<text"):
        if fragment.startswith(tag) and fragment[len(tag):len(tag) + 1] in (" ", ">", "/"):
            written = "".join(
                f' {name}="{escape_xml_attr(value)}"' for name, value in attributes.items()
            )
            return f"{tag}{written}{fragment[len(tag):]}"
    return fragment


#: How far apart the two axis factors may be and still count as one uniform scale.
#:
#: Exact equality is the wrong test, and a real deck says so.  Google Slides writes
#: ``ext``/``chExt`` pairs that are uniform by intent and not by arithmetic: the meal
#: planning template has a group at 3.796875 across and 3.7968797 down, and another at
#: 0.75 against 0.7499996 -- the residue of rounding a scale into two EMU pairs.  Four
#: such groups would otherwise take the folding path on float noise alone, changing how
#: an unmistakably uniform group is drawn.  1e-4 is three orders of magnitude above the
#: 1.3e-6 worst case measured there, and still far below any deliberate stretch (the
#: narrowest in the corpus is 0.7006 against 1.8185).  At 1e-4 the two axes cannot
#: disagree by even a tenth of a pixel across a 1000 px group.
UNIFORM_SCALE_TOLERANCE = 1e-4


def _is_uniform(scale_x: float, scale_y: float) -> bool:
    return abs(scale_x - scale_y) <= UNIFORM_SCALE_TOLERANCE * max(abs(scale_x), abs(scale_y))


def swaps_group_axes(transform: m.Transform) -> bool:
    """Does an enclosing group's ``x`` factor land on this child's *height*?

    A group's ``ext``/``chExt`` ratio is **not** a matrix that multiplies into a rotated
    child's transform.  PowerPoint gives each factor to one of the child's own axes,
    choosing by which slide axis that axis currently lies nearer to, and then rotates the
    result -- so the composite is a rotated *rectangle*, never a shear.  Past the 45
    degree mark the child's width has turned more vertical than horizontal and the two
    factors change places.

    Measured on a 43-slide probe deck exported by PowerPoint 16.x, with the drawn
    vertices read out of the PDF rather than off a rasterisation
    (``tools/make_group_shear_probe.py``, deck ``sweep``).  A square inside a 4:1 group
    came out 4:1 and square-cornered at every one of 20 angles -- interior angle 90.000
    degrees throughout, so **there is no shear at any angle**, and the two side lengths
    were exactly 1x and 4x the authored extent with nothing in between.  Which side got
    the 4 flipped between an authored 44.9 (width) and an authored 45.0 (height), and
    flipped back between 134 and 135; 180 behaved as 0, 225 as 45, 270 as 90, 315 as 135,
    405 as 45, -30 as 150 and -60 as 120.  The 44/45 boundary did not move when the ratio
    was changed to 2:1 or 10:1, so it is a property of the angle alone.  A 4:2 group put
    4 and 2 on the child's own axes the same way, and a non-square child was treated the
    same as a square one.

    **A first reading that the measurements refuted.**  Since ``|sin|`` and ``|cos|``
    cross at exactly 45 degrees, the obvious rule is "swap whenever ``|sin| > |cos|``",
    with flips irrelevant because a flip is a sign and this test takes magnitudes.  The
    ties say otherwise: at exactly 45 degrees an unflipped child swaps but one carrying
    ``flipH`` *or* ``flipV`` does not, and at exactly 135 it is the other way round.  One
    flip negates the angle the test is made on; two flips cancel and it does not.  Away
    from the two tie angles the two readings agree exactly, which is why 26 flipped
    probes at 30 and 60 degrees look like the flip does nothing.
    """
    angle = -transform.rotation if transform.flip_h != transform.flip_v else transform.rotation
    return 45.0 <= angle % 180.0 < 135.0


def _folded_child(
    element: m.SlideElement, scale_x: float, scale_y: float, child_x: float, child_y: float
) -> m.SlideElement:
    """The child with its enclosing group's coordinate scale baked into its own box.

    The centre moves by the group's scale on the slide axes; the extents grow by the same
    two factors assigned to the child's *own* axes, swapping when
    :func:`swaps_group_axes` says the child has turned far enough.  Rotation and flips are
    left alone -- they are applied to the grown box, after the scale, which is the whole
    point.
    """
    return replace(element, transform=_folded_transform(
        element.transform, scale_x, scale_y, child_x, child_y
    ))


def _folded_transform(
    transform: m.Transform, scale_x: float, scale_y: float, child_x: float, child_y: float
) -> m.Transform:
    width_scale, height_scale = (
        (scale_y, scale_x) if swaps_group_axes(transform) else (scale_x, scale_y)
    )
    width = transform.extent_width * width_scale
    height = transform.extent_height * height_scale
    centre_x = (transform.offset_x + transform.extent_width / 2 - child_x) * scale_x
    centre_y = (transform.offset_y + transform.extent_height / 2 - child_y) * scale_y
    return replace(
        transform,
        offset_x=centre_x - width / 2,
        offset_y=centre_y - height / 2,
        extent_width=width,
        extent_height=height,
    )


def render_group(group: "m.GroupElement | m.ChartElement", context: RenderContext) -> str:
    """Map the group's child coordinate space onto its on-slide box.

    A group declares both where it sits (``a:off``/``a:ext``) and what coordinate system
    its children were authored in (``a:chOff``/``a:chExt``).  The group's own placement is
    ``T(off) . R . F``, and nested groups compose outward-in.

    How the ``ext``/``chExt`` ratio itself reaches the children depends on whether it is
    uniform:

    * **Uniform.** It is emitted as an SVG ``scale()`` around the whole subtree, because a
      uniform scale commutes with rotation and the result is exact.  The factor is
      recorded on the context so the text renderer can undo it for the glyphs, which
      PowerPoint does not scale; see :attr:`RenderContext.group_scale`.
    * **Non-uniform.** A ``scale()`` around a *rotated* child would compose to a shear,
      and PowerPoint never draws one -- see :func:`swaps_group_axes`.  So the ratio is
      folded into each child's own box instead, leaving the child free to rotate the
      grown box afterwards.  Text then needs no counter-transform at all: the frame is
      already in the enclosing space and the point size was never touched.

    The one child that cannot be folded is a table, whose column widths and row heights
    are authored in EMU and do not follow its frame; it keeps the ``scale()`` wrapper, and
    with it the pre-existing approximation for a rotated table inside a stretched group.
    Anything else that turns out to have no transform to fold into keeps it too, rather
    than failing where it used to draw.
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

    if _is_uniform(scale_x, scale_y):
        parts.append(f"scale({num(scale_x)}, {num(scale_y)})")
        parts.append(f"translate({num(-child_x)}, {num(-child_y)})")
        children = _render_scaled(group.children, scale_x, scale_y, context)
        return f'<g transform="{" ".join(parts)}">{children}</g>'

    wrapper = f"scale({num(scale_x)}, {num(scale_y)}) translate({num(-child_x)}, {num(-child_y)})"
    rendered: list[str] = []
    offset_x = group.child_transform.offset_x
    offset_y = group.child_transform.offset_y
    for child in group.children:
        if isinstance(child, m.TableElement) or getattr(child, "transform", None) is None:
            inner = _render_scaled([child], scale_x, scale_y, context)
            if inner:
                rendered.append(f'<g transform="{wrapper}">{inner}</g>')
            continue
        folded = _folded_child(child, scale_x, scale_y, offset_x, offset_y)
        text_transform = getattr(child, "text_transform", None)
        if text_transform is not None:
            folded = replace(folded, text_transform=_folded_transform(
                text_transform, scale_x, scale_y, offset_x, offset_y
            ))
        drawn = render_element(folded, context)
        if drawn:
            rendered.append(drawn)
    return f'<g transform="{" ".join(parts)}">{"".join(rendered)}</g>'


def _render_scaled(
    children: "list[m.SlideElement]", scale_x: float, scale_y: float, context: RenderContext
) -> str:
    outer_scale = context.group_scale
    context.group_scale = (outer_scale[0] * scale_x, outer_scale[1] * scale_y)
    try:
        return "".join(
            rendered
            for rendered in (render_element(child, context) for child in children)
            if rendered
        )
    finally:
        context.group_scale = outer_scale
