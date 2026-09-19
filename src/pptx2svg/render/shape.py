"""Shape, connector, image and table rendering.

Every element is wrapped in a ``<g>`` carrying its transform, so geometry is generated in
a local space where the shape occupies ``(0, 0)`` to ``(width, height)`` -- which is what
the preset generators and custom-geometry paths assume.
"""

from __future__ import annotations

from dataclasses import replace

from .. import model as m
from ..units import emu_to_px
from .context import RenderContext, num
from .effect import render_blip_effects, render_effects
from .fill import render_fill_attrs, render_markers, render_outline_attrs, tile_pattern
from .geometry import render_geometry
from .text import compute_sp_autofit_height, render_text_body


def build_transform_attr(transform: m.Transform) -> str:
    """Translate, then rotate about the centre, then flip -- PowerPoint's own order."""
    x = emu_to_px(transform.offset_x)
    y = emu_to_px(transform.offset_y)
    w = emu_to_px(transform.extent_width)
    h = emu_to_px(transform.extent_height)

    parts = [f"translate({num(x)}, {num(y)})"]

    if transform.rotation:
        parts.append(f"rotate({num(transform.rotation)}, {num(w / 2)}, {num(h / 2)})")

    if transform.flip_h or transform.flip_v:
        # Mirror about the shape's own centre by translating, scaling, and relying on the
        # translate to bring the mirrored box back over the original.
        parts.append(
            f"translate({num(w if transform.flip_h else 0)}, {num(h if transform.flip_v else 0)})"
        )
        parts.append(f"scale({-1 if transform.flip_h else 1}, {-1 if transform.flip_v else 1})")

    return " ".join(parts)


def _styled(geometry_svg: str, attrs: str) -> str:
    """Splice fill/stroke attributes into the element a geometry generator produced."""
    if not geometry_svg.startswith("<"):
        return geometry_svg
    tag_end = 1
    while tag_end < len(geometry_svg) and (
        geometry_svg[tag_end].isalnum() or geometry_svg[tag_end] == ":"
    ):
        tag_end += 1
    return f"{geometry_svg[:tag_end]} {attrs}{geometry_svg[tag_end:]}"


def _render_shape_text(
    shape: m.ShapeElement, transform: m.Transform, context: RenderContext
) -> str:
    """Lay the text out in its own box when the shape gives it one.

    ``dsp:txXfrm`` is SmartArt's way of saying "the label goes *here*, not across the
    middle of the shape".  The text renderer positions relative to the shape group's
    origin and only reads the extent, so the box's own offset has to be applied as a
    translation of the difference between the two -- both are absolute in the same space.
    """
    box = shape.text_transform or transform
    text_svg = render_text_body(shape.text_body, box, context)
    if not text_svg or shape.text_transform is None:
        return text_svg

    dx = emu_to_px(shape.text_transform.offset_x - transform.offset_x)
    dy = emu_to_px(shape.text_transform.offset_y - transform.offset_y)
    if not dx and not dy:
        return text_svg
    return f'<g transform="translate({num(dx)}, {num(dy)})">{text_svg}</g>'


def render_shape(shape: m.ShapeElement, context: RenderContext) -> str:
    transform = shape.transform

    # spAutofit grows the shape box to fit its text before anything is positioned.
    if shape.text_body is not None and shape.text_body.body_properties.auto_fit == "spAutofit":
        required = compute_sp_autofit_height(shape.text_body, transform, context)
        if required is not None:
            transform = replace(transform, extent_height=required)

    width = emu_to_px(transform.extent_width)
    height = emu_to_px(transform.extent_height)

    fill_attrs = render_fill_attrs(shape.fill, context, (0, 0, width, height))
    outline_attrs = render_outline_attrs(shape.outline, context)
    filter_attr = render_effects(shape.effects, context)

    parts = [f'<g transform="{build_transform_attr(transform)}"' + (f" {filter_attr}" if filter_attr else "") + ">"]

    geometry_svg = render_geometry(shape.geometry, width, height)
    if geometry_svg:
        parts.append(_styled(geometry_svg, f"{fill_attrs} {outline_attrs}"))

    if shape.text_body is not None:
        parts.append(_render_shape_text(shape, transform, context))

    parts.append("</g>")
    return "".join(parts)


def render_connector(connector: m.ConnectorElement, context: RenderContext) -> str:
    width = emu_to_px(connector.transform.extent_width)
    height = emu_to_px(connector.transform.extent_height)

    outline_attrs = render_outline_attrs(connector.outline, context)
    filter_attr = render_effects(connector.effects, context)
    marker_attrs = render_markers(connector.outline, context)
    marker_suffix = f" {marker_attrs}" if marker_attrs else ""

    parts = [
        f'<g transform="{build_transform_attr(connector.transform)}"'
        + (f" {filter_attr}" if filter_attr else "")
        + ">"
    ]

    geometry_svg = render_geometry(connector.geometry, width, height)
    if geometry_svg:
        parts.append(_styled(geometry_svg, f'{outline_attrs} fill="none"{marker_suffix}'))
    else:
        parts.append(
            f'<line x1="0" y1="0" x2="{num(width)}" y2="{num(height)}" '
            f'{outline_attrs} fill="none"{marker_suffix}/>'
        )

    parts.append("</g>")
    return "".join(parts)


def render_image(image: m.ImageElement, context: RenderContext) -> str:
    width = emu_to_px(image.transform.extent_width)
    height = emu_to_px(image.transform.extent_height)
    href = f"data:{image.mime_type};base64,{image.image_data}"

    filter_attr = render_effects(image.effects, context)
    blip_attr = render_blip_effects(image.blip_effects, context)

    attrs: list[str] = []
    clip_attr = ""

    # A non-rectangular picture frame clips the bitmap to the shape outline.
    if image.geometry is not None and not _is_plain_rect(image.geometry):
        clip_id = context.new_id("clip")
        context.add_def(
            f'<clipPath id="{clip_id}">{render_geometry(image.geometry, width, height)}</clipPath>'
        )
        clip_attr = f' clip-path="url(#{clip_id})"'

    inner: list[str] = []
    if image.src_rect is not None and _has_crop(image.src_rect):
        # `a:srcRect` crops the source; scale the image up so the kept part fills the frame.
        crop = image.src_rect
        kept_w = max(1e-6, 1 - crop.left - crop.right)
        kept_h = max(1e-6, 1 - crop.top - crop.bottom)
        scaled_w = width / kept_w
        scaled_h = height / kept_h
        offset_x = -crop.left * scaled_w
        offset_y = -crop.top * scaled_h
        crop_id = context.new_id("crop")
        context.add_def(
            f'<clipPath id="{crop_id}">'
            f'<rect width="{num(width)}" height="{num(height)}"/></clipPath>'
        )
        inner.append(
            f'<g clip-path="url(#{crop_id})">'
            f'<image href="{href}" x="{num(offset_x)}" y="{num(offset_y)}" '
            f'width="{num(scaled_w)}" height="{num(scaled_h)}" preserveAspectRatio="none"'
            + (f" {blip_attr}" if blip_attr else "")
            + "/></g>"
        )
    elif image.tile is not None and (
        tiled := tile_pattern(
            (tile_id := context.new_id("imgtile")),
            href,
            image.image_data,
            image.tile,
            (0, 0, width, height),
            blip_attr,
        )
    ):
        # `a:tile` repeats the bitmap instead of stretching it, at the picture's **own**
        # size scaled by `sx`/`sy` -- not at a fraction of the frame, which is what this
        # drew until the fill path was measured.  The two paths carry the identical
        # `a:tile`, so they now share one builder rather than agreeing by hand; and when
        # the builder cannot read the picture's natural size it returns nothing and this
        # falls through to drawing one stretched copy, which is the same wrong as an
        # untiled fill rather than a tiling at an invented pitch.
        context.add_def(tiled)
        inner.append(
            f'<rect width="{num(width)}" height="{num(height)}" fill="url(#{tile_id})"/>'
        )
    else:
        # `a:stretch/a:fillRect` insets the bitmap from the frame's edges as a fraction
        # of the frame; the usual all-zero rect means "fill it", and a negative inset
        # pushes the bitmap outside, which the frame's clip then trims.
        left = top = 0.0
        draw_width, draw_height = width, height
        stretch = image.stretch
        if stretch is not None and _has_crop(stretch):
            left = stretch.left * width
            top = stretch.top * height
            draw_width = max(1e-6, width * (1 - stretch.left - stretch.right))
            draw_height = max(1e-6, height * (1 - stretch.top - stretch.bottom))
        inner.append(
            f'<image href="{href}" x="{num(left)}" y="{num(top)}" '
            f'width="{num(draw_width)}" height="{num(draw_height)}" '
            'preserveAspectRatio="none"'
            + (f" {blip_attr}" if blip_attr else "")
            + "/>"
        )

    if image.outline is not None:
        outline_attrs = render_outline_attrs(image.outline, context)
        geometry = image.geometry or m.PresetGeometry(preset="rect")
        inner.append(_styled(render_geometry(geometry, width, height), f'fill="none" {outline_attrs}'))

    attrs_str = "".join(attrs)
    group_open = (
        f'<g transform="{build_transform_attr(image.transform)}"'
        + (f" {filter_attr}" if filter_attr else "")
        + clip_attr
        + attrs_str
        + ">"
    )
    return group_open + "".join(inner) + "</g>"


def _is_plain_rect(geometry: m.Geometry) -> bool:
    return isinstance(geometry, m.PresetGeometry) and geometry.preset in ("rect", "flowChartProcess")


def _has_crop(rect: m.SrcRect | m.StretchFillRect) -> bool:
    """Does this relative rect actually inset anything?"""
    return any(value for value in (rect.left, rect.top, rect.right, rect.bottom))


def render_table(table: m.TableElement, context: RenderContext) -> str:
    """Tables draw as a grid of cell rectangles, borders and text bodies.

    Merged cells are skipped (the spanning cell already covers their area) and the
    spanning cell is widened/heightened by its ``gridSpan``/``rowSpan``.
    """
    data = table.table
    parts = [f'<g transform="{build_transform_attr(table.transform)}">']

    column_offsets: list[float] = [0.0]
    for column in data.columns:
        column_offsets.append(column_offsets[-1] + emu_to_px(column.width))

    row_heights = _row_heights(data, context)
    row_offsets: list[float] = [0.0]
    for height in row_heights:
        row_offsets.append(row_offsets[-1] + emu_to_px(height))

    # Fills and text first, then every border on top, so a neighbour's fill cannot
    # paint over a shared edge.
    borders: list[str] = []

    for row_index, row in enumerate(data.rows):
        for column_index, cell in enumerate(row.cells):
            if cell.h_merge or cell.v_merge:
                continue
            if column_index >= len(column_offsets) - 1:
                continue

            x = column_offsets[column_index]
            y = row_offsets[row_index]
            end_column = min(column_index + cell.grid_span, len(column_offsets) - 1)
            end_row = min(row_index + cell.row_span, len(row_offsets) - 1)
            width = column_offsets[end_column] - x
            height = row_offsets[end_row] - y

            if cell.fill is not None:
                fill_attrs = render_fill_attrs(cell.fill, context, (x, y, width, height))
                parts.append(
                    f'<rect x="{num(x)}" y="{num(y)}" width="{num(width)}" '
                    f'height="{num(height)}" {fill_attrs}/>'
                )

            if cell.borders is not None:
                borders.extend(_cell_borders(cell.borders, x, y, width, height, context))

            if cell.text_body is not None:
                cell_transform = m.Transform(
                    offset_x=0,
                    offset_y=0,
                    extent_width=table.table.columns[column_index].width * max(1, cell.grid_span),
                    extent_height=sum(
                        row_heights[row_index: row_index + max(1, cell.row_span)]
                    ),
                )
                text_svg = render_text_body(cell.text_body, cell_transform, context)
                if text_svg:
                    parts.append(
                        f'<g transform="translate({num(x)}, {num(y)})">{text_svg}</g>'
                    )

    parts.extend(borders)
    parts.append("</g>")
    return "".join(parts)


def _row_heights(data: m.TableData, context: RenderContext) -> list[float]:
    """Row heights in EMU, grown to fit their text.

    ``a:tr@h`` is a *minimum*: PowerPoint makes a row taller when its text needs the
    space, and everything below it moves down.  Taking the attribute as exact leaves the
    gridlines of a text-heavy table drifting further out of place with every row.

    Only cells that occupy a single row get a vote.  A cell spanning several rows has no
    one row to grow, and guessing how to share its height between them would do more
    harm than leaving it alone.
    """
    heights = [row.height for row in data.rows]
    for row_index, row in enumerate(data.rows):
        for column_index, cell in enumerate(row.cells):
            if cell.text_body is None or cell.h_merge or cell.v_merge:
                continue
            if cell.row_span > 1 or column_index >= len(data.columns):
                continue
            width = data.columns[column_index].width * max(1, cell.grid_span)
            required = compute_sp_autofit_height(
                cell.text_body,
                m.Transform(extent_width=width, extent_height=heights[row_index]),
                context,
            )
            if required is not None:
                heights[row_index] = required
    return heights


def _cell_borders(
    borders: m.CellBorders,
    x: float,
    y: float,
    width: float,
    height: float,
    context: RenderContext,
) -> list[str]:
    edges = (
        (borders.top, (x, y, x + width, y)),
        (borders.bottom, (x, y + height, x + width, y + height)),
        (borders.left, (x, y, x, y + height)),
        (borders.right, (x + width, y, x + width, y + height)),
    )
    lines: list[str] = []
    for outline, (x1, y1, x2, y2) in edges:
        if outline is None or outline.fill is None:
            continue
        attrs = render_outline_attrs(outline, context)
        lines.append(
            f'<line x1="{num(x1)}" y1="{num(y1)}" x2="{num(x2)}" y2="{num(y2)}" {attrs}/>'
        )
    return lines
