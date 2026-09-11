"""The resolver: source model in, render model out.

This is where PowerPoint's inheritance actually happens.  Five separate cascades meet
here, and they are independent of each other:

* **Element list** -- master shapes, then layout shapes, then slide shapes, in that
  z-order.  Template placeholders are dropped (the slide's own copy supersedes them) and
  empty slide placeholders are dropped (an unfilled "Click to add title" draws nothing).
* **Placeholder properties** -- a slide placeholder with no transform or geometry of its
  own borrows them from the matching layout shape, then the matching master shape.
* **Background** -- slide, else layout, else master; a ``p:bgRef`` indexes the theme's
  background fill list.
* **Shape formatting** -- explicit ``a:spPr`` fill/line/effects, else the theme format
  scheme entry named by ``a:style``'s ``fillRef``/``lnRef``/``effectRef``.
* **Text** -- run properties fall back through the paragraph's ``defRPr``, the shape's
  ``lstStyle``, the layout and master placeholder ``lstStyle``s, the master's
  ``txStyles`` for the placeholder's type, and finally the presentation's
  ``defaultTextStyle``.

Charts, SmartArt and EMF/WMF metafiles are out of scope: they resolve to a positioned
placeholder and a warning rather than being dropped silently.
"""

from __future__ import annotations

import base64
from dataclasses import dataclass, field, replace
from typing import Callable, Iterable, Sequence

from .. import model as m
from ..opc import OpcPackage
from ..parse import source as s
from ..units import ROTATION_UNIT
from .color import ColorContext, build_effective_color_map, resolve_color

#: Placeholder types that inherit from the master's ``body`` placeholder.
BODY_PLACEHOLDER_TYPES = frozenset(
    {"body", "subTitle", "obj", "chart", "clipArt", "dgm", "media", "pic", "tbl"}
)

#: Default line width when ``a:ln`` gives none (1 pt).
DEFAULT_LINE_WIDTH_EMU = 12700

MIME_BY_EXTENSION = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tiff": "image/tiff",
    "tif": "image/tiff",
    "svg": "image/svg+xml",
    "webp": "image/webp",
    "emf": "image/emf",
    "wmf": "image/wmf",
}

SUPPORTED_IMAGE_MIME_TYPES = frozenset(
    {
        "image/png",
        "image/jpeg",
        "image/gif",
        "image/bmp",
        "image/tiff",
        "image/svg+xml",
        "image/webp",
        "image/x-icon",
    }
)


@dataclass
class Warning:
    code: str
    message: str
    slide_number: int | None = None
    part_path: str | None = None

    def __str__(self) -> str:
        where = f" (slide {self.slide_number})" if self.slide_number else ""
        return f"[{self.code}]{where} {self.message}"


@dataclass
class ResolveContext:
    """Per-slide resolution state."""

    package: OpcPackage
    presentation: s.SourcePresentation
    slide: s.SourceSlide
    layout: s.SourceSlideLayout | None
    master: s.SourceSlideMaster | None
    theme: s.SourceTheme | None
    colors: ColorContext
    warnings: list[Warning] = field(default_factory=list)
    #: Part whose relationships resolve the element currently being processed.
    part_path: str = ""
    #: Fill inherited by ``a:grpFill`` children, innermost last.
    group_fill_stack: list[m.Fill] = field(default_factory=list)

    def warn(self, code: str, message: str) -> None:
        self.warnings.append(
            Warning(
                code=code,
                message=message,
                slide_number=self.slide.slide_number,
                part_path=self.part_path,
            )
        )


@dataclass
class ResolvedPresentation:
    slide_size: m.SlideSize
    slides: list[m.Slide]
    warnings: list[Warning] = field(default_factory=list)
    font_scheme: m.FontScheme = field(default_factory=m.FontScheme)


def resolve_presentation(
    package: OpcPackage,
    presentation: s.SourcePresentation,
    *,
    slide_numbers: Iterable[int] | None = None,
) -> ResolvedPresentation:
    wanted = set(slide_numbers) if slide_numbers is not None else None
    slide_size = m.SlideSize(width=presentation.slide_width, height=presentation.slide_height)
    warnings: list[Warning] = []
    slides: list[m.Slide] = []
    font_scheme = m.FontScheme()

    for source_slide in presentation.slides:
        if wanted is not None and source_slide.slide_number not in wanted:
            continue
        context = _build_context(package, presentation, source_slide)
        slides.append(resolve_slide(context))
        warnings.extend(context.warnings)
        if context.theme is not None:
            font_scheme = _font_scheme(context.theme)

    return ResolvedPresentation(
        slide_size=slide_size, slides=slides, warnings=warnings, font_scheme=font_scheme
    )


def _build_context(
    package: OpcPackage, presentation: s.SourcePresentation, slide: s.SourceSlide
) -> ResolveContext:
    layout = presentation.layouts.get(slide.layout_part_path or "")
    master = presentation.masters.get(layout.master_part_path or "") if layout else None
    theme = presentation.themes.get(master.theme_part_path or "") if master else None
    color_map = build_effective_color_map(
        master.color_map if master else None,
        layout.color_map_override if layout else None,
        slide.color_map_override,
    )
    return ResolveContext(
        package=package,
        presentation=presentation,
        slide=slide,
        layout=layout,
        master=master,
        theme=theme,
        colors=ColorContext(theme, color_map),
        part_path=slide.part_path,
    )


def _font_scheme(theme: s.SourceTheme) -> m.FontScheme:
    scheme = theme.font_scheme
    return m.FontScheme(
        major_font=scheme.major_latin or "Calibri Light",
        minor_font=scheme.minor_latin or "Calibri",
        major_font_ea=scheme.major_east_asian or None,
        minor_font_ea=scheme.minor_east_asian or None,
        major_font_cs=scheme.major_complex_script or None,
        minor_font_cs=scheme.minor_complex_script or None,
        major_font_jpan=scheme.major_japanese or None,
        minor_font_jpan=scheme.minor_japanese or None,
    )


# --------------------------------------------------------------------------------------
# Slide
# --------------------------------------------------------------------------------------


def resolve_slide(context: ResolveContext) -> m.Slide:
    slide, layout, master = context.slide, context.layout, context.master

    show_master = slide.show_master_shapes and (layout.show_master_shapes if layout else True)

    elements: list[m.SlideElement] = []
    if show_master and master is not None:
        elements.extend(_resolve_template_elements(context, master.shapes, master.part_path))
    if layout is not None:
        elements.extend(_resolve_template_elements(context, layout.shapes, layout.part_path))
    context.part_path = slide.part_path
    elements.extend(_resolve_slide_elements(context, slide.shapes, slide.part_path))

    return m.Slide(
        slide_number=slide.slide_number,
        background=_resolve_background(context),
        elements=elements,
        show_master_sp=show_master,
    )


def _resolve_template_elements(
    context: ResolveContext, shapes: Sequence[s.SourceShapeNode], part_path: str
) -> list[m.SlideElement]:
    """Master/layout shapes minus placeholders -- the slide supplies its own copies."""
    context.part_path = part_path
    resolved: list[m.SlideElement] = []
    for shape in shapes:
        if _node_placeholder(shape) is not None:
            continue
        element = resolve_element(context, shape)
        if element is not None:
            resolved.append(element)
    return resolved


def _resolve_slide_elements(
    context: ResolveContext, shapes: Sequence[s.SourceShapeNode], part_path: str
) -> list[m.SlideElement]:
    context.part_path = part_path
    resolved: list[m.SlideElement] = []
    for shape in shapes:
        if isinstance(shape, s.SourceShape) and _is_empty_placeholder(shape):
            continue
        element = resolve_element(context, shape)
        if element is not None:
            resolved.append(element)
    return resolved


def _is_empty_placeholder(shape: s.SourceShape) -> bool:
    """An unfilled placeholder ("Click to add title") must not draw its prompt text."""
    if shape.placeholder is None:
        return False
    body = shape.text_body
    if body is None or not body.paragraphs:
        return True
    return not any(run.text for para in body.paragraphs for run in para.runs)


def _resolve_background(context: ResolveContext) -> m.Background | None:
    """Background falls back slide -> layout -> master."""
    for source in (
        context.slide.background,
        context.layout.background if context.layout else None,
        context.master.background if context.master else None,
    ):
        if source is None:
            continue
        if source.fill is not None:
            return m.Background(fill=_resolve_fill(context, source.fill))
        if source.bg_ref is not None:
            return m.Background(fill=_resolve_fill_reference(context, source.bg_ref))
    return None


# --------------------------------------------------------------------------------------
# Elements
# --------------------------------------------------------------------------------------


def resolve_element(context: ResolveContext, node: s.SourceShapeNode) -> m.SlideElement | None:
    if isinstance(node, s.SourceShape):
        return _resolve_shape(context, node)
    if isinstance(node, s.SourceConnector):
        return _resolve_connector(context, node)
    if isinstance(node, s.SourceImage):
        return _resolve_image(context, node)
    if isinstance(node, s.SourceGroup):
        return _resolve_group(context, node)
    if isinstance(node, s.SourceTable):
        return _resolve_table(context, node)
    if isinstance(node, s.SourceUnsupported):
        return _resolve_unsupported(context, node)
    return None


def _resolve_shape(context: ResolveContext, shape: s.SourceShape) -> m.ShapeElement:
    layout_node, master_node = _placeholder_match(context, shape)
    layout_shape = layout_node if isinstance(layout_node, s.SourceShape) else None
    master_shape = master_node if isinstance(master_node, s.SourceShape) else None

    transform = _first(
        shape.transform, _node_transform(layout_node), _node_transform(master_node)
    )
    geometry = _first(
        shape.geometry,
        layout_shape.geometry if layout_shape else None,
        master_shape.geometry if master_shape else None,
    )

    placeholder_type = (
        (shape.placeholder.type or "obj") if shape.placeholder is not None else None
    )

    text_body = None
    if shape.text_body is not None:
        text_body = _resolve_text_body(
            context,
            shape.text_body,
            inherited=[
                layout_shape.text_body if layout_shape else None,
                master_shape.text_body if master_shape else None,
            ],
            placeholder_type=placeholder_type,
        )

    return m.ShapeElement(
        transform=_resolve_transform(context, transform),
        geometry=_resolve_geometry(geometry),
        fill=_resolve_shape_fill(context, shape.fill, shape.style),
        outline=_resolve_shape_outline(context, shape.outline, shape.style),
        text_body=text_body,
        effects=_resolve_shape_effects(context, shape.effects, shape.style),
        placeholder_type=placeholder_type,
        placeholder_idx=shape.placeholder.idx if shape.placeholder else None,
        alt_text=shape.alt_text or shape.name,
        hyperlink=_resolve_hyperlink(context, shape.hyperlink_rel_id),
    )


def _resolve_connector(context: ResolveContext, connector: s.SourceConnector) -> m.ConnectorElement:
    return m.ConnectorElement(
        transform=_resolve_transform(context, connector.transform),
        geometry=_resolve_geometry(connector.geometry),
        outline=_resolve_shape_outline(context, connector.outline, connector.style),
        effects=_resolve_shape_effects(context, connector.effects, connector.style),
        alt_text=connector.alt_text or connector.name,
    )


def _resolve_group(context: ResolveContext, group: s.SourceGroup) -> m.GroupElement:
    transform = _resolve_transform(context, group.transform)
    child_transform = (
        _resolve_transform(context, group.child_transform)
        if group.child_transform is not None
        else replace(transform)
    )

    group_fill = _resolve_fill(context, group.fill) if group.fill is not None else None
    if group_fill is not None:
        context.group_fill_stack.append(group_fill)
    try:
        children = [
            element
            for element in (resolve_element(context, child) for child in group.children)
            if element is not None
        ]
    finally:
        if group_fill is not None:
            context.group_fill_stack.pop()

    return m.GroupElement(
        transform=transform,
        child_transform=child_transform,
        children=children,
        effects=_resolve_effects(context, group.effects),
        alt_text=group.alt_text or group.name,
    )


def _resolve_image(context: ResolveContext, image: s.SourceImage) -> m.SlideElement | None:
    layout_node, master_node = _placeholder_match(context, image)
    transform = _first(
        image.transform, _node_transform(layout_node), _node_transform(master_node)
    )

    media = _load_media(context, image.blip_relationship_id)
    if media is None:
        context.warn(
            "unresolved-image",
            f"picture {image.name or image.shape_id!r} has no readable image part",
        )
        return None

    data, mime_type = media
    if mime_type in ("image/emf", "image/wmf"):
        context.warn(
            "metafile-image",
            f"picture {image.name or image.shape_id!r} is an EMF/WMF metafile, "
            "which is not rasterised; drawing a placeholder",
        )
        return m.ShapeElement(
            transform=_resolve_transform(context, transform),
            geometry=m.PresetGeometry(preset="rect"),
            fill=m.SolidFill(color=m.ResolvedColor(hex="#e0e0e0")),
            alt_text=image.alt_text or image.name,
        )

    return m.ImageElement(
        transform=_resolve_transform(context, transform),
        image_data=data,
        mime_type=mime_type,
        effects=_resolve_effects(context, image.effects),
        blip_effects=_resolve_blip_effects(context, image.blip_effects),
        src_rect=_rect(image.src_rect, m.SrcRect),
        alt_text=image.alt_text or image.name,
        stretch=_rect(image.stretch, m.StretchFillRect),
        tile=_tile(image.tile),
        hyperlink=_resolve_hyperlink(context, image.hyperlink_rel_id),
        geometry=_resolve_geometry(image.geometry) if image.geometry is not None else None,
        outline=_resolve_outline(context, image.outline),
    )


def _resolve_table(context: ResolveContext, table: s.SourceTable) -> m.TableElement:
    rows = [
        m.TableRow(
            height=row.height,
            cells=[_resolve_table_cell(context, cell) for cell in row.cells],
        )
        for row in table.rows
    ]
    return m.TableElement(
        transform=_resolve_transform(context, table.transform),
        table=m.TableData(
            rows=rows, columns=[m.TableColumn(width=width) for width in table.columns]
        ),
        alt_text=table.alt_text or table.name,
    )


def _resolve_table_cell(context: ResolveContext, cell: s.SourceTableCell) -> m.TableCell:
    text_body = None
    if cell.text_body is not None:
        text_body = _resolve_text_body(context, cell.text_body, inherited=[], placeholder_type=None)
        # Cell margins and anchor live on `a:tcPr`, not on the cell's `a:bodyPr`.
        body_properties = text_body.body_properties
        if cell.margin_left is not None:
            body_properties.margin_left = cell.margin_left
        if cell.margin_right is not None:
            body_properties.margin_right = cell.margin_right
        if cell.margin_top is not None:
            body_properties.margin_top = cell.margin_top
        if cell.margin_bottom is not None:
            body_properties.margin_bottom = cell.margin_bottom
        if cell.anchor is not None:
            body_properties.anchor = cell.anchor

    borders = m.CellBorders(
        top=_resolve_table_border(context, cell.border_top),
        bottom=_resolve_table_border(context, cell.border_bottom),
        left=_resolve_table_border(context, cell.border_left),
        right=_resolve_table_border(context, cell.border_right),
    )

    return m.TableCell(
        text_body=text_body,
        fill=_resolve_fill(context, cell.fill) if cell.fill is not None else None,
        borders=borders,
        grid_span=cell.grid_span,
        row_span=cell.row_span,
        h_merge=cell.h_merge,
        v_merge=cell.v_merge,
    )


def _resolve_table_border(
    context: ResolveContext, outline: s.SourceOutline | None
) -> m.Outline | None:
    border = _resolve_outline(context, outline)
    if border is None:
        return None
    if border.fill is None:
        border.fill = m.SolidFill(color=m.ResolvedColor(hex="#000000"))
    return border


def _resolve_unsupported(
    context: ResolveContext, node: s.SourceUnsupported
) -> m.SlideElement | None:
    """Charts / SmartArt / OLE: draw the embedded preview when there is one."""
    if node.fallback_rel_id is not None and node.what == "ole":
        media = _load_media(context, node.fallback_rel_id)
        if media is not None and media[1] in SUPPORTED_IMAGE_MIME_TYPES:
            return m.ImageElement(
                transform=_resolve_transform(context, node.transform),
                image_data=media[0],
                mime_type=media[1],
                alt_text=node.alt_text or node.name,
            )

    context.warn(
        "unsupported-graphic-frame",
        f"{node.what} {node.name or ''!r} is not rendered; drawing an empty frame",
    )
    return m.ShapeElement(
        transform=_resolve_transform(context, node.transform),
        geometry=m.PresetGeometry(preset="rect"),
        fill=None,
        outline=None,
        alt_text=node.alt_text or node.name,
    )


# --------------------------------------------------------------------------------------
# Placeholder matching
# --------------------------------------------------------------------------------------


def _node_placeholder(node: s.SourceShapeNode) -> s.SourcePlaceholder | None:
    return getattr(node, "placeholder", None)


def _node_transform(node: s.SourceShapeNode | None) -> s.SourceTransform | None:
    return getattr(node, "transform", None) if node is not None else None


def _placeholder_match(
    context: ResolveContext, shape: s.SourceShapeNode
) -> tuple[s.SourceShapeNode | None, s.SourceShapeNode | None]:
    """Find the layout and master shapes a slide placeholder inherits from.

    Matching is by ``idx`` against the layout (a unique match only -- an ambiguous one
    would inherit arbitrarily), then by placeholder *type* from layout to master, where
    every body-ish type collapses onto the master's single ``body`` placeholder.
    """
    placeholder = _node_placeholder(shape)
    if placeholder is None:
        return None, None

    layout_shapes = context.layout.shapes if context.layout else []
    index = placeholder.idx or 0
    candidates = [
        candidate
        for candidate in layout_shapes
        if (ph := _node_placeholder(candidate)) is not None and (ph.idx or 0) == index
    ]
    if len(candidates) != 1:
        return None, None
    layout_node = candidates[0]

    layout_type = (_node_placeholder(layout_node).type or "obj")  # type: ignore[union-attr]
    master_shapes = context.master.shapes if context.master else []
    master_candidates = [
        candidate
        for candidate in master_shapes
        if (ph := _node_placeholder(candidate)) is not None
        and _master_type_matches(ph.type or "obj", layout_type)
    ]
    master_node = master_candidates[0] if len(master_candidates) == 1 else None
    return layout_node, master_node


def _master_type_matches(master_type: str, layout_type: str) -> bool:
    if layout_type in ("title", "ctrTitle"):
        return master_type == "title"
    if layout_type in BODY_PLACEHOLDER_TYPES:
        return master_type == "body"
    return master_type == layout_type


# --------------------------------------------------------------------------------------
# Transform, geometry
# --------------------------------------------------------------------------------------


def _resolve_transform(
    context: ResolveContext, transform: s.SourceTransform | None
) -> m.Transform:
    if transform is None:
        context.warn("missing-transform", "element has no transform; using a zero-size fallback")
        return m.Transform()
    return m.Transform(
        offset_x=transform.offset_x,
        offset_y=transform.offset_y,
        extent_width=transform.width,
        extent_height=transform.height,
        rotation=transform.rotation / ROTATION_UNIT,
        flip_h=transform.flip_horizontal,
        flip_v=transform.flip_vertical,
    )


def _resolve_geometry(geometry: s.SourceGeometry | None) -> m.Geometry:
    if geometry is None:
        return m.PresetGeometry(preset="rect")
    if isinstance(geometry, s.SourceCustomGeometry):
        return m.CustomGeometry(paths=list(geometry.paths))
    return m.PresetGeometry(
        preset=geometry.preset, adjust_values=dict(geometry.adjust_values)
    )


# --------------------------------------------------------------------------------------
# Fill, outline, effects
# --------------------------------------------------------------------------------------


def _resolve_shape_fill(
    context: ResolveContext,
    fill: s.SourceFill | None,
    style: s.SourceShapeStyle | None,
) -> m.Fill | None:
    """Explicit ``a:spPr`` fill wins; otherwise the theme entry named by ``a:fillRef``."""
    if fill is not None:
        return _resolve_fill(context, fill)
    if style is not None and style.fill_ref is not None:
        return _resolve_fill_reference(context, style.fill_ref)
    return None


def _resolve_shape_outline(
    context: ResolveContext,
    outline: s.SourceOutline | None,
    style: s.SourceShapeStyle | None,
) -> m.Outline | None:
    resolved = _resolve_outline(context, outline)
    if resolved is not None and resolved.fill is not None:
        return resolved

    reference = _resolve_line_reference(context, style.line_ref) if style else None
    if reference is None:
        return resolved
    if resolved is None:
        return reference
    # A local `a:ln` that only sets width/dash still takes its colour from the theme.
    merged = replace(reference)
    if outline is not None:
        if outline.width is not None:
            merged.width = outline.width
        if outline.dash_style is not None:
            merged.dash_style = outline.dash_style
        if outline.custom_dash is not None:
            merged.custom_dash = list(outline.custom_dash)
        if outline.line_cap is not None:
            merged.line_cap = outline.line_cap
        if outline.line_join is not None:
            merged.line_join = outline.line_join
        if outline.head_end is not None:
            merged.head_end = outline.head_end
        if outline.tail_end is not None:
            merged.tail_end = outline.tail_end
    return merged


def _resolve_shape_effects(
    context: ResolveContext,
    effects: s.SourceEffectList | None,
    style: s.SourceShapeStyle | None,
) -> m.EffectList | None:
    if effects is not None:
        return _resolve_effects(context, effects)
    if style is not None and style.effect_ref is not None:
        return _resolve_effect_reference(context, style.effect_ref)
    return None


def _resolve_fill(context: ResolveContext, fill: s.SourceFill | None) -> m.Fill | None:
    if fill is None:
        return None

    if isinstance(fill, s.SourceNoFill):
        return m.NoFill()

    if isinstance(fill, s.SourceSolidFill):
        color = resolve_color(context.colors, fill.color)
        return m.SolidFill(color=color) if color is not None else None

    if isinstance(fill, s.SourceGradientFill):
        stops = []
        for stop in fill.stops:
            color = resolve_color(context.colors, stop.color)
            if color is not None:
                stops.append(m.GradientStop(position=stop.position, color=color))
        if not stops:
            return None
        return m.GradientFill(
            stops=stops,
            # OOXML measures gradient angle clockwise from the positive x axis in
            # 1/60000 degrees; SVG wants plain degrees in the same direction.
            angle=fill.angle / ROTATION_UNIT,
            gradient_type=fill.gradient_type,
            center_x=fill.center_x,
            center_y=fill.center_y,
        )

    if isinstance(fill, s.SourcePatternFill):
        foreground = resolve_color(context.colors, fill.foreground_color)
        background = resolve_color(context.colors, fill.background_color)
        if foreground is None or background is None:
            return None
        return m.PatternFill(
            preset=fill.preset, foreground_color=foreground, background_color=background
        )

    if isinstance(fill, s.SourceImageFill):
        media = _load_media(context, fill.blip_relationship_id)
        if media is None:
            return None
        data, mime_type = media
        if mime_type not in SUPPORTED_IMAGE_MIME_TYPES:
            context.warn("unsupported-fill-image", f"image fill of type {mime_type} not rendered")
            return None
        return m.ImageFill(image_data=data, mime_type=mime_type, tile=_tile(fill.tile))

    if isinstance(fill, s.SourceGroupFill):
        return context.group_fill_stack[-1] if context.group_fill_stack else None

    return None


def _resolve_outline(
    context: ResolveContext, outline: s.SourceOutline | None
) -> m.Outline | None:
    if outline is None:
        return None
    fill = _resolve_fill(context, outline.fill)
    if isinstance(fill, m.NoFill):
        return None
    return m.Outline(
        width=outline.width if outline.width is not None else DEFAULT_LINE_WIDTH_EMU,
        fill=fill if isinstance(fill, (m.SolidFill, m.GradientFill)) else None,
        dash_style=outline.dash_style or "solid",
        custom_dash=list(outline.custom_dash) if outline.custom_dash else None,
        line_cap=outline.line_cap,
        line_join=outline.line_join,
        head_end=outline.head_end,
        tail_end=outline.tail_end,
    )


def _resolve_fill_reference(
    context: ResolveContext, ref: s.SourceStyleReference
) -> m.Fill | None:
    """``a:fillRef``/``p:bgRef``: index into the theme format scheme, colour overridden."""
    if ref.idx == 0:
        return None
    scheme = context.theme.format_scheme if context.theme else None
    if scheme is None:
        return None

    # Indices >= 1000 address bgFillStyleLst, with 1000 meaning "no template".
    if ref.idx >= 1000:
        styles, array_index = scheme.bg_fill_styles, ref.idx - 1001
    else:
        styles, array_index = scheme.fill_styles, ref.idx - 1
    if array_index < 0 or array_index >= len(styles):
        return None

    resolved = _resolve_fill(context, styles[array_index])
    override = resolve_color(context.colors, ref.color)
    if resolved is None or override is None:
        return resolved

    if isinstance(resolved, m.SolidFill):
        return m.SolidFill(color=override)
    if isinstance(resolved, m.GradientFill):
        # The theme gradient keeps its stop positions and transforms; only the base
        # colour is replaced, so re-resolve each stop against the override.
        return m.GradientFill(
            stops=[
                m.GradientStop(position=stop.position, color=_blend_stop(stop.color, override))
                for stop in resolved.stops
            ],
            angle=resolved.angle,
            gradient_type=resolved.gradient_type,
            center_x=resolved.center_x,
            center_y=resolved.center_y,
        )
    return resolved


def _blend_stop(_original: m.ResolvedColor, override: m.ResolvedColor) -> m.ResolvedColor:
    """Theme gradient stops carry their own tint/shade; the override supplies the hue.

    The theme's stop colours are all the same scheme colour under different transforms,
    and those transforms were already baked in when the format scheme was resolved
    against the *theme's* colour, not the shape's.  Re-deriving them exactly would mean
    keeping the unresolved stop colours around; matching pptx-glimpse, the override
    simply replaces the stop.
    """
    return override


def _resolve_line_reference(
    context: ResolveContext, ref: s.SourceStyleReference | None
) -> m.Outline | None:
    if ref is None or ref.idx == 0:
        return None
    scheme = context.theme.format_scheme if context.theme else None
    if scheme is None or ref.idx - 1 >= len(scheme.line_styles) or ref.idx < 1:
        return None

    resolved = _resolve_outline(context, scheme.line_styles[ref.idx - 1])
    if resolved is None:
        return None
    override = resolve_color(context.colors, ref.color)
    if override is not None:
        resolved.fill = m.SolidFill(color=override)
    return resolved


def _resolve_effect_reference(
    context: ResolveContext, ref: s.SourceStyleReference
) -> m.EffectList | None:
    scheme = context.theme.format_scheme if context.theme else None
    if scheme is None or ref.idx >= len(scheme.effect_styles) or ref.idx < 0:
        return None
    return _resolve_effects(context, scheme.effect_styles[ref.idx])


def _resolve_effects(
    context: ResolveContext, effects: s.SourceEffectList | None
) -> m.EffectList | None:
    if effects is None:
        return None

    resolved = m.EffectList()
    if effects.outer_shadow is not None:
        color = resolve_color(context.colors, effects.outer_shadow.color)
        if color is not None:
            resolved.outer_shadow = m.OuterShadow(
                blur_radius=effects.outer_shadow.blur_radius,
                distance=effects.outer_shadow.distance,
                direction=effects.outer_shadow.direction / ROTATION_UNIT,
                color=color,
                alignment=effects.outer_shadow.alignment,
                rotate_with_shape=effects.outer_shadow.rotate_with_shape,
            )
    if effects.inner_shadow is not None:
        color = resolve_color(context.colors, effects.inner_shadow.color)
        if color is not None:
            resolved.inner_shadow = m.InnerShadow(
                blur_radius=effects.inner_shadow.blur_radius,
                distance=effects.inner_shadow.distance,
                direction=effects.inner_shadow.direction / ROTATION_UNIT,
                color=color,
            )
    if effects.glow is not None:
        color = resolve_color(context.colors, effects.glow.color)
        if color is not None:
            resolved.glow = m.Glow(radius=effects.glow.radius, color=color)
    if effects.soft_edge is not None:
        resolved.soft_edge = m.SoftEdge(radius=effects.soft_edge.radius)

    return None if resolved.is_empty() else resolved


def _resolve_blip_effects(
    context: ResolveContext, effects: s.SourceBlipEffects | None
) -> m.BlipEffects | None:
    if effects is None:
        return None
    duotone = None
    if effects.duotone is not None:
        first = resolve_color(context.colors, effects.duotone[0])
        second = resolve_color(context.colors, effects.duotone[1])
        if first is not None and second is not None:
            duotone = m.DuotoneEffect(color1=first, color2=second)
    change = None
    if effects.clr_change is not None:
        source = resolve_color(context.colors, effects.clr_change[0])
        target = resolve_color(context.colors, effects.clr_change[1])
        if source is not None and target is not None:
            change = m.ClrChangeEffect(clr_from=source, clr_to=target)

    return m.BlipEffects(
        grayscale=effects.grayscale,
        bi_level=m.BiLevelEffect(threshold=effects.bi_level)
        if effects.bi_level is not None
        else None,
        blur=m.BlurEffect(radius=effects.blur[0], grow=effects.blur[1])
        if effects.blur is not None
        else None,
        lum=m.LumEffect(brightness=effects.lum[0], contrast=effects.lum[1])
        if effects.lum is not None
        else None,
        duotone=duotone,
        clr_change=change,
    )


# --------------------------------------------------------------------------------------
# Media and hyperlinks
# --------------------------------------------------------------------------------------


def _load_media(context: ResolveContext, rel_id: str | None) -> tuple[str, str] | None:
    """Relationship id -> (base64 payload, MIME type)."""
    if rel_id is None:
        return None
    target = context.package.related_part(context.part_path, rel_id)
    if target is None:
        return None
    payload = context.package.read(target)
    if payload is None:
        return None
    mime_type = context.package.content_type(target)
    if mime_type is None:
        extension = target.rsplit(".", 1)[-1].lower()
        mime_type = MIME_BY_EXTENSION.get(extension, "image/png")
    return base64.b64encode(payload).decode("ascii"), mime_type


def _resolve_hyperlink(context: ResolveContext, rel_id: str | None) -> m.Hyperlink | None:
    if rel_id is None:
        return None
    relationship = context.package.relationships(context.part_path).get(rel_id)
    if relationship is None:
        return None
    return m.Hyperlink(url=relationship.target)


def _rect(values: tuple[float, float, float, float] | None, factory: Callable):
    if values is None:
        return None
    left, top, right, bottom = values
    return factory(left=left, top=top, right=right, bottom=bottom)


def _tile(tile: s.SourceImageFillTile | None):
    if tile is None:
        return None
    return m.ImageFillTile(
        tx=tile.tx, ty=tile.ty, sx=tile.sx, sy=tile.sy, flip=tile.flip, align=tile.align
    )


def _first(*values):
    for value in values:
        if value is not None:
            return value
    return None


# Text resolution lives in its own module to keep this one readable.
from .text import resolve_text_body as _resolve_text_body  # noqa: E402
