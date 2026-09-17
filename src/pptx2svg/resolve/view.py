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

SmartArt and EMF/WMF metafiles are handled here too, in both cases by finding the
pre-rendered copy PowerPoint already stored rather than reimplementing what made it: a
diagram's cached DrawingML shape tree, and a metafile's embedded PDF or bitmap preview.
Charts have no such cache and remain out of scope: they resolve to a positioned
placeholder and a warning rather than being dropped silently.
"""

from __future__ import annotations

import base64
from dataclasses import dataclass, field, replace
from typing import Callable, Iterable, Sequence

from .. import model as m
from ..fonts.embedded import NO_EMBEDDED_FONTS, EmbeddedFonts
from ..metafile import extract_metafile_preview
from ..metafile.pdf import PdfRasterizerNotAvailable, rasterise_pdf
from ..opc import OpcPackage
from ..parse import source as s
from ..parse.chart import flat_chart_kind, is_three_d_kind, parse_chart_space
from ..parse.drawing import parse_group_transforms
from ..parse.shapes import parse_shape_tree
from ..parse.table_styles_builtin import builtin_table_style
from ..text.fontmap import east_asian_family
from ..units import ROTATION_UNIT
from ..xmlutil import attr, child, descendants
from .chart import (
    CHART_TEXT_BODY,
    EMU_PER_POINT,
    ChartBuilder,
    ChartStyle,
    accent_colors,
    default_font_size,
)
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

METAFILE_MIME_TYPES = frozenset({"image/emf", "image/wmf", "image/x-emf", "image/x-wmf"})

#: Relationship type from a SmartArt data-model part to its cached DrawingML rendering.
#: Two spellings exist for the same relationship -- Microsoft's own and the ISO/IEC
#: transitional one that ``purl.oclc.org`` hosts -- and which one appears depends on
#: which Office version and which save format wrote the file, so both are accepted.
#: Chart groups the renderer can draw.  Everything else warns and draws an empty frame
#: rather than a wrong picture.
DRAWABLE_CHART_KINDS = frozenset(
    {
        "barChart",
        "lineChart",
        "areaChart",
        "scatterChart",
        "bubbleChart",
        "pieChart",
        "doughnutChart",
        "ofPieChart",
        "radarChart",
        "stockChart",
    }
)

DIAGRAM_DRAWING_REL_TYPES = (
    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
    "http://purl.oclc.org/ooxml/officeDocument/relationships/diagramDrawing",
)

#: Why a SmartArt frame can come out blank through no fault of the file.  PowerPoint
#: caches a laid-out DrawingML copy of every diagram, and reading that cache is the whole
#: of our SmartArt support: the layout algorithms in ``dgm:layoutDef`` are a diagram
#: engine and a project in their own right.  Office 2007 did not always write the cache,
#: and later versions sometimes write an empty one, so a perfectly valid deck can carry a
#: diagram nothing here can draw.
NO_CACHED_DRAWING = (
    "has no cached DrawingML rendering.  PowerPoint caches a laid-out copy of every "
    "diagram and that copy is what we draw; Office 2007 did not always write one.  "
    "Laying the diagram out from its layout definition is not implemented, so the frame "
    "is left empty"
)

#: A user-supplied EMF/WMF converter: ``(bytes, mime_type) -> (bytes, mime_type) | None``.
#: Returning ``None`` means "I cannot convert this", and resolution falls through to the
#: built-in preview extraction.
MetafileConverter = Callable[[bytes, str], "tuple[bytes, str] | None"]


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
    #: Namespace for element ids of the part currently being resolved.  Slide shapes are
    #: addressable and editable; layout and master shapes are inherited decoration, and the
    #: prefix says so rather than leaving a caller to guess from the id alone.
    id_prefix: str = ""
    #: Optional external EMF/WMF converter; see :data:`MetafileConverter`.
    metafile_converter: MetafileConverter | None = None

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
    #: Faces the deck carried in ``<p:embeddedFontLst>``, decoded and rights-checked.
    #: Filled by :func:`pptx2svg.convert_pptx_to_model` rather than by
    #: :func:`resolve_presentation`, which is what lets the decode be limited to the
    #: families the resolved slides actually ask for.
    embedded_fonts: EmbeddedFonts = field(default_factory=lambda: NO_EMBEDDED_FONTS)


def resolve_presentation(
    package: OpcPackage,
    presentation: s.SourcePresentation,
    *,
    slide_numbers: Iterable[int] | None = None,
    metafile_converter: MetafileConverter | None = None,
) -> ResolvedPresentation:
    wanted = set(slide_numbers) if slide_numbers is not None else None
    slide_size = m.SlideSize(width=presentation.slide_width, height=presentation.slide_height)
    warnings: list[Warning] = []
    slides: list[m.Slide] = []
    font_scheme = m.FontScheme()

    for source_slide in presentation.slides:
        if wanted is not None and source_slide.slide_number not in wanted:
            continue
        context = _build_context(
            package, presentation, source_slide, metafile_converter=metafile_converter
        )
        slides.append(resolve_slide(context))
        warnings.extend(context.warnings)
        if context.theme is not None:
            font_scheme = _font_scheme(context.theme)

    return ResolvedPresentation(
        slide_size=slide_size, slides=slides, warnings=warnings, font_scheme=font_scheme
    )


def _build_context(
    package: OpcPackage,
    presentation: s.SourcePresentation,
    slide: s.SourceSlide,
    *,
    metafile_converter: MetafileConverter | None = None,
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
        metafile_converter=metafile_converter,
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
        context.id_prefix = "mst:"
        elements.extend(_resolve_template_elements(context, master.shapes, master.part_path))
    if layout is not None:
        context.id_prefix = "lay:"
        elements.extend(_resolve_template_elements(context, layout.shapes, layout.part_path))
    context.part_path = slide.part_path
    context.id_prefix = f"{slide.slide_id}." if slide.slide_id is not None else ""
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
    for index, shape in enumerate(shapes):
        if _node_placeholder(shape) is not None:
            continue
        element = resolve_element(context, shape, (index,))
        if element is not None:
            resolved.append(element)
    return resolved


def _resolve_slide_elements(
    context: ResolveContext, shapes: Sequence[s.SourceShapeNode], part_path: str
) -> list[m.SlideElement]:
    context.part_path = part_path
    resolved: list[m.SlideElement] = []
    for index, shape in enumerate(shapes):
        if isinstance(shape, s.SourceShape) and _is_empty_placeholder(shape):
            continue
        element = resolve_element(context, shape, (index,))
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


def resolve_element(
    context: ResolveContext,
    node: s.SourceShapeNode,
    path: tuple[int, ...] = (),
) -> m.SlideElement | None:
    """Resolve one shape node.

    Being the single dispatch point for every element at every nesting depth, this is also
    where identity is attached -- doing it here rather than in each ``_resolve_*`` keeps groups
    and their descendants consistent for free.
    """
    # `p:cNvPr@hidden` is PowerPoint's "hide" in the selection pane: the shape is still
    # in the file, with all its formatting, and simply is not drawn.
    if getattr(node, "hidden", False):
        return None
    if isinstance(node, s.SourceShape):
        element = _resolve_shape(context, node)
    elif isinstance(node, s.SourceConnector):
        element = _resolve_connector(context, node)
    elif isinstance(node, s.SourceImage):
        element = _resolve_image(context, node)
    elif isinstance(node, s.SourceGroup):
        element = _resolve_group(context, node, path)
    elif isinstance(node, s.SourceTable):
        element = _resolve_table(context, node)
    elif isinstance(node, s.SourceUnsupported):
        element = _resolve_unsupported(context, node, path)
    else:
        return None

    if element is not None:
        shape_id = getattr(node, "shape_id", None)
        element.element_id = f"{context.id_prefix}{shape_id}" if shape_id else None
        element.element_path = "-".join(str(step) for step in path) or None
    return element


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
        text_transform=(
            _resolve_transform(context, shape.text_transform)
            if shape.text_transform is not None
            else None
        ),
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


def _resolve_group(
    context: ResolveContext, group: s.SourceGroup, path: tuple[int, ...] = ()
) -> m.GroupElement:
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
            for element in (
                resolve_element(context, child, path + (index,))
                for index, child in enumerate(group.children)
            )
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

    media = _load_media_bytes(context, image.blip_relationship_id)
    if media is None:
        context.warn(
            "unresolved-image",
            f"picture {image.name or image.shape_id!r} has no readable image part",
        )
        return None

    payload, mime_type = media
    if mime_type in METAFILE_MIME_TYPES:
        preview = _metafile_preview(
            context,
            payload,
            mime_type,
            width_emu=transform.width if transform is not None else None,
            described_as=f"picture {image.name or image.shape_id!r}",
        )
        if preview is None:
            return m.ShapeElement(
                transform=_resolve_transform(context, transform),
                geometry=m.PresetGeometry(preset="rect"),
                fill=m.SolidFill(color=m.ResolvedColor(hex="#e0e0e0")),
                alt_text=image.alt_text or image.name,
            )
        payload, mime_type = preview

    data = base64.b64encode(payload).decode("ascii")
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
    style = _table_style(context, table)
    column_count = len(table.columns)
    row_count = len(table.rows)

    rows = [
        m.TableRow(
            height=row.height,
            cells=[
                _resolve_table_cell(
                    context,
                    cell,
                    _cell_style(style, table, row_index, column_index, cell,
                                row_count, column_count),
                )
                for column_index, cell in enumerate(row.cells)
            ],
        )
        for row_index, row in enumerate(table.rows)
    ]
    return m.TableElement(
        transform=_resolve_transform(context, table.transform),
        table=m.TableData(
            rows=rows, columns=[m.TableColumn(width=width) for width in table.columns]
        ),
        alt_text=table.alt_text or table.name,
    )


def _table_style(context: ResolveContext, table: s.SourceTable) -> s.SourceTableStyle | None:
    """Find the table's style: its own id, else the presentation's default.

    A deck's ``tableStyles.xml`` is consulted first, since a custom style there may reuse
    a built-in's GUID, and the built-in catalogue second.

    An id neither of them has renders unstyled, which is also what PowerPoint does with
    an id it does not recognise -- so the output is defensible rather than wrong.  What
    is wrong is doing it quietly: an unstyled table looks exactly like a table whose
    style genuinely carries no fills, borders or bold, so nobody can tell a deck that
    rendered correctly from one that lost every band and header rule.  That is the same
    failure mode ``font-substituted`` exists to close, and it gets the same treatment
    here.
    """
    styles = context.presentation.table_styles
    own_id = table.style_id
    style_id = own_id or (styles.default_style_id if styles else None)
    if not style_id:
        return None
    if styles is not None and style_id in styles.styles:
        return styles.styles[style_id]
    style = builtin_table_style(style_id)
    if style is None:
        named = table.alt_text or table.name or "a table"
        source = "names" if own_id else "inherits the presentation's default"
        context.warn(
            "table-style-unknown",
            f"{named} {source} table style {style_id}, which is in neither the deck's "
            "tableStyles.xml nor the built-in catalogue; drawing it with its own cell "
            "formatting only. PowerPoint draws an unrecognised id unstyled too, so the "
            "render is not wrong -- but if that id names a style PowerPoint does know, "
            "every fill, border and bold it carries is missing here.",
        )
    return style


def _band(index: int, first: bool, last: bool, count: int) -> int | None:
    """Which banding stripe a row or column falls in, or ``None`` for neither.

    The header and footer rows are outside the banding and do not advance it, so with
    ``firstRow`` set the first body row is band 1, not band 2.  Verified against
    PowerPoint's own render.
    """
    start = 1 if first else 0
    end = count - (1 if last else 0)
    if not start <= index < end:
        return None
    return (index - start) % 2


def _cell_regions(
    style: s.SourceTableStyle,
    table: s.SourceTable,
    row_index: int,
    column_index: int,
    cell: s.SourceTableCell,
    row_count: int,
    column_count: int,
) -> list[tuple[s.SourceTableCellStyle, tuple[int, int], tuple[int, int]]]:
    """Every style region covering this cell, lowest precedence first.

    Each entry carries the region's own row and column extent, because a region's four
    named borders apply at *its* boundary and its inside borders within it: ``firstRow``
    puts its ``bottom`` under the header rather than under the table.
    """
    last_row = row_index + max(1, cell.row_span) >= row_count
    last_col = column_index + max(1, cell.grid_span) >= column_count
    is_first_row = table.first_row and row_index == 0
    is_last_row = table.last_row and last_row
    is_first_col = table.first_col and column_index == 0
    is_last_col = table.last_col and last_col

    all_rows = (0, row_count - 1)
    all_cols = (0, column_count - 1)
    regions: list[tuple[s.SourceTableCellStyle | None, tuple[int, int], tuple[int, int]]] = [
        (style.whole_table, all_rows, all_cols)
    ]

    if table.band_col:
        band = _band(column_index, table.first_col, table.last_col, column_count)
        if band is not None:
            region = style.band1_v if band == 0 else style.band2_v
            regions.append((region, all_rows, (column_index, column_index)))
    if table.band_row:
        band = _band(row_index, table.first_row, table.last_row, row_count)
        if band is not None:
            region = style.band1_h if band == 0 else style.band2_h
            regions.append((region, (row_index, row_index), all_cols))

    if is_last_col:
        regions.append((style.last_col, all_rows, (column_count - 1, column_count - 1)))
    if is_first_col:
        regions.append((style.first_col, all_rows, (0, 0)))
    if is_last_row:
        regions.append((style.last_row, (row_count - 1, row_count - 1), all_cols))
    if is_first_row:
        regions.append((style.first_row, (0, 0), all_cols))

    corner = None
    if is_first_row and is_first_col:
        corner = style.nw_cell
    elif is_first_row and is_last_col:
        corner = style.ne_cell
    elif is_last_row and is_first_col:
        corner = style.sw_cell
    elif is_last_row and is_last_col:
        corner = style.se_cell
    if corner is not None:
        regions.append((corner, (row_index, row_index), (column_index, column_index)))

    return [(region, rows, cols) for region, rows, cols in regions if region is not None]


@dataclass
class _CellStyle:
    """A table style flattened onto one cell, ready to sit under its own formatting."""

    fill: s.SourceFill | None = None
    fill_ref: s.SourceStyleReference | None = None
    text: s.SourceRunProperties | None = None
    border_top: s.SourceOutline | None = None
    border_bottom: s.SourceOutline | None = None
    border_left: s.SourceOutline | None = None
    border_right: s.SourceOutline | None = None


def _cell_style(
    style: s.SourceTableStyle | None,
    table: s.SourceTable,
    row_index: int,
    column_index: int,
    cell: s.SourceTableCell,
    row_count: int,
    column_count: int,
) -> _CellStyle:
    merged = _CellStyle()
    if style is None or not column_count or not row_count:
        return merged

    row_end = row_index + max(1, cell.row_span) - 1
    column_end = column_index + max(1, cell.grid_span) - 1

    for region, (first_row, last_row), (first_col, last_col) in _cell_regions(
        style, table, row_index, column_index, cell, row_count, column_count
    ):
        if region.fill is not None:
            merged.fill = region.fill
            merged.fill_ref = None
        elif region.fill_ref is not None:
            merged.fill_ref = region.fill_ref
            merged.fill = None
        if region.text is not None:
            merged.text = _merge_run_properties(merged.text, region.text)

        for edge, at_boundary, outer, inner in (
            ("border_top", row_index == first_row, region.border_top, region.border_inside_h),
            ("border_bottom", row_end == last_row, region.border_bottom, region.border_inside_h),
            ("border_left", column_index == first_col, region.border_left, region.border_inside_v),
            ("border_right", column_end == last_col, region.border_right, region.border_inside_v),
        ):
            outline = outer if at_boundary else inner
            if outline is not None:
                setattr(merged, edge, outline)

    return merged


def _merge_run_properties(
    base: s.SourceRunProperties | None, over: s.SourceRunProperties
) -> s.SourceRunProperties:
    if base is None:
        return replace(over)
    merged = replace(base)
    for name, value in vars(over).items():
        if value is not None:
            setattr(merged, name, value)
    return merged


def _resolve_table_cell(
    context: ResolveContext,
    cell: s.SourceTableCell,
    style: _CellStyle | None = None,
) -> m.TableCell:
    style = style or _CellStyle()
    text_body = None
    if cell.text_body is not None:
        text_body = _resolve_text_body(
            context,
            cell.text_body,
            inherited=[],
            placeholder_type=None,
            extra_defaults=style.text,
        )
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

    # The cell's own `a:lnL`/`a:lnR`/... win outright; the style only fills the gaps.
    borders = m.CellBorders(
        top=_resolve_table_border(context, cell.border_top or style.border_top),
        bottom=_resolve_table_border(context, cell.border_bottom or style.border_bottom),
        left=_resolve_table_border(context, cell.border_left or style.border_left),
        right=_resolve_table_border(context, cell.border_right or style.border_right),
    )

    fill = cell.fill
    if fill is None and style.fill is not None:
        fill = style.fill
    resolved_fill = _resolve_fill(context, fill) if fill is not None else None
    if resolved_fill is None and cell.fill is None and style.fill_ref is not None:
        resolved_fill = _resolve_fill_reference(context, style.fill_ref)

    return m.TableCell(
        text_body=text_body,
        fill=resolved_fill,
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
    context: ResolveContext, node: s.SourceUnsupported, path: tuple[int, ...] = ()
) -> m.SlideElement | None:
    """Charts / SmartArt / OLE: draw the embedded preview when there is one."""
    if node.fallback_rel_id is not None and node.what == "ole":
        media = _load_media_bytes(context, node.fallback_rel_id)
        if media is not None:
            payload, mime_type = media
            if mime_type in METAFILE_MIME_TYPES:
                preview = _metafile_preview(
                    context,
                    payload,
                    mime_type,
                    width_emu=node.transform.width if node.transform else None,
                    described_as=f"OLE preview {node.name or ''!r}",
                )
                payload, mime_type = preview if preview is not None else (b"", "")
            if mime_type in SUPPORTED_IMAGE_MIME_TYPES:
                return m.ImageElement(
                    transform=_resolve_transform(context, node.transform),
                    image_data=base64.b64encode(payload).decode("ascii"),
                    mime_type=mime_type,
                    alt_text=node.alt_text or node.name,
                )

    if node.what == "chart":
        drawn = _resolve_chart(context, node)
        if drawn is not None:
            return drawn
        # _resolve_chart has already said which link failed; an empty frame plus a
        # second, vaguer warning would only bury it.
        return _empty_graphic_frame(context, node)

    if node.what == "chartex":
        # Nothing here reads `cx:chartSpace`, and it is a different format rather than a
        # missing case in the chart reader: a different namespace, a different data model
        # and no `c:*Chart` group anywhere in it.
        context.warn(
            "chart-unsupported-type",
            f"chart {node.name or node.shape_id or ''!r} is an Office 2016 chart "
            "(treemap, sunburst, histogram, box-and-whisker, waterfall, funnel or map), "
            "which is not rendered; drawing an empty frame",
        )
        return _empty_graphic_frame(context, node)

    if node.what == "diagram":
        diagram = _resolve_diagram(context, node, path)
        if diagram is not None:
            return diagram
        # _resolve_diagram has already said which link in the chain failed, and "no
        # cached drawing" is a different fact from "failed to load" -- someone looking
        # at a blank rectangle needs to know which.  The vaguer warning on top of it
        # would only bury that.
        return _empty_graphic_frame(context, node)

    context.warn(
        "unsupported-graphic-frame",
        f"{node.what} {node.name or ''!r} is not rendered; drawing an empty frame",
    )
    return _empty_graphic_frame(context, node)


def _empty_graphic_frame(context: ResolveContext, node: s.SourceUnsupported) -> m.SlideElement:
    """A positioned but undrawn frame: the honest output for content we cannot render."""
    return m.ShapeElement(
        transform=_resolve_transform(context, node.transform),
        geometry=m.PresetGeometry(preset="rect"),
        fill=None,
        outline=None,
        alt_text=node.alt_text or node.name,
    )


# --------------------------------------------------------------------------------------
# Charts
# --------------------------------------------------------------------------------------


def _resolve_chart(context: ResolveContext, node: s.SourceUnsupported) -> m.SlideElement | None:
    """Read the chart part and lay it out.

    There is no cached drawing to fall back on -- PowerPoint stores a chart as data and
    re-draws it every time -- so this reads ``c:chartSpace`` and hands it to
    :mod:`pptx2svg.resolve.chart`, which computes the plot rectangle, the axis range and
    every bar.  Returns ``None`` when the part is missing or holds no chart type this
    can draw, having first said which; the caller then draws an empty frame.
    """
    label = f"chart {node.name or node.shape_id or ''!r}"

    def give_up(code: str, detail: str) -> None:
        context.warn(code, f"{label} {detail}")
        return None

    if node.fallback_rel_id is None:
        return give_up("chart-unreadable", "names no chart part")
    part = context.package.related_part(context.part_path, node.fallback_rel_id)
    if part is None or not context.package.has_part(part):
        return give_up("chart-unreadable", "points at a chart part that is not in the package")

    try:
        xml = context.package.read_xml(part)
    except Exception:
        return give_up("chart-unreadable", f"has a chart part ({part}) that is not well-formed XML")
    if xml is None:
        return give_up("chart-unreadable", f"has an unreadable chart part ({part})")

    source = parse_chart_space(xml)
    if source is None:
        return give_up("chart-unreadable", f"has a chart part ({part}) with no c:chart in it")

    plots = _drawable_plots(source)
    if not plots:
        kinds = ", ".join(sorted({p.kind for p in source.plots})) or "nothing"
        return give_up(
            "chart-unsupported-type",
            f"holds {kinds}, which is not rendered yet; drawing an empty frame",
        )
    plot = plots[0]

    three_d = sorted({p.kind for p in plots if is_three_d_kind(p.kind)})
    if three_d:
        # Not `chart-unsupported-type`: that code means "nothing was drawn".  This one is
        # the other thing a renderer can be, and the deck should not have to guess which
        # it got -- every category, value, label and axis in the picture is right, and the
        # scene it stands for is missing.
        context.warn(
            "chart-3d-flattened",
            f"{label} holds {', '.join(three_d)} and is drawn flat: no floor, back wall, "
            "depth or extrusion, and the c:view3D camera is not applied. Its data, "
            "categories, axis and legend are drawn in full",
        )

    transform = _resolve_transform(context, node.transform)
    if transform.extent_width <= 0 or transform.extent_height <= 0:
        return give_up("chart-unreadable", "has a zero-sized frame")

    chart_context = _chart_context(context, source, part)
    style = _chart_style(chart_context, source)

    builder = ChartBuilder(
        source,
        plot,
        width_pt=transform.extent_width / EMU_PER_POINT,
        height_pt=transform.extent_height / EMU_PER_POINT,
        style=style,
        resolve_fill=lambda fill: _resolve_fill(chart_context, fill),
        resolve_outline=lambda outline: _resolve_outline(chart_context, outline),
        resolve_text=lambda rich, text, size, align: _resolve_chart_title_text(
            chart_context, rich, text, size, align
        ),
        # `c:txPr` may name `+mn-lt` rather than a face; expanding it needs the theme.
        resolve_typeface=lambda typeface: _resolve_chart_typeface(chart_context, typeface),
        # Every drawable group, not only the first: a combo chart is several of them over
        # one plot area.  `ChartBuilder` decides which of them it can draw together.
        plots=plots,
    )
    children, data = builder.build()

    return m.ChartElement(
        transform=transform,
        chart=data,
        child_transform=m.Transform(
            offset_x=0,
            offset_y=0,
            extent_width=transform.extent_width,
            extent_height=transform.extent_height,
        ),
        children=children,
        alt_text=node.alt_text or node.name,
    )


def _drawable_plots(source) -> list:
    """Every plot group this renderer knows how to draw, in document order.

    ``barChart``, ``lineChart``, ``areaChart``, ``pieChart``, ``doughnutChart``,
    ``radarChart`` and the rest (and their 3-D spellings, drawn flat).  A chart holding
    several of them is a **combo**, and whether they can be drawn together is
    :meth:`~pptx2svg.resolve.chart.ChartBuilder._drawn_plots`' decision rather than this
    one -- it needs the axes, which are the chart's and not the group's.  What this
    guarantees is only that the *first* entry is drawable, so a chart whose first group
    is one we do not draw still draws the second.
    """
    return [plot for plot in source.plots if flat_chart_kind(plot.kind) in DRAWABLE_CHART_KINDS]


def _chart_context(context: ResolveContext, source, part: str) -> ResolveContext:
    """A resolution context scoped to the chart part.

    Two things have to change and nothing else.  ``part_path`` moves to the chart, because
    anything the chart relates to is related to *it* and not to the slide.  And the colour
    map becomes the chart's own: ``c:clrMapOvr`` is the innermost scope, and layering it
    on top of the slide's override -- the obvious thing to do -- is wrong, because a slide
    that remaps ``bg1``/``tx1`` for its own shapes does not remap them for a chart that
    declares its own mapping.  Both pptx-renderer and this project's roadmap flag it, and
    it is invisible until a deck does both at once.
    """
    if source.color_map_override is None:
        return replace(context, part_path=part)
    mapping = build_effective_color_map(
        context.master.color_map if context.master else None,
        None,
        source.color_map_override,
    )
    return replace(context, part_path=part, colors=ColorContext(context.theme, mapping))


def _chart_style(context: ResolveContext, source) -> ChartStyle:
    """The chart's text and series-colour defaults.

    Measured on ``authoring-integration.pptx``: with no ``c:txPr`` anywhere, PowerPoint
    drew every axis label and legend entry in **Aptos at 10 pt** -- the theme's minor
    latin face -- and black.  The title is deliberately not styled here; see
    :func:`_resolve_chart_title_text`.
    """
    theme = context.theme
    minor = (theme.font_scheme.minor_latin or None) if theme is not None else None
    # The East Asian face is a separate cascade, and the chart never states one: every
    # `c:txPr` in `real-financial-report.pptx` names `<a:latin typeface="Arial"/>` and
    # stops, while PowerPoint drew the Japanese category labels in the theme's
    # `<a:font script="Jpan" typeface="游ゴシック"/>`.  Body before heading, for the same
    # reason `_theme_body_latin` picks the minor face: the export used YuGothic-Regular
    # where the major entry is `游ゴシック Light`.
    minor_ea = (
        east_asian_family(
            theme.font_scheme.minor_east_asian,
            theme.font_scheme.minor_japanese,
            theme.font_scheme.major_east_asian,
            theme.font_scheme.major_japanese,
        )
        if theme is not None
        else None
    )
    text_color = resolve_color(context.colors, s.SchemeColor(scheme="tx1")) or m.ResolvedColor(
        hex="#000000"
    )
    accents = accent_colors(
        lambda key: resolve_color(context.colors, s.SchemeColor(scheme=key))
    )
    return ChartStyle(
        font_family=minor,
        font_size=default_font_size(source),
        color=text_color,
        accents=accents,
        font_family_ea=minor_ea,
    )


def _resolve_chart_typeface(context: ResolveContext, typeface: str | None) -> str | None:
    return _expand_theme_typeface(context, typeface)


def _resolve_chart_title_text(
    context: ResolveContext, rich, text: str, size: float, align: str
) -> m.TextBody:
    """A chart title, through the ordinary text cascade.

    This is the one piece of chart text that does *not* take the chart's 10 pt minor-face
    default.  ``c:title/c:tx/c:rich`` is plain DrawingML, and PowerPoint resolves it the
    way it resolves any unstyled text box: in ``authoring-integration.pptx`` the axis
    labels came out Aptos 10 pt and the title, whose ``a:rPr`` names nothing at all, came
    out Arial 18 pt -- the same face and size that file's plain text boxes get.  Running
    it through :func:`resolve_text_body` reproduces both halves of that for free.
    """
    if rich is None:
        return m.TextBody(
            paragraphs=[
                m.Paragraph(
                    runs=[m.TextRun(text=text, properties=m.RunProperties(font_size=size))],
                    properties=m.ParagraphProperties(alignment=align),
                )
            ],
            body_properties=CHART_TEXT_BODY,
        )
    body = _resolve_text_body(context, rich, [], None)
    for paragraph in body.paragraphs:
        paragraph.properties.alignment = align
    return replace(body, body_properties=CHART_TEXT_BODY)


# --------------------------------------------------------------------------------------
# SmartArt
# --------------------------------------------------------------------------------------


def _resolve_diagram(
    context: ResolveContext, node: s.SourceUnsupported, path: tuple[int, ...]
) -> m.SlideElement | None:
    """Render SmartArt from the DrawingML rendering PowerPoint already cached.

    A SmartArt diagram is authored as *data* -- a node tree plus a layout algorithm --
    and laying it out is a large piece of work.  It is also work PowerPoint has already
    done: every time it saves, it writes the fully-positioned result into a separate
    drawing part, so that other consumers do not have to run the layout engine.  That
    part is plain DrawingML, the same vocabulary as a slide's own shape tree, which is
    why this is forty lines and not a diagram engine.

    Finding that part is the fiddly bit; see :func:`_diagram_drawing_part`.

    Children are resolved with ``part_path`` pointed at the *drawing* part, because the
    pictures inside a diagram are related to it and not to the slide; resolving them
    against the slide would silently find the wrong image or none at all.

    Returns ``None`` when any link in the chain is missing, having first said *which*
    link and why; the caller then draws an empty frame.
    """
    label = f"SmartArt {node.name or node.shape_id or ''!r}"

    def give_up(code: str, detail: str) -> None:
        context.warn(code, f"{label} {detail}")
        return None

    data_part = context.package.related_part(context.part_path, node.fallback_rel_id)
    if data_part is None:
        return give_up(
            "diagram-unreadable", "points at a data-model part that is not in the package"
        )

    drawing_part = _diagram_drawing_part(context, data_part)
    if drawing_part is None:
        return give_up("diagram-no-cached-drawing", NO_CACHED_DRAWING)

    try:
        drawing = context.package.read_xml(drawing_part)
    except Exception:
        return give_up(
            "diagram-unreadable",
            f"has a cached drawing ({drawing_part}) that is not well-formed XML",
        )
    if drawing is None:
        return give_up(
            "diagram-unreadable", f"has an unreadable cached drawing ({drawing_part})"
        )

    sp_tree = child(drawing, "spTree")
    if sp_tree is None:
        return give_up(
            "diagram-unreadable",
            f"has a cached drawing ({drawing_part}) with no dsp:spTree in it",
        )

    frame_transform = _resolve_transform(context, node.transform)
    child_transform = _diagram_child_transform(sp_tree, frame_transform)

    # PowerPoint writes `id="0" name=""` on *every* shape in a cached diagram drawing --
    # identity there is carried by `modelId`, not by the DrawingML id.  So the ids are
    # rebuilt from the frame's id and the child's position, which is unique, stable
    # across runs, and keeps `data-pptx-id` addressable.
    outer_part, outer_prefix = context.part_path, context.id_prefix
    context.part_path = drawing_part
    children: list[m.SlideElement] = []
    try:
        for index, shape in enumerate(parse_shape_tree(sp_tree)):
            context.id_prefix = f"{outer_prefix}{node.shape_id or 'dgm'}/{index}/"
            element = resolve_element(context, shape, path + (index,))
            if element is not None:
                children.append(element)
    finally:
        context.part_path, context.id_prefix = outer_part, outer_prefix

    if not children:
        # PowerPoint writes this: a complete dsp:spTree holding nvGrpSpPr and grpSpPr and
        # no shapes at all.  Thirteen of the forty-six real decks measured look like this.
        # As far as output goes it is the same as having no drawing, and the reader needs
        # the same explanation.
        return give_up(
            "diagram-no-cached-drawing",
            f"has a cached drawing ({drawing_part}) whose shape tree is empty, so there "
            "is nothing to draw",
        )

    return m.GroupElement(
        transform=frame_transform,
        child_transform=child_transform,
        children=children,
        alt_text=node.alt_text or node.name,
    )


def _diagram_drawing_part(context: ResolveContext, data_part: str) -> str | None:
    """Find the cached DrawingML rendering that belongs to one diagram.

    This is not where the obvious reading of the schema puts it.  ``dgm:relIds`` on the
    graphic frame names four parts -- data model, layout, quick style, colours -- and
    conspicuously not the drawing, because the cached drawing was added to the format
    after ``relIds`` was specified.  Microsoft keyed it through an extension instead::

        slide rels --r:dm--------------> ppt/diagrams/data1.xml
            data1.xml dgm:extLst/dsp:dataModelExt@relId = "rId6"
                                                 |
        slide rels --rId6 (diagramDrawing)-------+--> ppt/diagrams/drawing1.xml

    So the relationship id is written in the *data* part but resolved against the
    *slide's* relationships.  Every one of the 46 real PowerPoint decks checked that has
    a cached drawing at all does it this way, and none of them has a
    ``ppt/diagrams/_rels/data1.xml.rels`` for it to hang off.

    Two fallbacks follow, in decreasing confidence:

    * the data part's own relationships, which is what the ISO/transitional layout would
      imply and what an independent producer might reasonably write;
    * failing that, a diagram-drawing relationship on the owning part -- but only when
      there is exactly one, since a slide with two SmartArt frames offers no way to tell
      which drawing belongs to which frame without the ``relId`` above.
    """
    owner = context.part_path

    relationship_id = _data_model_drawing_rel_id(context, data_part)
    if relationship_id is not None:
        target = context.package.related_part(owner, relationship_id)
        if target is not None and context.package.has_part(target):
            return target

    for rel_type in DIAGRAM_DRAWING_REL_TYPES:
        target = context.package.first_related_part(data_part, rel_type)
        if target is not None and context.package.has_part(target):
            return target

    candidates = [
        target
        for rel_type in DIAGRAM_DRAWING_REL_TYPES
        for target in context.package.related_parts_of_type(owner, rel_type)
        if context.package.has_part(target)
    ]
    return candidates[0] if len(candidates) == 1 else None


def _data_model_drawing_rel_id(context: ResolveContext, data_part: str) -> str | None:
    """``dsp:dataModelExt@relId`` out of the data model part, if it carries one."""
    try:
        data_model = context.package.read_xml(data_part)
    except Exception:
        return None
    if data_model is None:
        return None
    for node in descendants(data_model, "dataModelExt"):
        relationship_id = attr(node, "relId")
        if relationship_id:
            return relationship_id
    return None


def _diagram_child_transform(sp_tree, frame: m.Transform) -> m.Transform:
    """The coordinate space the cached diagram shapes were laid out in.

    PowerPoint writes ``dsp:spTree/dsp:grpSpPr/a:xfrm`` with ``chOff``/``chExt`` matching
    the graphic frame, so the normal group mapping scales the diagram into the frame.
    When the ``xfrm`` is absent -- some writers omit it -- the shapes are in a space whose
    origin is the frame's top-left and whose extent is the frame's, which is what this
    falls back to.  Note that it is *not* ``replace(frame)``: a group with no ``chOff``
    has children in absolute slide coordinates, whereas a diagram's are always relative
    to its own origin.
    """
    _, inner = parse_group_transforms(child(sp_tree, "grpSpPr"))
    if inner is None or not inner.width or not inner.height:
        return m.Transform(
            offset_x=0,
            offset_y=0,
            extent_width=frame.extent_width,
            extent_height=frame.extent_height,
        )
    return m.Transform(
        offset_x=inner.offset_x,
        offset_y=inner.offset_y,
        extent_width=inner.width,
        extent_height=inner.height,
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
        media = _load_media_bytes(context, fill.blip_relationship_id)
        if media is None:
            return None
        payload, mime_type = media
        if mime_type in METAFILE_MIME_TYPES:
            # A metafile is as legal a fill as it is a picture.  There is no on-slide
            # extent to size the raster by here -- the fill is scaled by whatever shape
            # it lands in -- so the preview is rendered at its own natural size.
            preview = _metafile_preview(
                context, payload, mime_type, width_emu=None, described_as="image fill"
            )
            if preview is None:
                return None
            payload, mime_type = preview
        if mime_type not in SUPPORTED_IMAGE_MIME_TYPES:
            context.warn("unsupported-fill-image", f"image fill of type {mime_type} not rendered")
            return None
        data = base64.b64encode(payload).decode("ascii")
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


def _load_media_bytes(context: ResolveContext, rel_id: str | None) -> tuple[bytes, str] | None:
    """Relationship id -> (raw payload, MIME type)."""
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
    return payload, mime_type


def _load_media(context: ResolveContext, rel_id: str | None) -> tuple[str, str] | None:
    """Relationship id -> (base64 payload, MIME type)."""
    media = _load_media_bytes(context, rel_id)
    if media is None:
        return None
    payload, mime_type = media
    return base64.b64encode(payload).decode("ascii"), mime_type


def _metafile_preview(
    context: ResolveContext,
    payload: bytes,
    mime_type: str,
    *,
    width_emu: float | None,
    described_as: str,
) -> tuple[bytes, str] | None:
    """Turn EMF/WMF bytes into something renderable, or ``None`` to draw a placeholder.

    Three routes, in descending order of the caller's authority over the result:

    1. A ``metafile_converter`` the caller installed.  It was configured deliberately, so
       it wins outright; it can return SVG, PNG, anything the renderer can embed.
    2. The preview Office already embedded -- a PDF, which still needs rasterising, or a
       DIB, which :mod:`pptx2svg.metafile.dib` turns into a PNG with no dependencies.
    3. Nothing, which keeps the historical grey rectangle and a warning.

    The distinction between "no preview in the file" and "preview found but no rasteriser
    installed" is kept in the warning text, because the fix differs: the first needs an
    external converter, the second needs ``pip install pptx2svg[metafile]``.
    """
    if context.metafile_converter is not None:
        try:
            converted = context.metafile_converter(payload, mime_type)
        except Exception as error:  # a user hook must not abort the whole conversion
            context.warn(
                "metafile-converter-failed",
                f"{described_as}: the configured metafile_converter raised {error!r}; "
                "falling back to the embedded preview",
            )
            converted = None
        if converted is not None:
            return converted

    preview = extract_metafile_preview(payload)
    if preview is None:
        context.warn(
            "metafile-image",
            f"{described_as} is an EMF/WMF metafile with no embedded preview, and "
            "vector metafile records are not interpreted; drawing a placeholder",
        )
        return None

    if preview.mime_type != "application/pdf":
        return preview.data, preview.mime_type

    try:
        png = rasterise_pdf(preview.data, width_emu=width_emu)
    except PdfRasterizerNotAvailable:
        context.warn(
            "metafile-rasterizer-missing",
            f"{described_as} carries an embedded PDF preview, but rendering it needs "
            "pypdfium2 (pip install pptx2svg[metafile]); drawing a placeholder",
        )
        return None
    if png is None:
        context.warn(
            "metafile-image",
            f"{described_as} carries an embedded PDF preview that could not be "
            "rendered; drawing a placeholder",
        )
        return None
    return png, "image/png"


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
from .text import _resolve_typeface as _expand_theme_typeface  # noqa: E402
from .text import resolve_text_body as _resolve_text_body  # noqa: E402
