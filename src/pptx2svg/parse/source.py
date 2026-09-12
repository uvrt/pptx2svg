"""The source model: OOXML as parsed, before any inheritance or theme resolution.

Deliberately *unresolved*.  Theme colours stay as ``SourceSchemeColor("accent1")`` with
their lumMod/tint/shade transforms unapplied, images stay as relationship ids, and
``+mn-lt`` stays as the literal string.  Resolution needs context the reader does not
have -- which theme, which colour map, which layout a placeholder inherits from -- so it
belongs in :mod:`pptx2svg.resolve`.

``None`` means "not specified here, inherit from the layer above"; that distinction is
what makes placeholder and list-style inheritance work, so no field gets a concrete
default at parse time.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Union

from ..model import (
    ArrowEndpoint,
    BulletType,
    CustomGeometryPath,
    DashStyle,
    LineCap,
    LineJoin,
    RectangleAlignment,
    SpacingValue,
    TabStop,
    TextVerticalType,
)

# --------------------------------------------------------------------------------------
# Colour
# --------------------------------------------------------------------------------------

ColorTransformKind = Literal["lumMod", "lumOff", "tint", "shade", "alpha", "satMod", "satOff"]


@dataclass(frozen=True)
class ColorTransform:
    kind: ColorTransformKind
    #: OOXML 1/1000-percent
    value: float


@dataclass
class SrgbColor:
    hex: str
    transforms: list[ColorTransform] = field(default_factory=list)
    kind: Literal["srgb"] = "srgb"


@dataclass
class SchemeColor:
    """A ``a:schemeClr`` reference; ``scheme`` is a colour-map slot such as ``tx1``."""

    scheme: str
    transforms: list[ColorTransform] = field(default_factory=list)
    kind: Literal["scheme"] = "scheme"


@dataclass
class SystemColor:
    value: str
    last_color: str | None = None
    transforms: list[ColorTransform] = field(default_factory=list)
    kind: Literal["system"] = "system"


SourceColor = Union[SrgbColor, SchemeColor, SystemColor]


# --------------------------------------------------------------------------------------
# Fill and line
# --------------------------------------------------------------------------------------


@dataclass
class SourceSolidFill:
    color: SourceColor
    kind: Literal["solid"] = "solid"


@dataclass
class SourceNoFill:
    kind: Literal["none"] = "none"


@dataclass
class SourceGradientStop:
    position: float
    color: SourceColor


@dataclass
class SourceGradientFill:
    stops: list[SourceGradientStop]
    gradient_type: Literal["linear", "radial"] = "linear"
    #: OOXML 1/60000 degrees
    angle: float = 0.0
    center_x: float | None = None
    center_y: float | None = None
    kind: Literal["gradient"] = "gradient"


@dataclass
class SourceImageFillTile:
    tx: float = 0.0
    ty: float = 0.0
    sx: float = 1.0
    sy: float = 1.0
    flip: Literal["none", "x", "y", "xy"] = "none"
    align: RectangleAlignment = "tl"


@dataclass
class SourceImageFill:
    blip_relationship_id: str
    tile: SourceImageFillTile | None = None
    src_rect: tuple[float, float, float, float] | None = None
    stretch: tuple[float, float, float, float] | None = None
    kind: Literal["image"] = "image"


@dataclass
class SourcePatternFill:
    preset: str
    foreground_color: SourceColor
    background_color: SourceColor
    kind: Literal["pattern"] = "pattern"


@dataclass
class SourceGroupFill:
    """``a:grpFill`` -- inherit the enclosing group's fill."""

    kind: Literal["group"] = "group"


SourceFill = Union[
    SourceSolidFill,
    SourceNoFill,
    SourceGradientFill,
    SourceImageFill,
    SourcePatternFill,
    SourceGroupFill,
]


@dataclass
class SourceOutline:
    width: float | None = None
    fill: SourceFill | None = None
    dash_style: DashStyle | None = None
    custom_dash: list[float] | None = None
    line_cap: LineCap | None = None
    line_join: LineJoin | None = None
    head_end: ArrowEndpoint | None = None
    tail_end: ArrowEndpoint | None = None


@dataclass
class SourceStyleReference:
    """``a:fillRef`` / ``a:lnRef`` / ``a:effectRef`` -- an index into the theme's fmtScheme."""

    idx: int
    color: SourceColor | None = None


@dataclass
class SourceShapeStyle:
    fill_ref: SourceStyleReference | None = None
    line_ref: SourceStyleReference | None = None
    effect_ref: SourceStyleReference | None = None
    font_ref: SourceStyleReference | None = None


# --------------------------------------------------------------------------------------
# Effects
# --------------------------------------------------------------------------------------


@dataclass
class SourceOuterShadow:
    blur_radius: float
    distance: float
    direction: float
    color: SourceColor
    alignment: RectangleAlignment = "b"
    rotate_with_shape: bool = True


@dataclass
class SourceInnerShadow:
    blur_radius: float
    distance: float
    direction: float
    color: SourceColor


@dataclass
class SourceGlow:
    radius: float
    color: SourceColor


@dataclass
class SourceSoftEdge:
    radius: float


@dataclass
class SourceEffectList:
    outer_shadow: SourceOuterShadow | None = None
    inner_shadow: SourceInnerShadow | None = None
    glow: SourceGlow | None = None
    soft_edge: SourceSoftEdge | None = None


@dataclass
class SourceBlipEffects:
    grayscale: bool = False
    bi_level: float | None = None
    blur: tuple[float, bool] | None = None
    lum: tuple[float, float] | None = None
    duotone: tuple[SourceColor, SourceColor] | None = None
    clr_change: tuple[SourceColor, SourceColor] | None = None


# --------------------------------------------------------------------------------------
# Geometry and transform
# --------------------------------------------------------------------------------------


@dataclass
class SourceTransform:
    offset_x: float
    offset_y: float
    width: float
    height: float
    #: OOXML 1/60000 degrees
    rotation: float = 0.0
    flip_horizontal: bool = False
    flip_vertical: bool = False


@dataclass
class SourcePresetGeometry:
    preset: str
    adjust_values: dict[str, float] = field(default_factory=dict)
    kind: Literal["preset"] = "preset"


@dataclass
class SourceCustomGeometry:
    paths: list[CustomGeometryPath] = field(default_factory=list)
    kind: Literal["custom"] = "custom"


SourceGeometry = Union[SourcePresetGeometry, SourceCustomGeometry]


# --------------------------------------------------------------------------------------
# Text
# --------------------------------------------------------------------------------------


@dataclass
class SourceRunProperties:
    bold: bool | None = None
    italic: bool | None = None
    underline: bool | None = None
    strikethrough: bool | None = None
    baseline: float | None = None
    #: points
    font_size: float | None = None
    typeface: str | None = None
    typeface_ea: str | None = None
    typeface_cs: str | None = None
    color: SourceColor | None = None
    highlight: SourceColor | None = None
    outline_width: float | None = None
    outline_color: SourceColor | None = None
    hyperlink_rel_id: str | None = None
    hyperlink_tooltip: str | None = None


@dataclass
class SourceTextRun:
    text: str
    properties: SourceRunProperties | None = None


@dataclass
class SourceParagraphProperties:
    align: Literal["l", "ctr", "r", "just"] | None = None
    level: int | None = None
    line_spacing: SpacingValue | None = None
    space_before: SpacingValue | None = None
    space_after: SpacingValue | None = None
    margin_left: float | None = None
    indent: float | None = None
    bullet: BulletType | None = None
    bullet_font: str | None = None
    bullet_color: SourceColor | None = None
    bullet_size_pct: float | None = None
    tab_stops: list[TabStop] | None = None
    default_run_properties: SourceRunProperties | None = None


@dataclass
class SourceParagraph:
    runs: list[SourceTextRun] = field(default_factory=list)
    properties: SourceParagraphProperties | None = None
    end_para_run_properties: SourceRunProperties | None = None


@dataclass
class SourceTextStyle:
    """``a:lstStyle`` / ``p:titleStyle`` -- per-outline-level default paragraph properties."""

    default_paragraph: SourceParagraphProperties | None = None
    #: index 0 == lvl1pPr ... index 8 == lvl9pPr
    levels: list[SourceParagraphProperties | None] = field(
        default_factory=lambda: [None] * 9
    )


@dataclass
class SourceTextBodyProperties:
    margin_left: float | None = None
    margin_right: float | None = None
    margin_top: float | None = None
    margin_bottom: float | None = None
    anchor: Literal["t", "ctr", "b"] | None = None
    wrap: Literal["square", "none"] | None = None
    auto_fit: Literal["noAutofit", "normAutofit", "spAutofit"] | None = None
    font_scale: float | None = None
    ln_spc_reduction: float | None = None
    num_col: int | None = None
    vert: TextVerticalType | None = None
    rotation: float | None = None


@dataclass
class SourceTextBody:
    paragraphs: list[SourceParagraph] = field(default_factory=list)
    properties: SourceTextBodyProperties | None = None
    list_style: SourceTextStyle | None = None


# --------------------------------------------------------------------------------------
# Shape tree nodes
# --------------------------------------------------------------------------------------


@dataclass
class SourcePlaceholder:
    type: str | None = None
    idx: int | None = None


@dataclass
class SourceShape:
    name: str | None = None
    shape_id: str | None = None
    alt_text: str | None = None
    placeholder: SourcePlaceholder | None = None
    transform: SourceTransform | None = None
    geometry: SourceGeometry | None = None
    fill: SourceFill | None = None
    outline: SourceOutline | None = None
    effects: SourceEffectList | None = None
    style: SourceShapeStyle | None = None
    text_body: SourceTextBody | None = None
    hyperlink_rel_id: str | None = None
    kind: Literal["shape"] = "shape"


@dataclass
class SourceConnector:
    name: str | None = None
    shape_id: str | None = None
    alt_text: str | None = None
    transform: SourceTransform | None = None
    geometry: SourceGeometry | None = None
    outline: SourceOutline | None = None
    effects: SourceEffectList | None = None
    style: SourceShapeStyle | None = None
    kind: Literal["connector"] = "connector"


@dataclass
class SourceImage:
    blip_relationship_id: str | None = None
    name: str | None = None
    shape_id: str | None = None
    alt_text: str | None = None
    placeholder: SourcePlaceholder | None = None
    transform: SourceTransform | None = None
    geometry: SourceGeometry | None = None
    outline: SourceOutline | None = None
    effects: SourceEffectList | None = None
    blip_effects: SourceBlipEffects | None = None
    src_rect: tuple[float, float, float, float] | None = None
    stretch: tuple[float, float, float, float] | None = None
    tile: SourceImageFillTile | None = None
    hyperlink_rel_id: str | None = None
    kind: Literal["image"] = "image"


@dataclass
class SourceTableCell:
    text_body: SourceTextBody | None = None
    fill: SourceFill | None = None
    border_top: SourceOutline | None = None
    border_bottom: SourceOutline | None = None
    border_left: SourceOutline | None = None
    border_right: SourceOutline | None = None
    grid_span: int = 1
    row_span: int = 1
    h_merge: bool = False
    v_merge: bool = False
    margin_left: float | None = None
    margin_right: float | None = None
    margin_top: float | None = None
    margin_bottom: float | None = None
    anchor: Literal["t", "ctr", "b"] | None = None


@dataclass
class SourceTableRow:
    height: float = 0.0
    cells: list[SourceTableCell] = field(default_factory=list)


@dataclass
class SourceTable:
    name: str | None = None
    shape_id: str | None = None
    alt_text: str | None = None
    transform: SourceTransform | None = None
    columns: list[float] = field(default_factory=list)
    rows: list[SourceTableRow] = field(default_factory=list)
    first_row: bool = False
    band_row: bool = False
    kind: Literal["table"] = "table"


@dataclass
class SourceGroup:
    name: str | None = None
    shape_id: str | None = None
    alt_text: str | None = None
    transform: SourceTransform | None = None
    child_transform: SourceTransform | None = None
    fill: SourceFill | None = None
    effects: SourceEffectList | None = None
    children: list["SourceShapeNode"] = field(default_factory=list)
    kind: Literal["group"] = "group"


@dataclass
class SourceUnsupported:
    """A graphic frame we can position but not draw (chart, SmartArt, OLE, media)."""

    what: str
    name: str | None = None
    shape_id: str | None = None
    alt_text: str | None = None
    transform: SourceTransform | None = None
    #: Relationship id of a rendered fallback, when the frame ships one.
    fallback_rel_id: str | None = None
    fallback_part: str | None = None
    kind: Literal["unsupported"] = "unsupported"


SourceShapeNode = Union[
    SourceShape, SourceConnector, SourceImage, SourceTable, SourceGroup, SourceUnsupported
]


# --------------------------------------------------------------------------------------
# Parts
# --------------------------------------------------------------------------------------


@dataclass
class SourceColorMap:
    """``p:clrMap`` / ``p:clrMapOvr`` -- slot name -> colour-scheme key."""

    mapping: dict[str, str] = field(default_factory=dict)


@dataclass
class SourceFormatScheme:
    fill_styles: list[SourceFill] = field(default_factory=list)
    line_styles: list[SourceOutline] = field(default_factory=list)
    effect_styles: list[SourceEffectList | None] = field(default_factory=list)
    bg_fill_styles: list[SourceFill] = field(default_factory=list)


@dataclass
class SourceFontScheme:
    major_latin: str | None = None
    minor_latin: str | None = None
    major_east_asian: str | None = None
    minor_east_asian: str | None = None
    major_complex_script: str | None = None
    minor_complex_script: str | None = None
    major_japanese: str | None = None
    minor_japanese: str | None = None


@dataclass
class SourceTheme:
    part_path: str
    color_scheme: dict[str, SourceColor] = field(default_factory=dict)
    font_scheme: SourceFontScheme = field(default_factory=SourceFontScheme)
    format_scheme: SourceFormatScheme = field(default_factory=SourceFormatScheme)


@dataclass
class SourceBackground:
    fill: SourceFill | None = None
    #: ``p:bgRef`` -- index into the theme's bgFillStyleLst plus an override colour.
    bg_ref: SourceStyleReference | None = None


@dataclass
class SourceSlideBase:
    part_path: str
    shapes: list[SourceShapeNode] = field(default_factory=list)
    background: SourceBackground | None = None
    color_map_override: SourceColorMap | None = None


@dataclass
class SourceSlideMaster(SourceSlideBase):
    theme_part_path: str | None = None
    color_map: SourceColorMap | None = None
    title_style: SourceTextStyle | None = None
    body_style: SourceTextStyle | None = None
    other_style: SourceTextStyle | None = None


@dataclass
class SourceSlideLayout(SourceSlideBase):
    master_part_path: str | None = None
    show_master_shapes: bool = True
    layout_type: str | None = None


@dataclass
class SourceSlide(SourceSlideBase):
    layout_part_path: str | None = None
    show_master_shapes: bool = True
    slide_number: int = 1
    #: ``p:sldId/@id`` from the presentation's slide list -- deck-unique and, unlike
    #: ``slide_number``, unaffected by reordering.
    slide_id: int | None = None


@dataclass
class SourcePresentation:
    part_path: str
    slide_width: float = 9144000
    slide_height: float = 6858000
    default_text_style: SourceTextStyle | None = None
    slides: list[SourceSlide] = field(default_factory=list)
    layouts: dict[str, SourceSlideLayout] = field(default_factory=dict)
    masters: dict[str, SourceSlideMaster] = field(default_factory=dict)
    themes: dict[str, SourceTheme] = field(default_factory=dict)
