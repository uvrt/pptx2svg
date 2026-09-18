"""The render model: a display-oriented, fully-resolved description of a slide.

This is the Python port of ``packages/renderer/src/model/*`` from pptx-glimpse.  Every
colour here is a concrete hex string (theme lookups and lumMod/tint/shade transforms
have already been applied), every font is a real typeface name (``+mn-lt`` resolved),
and every inherited placeholder property has been merged down.  The renderer consumes
this and nothing else -- it never sees the OOXML.

Lengths stay in EMU because shape geometry, text margins and line widths all arrive in
EMU and only become pixels at the moment they are written into the SVG.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal, Union

# --------------------------------------------------------------------------------------
# Colour and fill
# --------------------------------------------------------------------------------------

ImageMimeType = str
RectangleAlignment = Literal["tl", "t", "tr", "l", "ctr", "r", "bl", "b", "br"]


@dataclass(frozen=True)
class ResolvedColor:
    hex: str
    alpha: float = 1.0


@dataclass
class SolidFill:
    color: ResolvedColor
    type: Literal["solid"] = "solid"


@dataclass
class GradientStop:
    position: float
    color: ResolvedColor


@dataclass
class GradientFill:
    stops: list[GradientStop]
    angle: float = 0.0
    gradient_type: Literal["linear", "radial"] = "linear"
    center_x: float | None = None
    center_y: float | None = None
    type: Literal["gradient"] = "gradient"


@dataclass
class ImageFillTile:
    tx: float = 0.0
    ty: float = 0.0
    sx: float = 1.0
    sy: float = 1.0
    flip: Literal["none", "x", "y", "xy"] = "none"
    align: RectangleAlignment = "tl"


@dataclass
class ImageFill:
    #: base64-encoded payload, ready for a ``data:`` URI.
    image_data: str
    mime_type: ImageMimeType
    tile: ImageFillTile | None = None
    type: Literal["image"] = "image"


@dataclass
class PatternFill:
    preset: str
    foreground_color: ResolvedColor
    background_color: ResolvedColor
    type: Literal["pattern"] = "pattern"


@dataclass
class NoFill:
    type: Literal["none"] = "none"


Fill = Union[SolidFill, GradientFill, ImageFill, PatternFill, NoFill]


# --------------------------------------------------------------------------------------
# Line / outline
# --------------------------------------------------------------------------------------

ArrowType = Literal["none", "triangle", "stealth", "diamond", "oval", "arrow"]
ArrowSize = Literal["sm", "med", "lg"]
LineCap = Literal["butt", "round", "square"]
LineJoin = Literal["miter", "round", "bevel"]
DashStyle = Literal[
    "solid", "dash", "dot", "dashDot", "lgDash", "lgDashDot", "lgDashDotDot", "sysDash", "sysDot"
]


@dataclass
class ArrowEndpoint:
    type: ArrowType
    width: ArrowSize = "med"
    length: ArrowSize = "med"


@dataclass
class Outline:
    #: EMU
    width: float = 9525
    fill: Union[SolidFill, GradientFill, None] = None
    dash_style: DashStyle = "solid"
    custom_dash: list[float] | None = None
    line_cap: LineCap | None = None
    line_join: LineJoin | None = None
    head_end: ArrowEndpoint | None = None
    tail_end: ArrowEndpoint | None = None


# --------------------------------------------------------------------------------------
# Effects
# --------------------------------------------------------------------------------------


@dataclass
class OuterShadow:
    blur_radius: float
    distance: float
    #: degrees
    direction: float
    color: ResolvedColor
    alignment: RectangleAlignment = "b"
    rotate_with_shape: bool = True


@dataclass
class InnerShadow:
    blur_radius: float
    distance: float
    direction: float
    color: ResolvedColor


@dataclass
class Glow:
    radius: float
    color: ResolvedColor


@dataclass
class SoftEdge:
    radius: float


@dataclass
class EffectList:
    outer_shadow: OuterShadow | None = None
    inner_shadow: InnerShadow | None = None
    glow: Glow | None = None
    soft_edge: SoftEdge | None = None

    def is_empty(self) -> bool:
        return not (self.outer_shadow or self.inner_shadow or self.glow or self.soft_edge)


@dataclass
class BiLevelEffect:
    threshold: float


@dataclass
class BlurEffect:
    radius: float
    grow: bool = True


@dataclass
class LumEffect:
    brightness: float = 0.0
    contrast: float = 0.0


@dataclass
class DuotoneEffect:
    color1: ResolvedColor
    color2: ResolvedColor


@dataclass
class ClrChangeEffect:
    clr_from: ResolvedColor
    clr_to: ResolvedColor


@dataclass
class BlipEffects:
    grayscale: bool = False
    bi_level: BiLevelEffect | None = None
    blur: BlurEffect | None = None
    lum: LumEffect | None = None
    duotone: DuotoneEffect | None = None
    clr_change: ClrChangeEffect | None = None


# --------------------------------------------------------------------------------------
# Text
# --------------------------------------------------------------------------------------

TextVerticalType = Literal["horz", "vert", "vert270", "eaVert", "wordArtVert", "mongolianVert"]
AutoNumScheme = Literal[
    "arabicPeriod",
    "arabicParenR",
    "romanUcPeriod",
    "romanLcPeriod",
    "alphaUcPeriod",
    "alphaLcPeriod",
    "alphaLcParenR",
    "alphaUcParenR",
    "arabicPlain",
]


@dataclass(frozen=True)
class Hyperlink:
    url: str
    tooltip: str | None = None


@dataclass
class TextOutline:
    width: float
    color: ResolvedColor


@dataclass
class NoBullet:
    type: Literal["none"] = "none"


@dataclass
class CharBullet:
    char: str
    type: Literal["char"] = "char"


@dataclass
class AutoNumBullet:
    scheme: AutoNumScheme = "arabicPeriod"
    start_at: int = 1
    type: Literal["autoNum"] = "autoNum"


BulletType = Union[NoBullet, CharBullet, AutoNumBullet]


@dataclass
class PointsSpacing:
    """``a:spcPts`` -- 1/100 point."""

    value: float
    type: Literal["pts"] = "pts"


@dataclass
class PercentSpacing:
    """``a:spcPct`` -- 1/1000 percent (50000 == 50%)."""

    value: float
    type: Literal["pct"] = "pct"


SpacingValue = Union[PointsSpacing, PercentSpacing]


@dataclass
class TabStop:
    position: float
    alignment: Literal["l", "ctr", "r", "dec"] = "l"


@dataclass
class RunProperties:
    font_size: float | None = None
    font_family: str | None = None
    font_family_ea: str | None = None
    font_family_cs: str | None = None
    bold: bool = False
    italic: bool = False
    underline: bool = False
    #: ``a:rPr@u`` when it names something other than a plain single rule -- ``"dbl"``,
    #: ``"wavy"``, ``"dotted"``, ``"dotDash"`` and friends.
    underline_style: str | None = None
    strikethrough: bool = False
    color: ResolvedColor | None = None
    #: >0 superscript, <0 subscript
    baseline: float = 0.0
    hyperlink: Hyperlink | None = None
    outline: TextOutline | None = None
    highlight: ResolvedColor | None = None


@dataclass
class TextRun:
    text: str
    properties: RunProperties = field(default_factory=RunProperties)


@dataclass
class ParagraphProperties:
    alignment: Literal["l", "ctr", "r", "just"] | None = None
    line_spacing: SpacingValue | None = None
    space_before: SpacingValue = field(default_factory=lambda: PercentSpacing(0))
    space_after: SpacingValue = field(default_factory=lambda: PercentSpacing(0))
    level: int = 0
    bullet: BulletType | None = None
    bullet_font: str | None = None
    bullet_color: ResolvedColor | None = None
    bullet_size_pct: float | None = None
    margin_left: float | None = None
    indent: float | None = None
    tab_stops: list[TabStop] = field(default_factory=list)


@dataclass
class Paragraph:
    runs: list[TextRun] = field(default_factory=list)
    properties: ParagraphProperties = field(default_factory=ParagraphProperties)
    end_para_run_properties: RunProperties | None = None


@dataclass
class BodyProperties:
    anchor: Literal["t", "ctr", "b"] = "t"
    #: EMU; PowerPoint defaults are 0.1" left/right and 0.05" top/bottom.
    margin_left: float = 91440
    margin_right: float = 91440
    margin_top: float = 45720
    margin_bottom: float = 45720
    wrap: Literal["square", "none"] = "square"
    auto_fit: Literal["noAutofit", "normAutofit", "spAutofit"] = "noAutofit"
    font_scale: float = 1.0
    ln_spc_reduction: float = 0.0
    num_col: int = 1
    vert: TextVerticalType = "horz"
    #: ``a:bodyPr@rot`` in degrees -- the text rotates inside the shape, independently of
    #: the shape's own rotation, about the text box's centre.
    rotation: float = 0.0
    #: ``a:bodyPr@defTabSz`` -- spacing of the implicit tab stops, EMU.  PowerPoint's
    #: default is one inch.
    default_tab_size: float = 914400


@dataclass
class TextBody:
    paragraphs: list[Paragraph] = field(default_factory=list)
    body_properties: BodyProperties = field(default_factory=BodyProperties)


# --------------------------------------------------------------------------------------
# Geometry and shapes
# --------------------------------------------------------------------------------------


@dataclass
class Transform:
    #: all EMU
    offset_x: float = 0.0
    offset_y: float = 0.0
    extent_width: float = 0.0
    extent_height: float = 0.0
    #: degrees
    rotation: float = 0.0
    flip_h: bool = False
    flip_v: bool = False


@dataclass
class PresetGeometry:
    preset: str
    adjust_values: dict[str, float] = field(default_factory=dict)
    type: Literal["preset"] = "preset"


@dataclass
class CustomGeometryPath:
    width: float
    height: float
    #: SVG path data in the path's own coordinate space.
    commands: str


@dataclass
class CustomGeometry:
    paths: list[CustomGeometryPath] = field(default_factory=list)
    type: Literal["custom"] = "custom"


Geometry = Union[PresetGeometry, CustomGeometry]


@dataclass
class ShapeElement:
    transform: Transform
    geometry: Geometry
    fill: Fill | None = None
    outline: Outline | None = None
    text_body: TextBody | None = None
    #: ``dsp:txXfrm`` -- where the text box sits when SmartArt places it away from the
    #: shape.  Absolute, in the same space as :attr:`transform`; ``None`` means the text
    #: fills the shape, which is what every non-diagram shape does.
    text_transform: Transform | None = None
    effects: EffectList | None = None
    placeholder_type: str | None = None
    placeholder_idx: int | None = None
    alt_text: str | None = None
    hyperlink: Hyperlink | None = None
    #: Identity of the source shape, for cross-referencing rendered output back to the deck:
    #: ``"<sldId>.<cNvPr id>"`` for slide shapes, ``"lay:<id>"``/``"mst:<id>"`` for shapes
    #: inherited from the layout or master.  ``None`` when the source had no ``p:cNvPr``.
    element_id: str | None = None
    #: Index path through the shape tree, e.g. ``"3-1"`` for the second child of the fourth
    #: shape.  Disambiguates the id, which is *not* guaranteed unique in real decks.
    element_path: str | None = None
    type: Literal["shape"] = "shape"


@dataclass
class ConnectorElement:
    transform: Transform
    geometry: Geometry
    outline: Outline | None = None
    effects: EffectList | None = None
    alt_text: str | None = None
    #: Identity of the source shape, for cross-referencing rendered output back to the deck:
    #: ``"<sldId>.<cNvPr id>"`` for slide shapes, ``"lay:<id>"``/``"mst:<id>"`` for shapes
    #: inherited from the layout or master.  ``None`` when the source had no ``p:cNvPr``.
    element_id: str | None = None
    #: Index path through the shape tree, e.g. ``"3-1"`` for the second child of the fourth
    #: shape.  Disambiguates the id, which is *not* guaranteed unique in real decks.
    element_path: str | None = None
    type: Literal["connector"] = "connector"


@dataclass
class SrcRect:
    left: float = 0.0
    top: float = 0.0
    right: float = 0.0
    bottom: float = 0.0


@dataclass
class StretchFillRect:
    left: float = 0.0
    top: float = 0.0
    right: float = 0.0
    bottom: float = 0.0


@dataclass
class TileInfo:
    tx: float = 0.0
    ty: float = 0.0
    sx: float = 1.0
    sy: float = 1.0
    flip: Literal["none", "x", "y", "xy"] = "none"
    align: RectangleAlignment = "tl"


@dataclass
class ImageElement:
    transform: Transform
    image_data: str
    mime_type: ImageMimeType
    effects: EffectList | None = None
    blip_effects: BlipEffects | None = None
    src_rect: SrcRect | None = None
    alt_text: str | None = None
    stretch: StretchFillRect | None = None
    tile: TileInfo | None = None
    hyperlink: Hyperlink | None = None
    #: Non-rectangular picture frames clip the bitmap to the shape's geometry.
    geometry: Geometry | None = None
    outline: Outline | None = None
    #: Identity of the source shape, for cross-referencing rendered output back to the deck:
    #: ``"<sldId>.<cNvPr id>"`` for slide shapes, ``"lay:<id>"``/``"mst:<id>"`` for shapes
    #: inherited from the layout or master.  ``None`` when the source had no ``p:cNvPr``.
    element_id: str | None = None
    #: Index path through the shape tree, e.g. ``"3-1"`` for the second child of the fourth
    #: shape.  Disambiguates the id, which is *not* guaranteed unique in real decks.
    element_path: str | None = None
    type: Literal["image"] = "image"


@dataclass
class GroupElement:
    """A group viewport.

    Children are mapped with ``T(off) * R(ext center) * F(ext center) * S(ext / chExt) *
    T(-chOff)``; nested groups compose ancestor matrices from the outside inward.
    """

    transform: Transform
    child_transform: Transform
    children: list["SlideElement"] = field(default_factory=list)
    effects: EffectList | None = None
    alt_text: str | None = None
    #: Identity of the source shape, for cross-referencing rendered output back to the deck:
    #: ``"<sldId>.<cNvPr id>"`` for slide shapes, ``"lay:<id>"``/``"mst:<id>"`` for shapes
    #: inherited from the layout or master.  ``None`` when the source had no ``p:cNvPr``.
    element_id: str | None = None
    #: Index path through the shape tree, e.g. ``"3-1"`` for the second child of the fourth
    #: shape.  Disambiguates the id, which is *not* guaranteed unique in real decks.
    element_path: str | None = None
    type: Literal["group"] = "group"


@dataclass
class CellBorders:
    top: Outline | None = None
    bottom: Outline | None = None
    left: Outline | None = None
    right: Outline | None = None


@dataclass
class TableCell:
    text_body: TextBody | None = None
    fill: Fill | None = None
    borders: CellBorders | None = None
    grid_span: int = 1
    row_span: int = 1
    h_merge: bool = False
    v_merge: bool = False


@dataclass
class TableRow:
    height: float = 0.0
    cells: list[TableCell] = field(default_factory=list)


@dataclass
class TableColumn:
    width: float = 0.0


@dataclass
class TableData:
    rows: list[TableRow] = field(default_factory=list)
    columns: list[TableColumn] = field(default_factory=list)


@dataclass
class TableElement:
    transform: Transform
    table: TableData
    alt_text: str | None = None
    #: Identity of the source shape, for cross-referencing rendered output back to the deck:
    #: ``"<sldId>.<cNvPr id>"`` for slide shapes, ``"lay:<id>"``/``"mst:<id>"`` for shapes
    #: inherited from the layout or master.  ``None`` when the source had no ``p:cNvPr``.
    element_id: str | None = None
    #: Index path through the shape tree, e.g. ``"3-1"`` for the second child of the fourth
    #: shape.  Disambiguates the id, which is *not* guaranteed unique in real decks.
    element_path: str | None = None
    type: Literal["table"] = "table"


# --------------------------------------------------------------------------------------
# Charts
# --------------------------------------------------------------------------------------


@dataclass
class ChartSeries:
    """One plotted series, with its blanks preserved.

    ``values`` carries ``None`` where the workbook cell was empty, which is a different
    thing from a zero: ``c:dispBlanksAs`` draws a gap, a zero or an interpolated span,
    and a caller reading the data back needs to be able to tell which it had.
    """

    name: str | None = None
    values: list[float | None] = field(default_factory=list)
    categories: list[str] = field(default_factory=list)
    color: ResolvedColor | None = None
    #: The series' own ``c:formatCode``, used for its data labels.
    format_code: str | None = None


@dataclass
class ChartAxisScale:
    """The value axis as *drawn*, not as authored.

    OOXML usually leaves the range to the renderer, so these are the numbers this library
    chose; they are exposed because a caller re-plotting the data needs to know what the
    picture it is looking at actually shows.
    """

    minimum: float
    maximum: float
    major_unit: float


@dataclass
class Chart3DView:
    """``c:view3D`` -- the camera a 3-D chart's scene was authored with.

    Half spent, half carried.  The **plot rectangle** this camera implies is applied where
    it is measured -- a ``bar3DChart`` with right-angle axes is laid out in its scene's
    front face, displaced and shrunk by the depth, see
    :func:`~pptx2svg.resolve.chart.three_d_plot_rect` -- and the scene itself is still
    drawn flat (see :attr:`ChartData.three_d`).  So this is both the record of what the
    flattening threw away and the input whoever draws that scene will want.

    ``None`` fields are elements the file did not state; the defaults they take are
    PowerPoint's rather than the schema's, and which is which is documented on
    :class:`~pptx2svg.parse.chart.SourceChartView3D`.
    """

    #: Degrees of pitch and yaw.
    rot_x: float | None = None
    rot_y: float | None = None
    #: Depth and height, each as a percentage of the scene's width.
    depth_percent: float | None = None
    h_percent: float | None = None
    #: ``c:rAngAx`` -- right-angle axes, which turns the perspective off.
    right_angle_axes: bool | None = None
    perspective: float | None = None


@dataclass
class ChartData:
    """What a chart plots, independent of how it was drawn."""

    #: The OOXML group element, with 3-D variants already mapped to their 2-D equivalent
    #: (``bar3DChart`` -> ``barChart``).  :attr:`three_d` is what that mapping erases.
    kind: str
    series: list[ChartSeries] = field(default_factory=list)
    categories: list[str] = field(default_factory=list)
    title: str | None = None
    grouping: str | None = None
    #: ``col`` or ``bar``; ``None`` for chart types that have no bar direction.
    bar_direction: str | None = None
    value_axis: ChartAxisScale | None = None
    #: ``b`` / ``t`` / ``l`` / ``r`` / ``tr``, or ``None`` when there is no legend.
    legend_position: str | None = None
    #: Whether the group was authored as a 3-D spelling.  Such a chart's **scene** is
    #: still drawn flat -- no floor, no wall, no extrusion -- and it warns
    #: ``chart-3d-flattened`` when it is.  Its value axis and its plot rectangle are
    #: PowerPoint's own: a 3-D axis is not padded, and a ``bar3DChart``'s plot is the
    #: front face its camera puts inside the frame.  Both are measured, and both are why
    #: this flag has to survive the mapping to :attr:`kind`.
    three_d: bool = False
    #: ``c:view3D``, when the file states it.  The plot rectangle is laid out through it;
    #: the scene is not.  See :attr:`three_d`.
    view_3d: Chart3DView | None = None


@dataclass
class ChartElement:
    """A chart, both as data and as the primitives it was drawn with.

    ``children`` is an ordinary element list in the frame's own coordinate space, so a
    renderer needs no chart-specific code: the chart is lowered to rectangles, lines and
    text by the resolver, exactly as SmartArt is lowered to its cached shape tree.
    """

    transform: Transform
    chart: ChartData
    child_transform: Transform
    children: list["SlideElement"] = field(default_factory=list)
    alt_text: str | None = None
    #: Identity of the source graphic frame; see :class:`ShapeElement`.
    element_id: str | None = None
    element_path: str | None = None
    type: Literal["chart"] = "chart"


SlideElement = Union[
    ShapeElement, ImageElement, ConnectorElement, GroupElement, TableElement, ChartElement
]


# --------------------------------------------------------------------------------------
# Slide and presentation
# --------------------------------------------------------------------------------------


@dataclass
class Background:
    fill: Fill | None = None


@dataclass
class Slide:
    slide_number: int
    background: Background | None = None
    elements: list[SlideElement] = field(default_factory=list)
    show_master_sp: bool = True


@dataclass
class SlideSize:
    #: EMU
    width: float = 9144000
    height: float = 6858000


# --------------------------------------------------------------------------------------
# Theme
# --------------------------------------------------------------------------------------


@dataclass
class ColorScheme:
    dk1: str = "#000000"
    lt1: str = "#ffffff"
    dk2: str = "#44546a"
    lt2: str = "#e7e6e6"
    accent1: str = "#4472c4"
    accent2: str = "#ed7d31"
    accent3: str = "#a5a5a5"
    accent4: str = "#ffc000"
    accent5: str = "#5b9bd5"
    accent6: str = "#70ad47"
    hlink: str = "#0563c1"
    folHlink: str = "#954f72"


@dataclass
class FontScheme:
    major_font: str = "Calibri Light"
    minor_font: str = "Calibri"
    major_font_ea: str | None = None
    minor_font_ea: str | None = None
    major_font_cs: str | None = None
    minor_font_cs: str | None = None
    #: script-based font (Jpan) -- final fallback for CJK text
    major_font_jpan: str | None = None
    minor_font_jpan: str | None = None
