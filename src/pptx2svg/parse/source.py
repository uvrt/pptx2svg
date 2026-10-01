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

from ..model import (
    ArrowEndpoint,
    BulletType,
    CompoundLineType,
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
# DrawingML
# --------------------------------------------------------------------------------------

# The unresolved DrawingML types -- colour choices, fills, outlines, shape styles,
# effects, transforms and geometry -- moved to ooxml-common with the reader that makes
# them; they are the same classes here.
from ooxml_common.drawingml.model import (  # noqa: E402,F401
    ColorTransform,
    ColorTransformKind,
    SchemeColor,
    SourceColor,
    SrgbColor,
    SystemColor,
)
from ooxml_common.drawingml.source import (  # noqa: E402,F401
    SourceBlipEffects,
    SourceCustomGeometry,
    SourceEffectList,
    SourceFill,
    SourceFormatScheme,
    SourceGeometry,
    SourceGlow,
    SourceGradientFill,
    SourceGradientStop,
    SourceGroupFill,
    SourceImageFill,
    SourceImageFillTile,
    SourceInnerShadow,
    SourceNoFill,
    SourceOuterShadow,
    SourceOutline,
    SourcePatternFill,
    SourcePresetGeometry,
    SourceShapeStyle,
    SourceSoftEdge,
    SourceSolidFill,
    SourceStyleReference,
    SourceTransform,
)


# --------------------------------------------------------------------------------------
# Text and the shape tree
# --------------------------------------------------------------------------------------

# Moved to ooxml-common with the shape tree and text body readers, because a SmartArt
# diagram's cached drawing is a shape tree in a Word document too, and a chart's text is
# a text body.  Each name is the shared class.
from ooxml_common.drawingml.source_tree import (  # noqa: E402,F401
    SourceBlipBullet,
    SourceBulletType,
    SourceColorMap,
    SourceConnector,
    SourceGroup,
    SourceImage,
    SourceParagraph,
    SourceParagraphProperties,
    SourcePlaceholder,
    SourceRunProperties,
    SourceShape,
    SourceShapeNode,
    SourceTable,
    SourceTableCell,
    SourceTableRow,
    SourceTextBody,
    SourceTextBodyProperties,
    SourceTextRun,
    SourceTextStyle,
    SourceUnsupported,
)


#: ``a:tblStyle`` conditional regions, lowest precedence first.  A cell takes its
#: formatting from every region that covers it, with later entries winning -- so the
#: header row beats the banding, and a corner cell beats the header row.  The element
#: names are the OOXML ones; the attribute names are the Python ones.
TABLE_STYLE_REGIONS: tuple[tuple[str, str], ...] = (
    ("wholeTbl", "whole_table"),
    ("band2V", "band2_v"),
    ("band1V", "band1_v"),
    ("band2H", "band2_h"),
    ("band1H", "band1_h"),
    ("lastCol", "last_col"),
    ("firstCol", "first_col"),
    ("lastRow", "last_row"),
    ("firstRow", "first_row"),
    ("swCell", "sw_cell"),
    ("seCell", "se_cell"),
    ("nwCell", "nw_cell"),
    ("neCell", "ne_cell"),
)


@dataclass
class SourceTableCellStyle:
    """One conditional region of a table style.

    ``border_inside_h`` / ``border_inside_v`` are the edges *within* the region; the four
    named sides are the region's own outer boundary.  For ``wholeTbl`` that boundary is
    the table's outline and the inside borders are every gridline; for ``firstRow`` the
    boundary is the header row's four sides and ``insideV`` separates its cells.
    """

    fill: SourceFill | None = None
    #: ``a:tcStyle/a:fillRef`` -- an index into the theme's fill style list.
    fill_ref: SourceStyleReference | None = None
    #: ``a:tcTxStyle`` as run properties, so it can join the text cascade unchanged.
    text: SourceRunProperties | None = None
    border_left: SourceOutline | None = None
    border_right: SourceOutline | None = None
    border_top: SourceOutline | None = None
    border_bottom: SourceOutline | None = None
    border_inside_h: SourceOutline | None = None
    border_inside_v: SourceOutline | None = None


@dataclass
class SourceTableStyle:
    style_id: str
    name: str | None = None
    whole_table: SourceTableCellStyle | None = None
    band1_h: SourceTableCellStyle | None = None
    band2_h: SourceTableCellStyle | None = None
    band1_v: SourceTableCellStyle | None = None
    band2_v: SourceTableCellStyle | None = None
    first_row: SourceTableCellStyle | None = None
    last_row: SourceTableCellStyle | None = None
    first_col: SourceTableCellStyle | None = None
    last_col: SourceTableCellStyle | None = None
    nw_cell: SourceTableCellStyle | None = None
    ne_cell: SourceTableCellStyle | None = None
    sw_cell: SourceTableCellStyle | None = None
    se_cell: SourceTableCellStyle | None = None


@dataclass
class SourceTableStyles:
    """``ppt/tableStyles.xml`` -- the deck's custom styles and its default style id."""

    default_style_id: str | None = None
    styles: dict[str, SourceTableStyle] = field(default_factory=dict)


# --------------------------------------------------------------------------------------
# Parts
# --------------------------------------------------------------------------------------


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


@dataclass(frozen=True)
class SourceEmbeddedFont:
    """One ``<p:embeddedFont>``: a family name and the parts carrying its four cuts.

    Relationship ids rather than part paths, because that is what the element holds and
    resolving them needs the package.  ``part_path`` is the part whose relationships they
    belong to -- ``ppt/presentation.xml`` -- carried along so the reader does not have to
    be told twice.

    The four slots are an assertion by the deck about which file plays which role, and
    the file behind a slot need not agree with it: PowerPoint substitutes when a family
    has no cut for a slot, so ``bold`` can point at a regular-weight face.  The deck's
    claim is the one that governs rendering -- see :func:`pptx2svg.fonts.sfnt.relabel`.
    """

    typeface: str
    part_path: str
    regular: str | None = None
    bold: str | None = None
    italic: str | None = None
    bold_italic: str | None = None


@dataclass
class SourcePresentation:
    part_path: str
    slide_width: float = 9144000
    slide_height: float = 6858000
    default_text_style: SourceTextStyle | None = None
    table_styles: SourceTableStyles | None = None
    slides: list[SourceSlide] = field(default_factory=list)
    layouts: dict[str, SourceSlideLayout] = field(default_factory=dict)
    masters: dict[str, SourceSlideMaster] = field(default_factory=dict)
    themes: dict[str, SourceTheme] = field(default_factory=dict)
    #: ``<p:embeddedFontLst>``, in document order.  Empty for the great majority of decks.
    embedded_fonts: list[SourceEmbeddedFont] = field(default_factory=list)
    #: ``docProps/app.xml``'s ``AppVersion`` -- the version of the application that wrote
    #: the deck, ``12.0000`` for Office 2007 -- or ``None`` when the deck does not say.
    #: PowerPoint reads a chart's missing elements differently for a 2007 file; see
    #: :data:`pptx2svg.resolve.chart.OFFICE_2007_TICK_MARKS`.
    app_version: str | None = None
