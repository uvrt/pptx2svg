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

# DrawingML's value types -- colour, fills, outline, effects, transform, geometry, picture
# tiling and the theme's colour scheme -- moved to ooxml-common, so that docx2svg draws
# with the same types and the same renderers.  Every name stays importable from here, and
# each *is* the shared class, so ``isinstance`` and identity hold through either path.
from ooxml_common.drawingml.model import (  # noqa: F401
    ArrowEndpoint,
    ArrowSize,
    ArrowType,
    BiLevelEffect,
    BlipEffects,
    BlurEffect,
    ClrChangeEffect,
    ColorScheme,
    CompoundLineType,
    CustomGeometry,
    CustomGeometryPath,
    DashStyle,
    DuotoneEffect,
    EffectList,
    Fill,
    Geometry,
    Glow,
    GradientFill,
    GradientStop,
    ImageFill,
    ImageFillTile,
    ImageMimeType,
    InnerShadow,
    LineCap,
    LineJoin,
    LumEffect,
    NoFill,
    OuterShadow,
    Outline,
    PatternFill,
    PresetGeometry,
    RectangleAlignment,
    ResolvedColor,
    SoftEdge,
    SolidFill,
    SrcRect,
    StretchFillRect,
    TileInfo,
    Transform,
)

# The drawable scene -- text bodies, shapes, connectors, pictures, groups, tables and
# charts -- moved to ooxml-common with the renderers that draw it, so that docx2svg draws a
# chart or a SmartArt diagram with them.  Each name is the shared class.
from ooxml_common.drawingml.scene import (  # noqa: E402,F401
    AutoNumBullet,
    AutoNumScheme,
    BlipBullet,
    BodyProperties,
    BulletType,
    CellBorders,
    CharBullet,
    Chart3DView,
    ChartAxisScale,
    ChartData,
    ChartElement,
    ChartSeries,
    ConnectorElement,
    GroupElement,
    Hyperlink,
    ImageElement,
    NoBullet,
    Paragraph,
    ParagraphProperties,
    PercentSpacing,
    PointsSpacing,
    RunProperties,
    ShapeElement,
    SlideElement,
    SpacingValue,
    TabStop,
    TableCell,
    TableColumn,
    TableData,
    TableElement,
    TableRow,
    TextBody,
    TextOutline,
    TextRun,
    TextVerticalType,
)


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
