"""Chart layout and drawing: ``SourceChart`` in, ordinary slide elements out.

Unlike SmartArt there is nothing cached to lean on -- ``c:chartSpace`` is data plus
styling, and every position in the picture has to be computed.  The output is a plain
list of :class:`~pptx2svg.model.ShapeElement`, so the SVG writer, the text engine, the
fill resolver and the font substitution all apply to a chart exactly as they do to a
slide.  Nothing here knows about SVG.

**Every constant below was measured out of PowerPoint's own PDF export**, not reasoned
about.  Twelve probe charts differing in one input each -- title, legend position on all
four sides, font size at 8/10/14 pt, gap width, tick marks, series count -- were laid out
at an identical frame size and exported by PowerPoint 16.106; the plot rectangle, the
bars and every text baseline come out of the PDF as exact vector coordinates, so the
formulas could be fitted rather than guessed.  Each one records its residual against the
measurement.  Two font families are represented (Aptos and Arial, from the corpus), which
is what allowed the size-proportional and metric-proportional terms to be separated.

The layout PowerPoint actually performs, as far as the measurements can tell:

    frame
    ├─ title band            (only when there is a title)
    ├─ legend band           (only for legendPos t)
    │
    │   value labels │ plot area                     │ legend band (legendPos l/r)
    │
    ├─ category label band
    └─ legend band           (only for legendPos b)

Three things that look like bugs and are not:

* **The tick-mark allowance is reserved whether or not tick marks are drawn.**  A probe
  with ``majorTickMark="none"`` and one with ``"out"`` produced *byte-identical* plot
  rectangles, so the space is part of the layout rather than part of the tick.
* **The default axis and gridline colour is black, not grey**, and the default width is
  0.5 pt.  Every probe drew ``0 0 0 SC`` at ``6350 w``.  Charts written by modern
  PowerPoint carry a ``c:style`` or a chart-style part that overrides this to grey; none
  of the decks measured here does, and the measurement wins.
* **Axis labels default to 10 pt in the theme's minor font, but a ``c:rich`` title does
  not.**  In ``authoring-integration.pptx`` PowerPoint drew the axis labels in Aptos (the
  theme minor face) at 10 pt and the title, whose ``a:rPr`` names nothing at all, in
  Arial at 18 pt -- the same fallback it gives any unstyled DrawingML text.  So the title
  is resolved through the ordinary text cascade and the rest is not.
"""

from __future__ import annotations

import math
from dataclasses import dataclass, field, replace

from .. import model as m
from ..parse import chart as c
from ..parse import source as s
from ..text.fontmap import metrics_for
from ..text.measure import DEFAULT_LINE_HEIGHT_RATIO, is_cjk

EMU_PER_POINT = 12700.0

#: Padding between the frame edge and the outermost label block, in points.  Measured as
#: the left edge of the value-label column in all three decks and every probe: exactly
#: 6.5 pt, independent of frame size and font size.
FRAME_PADDING_PT = 6.5

#: The plot area's inset from the frame on a side with nothing on it.  Measured at
#: 10.9995 pt on the right of every probe and of ``authoring-integration.pptx``; it is
#: also the floor on the top inset (see :func:`_top_inset`).
EDGE_INSET_PT = 11.0

#: How far above its baseline a one-line label's optical centre sits, in ems.  Fitted;
#: see :attr:`FontBox.ink_centre` for the five measurements and why no exact rule emerged.
LABEL_INK_CENTRE_EM = 0.27

#: Gap between the right edge of the value labels and the value axis, over and above the
#: font's descent.  Fitted to 0.645 em across Aptos at 8/10/14 pt and Arial at 12 pt with
#: a residual under 0.05 pt in all four -- the tightest fit in this file.
VALUE_LABEL_GAP_EM = 0.645

#: Slack between the category-label line and the plot area, over and above the label's
#: own line height.  Fitted to 0.615 em over the same four cases; residual under 0.16 pt.
CATEGORY_LABEL_GAP_EM = 0.615

#: The plot area's inset above the topmost value label.  ``max(11.0, 5.0 + lineHeight/2)``
#: reproduces all four measurements to within 0.02 pt, including the 8 pt probe where the
#: floor is what binds.
TOP_INSET_BASE_PT = 5.0

#: A legend row's height, as a multiple of the line height.  Measured 24.083 pt for a
#: 10 pt Aptos legend against a 12.207 pt line height (1.973x); 2.0 is within 0.33 pt and
#: the same band was measured for ``legendPos`` ``t`` and ``b``.
LEGEND_BAND_LINES = 2.0

#: The vertical pitch between stacked legend entries, in ems.  Measured 18.0 pt for a
#: 10 pt legend on all three right-hand legends in real-financial-report.pptx -- which is
#: *not* the same as the horizontal band's height above, so the two are separate numbers.
LEGEND_ROW_PITCH_EM = 1.8

#: Where a horizontal legend's baseline sits, in ems from the frame edge it hugs.  Taken
#: straight off the probes rather than derived from the band: 12.913 pt above the frame
#: bottom for ``legendPos="b"`` and 17.133 pt below the frame top for ``"t"``, both at
#: 10 pt.  The two are not symmetric and no rule was found that makes them so.
LEGEND_BOTTOM_BASELINE_EM = 1.291
LEGEND_TOP_BASELINE_EM = 1.713

#: Legend swatch side and the gap after it, in ems.  Measured 5.4923 pt and 2.3711 pt at
#: 10 pt.
LEGEND_SWATCH_EM = 0.549
LEGEND_SWATCH_GAP_EM = 0.237
#: Padding either side of a side legend, and between entries in a horizontal one.
LEGEND_SIDE_LEAD_EM = 1.60
LEGEND_SIDE_TRAIL_EM = 1.01
LEGEND_ENTRY_GAP_EM = 0.5

#: A horizontal legend's run of entries is centred on the frame with this much lead-in
#: counted as part of it, which shifts the visible entries half of it to the right.
#: Measured at 1.93 pt of shift for a 10 pt legend, identically on the legend-b and
#: legend-t probes.
LEGEND_HORIZONTAL_LEAD_EM = 0.386

#: The title band, and its baseline inside it, as multiples of the line height and the
#: ascent.  Only one title was measurable (18 pt Arial, in two probes and the fixture, all
#: agreeing): the band is 29.70 pt against a 20.109 pt line height, and the baseline sits
#: 24.52 pt below the frame top against a 16.295 pt ascent.  These are the measured
#: ratios rather than the tidy 1.5 both are close to -- rounding the band cost 0.46 pt of
#: plot height, which moved every gridline by a pixel.  They are a one-font fit and should
#: be re-measured if a chart with a differently sized title ever disagrees.
TITLE_BAND_LINES = 1.4769
TITLE_BASELINE_ASCENTS = 1.5046

#: Default chart text size, in points.  ECMA-376's chart default and what PowerPoint drew
#: for every axis label and legend entry with no ``c:txPr``.
DEFAULT_CHART_FONT_PT = 10.0

#: Axis, tick and gridline defaults when no ``c:spPr`` says otherwise.
DEFAULT_AXIS_LINE_EMU = 6350.0
DEFAULT_AXIS_COLOR = "#000000"

#: A line series' stroke when its ``a:ln`` states no width -- 1.5 pt, measured.  A series
#: that *does* state one is taken literally: the real line chart in
#: ``real-financial-report.pptx`` says ``w="25400"`` and PowerPoint drew 2 pt.
DEFAULT_LINE_SERIES_WIDTH_EMU = 19050.0

#: A line series' marker outline when it states none -- 0.5 pt in the series' own colour.
DEFAULT_MARKER_OUTLINE_EMU = 6350.0

#: Marker symbols a line series cycles through when it states no ``c:marker`` of its own.
#: Measured only for series 0, which came out a **diamond** -- not the circle most
#: implementations assume.  The rest of the cycle is from the specification and is
#: **not measured**.
DEFAULT_MARKER_CYCLE = ("diamond", "square", "triangle", "x", "star", "dot")

#: Marker side when ``c:size`` is absent, in points (ECMA-376's default).
DEFAULT_MARKER_SIZE_PT = 7.0

#: Gap between a bar's edge and the *line box* of the data label beside it, in points.
#: Measured 4.86 pt at 10 pt and 4.70 pt at 14 pt for ``outEnd``, and 4.91 pt at 10 pt for
#: ``inBase`` -- so it is a fixed distance and **not** proportional to the font.
DATA_LABEL_GAP_PT = 4.85

#: The same for ``inEnd``, which sits closer to the bar's end.  One measurement, at 10 pt.
DATA_LABEL_INNER_GAP_PT = 4.05

#: Gap between a line chart's marker edge and its data label, in ems.  The label is
#: centred on the point vertically and sits to its right -- ECMA's ``r`` default, which is
#: what PowerPoint drew.  Measured once, with a 7 pt marker.
DATA_LABEL_LINE_GAP_EM = 0.6

#: The outline on a negative bar drawn hollow by ``c:invertIfNegative`` -- 0.75 pt.
INVERTED_BAR_OUTLINE_EMU = 9525.0

#: Floor on the divisor that turns a category band into one bar's width.  ``c:gapWidth``
#: is schema-bounded to 0..500 and files need not obey; -100 on a single series makes the
#: divisor zero exactly.
MIN_BAR_SLOTS = 0.01

#: A hand-written ``c:majorUnit`` can ask for an unbounded number of gridlines, so the
#: tick list is capped rather than trusted.
MAX_MAJOR_TICKS = 1000

#: ``c:gapWidth`` when absent, in percent of one bar's width (ECMA-376 default).
DEFAULT_GAP_WIDTH = 150.0

#: The six theme accents a series cycles through when it has no fill of its own.
ACCENT_KEYS = ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6")

#: Below this many major units of span, the plain power of ten is halved.  See
#: :func:`nice_axis_scale`; the threshold is somewhere in (1.842, 4.285] and 2 is the
#: round number inside it.
AXIS_HALVING_RATIO = 2.0

#: A horizontal bar chart's value axis comes out coarser than a vertical one's for the
#: same data and the same axis length, so once the interval is chosen it is stepped up
#: until the axis holds no more than this many of them.  **One measurement only** -- the
#: horizontal probe, whose 0..5 data PowerPoint drew as 0..6 by 2 where the identical
#: data on a vertical axis of almost the same length (151.4 pt against 145.0 pt) came out
#: 0..6 by 1.  It is therefore not a density limit, and what it really is remains unknown.
HORIZONTAL_MAX_INTERVALS = 5


# --------------------------------------------------------------------------------------
# Axis scaling
# --------------------------------------------------------------------------------------


def nice_axis_scale(
    data_minimum: float, data_maximum: float, horizontal: bool = False
) -> tuple[float, float, float]:
    """``(minimum, maximum, major_unit)`` for a value axis PowerPoint would draw itself.

    The major unit is the plain **power of ten** just below the span, halved when the span
    is less than :data:`AXIS_HALVING_RATIO` of it.  That is not the "aim for N ticks" rule
    every charting library uses, and the difference is not cosmetic -- N ticks cannot
    produce both of these, which PowerPoint does:

    ===========  ==============  ==========
    data         PowerPoint      intervals
    ===========  ==============  ==========
    0..5         0..6 by 1       6
    0..9         0..10 by 1      10
    -2..5        -3..6 by 1      9
    0..1842      0..2000 by 500  4
    0..4285      0..5000 by 1000 5
    ===========  ==============  ==========

    The domain always includes zero -- a bar that does not start at its axis is a
    different picture -- and both ends are rounded *strictly* outwards, so a series
    topping out at exactly 5 gets an axis to 6 rather than one whose last bar touches the
    frame.  Both bumps are measured: the first is what ``authoring-integration.pptx``
    does, the second is the -3 on the negative-value probe whose data floor is -2.
    """
    low = min(0.0, data_minimum)
    high = max(0.0, data_maximum)
    span = high - low

    if span <= 0 or not math.isfinite(span):
        # Every value zero (or unusable).  PowerPoint still draws an axis; 0..1 is the
        # smallest one that shows anything.
        return 0.0, 1.0, 1.0

    unit = 10.0 ** math.floor(math.log10(span))
    # A denormal span underflows the power of ten to zero; a span at the other end
    # overflows the strictly-outward rounding below to infinity.  Neither is a chart
    # anyone drew on purpose, and both used to raise out of the conversion.
    if not math.isfinite(unit) or unit <= 0:
        return 0.0, 1.0, 1.0
    if span / unit < AXIS_HALVING_RATIO:
        unit /= 2

    minimum, maximum = _axis_extent(unit, low, high, data_minimum, data_maximum)
    if horizontal:
        # Counted on the *rounded* extent, not the data span: 0..5 of data becomes a
        # 0..6 axis, and it is the six intervals in that which PowerPoint coarsens.
        while (maximum - minimum) / unit > HORIZONTAL_MAX_INTERVALS:
            stepped = _next_nice_unit(unit)
            if not math.isfinite(stepped) or stepped <= unit:
                break
            unit = stepped
            minimum, maximum = _axis_extent(unit, low, high, data_minimum, data_maximum)
    if not (math.isfinite(minimum) and math.isfinite(maximum) and maximum > minimum):
        return 0.0, 1.0, 1.0
    return minimum, maximum, unit


def _axis_extent(
    unit: float, low: float, high: float, data_minimum: float, data_maximum: float
) -> tuple[float, float]:
    """Round the domain outwards to whole units, strictly past the data at both ends."""
    maximum = math.ceil(high / unit) * unit
    if maximum <= data_maximum:
        maximum += unit
    minimum = math.floor(low / unit) * unit
    if data_minimum < 0 and minimum >= data_minimum:
        minimum -= unit
    return minimum, maximum


def _next_nice_unit(unit: float) -> float:
    """The next step up the 1-2-5 ladder from a unit already on it."""
    magnitude = 10.0 ** math.floor(math.log10(unit))
    mantissa = round(unit / magnitude, 6)
    if mantissa < 2:
        return 2 * magnitude
    if mantissa < 5:
        return 5 * magnitude
    return 10 * magnitude


# --------------------------------------------------------------------------------------
# Number formatting
# --------------------------------------------------------------------------------------


def format_number(value: float, format_code: str | None) -> str:
    """Render one number the way its ``c:formatCode`` asks.

    A small subset of the Excel format language: enough for the codes that appear on real
    axes (``General``, ``#,##0``, ``0.0%``, ``0.00``) and a graceful fall-through for the
    rest.  A full implementation is a project of its own and belongs nowhere near here.
    """
    if format_code is None or format_code in ("General", "@"):
        return _general(value)

    sections = _sections(format_code)
    # Excel's sections are positive; negative; zero; text.  A negative value uses the
    # second only when there *is* one -- with a single section it is formatted by that
    # one and keeps its own minus sign.
    negative_section = value < 0 and len(sections) > 1
    section = sections[1] if negative_section else sections[0]

    if "%" in section:
        decimals = _decimals(section)
        magnitude = abs(value) if negative_section else value
        return _wrap_negative(f"{magnitude * 100:.{decimals}f}%", section, negative_section)

    if not any(ch in section for ch in "#0"):
        return _general(value)

    decimals = _decimals(section)
    grouped = "," in _strip_literals(section)
    # The negative section states its own sign -- "(#,##0)" or "-#,##0" -- so the value
    # goes in unsigned and the section's own decoration is put back around it.
    magnitude = abs(value) if negative_section else value
    text = f"{magnitude:,.{decimals}f}" if grouped else f"{magnitude:.{decimals}f}"
    return _wrap_negative(text, section, negative_section)


def _wrap_negative(text: str, section: str, negative_section: bool) -> str:
    if not negative_section:
        return text
    if "(" in section and ")" in section:
        return f"({text})"
    return f"-{text}" if "-" in section else text


def _general(value: float) -> str:
    if value == int(value):
        return str(int(value))
    return f"{round(value, 10):g}"


def _sections(format_code: str) -> list[str]:
    """Split on ``;`` outside quotes -- positive, negative, zero, text."""
    sections: list[str] = []
    current: list[str] = []
    quoted = False
    for char in format_code:
        if char == '"':
            quoted = not quoted
        if char == ";" and not quoted:
            sections.append("".join(current))
            current = []
            continue
        current.append(char)
    sections.append("".join(current))
    return sections


def _strip_literals(section: str) -> str:
    """Drop ``[red]`` directives, ``"text"`` and escapes so only the numeric shape is left."""
    out: list[str] = []
    index = 0
    while index < len(section):
        char = section[index]
        if char == "[":
            index = section.find("]", index)
            if index < 0:
                break
            index += 1
            continue
        if char == '"':
            index = section.find('"', index + 1)
            if index < 0:
                break
            index += 1
            continue
        if char == "\\":
            index += 2
            continue
        if char in "_*":
            index += 2
            continue
        out.append(char)
        index += 1
    return "".join(out)


def _decimals(section: str) -> int:
    stripped = _strip_literals(section)
    if "." not in stripped:
        return 0
    tail = stripped.split(".", 1)[1]
    count = 0
    for char in tail:
        if char in "0#":
            count += 1
        else:
            break
    return count


# --------------------------------------------------------------------------------------
# Text metrics
# --------------------------------------------------------------------------------------


@dataclass(frozen=True)
class ChartFont:
    """The face a piece of chart text is drawn in, plus its metrics at that size.

    Three things in a chart can name a face and a size independently -- the chart's own
    ``c:txPr``, an axis's, and the legend's -- and they routinely disagree.  Carrying the
    pair together is what stops a label being *measured* in one face and *drawn* in
    another, which is the mistake ``text/metrics.py`` exists to prevent.
    """

    family: str | None
    box: "FontBox"

    @property
    def size(self) -> float:
        return self.box.size

    def width(self, text: str) -> float:
        return text_width(text, self.family, self.box.size)


@dataclass(frozen=True)
class FontBox:
    """The vertical metrics of one face at one size, in points."""

    size: float
    ascent: float
    descent: float

    @property
    def line_height(self) -> float:
        return self.ascent + self.descent

    @property
    def ink_centre(self) -> float:
        """How far above its baseline a label's optical centre sits.

        Used wherever PowerPoint centres a one-line label on something: a value-axis tick,
        or a legend swatch.  **The rule behind it was not identified.**  Measured offsets
        are 0.218 em (Aptos 10 pt), 0.2975 em (Aptos 8 pt), 0.2687 em (Aptos 14 pt),
        0.3208 em (Arial 12 pt) and 0.213 em (a 10 pt Aptos legend swatch), and that set
        is consistent with *none* of the obvious candidates -- half the cap height, half
        the x-height, half the line box, or the centre of the digits' own ink bounding
        box, each of which is out by 0.4 to 1.2 pt and in inconsistent directions.  So
        this is the fitted mean of the five, whose worst residual is 0.61 pt (about one
        pixel at the 1280 px the fidelity harness scores at).  Two of the five are 8 and
        10 pt, where PowerPoint's own 0.12 pt coordinate quantisation is +/-0.11 pt, so
        part of the spread is measurement noise rather than a missing term.
        """
        return LABEL_INK_CENTRE_EM * self.size

    @property
    def first_baseline(self) -> float:
        """Where our own text engine puts the first baseline below a box's top.

        Chart text is positioned by *baseline* here, but drawn by the ordinary text
        renderer, which positions by box.  Subtracting this converts one to the other, so
        the two stay in step even if the line-box rule changes.
        """
        return (DEFAULT_LINE_HEIGHT_RATIO - self.descent / self.size) * self.size


#: Fallback vertical metrics, as fractions of the em, for a face with no metrics table.
#: The descent is Calibri's, the commonest chart face after the theme's own; the ascent is
#: then whatever makes the line box the 1.2 em the renderer actually lays out, because a
#: shorter one would make an unmetricked title's band come out short of the text in it.
FALLBACK_DESCENT = 0.25
FALLBACK_ASCENT = DEFAULT_LINE_HEIGHT_RATIO - FALLBACK_DESCENT


def font_box(family: str | None, size: float) -> FontBox:
    metrics = metrics_for(family)
    if metrics is None:
        return FontBox(
            size=size, ascent=FALLBACK_ASCENT * size, descent=FALLBACK_DESCENT * size
        )
    units = metrics.units_per_em
    return FontBox(
        size=size,
        ascent=metrics.ascender / units * size,
        descent=abs(metrics.descender) / units * size,
    )


def text_width(text: str, family: str | None, size: float) -> float:
    """One line's advance width, in points.

    The CJK branch is not decoration: ``real-financial-report.pptx`` legends its series
    in Japanese, and measuring those with the Latin mean advance under-counted the legend
    band by 32 pt -- a quarter of the chart's width.  The rule is the same one
    :mod:`pptx2svg.text.measure` uses, so chart text is measured exactly as slide text is.
    """
    metrics = metrics_for(family)
    if metrics is None:
        return 0.5 * size * len(text)
    total = 0.0
    for char in text:
        width = metrics.widths.get(char)
        if width is None:
            width = metrics.cjk_width if is_cjk(ord(char)) else metrics.default_width
        total += width
    return total / metrics.units_per_em * size


# --------------------------------------------------------------------------------------
# The resolved chart
# --------------------------------------------------------------------------------------


@dataclass
class ChartStyle:
    """Everything the drawing needs that is not geometry."""

    font_family: str | None
    font_size: float
    color: m.ResolvedColor
    accents: list[m.ResolvedColor]


@dataclass
class _Rect:
    left: float
    top: float
    right: float
    bottom: float

    @property
    def width(self) -> float:
        return self.right - self.left

    @property
    def height(self) -> float:
        return self.bottom - self.top


@dataclass
class _Labels:
    """``c:dLbls`` resolved down to what actually gets printed."""

    show_value: bool = False
    show_category: bool = False
    show_series: bool = False
    show_percent: bool = False
    position: str | None = None
    number_format: str | None = None
    font: "ChartFont | None" = None
    color: m.ResolvedColor | None = None

    @property
    def anything(self) -> bool:
        return (
            self.show_value or self.show_category or self.show_series or self.show_percent
        )


@dataclass
class _Series:
    name: str | None
    values: list[float | None]
    color: m.ResolvedColor
    fill: m.Fill | None
    outline: m.Outline | None
    format_code: str | None
    invert_if_negative: bool
    #: Per-point ``c:dPt`` overrides.  Presence here also means "this point states its
    #: own formatting", which is what stops `c:invertIfNegative` overriding it.
    point_fills: dict[int, m.Fill] = field(default_factory=dict)
    point_outlines: dict[int, m.Outline] = field(default_factory=dict)
    #: ``c:varyColors`` fills, one per point.  Not a `c:dPt`, so inversion still applies.
    vary_fills: list[m.Fill] = field(default_factory=list)
    #: Line-chart only: the stroke along the points, and the marker drawn at each.
    line: m.Outline | None = None
    marker_symbol: str | None = None
    marker_size: float = DEFAULT_MARKER_SIZE_PT
    marker_fill: m.Fill | None = None
    marker_outline: m.Outline | None = None
    smooth: bool = False
    labels: _Labels | None = None
    #: ``c:dLbl`` overrides, keyed by point index.
    point_labels: dict[int, _Labels] = field(default_factory=dict)


class ChartBuilder:
    """Lowers one bar chart to slide elements, in the frame's own coordinate space.

    Coordinates are EMU with the frame's top-left at the origin, which is what
    :class:`~pptx2svg.model.GroupElement`'s child transform expects.  Internally
    everything is points, because every measured constant is.
    """

    def __init__(
        self,
        chart: c.SourceChart,
        plot: c.SourceChartPlot,
        *,
        width_pt: float,
        height_pt: float,
        style: ChartStyle,
        resolve_fill,
        resolve_outline,
        resolve_text,
        resolve_typeface=lambda typeface: typeface,
    ) -> None:
        self.chart = chart
        self.plot = plot
        self.frame = _Rect(0.0, 0.0, width_pt, height_pt)
        self.style = style
        self._resolve_fill = resolve_fill
        self._resolve_outline = resolve_outline
        self._resolve_text = resolve_text
        self._resolve_typeface = resolve_typeface
        self.elements: list[m.SlideElement] = []
        self._title_cache: "tuple[m.TextBody, FontBox] | None | object" = _UNSET

    # -- public -------------------------------------------------------------------------

    def build(self) -> tuple[list[m.SlideElement], m.ChartData]:
        series = self._series()
        categories = self._categories(series)
        value_axis = self._axis_for(1) or self._axis_of_kind("valAx")
        category_axis = self._axis_for(0) or self._axis_of_kind("catAx")

        scale = self._scale(series, value_axis)
        # Each axis carries its own `c:txPr`, and they disagree in real files.
        value_font = self._label_font(value_axis)
        category_font = self._label_font(category_axis)
        tick_texts = self._tick_texts(scale, value_axis)

        plot_rect = self._plot_rect(
            tick_texts, categories, value_font, category_font, scale
        )

        self._draw_background(plot_rect)
        self._draw_title()
        self._draw_gridlines(plot_rect, scale, value_axis)
        if self._is_line:
            self._draw_lines(plot_rect, series, categories, scale)
        else:
            self._draw_bars(plot_rect, series, categories, scale)
        self._draw_axis_lines(plot_rect, scale, value_axis, category_axis)
        self._draw_labels(
            plot_rect,
            scale,
            tick_texts,
            categories,
            value_axis,
            category_axis,
            value_font,
            category_font,
        )
        self._draw_data_labels(plot_rect, series, categories, scale)
        self._draw_legend(plot_rect, series)

        data = m.ChartData(
            kind=c.flat_chart_kind(self.plot.kind),
            series=[
                m.ChartSeries(
                    name=item.name,
                    values=list(item.values),
                    categories=list(categories),
                    color=item.color,
                    format_code=item.format_code,
                )
                for item in series
            ],
            categories=list(categories),
            title=self._title_text(),
            grouping=self.plot.grouping,
            bar_direction=self.plot.bar_direction or "col",
            value_axis=m.ChartAxisScale(
                minimum=scale[0], maximum=scale[1], major_unit=scale[2]
            ),
            legend_position=self._legend_position(),
        )
        return self.elements, data

    # -- model --------------------------------------------------------------------------

    def _vary_colors(self) -> bool:
        """Whether each *point* takes its own colour rather than the series' one.

        Only meaningful for a single unstacked series; with several series the colours
        already vary by series.
        """
        if not self.plot.vary_colors or len(self.plot.series) != 1:
            return False
        return (self.plot.grouping or "clustered") not in ("stacked", "percentStacked")

    def _series(self) -> list[_Series]:
        out: list[_Series] = []
        for index, source in enumerate(self.plot.series):
            fill = self._resolve_fill(source.fill)
            color = self._series_color(fill, index)
            if fill is None:
                fill = m.SolidFill(color=color)
            outline = self._resolve_outline(source.outline)
            item = _Series(
                name=source.name.plain if source.name else None,
                values=list(source.values),
                color=color,
                fill=fill,
                outline=outline,
                format_code=source.format_code,
                # ECMA-376 makes inversion the default; a file that does not want it says
                # so explicitly, and every chart in the corpus does.
                invert_if_negative=(
                    True if source.invert_if_negative is None else source.invert_if_negative
                ),
            )
            if self._vary_colors() and source.fill is None and self.style.accents:
                # Measured on the varyColors probe: points take accent1, accent2, accent3
                # *exactly*.  pptx-renderer darkens them to 88%, which PowerPoint does not.
                # Kept apart from `point_fills`, which means "this point has a `c:dPt` of
                # its own" and is what suppresses the negative-bar inversion.
                item.vary_fills = [
                    m.SolidFill(color=self.style.accents[index % len(self.style.accents)])
                    for index in range(len(item.values))
                ]
            if self._is_line:
                self._read_line_style(item, source, index)
            item.labels = self._read_labels(source.data_labels, self.plot.data_labels)
            if source.data_labels is not None:
                for point_index, override in source.data_labels.overrides.items():
                    item.point_labels[point_index] = self._read_labels(
                        override, source.data_labels, self.plot.data_labels
                    )
            for point in source.data_points:
                point_fill = self._resolve_fill(point.fill)
                if point_fill is not None:
                    item.point_fills[point.index] = point_fill
                point_outline = self._resolve_outline(point.outline)
                if point_outline is not None:
                    item.point_outlines[point.index] = point_outline
            out.append(item)
        return out

    def _read_labels(self, *sources: c.SourceChartDataLabels | None) -> _Labels:
        """``c:dLbls`` innermost first: a point's, then its series', then the group's.

        Every flag is optional at every level, so each is resolved separately rather than
        taking the first block whole.  ``c:delete`` on a point silences it outright.
        """
        present = [source for source in sources if source is not None]
        if any(source.delete for source in present):
            return _Labels()

        def flag(name: str) -> bool:
            for source in present:
                value = getattr(source, name)
                if value is not None:
                    return value
            return False

        def first(name: str):
            for source in present:
                value = getattr(source, name)
                if value:
                    return value
            return None

        labels = _Labels(
            show_value=flag("show_value"),
            show_category=flag("show_category_name"),
            show_series=flag("show_series_name"),
            show_percent=flag("show_percent"),
            position=first("position"),
            number_format=first("number_format"),
        )
        if labels.anything:
            labels.font = self._font(*(source.text_properties for source in present))
            labels.color = self.style.color
        return labels

    @property
    def _is_line(self) -> bool:
        return c.flat_chart_kind(self.plot.kind) == "lineChart"

    def _read_line_style(
        self, item: _Series, source: c.SourceChartSeries, index: int
    ) -> None:
        """A line series' stroke and marker.

        **A line's colour comes from its ``a:ln``, not from ``a:solidFill``.**  Measured:
        a probe series stating only ``<a:solidFill><a:srgbClr val="F97316"/></a:solidFill>``
        was drawn by PowerPoint in accent1, its bare fill ignored -- so a bar chart's
        colour rule cannot simply be reused here.
        """
        outline = self._resolve_outline(source.outline)
        if outline is not None and isinstance(outline.fill, m.SolidFill):
            item.color = outline.fill.color
        elif self.style.accents:
            item.color = self.style.accents[index % len(self.style.accents)]
        if outline is None or outline.fill is None:
            outline = m.Outline(
                width=DEFAULT_LINE_SERIES_WIDTH_EMU,
                fill=m.SolidFill(color=item.color),
            )
        if outline.line_cap is None:
            # Measured: PowerPoint strokes a line series with `1 J`, a round cap, whether
            # or not the `a:ln` says so.
            outline = replace(outline, line_cap="round")
        item.line = outline

        marker = source.marker
        if marker is None:
            # No `c:marker` at all: PowerPoint draws one anyway, from a per-series cycle.
            item.marker_symbol = DEFAULT_MARKER_CYCLE[index % len(DEFAULT_MARKER_CYCLE)]
        elif marker.symbol in (None, "auto"):
            item.marker_symbol = DEFAULT_MARKER_CYCLE[index % len(DEFAULT_MARKER_CYCLE)]
        elif marker.symbol != "none":
            item.marker_symbol = marker.symbol
        if marker is not None and marker.size:
            item.marker_size = marker.size
        item.marker_fill = (
            self._resolve_fill(marker.fill) if marker is not None else None
        ) or m.SolidFill(color=item.color)
        item.marker_outline = (
            self._resolve_outline(marker.outline) if marker is not None else None
        ) or m.Outline(
            width=DEFAULT_MARKER_OUTLINE_EMU, fill=m.SolidFill(color=item.color)
        )
        item.smooth = bool(source.smooth)

    def _series_color(self, fill: m.Fill | None, index: int) -> m.ResolvedColor:
        """One flat colour for the series, for its legend swatch and its fallback fill.

        A series that states no fill takes the next theme accent, cycling; Office's own
        accent1 is the last resort when the theme has none.
        """
        if isinstance(fill, m.SolidFill):
            return fill.color
        if self.style.accents:
            return self.style.accents[index % len(self.style.accents)]
        return m.ResolvedColor(hex="#4472C4")

    def _categories(self, series: list[_Series]) -> list[str]:
        """One label per category, as long as the longest series.

        The bars are indexed by this list, so a series longer than the labelled one would
        lose its tail.  Series in one group need not agree on length -- and a series may
        carry no ``c:cat`` at all -- so the labelled list is padded rather than trusted for
        its length.  PowerPoint numbers unlabelled categories from 1.
        """
        longest = max((len(item.values) for item in series), default=0)
        labels: list[str] = []
        for source in self.plot.series:
            if any(source.categories):
                labels = list(source.categories)
                break
        if len(labels) >= longest:
            return labels
        return labels + [str(index + 1) for index in range(len(labels), longest)]

    def _axis_for(self, position: int) -> c.SourceChartAxis | None:
        """The axis this plot group names in its ``c:axId`` list, by position.

        Index 0 is the category axis and index 1 the value axis.  Going through the ids
        rather than through the first ``c:valAx`` in the plot area is what makes a
        secondary axis land on the right series -- and it also copes with
        ``real-financial-report.pptx``, which names a third id no axis element declares.
        """
        if position >= len(self.plot.axis_ids):
            return None
        wanted = self.plot.axis_ids[position]
        for axis in self.chart.axes:
            if axis.axis_id == wanted:
                return axis
        return None

    def _axis_of_kind(self, kind: str) -> c.SourceChartAxis | None:
        for axis in self.chart.axes:
            if axis.kind == kind:
                return axis
        return None

    def _scale(
        self, series: list[_Series], axis: c.SourceChartAxis | None
    ) -> tuple[float, float, float]:
        stacked = (self.plot.grouping or "clustered") in ("stacked", "percentStacked")
        if (self.plot.grouping or "") == "percentStacked":
            # Measured: PowerPoint labels 0%, 10% ... 100%.
            return 0.0, 1.0, 0.1

        numbers: list[float] = []
        if stacked:
            # A stacked bar reaches the sum of its category, and the positive and negative
            # halves of that category stack away from zero independently.
            length = max((len(item.values) for item in series), default=0)
            for index in range(length):
                column = [_at(item.values, index) for item in series]
                numbers.append(sum(value for value in column if value and value > 0))
                numbers.append(sum(value for value in column if value and value < 0))
        else:
            numbers = [value for item in series for value in item.values if value is not None]

        if not numbers:
            numbers = [0.0]
        minimum, maximum, unit = nice_axis_scale(
            min(numbers), max(numbers), horizontal=(self.plot.bar_direction or "col") == "bar"
        )

        if axis is not None:
            if axis.minimum is not None:
                minimum = axis.minimum
            if axis.maximum is not None:
                maximum = axis.maximum
            if axis.major_unit is not None and axis.major_unit > 0:
                unit = axis.major_unit
        if maximum <= minimum:
            maximum = minimum + (unit or 1.0)
        return minimum, maximum, unit

    def _tick_texts(
        self, scale: tuple[float, float, float], axis: c.SourceChartAxis | None
    ) -> list[tuple[float, str]]:
        minimum, maximum, unit = scale
        format_code = self._axis_format(axis)
        if (self.plot.grouping or "") == "percentStacked":
            format_code = "0%"
        values = self._tick_values(scale)
        if not values:
            return [(minimum, format_number(minimum, format_code))]
        return [(value, format_number(value, format_code)) for value in values]

    def _axis_format(self, axis: c.SourceChartAxis | None) -> str | None:
        """The format code the value axis' labels are printed with.

        ``c:numFmt@sourceLinked`` means "take the cell's format", and the cell's format is
        what the value cache's own ``c:formatCode`` records -- so a source-linked axis
        falls through to the series.  An axis that says ``sourceLinked="0"`` has *chosen*
        its format, including when that choice is ``General``, and must not inherit.
        """
        if axis is None:
            return None
        if axis.number_format and axis.number_format != "General":
            return axis.number_format
        if axis.number_format_source_linked is False:
            return None
        for source in self.plot.series:
            if source.format_code and source.format_code != "General":
                return source.format_code
        return None

    # -- layout -------------------------------------------------------------------------

    def _font(self, *sources: "s.SourceTextBody | None") -> ChartFont:
        """The innermost ``c:txPr`` that names a size or a face wins, per property.

        Sources are given innermost first.  A ``c:txPr`` may name only one of the two --
        every axis in ``real-financial-report.pptx`` names ``Arial`` and a size while its
        legend names neither -- so size and face resolve independently rather than as a
        unit.
        """
        size = None
        typeface = None
        for source in (*sources, self.chart.text_properties):
            if size is None:
                size = _text_size(source)
            if typeface is None:
                typeface = _text_typeface(source)
        family = self._resolve_typeface(typeface) if typeface else None
        family = family or self.style.font_family
        return ChartFont(family=family, box=font_box(family, size or self.style.font_size))

    def _label_font(self, axis: c.SourceChartAxis | None) -> ChartFont:
        return self._font(axis.text_properties if axis is not None else None)

    def _legend_font(self) -> ChartFont:
        legend = self.chart.legend
        return self._font(legend.text_properties if legend is not None else None)

    def _plot_rect(
        self,
        tick_texts: list[tuple[float, str]],
        categories: list[str],
        value_font: ChartFont,
        category_font: ChartFont,
        scale: tuple[float, float, float],
    ) -> _Rect:
        frame = self.frame
        value_axis = self._axis_for(1) or self._axis_of_kind("valAx")
        category_axis = self._axis_for(0) or self._axis_of_kind("catAx")

        show_values = _labels_shown(value_axis)
        show_categories = _labels_shown(category_axis)
        horizontal = (self.plot.bar_direction or "col") == "bar"

        # The left column and the band under the plot each hold one axis' labels, and
        # `barDir` decides which.  Measured on the horizontal probe: its left inset,
        # 55.41 pt, is the same formula as a vertical chart's but fed the widest
        # *category* label instead of the widest tick.
        tick_labels = [text for _, text in tick_texts]
        down_left = categories if horizontal else tick_labels
        show_left = show_categories if horizontal else show_values
        show_bottom = show_values if horizontal else show_categories
        left_font = category_font if horizontal else value_font
        bottom_font = value_font if horizontal else category_font
        # `tickLblPos="nextTo"` means next to the *axis*, and a chart with negative values
        # has its category axis floating above the plot's lower edge.  PowerPoint then
        # reserves no band under the plot at all -- the negative probe's bottom inset is
        # 11.103 pt, the same half-label allowance as its top -- and prints the category
        # labels inside the plot, just under the zero line.
        labels_under_plot = show_bottom and (
            horizontal
            or scale[0] >= 0
            or (category_axis is not None and category_axis.tick_label_position == "low")
        )

        left = frame.left + EDGE_INSET_PT
        if show_left:
            widest = max((left_font.width(text) for text in down_left), default=0.0)
            left = (
                frame.left
                + FRAME_PADDING_PT
                + widest
                + left_font.box.descent
                + VALUE_LABEL_GAP_EM * left_font.size
            )

        right = frame.right - EDGE_INSET_PT
        if horizontal and show_values:
            # The value axis runs along the bottom now, and its last label is centred on
            # the plot's right edge, so half of it hangs outside.  Measured 13.67 pt
            # against an 11.0 pt inset and a 5.34 pt label.
            right -= max((value_font.width(text) for text in tick_labels), default=0.0) / 2
        # Nothing overhangs the top of a horizontal chart, so it takes the plain inset.
        top = frame.top + (EDGE_INSET_PT if horizontal else self._top_inset(value_font.box))
        if labels_under_plot:
            bottom = frame.bottom - (
                FRAME_PADDING_PT
                + bottom_font.box.line_height
                + CATEGORY_LABEL_GAP_EM * bottom_font.size
            )
        else:
            bottom = frame.bottom - self._top_inset(value_font.box)

        title = self._title_box()
        if title is not None:
            top += TITLE_BAND_LINES * title.line_height

        legend = self._legend_position()
        # `c:overlay` draws the legend on top of the plot rather than beside it, so an
        # overlaid legend takes no space away.  Implemented from the schema; no chart
        # measured against PowerPoint sets it.
        if legend is not None and not self._legend_overlays():
            legend_font = self._legend_font()
            band = LEGEND_BAND_LINES * legend_font.box.line_height
            if legend in ("b",):
                bottom -= band
            elif legend in ("t", "tr"):
                top += band
            elif legend == "r":
                # The side band *replaces* the plain edge inset rather than adding to it:
                # it already ends in its own trailing pad.  Measured on
                # real-financial-report's two bar charts, whose legends are Japanese and
                # of different lengths -- both were over by exactly 11.0 pt, the inset.
                right = frame.right - self._legend_side_width(legend_font)
            elif legend == "l":
                # On the left the value-label column follows the legend instead of the
                # frame edge, so the band contributes one edge inset less.  Measured:
                # the legend-l probe's left inset is 85.067 pt and the legend-r probe's
                # right inset 74.994 pt for the same entry -- a difference of exactly the
                # 11.0 pt inset, with the 21.07 pt label column on top.
                left += self._legend_side_width(legend_font) - EDGE_INSET_PT

        if right - left < 1.0:
            right = left + 1.0
        if bottom - top < 1.0:
            bottom = top + 1.0
        return _Rect(left, top, right, bottom)

    def _top_inset(self, label: FontBox) -> float:
        """Space above the plot area for the topmost value label to sit in.

        ``max(11.0, 5.0 + lineHeight/2)`` fits Aptos at 8, 10 and 14 pt and Arial at 12 pt
        to within 0.02 pt.  The floor is what binds at 8 pt, which is why a purely
        proportional rule does not work.
        """
        return max(EDGE_INSET_PT, TOP_INSET_BASE_PT + label.line_height / 2)

    def _legend_side_width(self, font: ChartFont) -> float:
        # Only the entries actually drawn: a series struck out by `c:legendEntry` would
        # otherwise reserve width for a label nobody sees, shifting the plot rectangle.
        deleted = self.chart.legend.deleted_entries if self.chart.legend else set()
        widest = max(
            (
                font.width(source.name.plain or "")
                for index, source in enumerate(self.plot.series)
                if source.name is not None and index not in deleted
            ),
            default=0.0,
        )
        return (
            widest
            + (LEGEND_SIDE_LEAD_EM + LEGEND_SWATCH_EM + LEGEND_SWATCH_GAP_EM + LEGEND_SIDE_TRAIL_EM)
            * font.size
        )

    def _legend_overlays(self) -> bool:
        return self.chart.legend is not None and self.chart.legend.overlay

    def _legend_position(self) -> str | None:
        legend = self.chart.legend
        if legend is None:
            return None
        position = legend.position or "r"
        return position if position in ("b", "t", "l", "r", "tr") else "r"

    def _title(self) -> tuple[m.TextBody, FontBox] | None:
        """The title's resolved text body and the metrics of the face it will be drawn in.

        The face matters to the *layout*, not just the drawing: the title band is a
        multiple of its line height, and a 3 pt error there moves every bar.  So the body
        is resolved before the plot rectangle is computed and the metrics are read back
        off it, rather than being guessed from a default.  Resolving a chart title is a
        full trip through the text cascade, so the result is cached.
        """
        if self._title_cache is _UNSET:
            self._title_cache = self._build_title()
        return self._title_cache

    def _build_title(self) -> "tuple[m.TextBody, FontBox] | None":
        text = self._title_text()
        if text is None or self.chart.title is None:
            return None
        size = _title_size(self.chart.title)
        body = self._resolve_text(self.chart.title.rich, text, size, align="ctr")
        family, resolved_size = _first_run_font(body)
        size = resolved_size or size
        # A run that named no size would otherwise take the *renderer's* default, which
        # is a second place the number lives.  Stamping it makes the size the layout used
        # and the size drawn the same number by construction.
        for paragraph in body.paragraphs:
            for run in paragraph.runs:
                if run.properties.font_size is None:
                    run.properties.font_size = size
        return body, font_box(family, size)

    def _title_box(self) -> FontBox | None:
        title = self._title()
        return None if title is None else title[1]

    def _title_text(self) -> str | None:
        if self.chart.auto_title_deleted or self.chart.title is None:
            return None
        return self.chart.title.plain or None

    # -- drawing ------------------------------------------------------------------------

    def _draw_background(self, rect: _Rect) -> None:
        """The chart frame's own fill, then the plot rectangle's.

        A chart with no ``c:spPr`` at all is transparent -- the slide shows through, which
        is what PowerPoint drew for ``authoring-integration.pptx`` -- so an absent fill is
        not the same as a white one and nothing is emitted for it.
        """
        fill = self._resolve_fill(self.chart.fill)
        outline = self._resolve_outline(self.chart.outline)
        if (fill is not None and not isinstance(fill, m.NoFill)) or outline is not None:
            self._rect(self.frame, fill=fill, outline=outline)

        plot_fill = self._resolve_fill(self.chart.plot_area_fill)
        if plot_fill is not None and not isinstance(plot_fill, m.NoFill):
            self._rect(rect, fill=plot_fill, outline=None)

    def _draw_title(self) -> None:
        title = self._title()
        if title is None:
            return
        body, box = title
        baseline = self.frame.top + TITLE_BASELINE_ASCENTS * box.ascent
        self._text(
            body,
            left=self.frame.left,
            width=self.frame.width,
            baseline=baseline,
            box=box,
        )

    def _draw_gridlines(
        self,
        rect: _Rect,
        scale: tuple[float, float, float],
        axis: c.SourceChartAxis | None,
    ) -> None:
        if axis is None or not axis.major_gridlines:
            return
        outline = self._axis_outline(axis.major_gridline_outline)
        horizontal = (self.plot.bar_direction or "col") == "bar"
        # A gridline runs *across* the value axis, so `barDir="bar"` turns them vertical.
        # Measured on the horizontal probe: three vertical lines at the ticks for 2, 4 and
        # 6, with the one for 0 left out because the category axis already draws it.
        axis_position = self._category_axis_position(rect, scale, None)
        for value in self._tick_values(scale):
            if horizontal:
                x = self._value_to_x(rect, value, scale)
                if abs(x - axis_position) < 0.01:
                    continue
                self._line(x, rect.top, x, rect.bottom, outline)
            else:
                y = self._value_to_y(rect, value, scale)
                if abs(y - axis_position) < 0.01:
                    continue
                self._line(rect.left, y, rect.right, y, outline)

    def _tick_values(self, scale: tuple[float, float, float]) -> list[float]:
        """Every major tick on the value axis, or nothing when the unit is unusable.

        `c:majorUnit` is attacker-controlled, so an interval small enough to ask for
        millions of gridlines returns an empty list rather than trying to draw them.
        """
        minimum, maximum, unit = scale
        span = maximum - minimum
        if unit <= 0 or not math.isfinite(span) or not math.isfinite(unit):
            return []
        steps = span / unit
        if not math.isfinite(steps) or steps <= 0 or steps > MAX_MAJOR_TICKS:
            return []
        return [minimum + index * unit for index in range(int(round(steps)) + 1)]

    def _draw_bars(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        if not series or not categories:
            return
        horizontal = (self.plot.bar_direction or "col") == "bar"
        grouping = self.plot.grouping or "clustered"
        stacked = grouping in ("stacked", "percentStacked")

        band = (rect.height if horizontal else rect.width) / len(categories)
        gap_width = self.plot.gap_width
        if gap_width is None:
            gap_width = DEFAULT_GAP_WIDTH
        slots = 1 if stacked else len(series)
        # `gapWidth` is the gap between category groups expressed as a percentage of one
        # bar's width, so the band holds `slots` bars plus `gapWidth/100` of one more.
        # The schema bounds it to 0..500 and a file is free to ignore that; at -100 on a
        # single series the divisor is exactly zero, which used to abort the conversion.
        bar_size = band / max(slots + gap_width / 100.0, MIN_BAR_SLOTS)
        overlap = self.plot.overlap if self.plot.overlap is not None else 0.0
        step = bar_size * (1.0 - overlap / 100.0)
        cluster = bar_size + step * (slots - 1)

        percent = grouping == "percentStacked"
        totals = _percent_totals(series) if percent else None

        for point in range(len(categories)):
            # A horizontal bar chart runs its category axis bottom-to-top, so category 0
            # is the *lowest* band.  Measured: "Reader" labels the bottom bar.
            band_start = (
                rect.bottom - (point + 1) * band if horizontal else rect.left + point * band
            )
            centre = band_start + band / 2
            positive_base = 0.0
            negative_base = 0.0
            for order, item in enumerate(series):
                value = _at(item.values, point)
                if value is None:
                    if self.chart.display_blanks_as == "zero":
                        value = 0.0
                    else:
                        continue
                if percent and totals is not None:
                    total = totals[point]
                    value = 0.0 if total == 0 else value / total

                if stacked:
                    start = positive_base if value >= 0 else negative_base
                    end = start + value
                    if value >= 0:
                        positive_base = end
                    else:
                        negative_base = end
                    slot = 0
                else:
                    start, end = 0.0, value
                    slot = order

                offset = centre - cluster / 2 + slot * step
                self._bar(rect, item, point, offset, bar_size, start, end, scale, horizontal)

    def _draw_lines(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        """One polyline per series, with a marker at each point.

        Points sit at the **centre of their category band**, exactly where a bar would be
        -- measured on both the probe and the real line chart, whose `c:crossBetween` says
        ``between``.  ``midCat`` would put them on the band edges; no chart measured here
        uses it, so that spelling is **implemented from the schema and unverified**.
        """
        if not categories:
            return
        band = rect.width / len(categories)
        midcat = self._cross_between() == "midCat"
        blanks = self.chart.display_blanks_as or "gap"

        for item in series:
            points: list[tuple[float, float] | None] = []
            for index in range(len(categories)):
                value = _at(item.values, index)
                if value is None:
                    if blanks == "zero":
                        value = 0.0
                    elif blanks == "span":
                        # "span" bridges the gap: the point is dropped and its neighbours
                        # join up, which is what leaving it out of the run does.
                        continue
                    else:
                        points.append(None)
                        continue
                x = (
                    rect.left + index * (rect.width / max(len(categories) - 1, 1))
                    if midcat
                    else rect.left + (index + 0.5) * band
                )
                points.append((x, self._value_to_y(rect, value, scale)))

            for run in _split_runs(points):
                if len(run) > 1 and item.line is not None:
                    self._polyline(run, item.line, smooth=item.smooth)
            for point in points:
                if point is not None and item.marker_symbol:
                    self._marker(point, item)

    def _cross_between(self) -> str:
        axis = self._axis_for(1) or self._axis_of_kind("valAx")
        return (axis.cross_between if axis is not None else None) or "between"

    def _marker(self, centre: tuple[float, float], item: _Series) -> None:
        """One marker, centred on its data point.

        ``c:size`` is the marker's **diameter in points** -- a 7 pt circle measured 6.96 pt
        across in the probe.
        """
        half = item.marker_size / 2
        x, y = centre
        box = _Rect(x - half, y - half, x + half, y + half)
        self._rect(
            box,
            fill=item.marker_fill,
            outline=item.marker_outline,
            preset=_MARKER_PRESETS.get(item.marker_symbol or "", "ellipse"),
        )

    def _bar(
        self,
        rect: _Rect,
        item: _Series,
        point: int,
        offset: float,
        size: float,
        start: float,
        end: float,
        scale: tuple[float, float, float],
        horizontal: bool,
    ) -> None:
        fill = item.point_fills.get(point)
        if fill is None and point < len(item.vary_fills):
            fill = item.vary_fills[point]
        if fill is None:
            fill = item.fill
        outline = item.point_outlines.get(point, item.outline)
        if end < start and item.invert_if_negative and point not in item.point_fills:
            # Measured: PowerPoint draws a negative bar white with a black 0.75 pt
            # outline, and draws it in the series colour when the file sets
            # `invertIfNegative` to 0.
            fill = m.SolidFill(color=m.ResolvedColor(hex="#FFFFFF"))
            outline = outline or m.Outline(
                width=INVERTED_BAR_OUTLINE_EMU,
                fill=m.SolidFill(color=m.ResolvedColor(hex=DEFAULT_AXIS_COLOR)),
            )

        if horizontal:
            x0 = self._value_to_x(rect, start, scale)
            x1 = self._value_to_x(rect, end, scale)
            box = _Rect(min(x0, x1), offset, max(x0, x1), offset + size)
        else:
            y0 = self._value_to_y(rect, start, scale)
            y1 = self._value_to_y(rect, end, scale)
            box = _Rect(offset, min(y0, y1), offset + size, max(y0, y1))
        if box.width <= 0 or box.height <= 0:
            return
        self._rect(box, fill=fill, outline=outline)

    def _category_axis_position(
        self,
        rect: _Rect,
        scale: tuple[float, float, float],
        axis: c.SourceChartAxis | None,
    ) -> float:
        """Where the category axis crosses the value axis, in frame coordinates.

        A ``y`` for a column chart and an ``x`` for a bar one -- it is a position *along*
        the value axis either way.  ``c:crosses="autoZero"`` -- the default and what every
        chart in the corpus says -- puts it at value zero, which is the plot's low edge
        only while nothing is negative.  ``tickLblPos="low"`` pins it to the low end
        regardless; that spelling appears in ``real-financial-report.pptx`` but only over
        positive data, so its behaviour under a negative minimum is **implemented from the
        schema and not measured**.
        """
        horizontal = (self.plot.bar_direction or "col") == "bar"
        low, high = (rect.left, rect.right) if horizontal else (rect.bottom, rect.top)

        if axis is not None and axis.tick_label_position == "low":
            return low
        crosses = axis.crosses if axis is not None else None
        if crosses == "max":
            return high
        if crosses == "min":
            return low
        value = axis.crosses_at if axis is not None and crosses == "val" else 0.0
        if value is None:
            value = 0.0
        along = (
            self._value_to_x(rect, value, scale)
            if horizontal
            else self._value_to_y(rect, value, scale)
        )
        return min(max(along, min(low, high)), max(low, high))

    def _draw_axis_lines(
        self,
        rect: _Rect,
        scale: tuple[float, float, float],
        value_axis: c.SourceChartAxis | None,
        category_axis: c.SourceChartAxis | None,
    ) -> None:
        # `barDir` swaps which axis is the horizontal line and which the vertical one, so
        # `c:delete` and `c:spPr` have to follow their own axis rather than a fixed side.
        # Measured on the horizontal probe: the value axis is the line along the bottom
        # and the category axis the one up the left.
        horizontal = (self.plot.bar_direction or "col") == "bar"
        crossing = self._category_axis_position(rect, scale, category_axis)

        if value_axis is not None and not value_axis.delete:
            outline = self._axis_outline(value_axis.outline)
            if horizontal:
                self._line(rect.left, rect.bottom, rect.right, rect.bottom, outline)
            else:
                self._line(rect.left, rect.top, rect.left, rect.bottom, outline)
        if category_axis is not None and not category_axis.delete:
            # The category axis sits where it crosses, which is the zero line and not the
            # plot's edge once anything is negative.
            outline = self._axis_outline(category_axis.outline)
            if horizontal:
                self._line(crossing, rect.top, crossing, rect.bottom, outline)
            else:
                self._line(rect.left, crossing, rect.right, crossing, outline)

    def _draw_labels(
        self,
        rect: _Rect,
        scale: tuple[float, float, float],
        tick_texts: list[tuple[float, str]],
        categories: list[str],
        value_axis: c.SourceChartAxis | None,
        category_axis: c.SourceChartAxis | None,
        value_font: ChartFont,
        category_font: ChartFont,
    ) -> None:
        """Draw both axes' labels, on whichever side ``barDir`` puts them.

        A ``col`` chart labels values down the left and categories along the bottom; a
        ``bar`` chart does the opposite.  Both sides use the same two placements, so the
        orientation only decides which set of strings goes where.
        """
        horizontal = (self.plot.bar_direction or "col") == "bar"
        if _labels_shown(value_axis):
            if horizontal:
                self._labels_along_bottom(
                    rect,
                    [(self._value_to_x(rect, value, scale), text) for value, text in tick_texts],
                    value_font,
                    axis_y=rect.bottom,
                    centred_on_position=True,
                )
            else:
                self._labels_down_left(
                    rect,
                    [(self._value_to_y(rect, value, scale), text) for value, text in tick_texts],
                    value_font,
                )
        if _labels_shown(category_axis) and categories:
            if horizontal:
                band = rect.height / len(categories)
                self._labels_down_left(
                    rect,
                    # Category 0 is the lowest band on a horizontal chart.
                    [
                        (rect.bottom - (index + 0.5) * band, text)
                        for index, text in enumerate(categories)
                    ],
                    category_font,
                )
            else:
                band = rect.width / len(categories)
                self._labels_along_bottom(
                    rect,
                    [
                        (rect.left + index * band, text)
                        for index, text in enumerate(categories)
                    ],
                    category_font,
                    axis_y=self._category_axis_position(rect, scale, category_axis),
                    width=band,
                )

    def _labels_down_left(
        self, rect: _Rect, labels: list[tuple[float, str]], font: ChartFont
    ) -> None:
        """Right-aligned in the column left of the plot, each centred on its own y."""
        box = font.box
        width = rect.left - self.frame.left - box.descent - VALUE_LABEL_GAP_EM * box.size
        for y, text in labels:
            if not text:
                continue
            self._text(
                self._label_body(text, font, align="r"),
                left=self.frame.left,
                width=width,
                baseline=y + box.ink_centre,
                box=box,
            )

    def _labels_along_bottom(
        self,
        rect: _Rect,
        labels: list[tuple[float, str]],
        font: ChartFont,
        *,
        axis_y: float,
        width: float | None = None,
        centred_on_position: bool = False,
    ) -> None:
        """One line below the axis: category labels centred in their band, ticks on theirs.

        The baseline hangs off the *category axis*, not the frame, because ``nextTo`` means
        what it says: on a chart with negative values the axis floats above the plot's
        lower edge and the labels follow it.  ``ascent + 0.615 em`` reproduces all five
        measurements -- Aptos at 8/10/14 pt, Arial at 12 pt, and the negative probe --
        with a worst residual of 0.63 pt, and is the same number as hanging the line's
        descender one frame padding above the frame whenever the axis *is* at the foot.
        """
        box = font.box
        baseline = axis_y + box.ascent + CATEGORY_LABEL_GAP_EM * box.size
        for position, text in labels:
            if not text:
                continue
            if centred_on_position:
                # A value tick's label is centred on the tick, so the box is opened wide
                # either side of it and the text centred in that.
                span = font.width(text) + box.size
                left, box_width = position - span / 2, span
            else:
                left, box_width = position, width or box.size
            self._text(
                self._label_body(text, font, align="ctr"),
                left=left,
                width=box_width,
                baseline=baseline,
                box=box,
            )

    def _draw_data_labels(
        self,
        rect: _Rect,
        series: list[_Series],
        categories: list[str],
        scale: tuple[float, float, float],
    ) -> None:
        """Print `c:dLbls` beside each point.

        Positions are measured against a probe: ``outEnd`` (the bar default) puts the
        label's line box one :data:`DATA_LABEL_GAP_PT` beyond the bar's end, ``inBase``
        the same distance inside its base, ``inEnd`` just inside its end, and ``ctr`` on
        the bar's middle.  A line chart's default is ECMA's ``r``: centred on the point
        vertically, one marker radius plus a gap to its right.
        """
        if not categories:
            return
        horizontal = (self.plot.bar_direction or "col") == "bar"
        percent_totals = _percent_totals(series)

        for order, item in enumerate(series):
            for point in range(len(categories)):
                labels = item.point_labels.get(point, item.labels)
                if labels is None or not labels.anything or labels.font is None:
                    continue
                value = _at(item.values, point)
                if value is None:
                    continue
                text = self._label_text(
                    labels, item, categories, point, value, percent_totals
                )
                if not text:
                    continue
                geometry = self._label_anchor(
                    rect, series, order, item, point, value, scale, horizontal
                )
                if geometry is None:
                    continue
                self._place_label(text, labels, geometry, horizontal)

    def _label_text(
        self,
        labels: _Labels,
        item: _Series,
        categories: list[str],
        point: int,
        value: float,
        percent_totals: list[float],
    ) -> str:
        """The label's lines, top to bottom.

        Measured on the multi-part probe: PowerPoint stacks series name, category name and
        value on **separate lines**, in that order, rather than joining them with the
        ``c:separator`` a single-line label would use.
        """
        parts: list[str] = []
        if labels.show_series and item.name:
            parts.append(item.name)
        if labels.show_category and point < len(categories) and categories[point]:
            parts.append(categories[point])
        if labels.show_percent:
            total = percent_totals[point] if point < len(percent_totals) else 0.0
            parts.append(format_number(value / total if total else 0.0, "0%"))
        if labels.show_value:
            parts.append(
                format_number(value, labels.number_format or item.format_code)
            )
        return "\n".join(parts)

    def _label_anchor(
        self,
        rect: _Rect,
        series: list[_Series],
        order: int,
        item: _Series,
        point: int,
        value: float,
        scale: tuple[float, float, float],
        horizontal: bool,
    ) -> "tuple[float, float, str] | None":
        """``(x, y, placement)`` for one label, in frame points."""
        if self._is_line:
            band = rect.width / max(len(item.values), 1)
            x = rect.left + (point + 0.5) * band
            return x + item.marker_size / 2, self._value_to_y(rect, value, scale), "right"

        box = self._bar_box(rect, series, order, item, point, value, scale, horizontal)
        if box is None:
            return None
        position = (item.labels.position if item.labels else None) or "outEnd"
        centre_x = (box.left + box.right) / 2
        centre_y = (box.top + box.bottom) / 2
        if horizontal:
            # The bar runs sideways, so "end" is its far edge in x.
            end, base = (box.right, box.left) if value >= 0 else (box.left, box.right)
            if position == "ctr":
                return centre_x, centre_y, "centre"
            if position == "inEnd":
                return end, centre_y, "inside-x"
            if position == "inBase":
                return base, centre_y, "outside-x-flip"
            return end, centre_y, "outside-x"
        end, base = (box.top, box.bottom) if value >= 0 else (box.bottom, box.top)
        if position == "ctr":
            return centre_x, centre_y, "centre"
        if position == "inEnd":
            return centre_x, end, "inside-y"
        if position == "inBase":
            return centre_x, base, "inside-base-y"
        return centre_x, end, "outside-y"

    def _place_label(
        self,
        text: str,
        labels: _Labels,
        geometry: tuple[float, float, str],
        horizontal: bool,
    ) -> None:
        x, y, placement = geometry
        font = labels.font
        assert font is not None
        box = font.box
        lines = text.split("\n")
        width = max((font.width(line) for line in lines), default=0.0) + box.size
        block = box.line_height * len(lines)

        if placement == "centre":
            baseline = y + box.ink_centre - block + box.line_height
            left, align = x - width / 2, "ctr"
        elif placement == "outside-y":
            baseline = y - DATA_LABEL_GAP_PT - box.descent
            left, align = x - width / 2, "ctr"
        elif placement == "inside-y":
            baseline = y + DATA_LABEL_INNER_GAP_PT + box.ascent
            left, align = x - width / 2, "ctr"
        elif placement == "inside-base-y":
            baseline = y - DATA_LABEL_GAP_PT - box.descent
            left, align = x - width / 2, "ctr"
        elif placement == "right":
            baseline = y + box.ink_centre
            left = x + DATA_LABEL_LINE_GAP_EM * box.size
            align = "l"
        elif placement == "inside-x":
            baseline = y + box.ink_centre
            left = x - DATA_LABEL_INNER_GAP_PT - width
            align = "r"
        elif placement == "outside-x-flip":
            baseline = y + box.ink_centre
            left = x + DATA_LABEL_GAP_PT
            align = "l"
        else:  # outside-x
            baseline = y + box.ink_centre
            left = x + DATA_LABEL_GAP_PT
            align = "l"

        # Multi-line labels stack upwards from the anchor, so the *last* line is the one
        # nearest the bar; walk them in order from the first baseline.
        first = baseline - box.line_height * (len(lines) - 1)
        for index, line in enumerate(lines):
            self._text(
                self._label_body(line, font, align=align),
                left=left,
                width=width,
                baseline=first + index * box.line_height,
                box=box,
            )

    def _bar_box(
        self,
        rect: _Rect,
        series: list[_Series],
        order: int,
        item: _Series,
        point: int,
        value: float,
        scale: tuple[float, float, float],
        horizontal: bool,
    ) -> "_Rect | None":
        """The rectangle one bar occupies, recomputed for the label that sits on it."""
        categories = max((len(other.values) for other in series), default=0)
        if categories <= 0:
            return None
        grouping = self.plot.grouping or "clustered"
        stacked = grouping in ("stacked", "percentStacked")
        band = (rect.height if horizontal else rect.width) / categories
        gap_width = self.plot.gap_width
        if gap_width is None:
            gap_width = DEFAULT_GAP_WIDTH
        slots = 1 if stacked else len(series)
        size = band / max(slots + gap_width / 100.0, MIN_BAR_SLOTS)
        overlap = self.plot.overlap if self.plot.overlap is not None else 0.0
        step = size * (1.0 - overlap / 100.0)
        cluster = size + step * (slots - 1)
        band_start = (rect.top if horizontal else rect.left) + (
            (categories - 1 - point) if horizontal else point
        ) * band
        centre = band_start + band / 2
        offset = centre - cluster / 2 + (0 if stacked else order) * step

        start = 0.0
        if stacked:
            for earlier in series[:order]:
                earlier_value = _at(earlier.values, point) or 0.0
                if (earlier_value >= 0) == (value >= 0):
                    start += earlier_value
        end = start + value

        if horizontal:
            x0 = self._value_to_x(rect, start, scale)
            x1 = self._value_to_x(rect, end, scale)
            return _Rect(min(x0, x1), offset, max(x0, x1), offset + size)
        y0 = self._value_to_y(rect, start, scale)
        y1 = self._value_to_y(rect, end, scale)
        return _Rect(offset, min(y0, y1), offset + size, max(y0, y1))

    def _draw_legend(self, rect: _Rect, series: list[_Series]) -> None:
        position = self._legend_position()
        if position is None or not series:
            return
        legend = self.chart.legend
        deleted = legend.deleted_entries if legend else set()
        entries = [
            (index, item)
            for index, item in enumerate(series)
            if index not in deleted and item.name
        ]
        if not entries:
            return

        font = self._legend_font()
        box = font.box
        swatch = LEGEND_SWATCH_EM * box.size
        gap = LEGEND_SWATCH_GAP_EM * box.size

        if position in ("b", "t", "tr"):
            widths = [swatch + gap + font.width(item.name or "") for _, item in entries]
            total = (
                sum(widths)
                + LEGEND_ENTRY_GAP_EM * box.size * (len(entries) - 1)
                + LEGEND_HORIZONTAL_LEAD_EM * box.size
            )
            baseline = (
                self.frame.bottom - LEGEND_BOTTOM_BASELINE_EM * box.size
                if position == "b"
                else self.frame.top + LEGEND_TOP_BASELINE_EM * box.size
            )
            x = (
                self.frame.left
                + (self.frame.width - total) / 2
                + LEGEND_HORIZONTAL_LEAD_EM * box.size
            )
            for (_, item), width in zip(entries, widths):
                self._legend_entry(item, x, baseline, swatch, gap, font)
                x += width + LEGEND_ENTRY_GAP_EM * box.size
            return

        # A side legend sits one lead gap outside the plot area.  Measured 15.996 pt at
        # 10 pt with the legend on the right, and the band on the left came out exactly
        # the same width, so the left case mirrors it against the frame edge.
        if position == "l":
            # Measured 10.996 pt from the frame's left edge in the legend-l probe, which
            # is the plain edge inset and not the 6.5 pt the label column starts at.
            x = self.frame.left + EDGE_INSET_PT
        else:
            x = rect.right + LEGEND_SIDE_LEAD_EM * box.size
        # A stacked legend is centred on the frame and each entry is centred in its row.
        # Measured against both bar charts in real-financial-report.pptx: baselines land
        # within 0.18 pt, where treating the row like the horizontal band's off-centre
        # line was 5.7 pt out.
        pitch = LEGEND_ROW_PITCH_EM * box.size
        y = self.frame.top + (self.frame.height - pitch * len(entries)) / 2
        for _, item in entries:
            self._legend_entry(item, x, y + pitch / 2 + box.ink_centre, swatch, gap, font)
            y += pitch

    def _legend_entry(
        self,
        item: _Series,
        x: float,
        baseline: float,
        swatch: float,
        gap: float,
        font: ChartFont,
    ) -> None:
        box = font.box
        centre = baseline - box.ink_centre
        self._rect(
            _Rect(x, centre - swatch / 2, x + swatch, centre + swatch / 2),
            fill=item.fill,
            outline=None,
        )
        body = self._label_body(item.name or "", font, align="l")
        self._text(
            body,
            left=x + swatch + gap,
            width=font.width(item.name or "") + box.size,
            baseline=baseline,
            box=box,
        )

    # -- primitives ---------------------------------------------------------------------

    def _value_to_y(self, rect: _Rect, value: float, scale: tuple[float, float, float]) -> float:
        minimum, maximum, _ = scale
        span = maximum - minimum
        if span <= 0:
            return rect.bottom
        return rect.bottom - (value - minimum) / span * rect.height

    def _value_to_x(self, rect: _Rect, value: float, scale: tuple[float, float, float]) -> float:
        minimum, maximum, _ = scale
        span = maximum - minimum
        if span <= 0:
            return rect.left
        return rect.left + (value - minimum) / span * rect.width

    def _axis_outline(self, outline: m.Outline | s.SourceOutline | None) -> m.Outline:
        resolved = self._resolve_outline(outline) if outline is not None else None
        if resolved is not None and resolved.fill is not None:
            return resolved
        return m.Outline(
            width=DEFAULT_AXIS_LINE_EMU,
            fill=m.SolidFill(color=m.ResolvedColor(hex=DEFAULT_AXIS_COLOR)),
        )

    def _rect(
        self,
        box: _Rect,
        *,
        fill: m.Fill | None,
        outline: m.Outline | None,
        preset: str = "rect",
    ) -> None:
        self.elements.append(
            m.ShapeElement(
                transform=m.Transform(
                    offset_x=box.left * EMU_PER_POINT,
                    offset_y=box.top * EMU_PER_POINT,
                    extent_width=box.width * EMU_PER_POINT,
                    extent_height=box.height * EMU_PER_POINT,
                ),
                geometry=m.PresetGeometry(preset=preset),
                fill=fill,
                outline=outline,
            )
        )

    def _polyline(
        self, points: list[tuple[float, float]], outline: m.Outline, *, smooth: bool
    ) -> None:
        """A series' line, as one custom-geometry path in the frame's own space.

        Emitted as a single element rather than a segment per pair so the join between
        segments is a real line join -- drawing them separately leaves a notch at every
        vertex once the stroke is 2 pt wide, which is what a real chart uses.
        """
        left = min(x for x, _ in points)
        top = min(y for _, y in points)
        right = max(x for x, _ in points)
        bottom = max(y for _, y in points)
        width = max(right - left, 1e-6)
        height = max(bottom - top, 1e-6)

        local = [((x - left), (y - top)) for x, y in points]
        commands = _path_commands(local, smooth)
        self.elements.append(
            m.ShapeElement(
                transform=m.Transform(
                    offset_x=left * EMU_PER_POINT,
                    offset_y=top * EMU_PER_POINT,
                    extent_width=width * EMU_PER_POINT,
                    extent_height=height * EMU_PER_POINT,
                ),
                geometry=m.CustomGeometry(
                    paths=[
                        # The path's own space, which the renderer scales onto the shape
                        # box: the commands below are in points, so the extents must be
                        # too.  Giving them in EMU collapses the line to nothing.
                        m.CustomGeometryPath(
                            width=width, height=height, commands=commands
                        )
                    ]
                ),
                fill=m.NoFill(),
                outline=outline,
            )
        )

    def _line(self, x0: float, y0: float, x1: float, y1: float, outline: m.Outline) -> None:
        self.elements.append(
            m.ConnectorElement(
                transform=m.Transform(
                    offset_x=min(x0, x1) * EMU_PER_POINT,
                    offset_y=min(y0, y1) * EMU_PER_POINT,
                    extent_width=abs(x1 - x0) * EMU_PER_POINT,
                    extent_height=abs(y1 - y0) * EMU_PER_POINT,
                ),
                geometry=m.PresetGeometry(preset="line"),
                outline=outline,
            )
        )

    def _label_body(self, text: str, font: ChartFont, *, align: str) -> m.TextBody:
        return m.TextBody(
            paragraphs=[
                m.Paragraph(
                    runs=[
                        m.TextRun(
                            text=text,
                            properties=m.RunProperties(
                                font_size=font.size,
                                font_family=font.family,
                                color=self.style.color,
                            ),
                        )
                    ],
                    properties=m.ParagraphProperties(alignment=align),  # type: ignore[arg-type]
                )
            ],
            body_properties=CHART_TEXT_BODY,
        )

    def _text(
        self, body: m.TextBody, *, left: float, width: float, baseline: float, box: FontBox
    ) -> None:
        """Place a one-line text box so its baseline lands where the layout asked.

        The box's own top is derived from the renderer's first-baseline rule rather than
        assumed, so chart text stays aligned with the rest of the deck if that rule moves.
        """
        top = baseline - box.first_baseline
        self.elements.append(
            m.ShapeElement(
                transform=m.Transform(
                    offset_x=left * EMU_PER_POINT,
                    offset_y=top * EMU_PER_POINT,
                    extent_width=max(width, 1.0) * EMU_PER_POINT,
                    extent_height=(box.line_height * 1.5) * EMU_PER_POINT,
                ),
                geometry=m.PresetGeometry(preset="rect"),
                fill=None,
                outline=None,
                text_body=body,
            )
        )


#: Chart text sits in a box with no inset and no wrapping: the layout already decided
#: where every string goes, so letting the text engine re-wrap it would move it.
CHART_TEXT_BODY = m.BodyProperties(
    anchor="t",
    margin_left=0,
    margin_right=0,
    margin_top=0,
    margin_bottom=0,
    wrap="none",
)


#: Distinguishes "no title" from "not resolved yet" in the title cache.
_UNSET = object()


def _first_run_font(body: m.TextBody) -> tuple[str | None, float | None]:
    for paragraph in body.paragraphs:
        for run in paragraph.runs:
            return run.properties.font_family, run.properties.font_size
    return None, None


def _labels_shown(axis: c.SourceChartAxis | None) -> bool:
    if axis is None:
        return True
    if axis.delete:
        return False
    return (axis.tick_label_position or "nextTo") != "none"


#: ``c:symbol`` -> the preset geometry that draws it.  PowerPoint's marker shapes are
#: simple enough that the preset catalogue already has every one.
_MARKER_PRESETS = {
    "circle": "ellipse",
    "dot": "ellipse",
    "square": "rect",
    "diamond": "diamond",
    "triangle": "triangle",
    "x": "mathMultiply",
    "plus": "mathPlus",
    "star": "star5",
    "dash": "rect",
}


def _split_runs(
    points: list["tuple[float, float] | None"],
) -> list[list[tuple[float, float]]]:
    """Split a series at its blanks, so ``dispBlanksAs="gap"`` really leaves a gap."""
    runs: list[list[tuple[float, float]]] = []
    current: list[tuple[float, float]] = []
    for point in points:
        if point is None:
            if current:
                runs.append(current)
            current = []
        else:
            current.append(point)
    if current:
        runs.append(current)
    return runs


def _path_commands(points: list[tuple[float, float]], smooth: bool) -> str:
    """SVG path data for one run of points, straight or smoothed.

    ``c:smooth`` is drawn as a Catmull-Rom spline converted to cubic Béziers, which is
    what PowerPoint's own curve through the points looks like; the probe's smoothed line
    has control points that are *not* collinear with its vertices, so it is a real spline
    and not a polyline.  The exact tension PowerPoint uses was **not** measured.
    """
    parts = [f"M {points[0][0]:.4f} {points[0][1]:.4f}"]
    if not smooth or len(points) < 3:
        for x, y in points[1:]:
            parts.append(f"L {x:.4f} {y:.4f}")
        return " ".join(parts)

    for index in range(len(points) - 1):
        p0 = points[index - 1] if index > 0 else points[index]
        p1 = points[index]
        p2 = points[index + 1]
        p3 = points[index + 2] if index + 2 < len(points) else p2
        c1 = (p1[0] + (p2[0] - p0[0]) / 6, p1[1] + (p2[1] - p0[1]) / 6)
        c2 = (p2[0] - (p3[0] - p1[0]) / 6, p2[1] - (p3[1] - p1[1]) / 6)
        parts.append(
            f"C {c1[0]:.4f} {c1[1]:.4f} {c2[0]:.4f} {c2[1]:.4f} {p2[0]:.4f} {p2[1]:.4f}"
        )
    return " ".join(parts)


def _at(values: list[float | None], index: int) -> float | None:
    """One series' value at a category index; series need not be the same length."""
    return values[index] if index < len(values) else None


def _percent_totals(series: list[_Series]) -> list[float]:
    length = max((len(item.values) for item in series), default=0)
    totals: list[float] = []
    for index in range(length):
        total = 0.0
        for item in series:
            value = _at(item.values, index)
            if value is not None:
                total += abs(value)
        totals.append(total)
    return totals


def _default_run(body: s.SourceTextBody | None) -> s.SourceRunProperties | None:
    """``c:txPr``'s first ``a:defRPr``.

    A ``c:txPr`` is a one-paragraph text body whose only purpose is to carry defaults, so
    the first ``a:pPr/a:defRPr`` is the whole of it.
    """
    if body is None:
        return None
    for paragraph in body.paragraphs:
        properties = paragraph.properties
        if properties is not None and properties.default_run_properties is not None:
            return properties.default_run_properties
    return None


def _text_size(body: s.SourceTextBody | None) -> float | None:
    """``c:txPr``'s ``a:defRPr@sz``, in points."""
    run = _default_run(body)
    return run.font_size if run is not None and run.font_size else None


def _text_typeface(body: s.SourceTextBody | None) -> str | None:
    """``c:txPr``'s ``a:defRPr/a:latin@typeface``, unexpanded.

    Every axis in ``real-financial-report.pptx`` names ``Arial`` here while the theme's
    minor face is something else, and measuring the labels in the theme face instead put
    the plot area 1.7 pt off.  A ``+mn-lt``-style pointer comes back as-is; expanding it
    needs the theme and happens in the resolver.
    """
    run = _default_run(body)
    return run.typeface if run is not None else None


def _title_size(title: c.SourceChartText | None) -> float:
    """A chart title's point size.

    PowerPoint draws an unstyled ``c:rich`` title at DrawingML's own default of 18 pt --
    measured, and the same size it gives a bare ``a:t`` on a slide -- not at the 10 pt the
    rest of a chart's text defaults to.
    """
    if title is not None and title.rich is not None:
        for paragraph in title.rich.paragraphs:
            for run in paragraph.runs:
                if run.properties is not None and run.properties.font_size:
                    return run.properties.font_size
            if (
                paragraph.properties is not None
                and paragraph.properties.default_run_properties is not None
                and paragraph.properties.default_run_properties.font_size
            ):
                return paragraph.properties.default_run_properties.font_size
    return 18.0


def accent_colors(resolve) -> list[m.ResolvedColor]:
    """The theme accent cycle a series falls back to when it has no fill of its own."""
    out: list[m.ResolvedColor] = []
    for key in ACCENT_KEYS:
        color = resolve(key)
        if color is not None:
            out.append(color)
    return out


def default_font_size(chart: c.SourceChart) -> float:
    return _text_size(chart.text_properties) or DEFAULT_CHART_FONT_PT


__all__ = [
    "CHART_TEXT_BODY",
    "ChartBuilder",
    "ChartStyle",
    "accent_colors",
    "default_font_size",
    "font_box",
    "format_number",
    "nice_axis_scale",
    "text_width",
]
