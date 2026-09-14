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
from ..text.measure import DEFAULT_LINE_HEIGHT_RATIO

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

#: Where the legend text's baseline sits inside that band, from its top.  Measured
#: 1.448 line heights with the legend at the bottom and 1.403 at the top; 1.43 splits them
#: to within 0.35 pt.
LEGEND_BASELINE_LINES = 1.43

#: Legend swatch side and the gap after it, in ems.  Measured 5.4923 pt and 2.3711 pt at
#: 10 pt.
LEGEND_SWATCH_EM = 0.549
LEGEND_SWATCH_GAP_EM = 0.237
#: Padding either side of a side legend, and between entries in a horizontal one.
LEGEND_SIDE_LEAD_EM = 1.60
LEGEND_SIDE_TRAIL_EM = 1.01
LEGEND_ENTRY_GAP_EM = 0.5

#: The title band, and its baseline inside it, as multiples of the line height and the
#: ascent.  Only one title was measurable (18 pt Arial, in two probes and the fixture, all
#: agreeing): band 29.70 pt against a 20.11 pt line height, baseline 24.50 pt below the
#: frame top against a 16.30 pt ascent.  The two coefficients below reproduce those to
#: 0.46 pt and 0.06 pt, but they are a one-font fit and should be re-measured if a chart
#: with a differently sized title ever disagrees.
TITLE_BAND_LINES = 1.5
TITLE_BASELINE_ASCENTS = 1.5

#: Default chart text size, in points.  ECMA-376's chart default and what PowerPoint drew
#: for every axis label and legend entry with no ``c:txPr``.
DEFAULT_CHART_FONT_PT = 10.0

#: Axis, tick and gridline defaults when no ``c:spPr`` says otherwise.
DEFAULT_AXIS_LINE_EMU = 6350.0
DEFAULT_AXIS_COLOR = "#000000"

#: ``c:gapWidth`` when absent, in percent of one bar's width (ECMA-376 default).
DEFAULT_GAP_WIDTH = 150.0

#: The six theme accents a series cycles through when it has no fill of its own.
ACCENT_KEYS = ("accent1", "accent2", "accent3", "accent4", "accent5", "accent6")

#: How many major intervals PowerPoint aims for on an automatic value axis.  Fitted to
#: three real charts: 0..5 -> 0..6 by 1, 0..4285 -> 0..5000 by 1000, 0..1842 -> 0..2000
#: by 500.  All three need 5 and no other value reproduces all three.
DESIRED_TICKS = 5


# --------------------------------------------------------------------------------------
# Axis scaling
# --------------------------------------------------------------------------------------


def nice_axis_scale(
    data_minimum: float, data_maximum: float, desired_ticks: int = DESIRED_TICKS
) -> tuple[float, float, float]:
    """``(minimum, maximum, major_unit)`` for a value axis PowerPoint would draw itself.

    The domain always includes zero -- a bar that does not start at its axis is a
    different picture -- and the maximum is rounded *strictly* up, so a series topping out
    at exactly 5 gets an axis to 6 rather than one whose last bar touches the frame.  That
    last rule is not cosmetic: it is what ``authoring-integration.pptx`` does.
    """
    low = min(0.0, data_minimum)
    high = max(0.0, data_maximum)
    span = high - low

    if span <= 0 or not math.isfinite(span):
        # Every value zero (or unusable).  PowerPoint still draws an axis; 0..1 is the
        # smallest one that shows anything.
        return 0.0, 1.0, 1.0

    unit = _nice_number(span / max(1, desired_ticks))
    maximum = math.ceil(high / unit) * unit
    if maximum <= data_maximum:
        maximum += unit
    minimum = math.floor(low / unit) * unit
    if data_minimum < 0 and minimum >= data_minimum:
        minimum -= unit
    return minimum, maximum, unit


def _nice_number(raw: float) -> float:
    """Round an interval up to 1, 2, 5 or 10 times a power of ten."""
    magnitude = 10.0 ** math.floor(math.log10(raw))
    residual = raw / magnitude
    if residual <= 1:
        return magnitude
    if residual <= 2:
        return 2 * magnitude
    if residual <= 5:
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

    section = _format_section(format_code, value)
    if "%" in section:
        decimals = _decimals(section)
        return f"{value * 100:.{decimals}f}%"

    if not any(ch in section for ch in "#0"):
        return _general(value)

    decimals = _decimals(section)
    grouped = "," in _strip_literals(section)
    magnitude = abs(value) if section is not _first_section(format_code) else value
    text = f"{magnitude:,.{decimals}f}" if grouped else f"{magnitude:.{decimals}f}"

    if section is not _first_section(format_code) and value < 0:
        if "(" in section and ")" in section:
            return f"({text})"
        return f"-{text}" if "-" in section else text
    return text


def _general(value: float) -> str:
    if value == int(value):
        return str(int(value))
    return f"{round(value, 10):g}"


def _first_section(format_code: str) -> str:
    return _sections(format_code)[0]


def _format_section(format_code: str, value: float) -> str:
    sections = _sections(format_code)
    if value < 0 and len(sections) > 1:
        return sections[1]
    return sections[0]


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
#: Calibri's, which is the commonest chart face after the theme's own.
FALLBACK_ASCENT = 0.75
FALLBACK_DESCENT = 0.25


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
    metrics = metrics_for(family)
    if metrics is None:
        return 0.5 * size * len(text)
    total = 0.0
    for char in text:
        total += metrics.widths.get(char, metrics.default_width)
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
class _Series:
    name: str | None
    values: list[float | None]
    color: m.ResolvedColor
    fill: m.Fill | None
    outline: m.Outline | None
    format_code: str | None
    invert_if_negative: bool
    point_fills: dict[int, m.Fill] = field(default_factory=dict)
    point_outlines: dict[int, m.Outline] = field(default_factory=dict)


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
    ) -> None:
        self.chart = chart
        self.plot = plot
        self.frame = _Rect(0.0, 0.0, width_pt, height_pt)
        self.style = style
        self._resolve_fill = resolve_fill
        self._resolve_outline = resolve_outline
        self._resolve_text = resolve_text
        self.elements: list[m.SlideElement] = []
        self._title_cache: "tuple[m.TextBody, FontBox] | None | object" = _UNSET
        #: Set by _draw_background, consumed by _draw_gridlines: the plot rectangle is
        #: not known until the labels have been measured, so the fill has to wait.
        self._plot_area_fill: m.Fill | None = None

    # -- public -------------------------------------------------------------------------

    def build(self) -> tuple[list[m.SlideElement], m.ChartData]:
        series = self._series()
        categories = self._categories(series)
        value_axis = self._axis_for(1) or self._axis_of_kind("valAx")
        category_axis = self._axis_for(0) or self._axis_of_kind("catAx")

        scale = self._scale(series, value_axis)
        label_size = self._label_size(value_axis)
        tick_texts = self._tick_texts(scale, value_axis)

        plot_rect = self._plot_rect(tick_texts, categories, label_size)

        self._draw_background()
        self._draw_title()
        self._draw_gridlines(plot_rect, scale, value_axis)
        self._draw_bars(plot_rect, series, categories, scale)
        self._draw_axis_lines(plot_rect, value_axis, category_axis)
        self._draw_value_labels(plot_rect, scale, tick_texts, value_axis, label_size)
        self._draw_category_labels(plot_rect, categories, category_axis, label_size)
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

    def _series(self) -> list[_Series]:
        out: list[_Series] = []
        for index, source in enumerate(self.plot.series):
            color = self._series_color(source, index)
            fill = self._resolve_fill(source.fill)
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
            for point in source.data_points:
                point_fill = self._resolve_fill(point.fill)
                if point_fill is not None:
                    item.point_fills[point.index] = point_fill
                point_outline = self._resolve_outline(point.outline)
                if point_outline is not None:
                    item.point_outlines[point.index] = point_outline
            out.append(item)
        return out

    def _series_color(self, source: c.SourceChartSeries, index: int) -> m.ResolvedColor:
        fill = self._resolve_fill(source.fill)
        if isinstance(fill, m.SolidFill):
            return fill.color
        if self.style.accents:
            return self.style.accents[index % len(self.style.accents)]
        return m.ResolvedColor(hex="#4472C4")

    def _categories(self, series: list[_Series]) -> list[str]:
        for source in self.plot.series:
            if any(source.categories):
                return list(source.categories)
        longest = max((len(item.values) for item in series), default=0)
        # PowerPoint numbers unlabelled categories from 1.
        return [str(index + 1) for index in range(longest)]

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
            return 0.0, 1.0, 0.2

        numbers: list[float] = []
        if stacked:
            length = max((len(item.values) for item in series), default=0)
            for index in range(length):
                positive = sum(
                    value for item in series
                    for value in [item.values[index] if index < len(item.values) else None]
                    if value is not None and value > 0
                )
                negative = sum(
                    value for item in series
                    for value in [item.values[index] if index < len(item.values) else None]
                    if value is not None and value < 0
                )
                numbers.extend([positive, negative])
        else:
            numbers = [value for item in series for value in item.values if value is not None]

        if not numbers:
            numbers = [0.0]
        minimum, maximum, unit = nice_axis_scale(min(numbers), max(numbers))

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
        out: list[tuple[float, str]] = []
        # Guard against a hand-written major unit that would generate millions of ticks.
        count = int(round((maximum - minimum) / unit)) if unit > 0 else 0
        if count <= 0 or count > 1000:
            return [(minimum, format_number(minimum, format_code))]
        for index in range(count + 1):
            value = minimum + index * unit
            out.append((value, format_number(value, format_code)))
        return out

    def _axis_format(self, axis: c.SourceChartAxis | None) -> str | None:
        if axis is None:
            return None
        if axis.number_format and axis.number_format != "General":
            return axis.number_format
        # `sourceLinked` means "use the cell's format"; the series cache carries it.
        for source in self.plot.series:
            if source.format_code and source.format_code != "General":
                return source.format_code
        return None

    # -- layout -------------------------------------------------------------------------

    def _label_size(self, axis: c.SourceChartAxis | None) -> FontBox:
        size = self.style.font_size
        if axis is not None and axis.text_properties is not None:
            size = _text_size(axis.text_properties) or size
        return font_box(self.style.font_family, size)

    def _plot_rect(
        self,
        tick_texts: list[tuple[float, str]],
        categories: list[str],
        label: FontBox,
    ) -> _Rect:
        frame = self.frame
        value_axis = self._axis_for(1) or self._axis_of_kind("valAx")
        category_axis = self._axis_for(0) or self._axis_of_kind("catAx")

        show_values = _labels_shown(value_axis)
        show_categories = _labels_shown(category_axis)

        left = frame.left + EDGE_INSET_PT
        if show_values:
            widest = max(
                (text_width(text, self.style.font_family, label.size) for _, text in tick_texts),
                default=0.0,
            )
            left = (
                frame.left
                + FRAME_PADDING_PT
                + widest
                + label.descent
                + VALUE_LABEL_GAP_EM * label.size
            )

        right = frame.right - EDGE_INSET_PT
        top = frame.top + self._top_inset(label)
        bottom = frame.bottom - (
            FRAME_PADDING_PT + label.line_height + CATEGORY_LABEL_GAP_EM * label.size
            if show_categories
            else EDGE_INSET_PT
        )

        title = self._title_box()
        if title is not None:
            top += TITLE_BAND_LINES * title.line_height

        legend = self._legend_position()
        if legend is not None:
            legend_box = font_box(self.style.font_family, self._legend_size())
            band = LEGEND_BAND_LINES * legend_box.line_height
            if legend in ("b",):
                bottom -= band
            elif legend in ("t", "tr"):
                top += band
            elif legend == "r":
                right -= self._legend_side_width(legend_box)
            elif legend == "l":
                left += self._legend_side_width(legend_box)

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

    def _legend_side_width(self, box: FontBox) -> float:
        widest = max(
            (
                text_width(source.name.plain or "", self.style.font_family, box.size)
                for source in self.plot.series
                if source.name is not None
            ),
            default=0.0,
        )
        return (
            widest
            + (LEGEND_SIDE_LEAD_EM + LEGEND_SWATCH_EM + LEGEND_SWATCH_GAP_EM + LEGEND_SIDE_TRAIL_EM)
            * box.size
        )

    def _legend_position(self) -> str | None:
        legend = self.chart.legend
        if legend is None:
            return None
        position = legend.position or "r"
        return position if position in ("b", "t", "l", "r", "tr") else "r"

    def _legend_size(self) -> float:
        legend = self.chart.legend
        if legend is not None and legend.text_properties is not None:
            return _text_size(legend.text_properties) or self.style.font_size
        return self.style.font_size

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

    def _draw_background(self) -> None:
        fill = self._resolve_fill(self.chart.fill)
        outline = self._resolve_outline(self.chart.outline)
        if fill is not None and not isinstance(fill, m.NoFill) or outline is not None:
            self._rect(self.frame, fill=fill, outline=outline)

        plot_fill = self._resolve_fill(self.chart.plot_area_fill)
        if plot_fill is not None and not isinstance(plot_fill, m.NoFill):
            # Drawn later, once the rectangle is known; recorded here would be wrong.
            self._plot_area_fill = plot_fill

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
        if self._plot_area_fill is not None:
            self._rect(rect, fill=self._plot_area_fill, outline=None)
        if axis is None or not axis.major_gridlines:
            return
        outline = self._axis_outline(axis.major_gridline_outline)
        minimum, maximum, unit = scale
        count = int(round((maximum - minimum) / unit)) if unit > 0 else 0
        if count <= 0 or count > 1000:
            return
        for index in range(count + 1):
            value = minimum + index * unit
            y = self._value_to_y(rect, value, scale)
            # The axis line is drawn separately and would double up on the zero gridline.
            if abs(y - rect.bottom) < 0.01:
                continue
            self._line(rect.left, y, rect.right, y, outline)

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
        bar_size = band / (slots + gap_width / 100.0)
        overlap = self.plot.overlap if self.plot.overlap is not None else 0.0
        step = bar_size * (1.0 - overlap / 100.0)
        cluster = bar_size + step * (slots - 1)

        percent = grouping == "percentStacked"
        totals = _percent_totals(series) if percent else None

        for point in range(len(categories)):
            band_start = (rect.top if horizontal else rect.left) + point * band
            centre = band_start + band / 2
            positive_base = 0.0
            negative_base = 0.0
            for order, item in enumerate(series):
                value = item.values[point] if point < len(item.values) else None
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
        fill = item.point_fills.get(point, item.fill)
        outline = item.point_outlines.get(point, item.outline)
        if end < start and item.invert_if_negative and point not in item.point_fills:
            # PowerPoint draws a negative bar hollow unless the file opts out.
            fill = m.SolidFill(color=m.ResolvedColor(hex="#FFFFFF"))
            outline = outline or m.Outline(
                width=DEFAULT_AXIS_LINE_EMU,
                fill=m.SolidFill(color=item.color),
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

    def _draw_axis_lines(
        self,
        rect: _Rect,
        value_axis: c.SourceChartAxis | None,
        category_axis: c.SourceChartAxis | None,
    ) -> None:
        if value_axis is not None and not value_axis.delete:
            self._line(
                rect.left, rect.top, rect.left, rect.bottom, self._axis_outline(value_axis.outline)
            )
        if category_axis is not None and not category_axis.delete:
            self._line(
                rect.left,
                rect.bottom,
                rect.right,
                rect.bottom,
                self._axis_outline(category_axis.outline),
            )

    def _draw_value_labels(
        self,
        rect: _Rect,
        scale: tuple[float, float, float],
        tick_texts: list[tuple[float, str]],
        axis: c.SourceChartAxis | None,
        box: FontBox,
    ) -> None:
        if not _labels_shown(axis):
            return
        for value, text in tick_texts:
            if not text:
                continue
            y = self._value_to_y(rect, value, scale)
            # The label's own ink is centred on the tick.  Digits have no descender, so
            # their ink runs from the baseline to the cap height and half of that is the
            # offset.  Measured against PowerPoint this lands within 0.7 pt.
            baseline = y + box.ink_centre
            body = self._label_body(text, box.size, align="r")
            self._text(
                body,
                left=self.frame.left,
                width=rect.left - self.frame.left - box.descent
                - VALUE_LABEL_GAP_EM * box.size,
                baseline=baseline,
                box=box,
            )

    def _draw_category_labels(
        self,
        rect: _Rect,
        categories: list[str],
        axis: c.SourceChartAxis | None,
        box: FontBox,
    ) -> None:
        if not _labels_shown(axis) or not categories:
            return
        band = rect.width / len(categories)
        # The label line's descender bottom sits one frame padding above whatever is below
        # it -- the frame edge, or the legend band.  Measured to within 0.55 pt.
        region_bottom = self.frame.bottom - FRAME_PADDING_PT
        legend = self._legend_position()
        if legend == "b":
            legend_box = font_box(self.style.font_family, self._legend_size())
            region_bottom -= LEGEND_BAND_LINES * legend_box.line_height
        baseline = region_bottom - box.descent
        for index, text in enumerate(categories):
            if not text:
                continue
            body = self._label_body(text, box.size, align="ctr")
            self._text(
                body,
                left=rect.left + index * band,
                width=band,
                baseline=baseline,
                box=box,
            )

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

        box = font_box(self.style.font_family, self._legend_size())
        swatch = LEGEND_SWATCH_EM * box.size
        gap = LEGEND_SWATCH_GAP_EM * box.size
        band = LEGEND_BAND_LINES * box.line_height

        if position in ("b", "t", "tr"):
            widths = [
                swatch + gap + text_width(item.name or "", self.style.font_family, box.size)
                for _, item in entries
            ]
            total = sum(widths) + LEGEND_ENTRY_GAP_EM * box.size * (len(entries) - 1)
            band_top = (
                self.frame.bottom - FRAME_PADDING_PT - band
                if position == "b"
                else self.frame.top
            )
            baseline = band_top + LEGEND_BASELINE_LINES * box.line_height
            x = self.frame.left + (self.frame.width - total) / 2
            for (_, item), width in zip(entries, widths):
                self._legend_entry(item, x, baseline, swatch, gap, box)
                x += width + LEGEND_ENTRY_GAP_EM * box.size
            return

        # A side legend sits one lead gap outside the plot area.  Measured 15.996 pt at
        # 10 pt with the legend on the right, and the band on the left came out exactly
        # the same width, so the left case mirrors it against the frame edge.
        if position == "l":
            x = self.frame.left + FRAME_PADDING_PT
        else:
            x = rect.right + LEGEND_SIDE_LEAD_EM * box.size
        height = band * len(entries)
        y = self.frame.top + (self.frame.height - height) / 2
        for _, item in entries:
            baseline = y + LEGEND_BASELINE_LINES * box.line_height
            self._legend_entry(item, x, baseline, swatch, gap, box)
            y += band

    def _legend_entry(
        self,
        item: _Series,
        x: float,
        baseline: float,
        swatch: float,
        gap: float,
        box: FontBox,
    ) -> None:
        centre = baseline - box.ink_centre
        self._rect(
            _Rect(x, centre - swatch / 2, x + swatch, centre + swatch / 2),
            fill=item.fill,
            outline=None,
        )
        body = self._label_body(item.name or "", box.size, align="l")
        self._text(
            body,
            left=x + swatch + gap,
            width=text_width(item.name or "", self.style.font_family, box.size) + box.size,
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

    def _rect(self, box: _Rect, *, fill: m.Fill | None, outline: m.Outline | None) -> None:
        self.elements.append(
            m.ShapeElement(
                transform=m.Transform(
                    offset_x=box.left * EMU_PER_POINT,
                    offset_y=box.top * EMU_PER_POINT,
                    extent_width=box.width * EMU_PER_POINT,
                    extent_height=box.height * EMU_PER_POINT,
                ),
                geometry=m.PresetGeometry(preset="rect"),
                fill=fill,
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

    def _label_body(self, text: str, size: float, *, align: str) -> m.TextBody:
        return m.TextBody(
            paragraphs=[
                m.Paragraph(
                    runs=[
                        m.TextRun(
                            text=text,
                            properties=m.RunProperties(
                                font_size=size,
                                font_family=self.style.font_family,
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


def _percent_totals(series: list[_Series]) -> list[float]:
    length = max((len(item.values) for item in series), default=0)
    totals: list[float] = []
    for index in range(length):
        total = 0.0
        for item in series:
            value = item.values[index] if index < len(item.values) else None
            if value is not None:
                total += abs(value)
        totals.append(total)
    return totals


def _text_size(body: s.SourceTextBody | None) -> float | None:
    """``c:txPr``'s ``a:defRPr@sz``, in points."""
    if body is None:
        return None
    for paragraph in body.paragraphs:
        properties = paragraph.properties
        if properties is None or properties.default_run_properties is None:
            continue
        size = properties.default_run_properties.font_size
        if size:
            return size
    return None


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
