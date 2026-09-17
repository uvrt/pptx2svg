"""Reader for ``c:chartSpace`` -- a chart part, as data plus styling.

Unlike SmartArt, **PowerPoint caches nothing**.  A ``c:chartSpace`` holds the numbers,
the series styling and the axis *preferences*; the axis range, the tick interval, the
plot rectangle and the legend placement are all computed by whoever draws it.  So this
reader's job is to surface every fact the layout needs and to invent none of them --
which is why almost every field here is ``None``-able: ``<c:overlap/>`` absent and
``<c:overlap val="0"/>`` mean different things to a bar chart.

Three traps that cost real time, all of them visible in the fixtures:

* **Chart booleans default to true.**  ``<c:delete/>`` with no ``val`` deletes the axis,
  and ``<c:overlay/>`` overlays the legend.  This is the opposite of the ``is_true``
  used everywhere else in this reader, hence :func:`_flag`.
* **Categories are not always ``c:strRef``.**  ``real-financial-report.pptx`` writes
  ``c:multiLvlStrRef``, and a category axis of dates or numbers arrives as ``c:numRef``.
  All four spellings resolve through :func:`_string_points`.
* **Always read the cache, never the formula.**  ``c:f`` names a range in the embedded
  workbook; ``c:numCache``/``c:strCache`` is the value PowerPoint last saw there, and
  it is what PowerPoint itself draws.  Parsing the workbook is a separate project.

Missing points are *not* zeros.  ``c:ptCount`` declares the length and the ``c:pt``
children are sparse, so a gap is a blank cell -- which ``c:dispBlanksAs`` says to draw as
a gap, a zero or a span, three visibly different pictures.  Blanks therefore survive as
``None`` in :attr:`SourceChartSeries.values` rather than being flattened early.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Literal
from xml.etree.ElementTree import Element

from ..xmlutil import (
    attr,
    child,
    children,
    int_attr,
    local_name,
    num_attr,
)
from .drawing import parse_fill, parse_line, parse_outline
from .source import (
    SourceColorMap,
    SourceFill,
    SourceOutline,
    SourceTextBody,
)
from .text import parse_text_body

#: The ``c:*Chart`` group elements this reader recognises.  3-D variants are read as
#: their 2-D equivalents by :func:`flat_chart_kind`; the renderer decides what to do
#: with that, and a flat bar chart beats an empty frame.
CHART_GROUP_ELEMENTS = (
    "barChart",
    "bar3DChart",
    "lineChart",
    "line3DChart",
    "areaChart",
    "area3DChart",
    "pieChart",
    "pie3DChart",
    "doughnutChart",
    "ofPieChart",
    "radarChart",
    "scatterChart",
    "bubbleChart",
    "stockChart",
    "surfaceChart",
    "surface3DChart",
)

_THREE_D_EQUIVALENT = {
    "bar3DChart": "barChart",
    "line3DChart": "lineChart",
    "area3DChart": "areaChart",
    "pie3DChart": "pieChart",
    "surface3DChart": "surfaceChart",
}

#: Every group element that is a 3-D spelling.  :func:`flat_chart_kind` erases this, and
#: the difference is not cosmetic: a 3-D value axis is **not padded**, so a chart that
#: forgets which spelling it came from draws the wrong numbers.  See
#: :func:`~pptx2svg.resolve.chart.nice_axis_scale` and ROADMAP.md 3.4.
THREE_D_CHART_KINDS = frozenset(_THREE_D_EQUIVALENT)

#: ``c:catAx`` / ``c:valAx`` / ``c:dateAx`` / ``c:serAx`` -- the four axis elements.
AXIS_ELEMENTS = ("catAx", "valAx", "dateAx", "serAx")


def flat_chart_kind(kind: str) -> str:
    """The 2-D chart a 3-D one degrades to; every other name is returned unchanged."""
    return _THREE_D_EQUIVALENT.get(kind, kind)


def is_three_d_kind(kind: str) -> bool:
    """Whether this group element was authored as a 3-D spelling.

    Five of them, of which four draw: ``surface3DChart`` is deferred along with the 2-D
    surface it maps to, and refuses rather than flattening.
    """
    return kind in THREE_D_CHART_KINDS


# --------------------------------------------------------------------------------------
# Source types
# --------------------------------------------------------------------------------------


@dataclass
class SourceChartText:
    """``c:tx`` / ``c:title`` rich text, or the cached string behind a reference.

    A title may be either literal DrawingML (``c:rich``) or a spreadsheet reference whose
    cached value is a plain string, and a series name is nearly always the latter.  Both
    are carried so the resolver can style rich text properly and still have something to
    draw when only the cache exists.
    """

    #: ``c:rich`` -- full DrawingML, styled like any other text body.
    rich: SourceTextBody | None = None
    #: The cached plain string from ``c:strRef``/``c:v``.
    cached: str | None = None

    @property
    def plain(self) -> str | None:
        if self.cached is not None:
            return self.cached
        if self.rich is None:
            return None
        text = "".join(
            run.text for paragraph in self.rich.paragraphs for run in paragraph.runs
        )
        return text or None


@dataclass
class SourceChartDataPoint:
    """``c:dPt`` -- a per-point override of the series' own styling."""

    index: int
    fill: SourceFill | None = None
    outline: SourceOutline | None = None
    invert_if_negative: bool | None = None
    #: ``c:explosion`` -- pie/doughnut slice offset, in percent of the radius.
    explosion: float | None = None


@dataclass
class SourceChartDataLabels:
    """``c:dLbls`` -- what to print on or beside each point."""

    delete: bool | None = None
    show_value: bool | None = None
    show_category_name: bool | None = None
    show_series_name: bool | None = None
    show_percent: bool | None = None
    show_bubble_size: bool | None = None
    show_legend_key: bool | None = None
    show_leader_lines: bool | None = None
    position: str | None = None
    number_format: str | None = None
    number_format_source_linked: bool | None = None
    text_properties: SourceTextBody | None = None
    fill: SourceFill | None = None
    outline: SourceOutline | None = None
    #: ``c:dLbl`` children, keyed by ``c:idx``.
    overrides: dict[int, "SourceChartDataLabels"] = field(default_factory=dict)


@dataclass
class SourceChartMarker:
    """``c:marker`` -- the symbol drawn at each point of a line or scatter series."""

    symbol: str | None = None
    #: ``c:size`` in points.
    size: float | None = None
    fill: SourceFill | None = None
    outline: SourceOutline | None = None


@dataclass
class SourceChartLines:
    """One ``c:serLines`` / ``c:hiLowLines`` / ``c:dropLines``.

    The element's *presence* is the switch and its ``c:spPr`` is only the styling, so a
    bare ``<c:serLines/>`` has to be told apart from an absent one: an ofPie probe with no
    element drew no connector at all, and one with a bare element drew two black 0.5 pt
    tangents.  Reading the outline alone would collapse the two.
    """

    outline: SourceOutline | None = None


@dataclass
class SourceChartUpDownBars:
    """``c:upDownBars`` -- the open-to-close body of a stock chart."""

    gap_width: float | None = None
    up_fill: SourceFill | None = None
    up_outline: SourceOutline | None = None
    down_fill: SourceFill | None = None
    down_outline: SourceOutline | None = None


@dataclass
class SourceChartSeries:
    """One ``c:ser``.

    ``values`` is dense and ``ptCount``-long with ``None`` for blank cells; ``categories``
    is the same length, padded with ``""``.  Callers that want "the numbers" can filter
    the ``None``s, and callers that care about gaps still have them.
    """

    index: int
    order: int
    name: SourceChartText | None = None
    categories: list[str] = field(default_factory=list)
    values: list[float | None] = field(default_factory=list)
    #: ``c:xVal`` / ``c:yVal`` / ``c:bubbleSize`` -- scatter and bubble carry their own
    #: coordinate lists rather than a category axis.
    x_values: list[float | None] | None = None
    bubble_sizes: list[float | None] | None = None
    #: ``c:formatCode`` off the value cache -- what PowerPoint prints in data labels.
    format_code: str | None = None
    category_format_code: str | None = None
    fill: SourceFill | None = None
    outline: SourceOutline | None = None
    invert_if_negative: bool | None = None
    smooth: bool | None = None
    marker: SourceChartMarker | None = None
    explosion: float | None = None
    data_points: list[SourceChartDataPoint] = field(default_factory=list)
    data_labels: SourceChartDataLabels | None = None


@dataclass
class SourceChartAxis:
    """One ``c:catAx`` / ``c:valAx`` / ``c:dateAx`` / ``c:serAx``."""

    kind: Literal["catAx", "valAx", "dateAx", "serAx"]
    axis_id: str | None = None
    position: str | None = None
    #: ``c:delete`` -- true means the axis exists for scaling but is not drawn.
    delete: bool = False
    orientation: str = "minMax"
    minimum: float | None = None
    maximum: float | None = None
    major_unit: float | None = None
    minor_unit: float | None = None
    log_base: float | None = None
    number_format: str | None = None
    number_format_source_linked: bool | None = None
    major_tick_mark: str | None = None
    minor_tick_mark: str | None = None
    tick_label_position: str | None = None
    #: ``c:majorGridlines`` present, and its line styling when it carries any.
    major_gridlines: bool = False
    major_gridline_outline: SourceOutline | None = None
    minor_gridlines: bool = False
    minor_gridline_outline: SourceOutline | None = None
    cross_axis_id: str | None = None
    crosses: str | None = None
    crosses_at: float | None = None
    #: ``c:crossBetween`` -- ``between`` puts categories in the middle of their band,
    #: ``midCat`` puts them on the tick.  Only meaningful on a value axis.
    cross_between: str | None = None
    #: ``c:lblOffset`` in percent; ``c:lblAlgn``.
    label_offset: float | None = None
    label_align: str | None = None
    tick_label_skip: int | None = None
    tick_mark_skip: int | None = None
    title: SourceChartText | None = None
    text_properties: SourceTextBody | None = None
    outline: SourceOutline | None = None


@dataclass
class SourceChartPlot:
    """One ``c:barChart``/``c:lineChart``/... group inside ``c:plotArea``.

    A plot area may hold several: that is what a combo chart is, and it is also how a
    secondary value axis arrives.  ``axis_ids`` is the group's own ``c:axId`` list, in
    document order, which is what ties it to its axes.
    """

    kind: str
    series: list[SourceChartSeries] = field(default_factory=list)
    axis_ids: list[str] = field(default_factory=list)
    grouping: str | None = None
    vary_colors: bool | None = None
    #: ``c:barDir`` -- ``col`` (vertical) or ``bar`` (horizontal).
    bar_direction: str | None = None
    #: ``c:gapWidth`` in percent of one bar's width.
    gap_width: float | None = None
    #: ``c:overlap`` in percent; negative separates clustered bars.
    overlap: float | None = None
    #: ``c:holeSize`` in percent, for doughnuts.
    hole_size: float | None = None
    first_slice_angle: float | None = None
    scatter_style: str | None = None
    radar_style: str | None = None
    marker: bool | None = None
    data_labels: SourceChartDataLabels | None = None
    #: ``c:bubbleScale`` in percent, ``c:sizeRepresents`` (``area`` or ``w``),
    #: ``c:showNegBubbles`` and ``c:bubble3D``.
    bubble_scale: float | None = None
    size_represents: str | None = None
    show_negative_bubbles: bool | None = None
    bubble_3d: bool | None = None
    #: ``c:ofPieType`` -- ``pie`` or ``bar`` -- and how the points are divided between the
    #: two plots: ``c:splitType`` with ``c:splitPos``, or ``c:custSplit``'s explicit list.
    of_pie_type: str | None = None
    split_type: str | None = None
    split_position: float | None = None
    custom_split: list[int] = field(default_factory=list)
    #: ``c:secondPieSize`` in percent of the first plot's radius.
    second_pie_size: float | None = None
    #: ``c:serLines`` -- the connector between an ofPie's two plots.
    series_lines: SourceChartLines | None = None
    #: ``c:hiLowLines`` and ``c:upDownBars`` -- a stock chart's two decorations.
    hi_low_lines: SourceChartLines | None = None
    up_down_bars: SourceChartUpDownBars | None = None


@dataclass
class SourceChartView3D:
    """``c:view3D`` -- the camera a 3-D chart's scene is drawn through.

    Every field is ``None``-able because **absent is not defaulted here either**, and for
    this element the difference is visible: a chart with no ``c:view3D`` at all and one
    with an empty ``<c:view3D/>`` are drawn differently by PowerPoint.  Measured on
    ``view3d-view``'s ``absent`` and ``empty`` probes -- the same chart on the same frame
    came back with a 94.08 pt value axis and a 155.04 pt one.

    What the element's absence means, measured on that deck rather than read off the
    schema: the ``absent`` probe is identical to 0.001 pt to the ``rotX=15 rotY=20
    depthPercent=100 rAngAx=0`` probe and to that same probe with ``perspective=30``,
    which is PowerPoint's own 3-D default *except* for ``rAngAx`` -- ECMA-376 gives that
    one a default of 1, and the picture PowerPoint draws is the one a 0 draws.

    Nothing reads these yet.  They are carried because the geometry they decide is
    measurable and unmeasured: see ROADMAP.md 3.4 for what each does to the plot
    rectangle, and why the count that follows from it is still open.
    """

    #: ``c:rotX`` -- pitch, in degrees, -90..90.  Positive tips the floor towards the
    #: viewer, and it is what takes the plot's height: a 195 pt frame's value axis ran
    #: 147.6 pt at ``rotX=0`` and 84.96 pt at ``rotX=90``.
    rot_x: float | None = None
    #: ``c:rotY`` -- yaw, in degrees, 0..360.  Measured to move the scene sideways only:
    #: seven values from 0 to 340 drew the same 127.4-127.7 pt axis.
    rot_y: float | None = None
    #: ``c:depthPercent`` -- the scene's depth as a percentage of its width, 20..2000.
    depth_percent: float | None = None
    #: ``c:hPercent`` -- the scene's height as a percentage of its width, 5..500.
    #: Measured to be exactly that ratio: 20/50/100/200 came back as 0.1995, 0.4975,
    #: 0.991 and 1.965 of the drawn width.  Absent, PowerPoint computes one from the
    #: frame, which is the open part of the geometry.
    h_percent: float | None = None
    #: ``c:rAngAx`` -- right-angle axes, which turns the perspective off.
    right_angle_axes: bool | None = None
    #: ``c:perspective`` -- 0..240, and ignored while :attr:`right_angle_axes` is true.
    #: Measured: ``rAngAx=1`` with ``perspective=120`` is identical to ``rAngAx=1``
    #: alone, and with ``rAngAx=0`` the same 120 takes the axis from 94.08 to 61.68 pt.
    perspective: float | None = None


@dataclass
class SourceChartLegend:
    position: str | None = None
    overlay: bool = False
    text_properties: SourceTextBody | None = None
    fill: SourceFill | None = None
    outline: SourceOutline | None = None
    #: ``c:legendEntry`` indices marked ``c:delete``.
    deleted_entries: set[int] = field(default_factory=set)


@dataclass
class SourceChart:
    """A whole ``c:chartSpace``."""

    plots: list[SourceChartPlot] = field(default_factory=list)
    axes: list[SourceChartAxis] = field(default_factory=list)
    title: SourceChartText | None = None
    auto_title_deleted: bool = False
    legend: SourceChartLegend | None = None
    #: ``c:view3D``.  Present only when the file states the element.
    view_3d: SourceChartView3D | None = None
    #: ``c:dispBlanksAs`` -- ``gap`` (default), ``zero`` or ``span``.
    display_blanks_as: str | None = None
    plot_visible_only: bool | None = None
    #: ``c:chartSpace/c:spPr`` -- the frame behind the whole chart.
    fill: SourceFill | None = None
    outline: SourceOutline | None = None
    #: ``c:plotArea/c:spPr`` -- the frame behind the plot rectangle only.
    plot_area_fill: SourceFill | None = None
    plot_area_outline: SourceOutline | None = None
    #: ``c:chartSpace/c:txPr`` -- the chart-wide text default.
    text_properties: SourceTextBody | None = None
    round_corners: bool | None = None
    #: ``c:clrMapOvr``.  A chart carries its own colour map and it must **not** be
    #: merged into the slide's: a ``schemeClr`` inside the chart resolves through this
    #: one or through the master's, never through a mixture of the two.
    color_map_override: SourceColorMap | None = None
    #: ``c:externalData@r:id`` -- the embedded workbook.  Recorded, not read.
    external_data_rel_id: str | None = None


# --------------------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------------------


def parse_chart_space(chart_space: Element | None) -> SourceChart | None:
    """``c:chartSpace`` -> :class:`SourceChart`, or ``None`` when there is no ``c:chart``."""
    if chart_space is None:
        return None
    chart = child(chart_space, "chart")
    if chart is None:
        return None

    plot_area = child(chart, "plotArea")
    title = child(chart, "title")

    return SourceChart(
        plots=[
            _plot(node)
            for node in children(plot_area)
            if local_name(node.tag) in CHART_GROUP_ELEMENTS
        ],
        axes=[
            _axis(node)
            for node in children(plot_area)
            if local_name(node.tag) in AXIS_ELEMENTS
        ],
        title=_text(title) if title is not None else None,
        # `c:autoTitleDeleted` is the only way to tell "no title" from "automatic title",
        # and an automatic title is what a single-series chart gets for free.
        auto_title_deleted=_flag(child(chart, "autoTitleDeleted"), default=False),
        legend=_legend(child(chart, "legend")),
        view_3d=_view_3d(child(chart, "view3D")),
        display_blanks_as=attr(child(chart, "dispBlanksAs"), "val"),
        plot_visible_only=_optional_flag(child(chart, "plotVisOnly")),
        fill=parse_fill(child(chart_space, "spPr")),
        outline=parse_outline(child(chart_space, "spPr")),
        plot_area_fill=parse_fill(child(plot_area, "spPr")),
        plot_area_outline=parse_outline(child(plot_area, "spPr")),
        text_properties=parse_text_body(child(chart_space, "txPr")),
        round_corners=_optional_flag(child(chart_space, "roundedCorners")),
        color_map_override=_color_map_override(child(chart_space, "clrMapOvr")),
        external_data_rel_id=_rel_id(child(chart_space, "externalData")),
    )


def _plot(node: Element) -> SourceChartPlot:
    kind = local_name(node.tag)
    return SourceChartPlot(
        kind=kind,
        series=sorted(
            (_series(ser, fallback_index) for fallback_index, ser in enumerate(children(node, "ser"))),
            key=lambda series: series.order,
        ),
        axis_ids=[value for value in (attr(ax, "val") for ax in children(node, "axId")) if value],
        grouping=attr(child(node, "grouping"), "val"),
        vary_colors=_optional_flag(child(node, "varyColors")),
        bar_direction=attr(child(node, "barDir"), "val"),
        gap_width=num_attr(child(node, "gapWidth"), "val"),
        overlap=num_attr(child(node, "overlap"), "val"),
        hole_size=num_attr(child(node, "holeSize"), "val"),
        first_slice_angle=num_attr(child(node, "firstSliceAng"), "val"),
        scatter_style=attr(child(node, "scatterStyle"), "val"),
        radar_style=attr(child(node, "radarStyle"), "val"),
        marker=_optional_flag(child(node, "marker")),
        data_labels=_data_labels(child(node, "dLbls")),
        bubble_scale=num_attr(child(node, "bubbleScale"), "val"),
        size_represents=attr(child(node, "sizeRepresents"), "val"),
        show_negative_bubbles=_optional_flag(child(node, "showNegBubbles")),
        bubble_3d=_optional_flag(child(node, "bubble3D")),
        of_pie_type=attr(child(node, "ofPieType"), "val"),
        split_type=attr(child(node, "splitType"), "val"),
        split_position=num_attr(child(node, "splitPos"), "val"),
        custom_split=_custom_split(child(node, "custSplit")),
        second_pie_size=num_attr(child(node, "secondPieSize"), "val"),
        series_lines=_chart_lines(child(node, "serLines")),
        hi_low_lines=_chart_lines(child(node, "hiLowLines")),
        up_down_bars=_up_down_bars(child(node, "upDownBars")),
    )


def _chart_lines(node: Element | None) -> SourceChartLines | None:
    """``c:serLines`` / ``c:hiLowLines``: present means draw, ``c:spPr`` only styles."""
    if node is None:
        return None
    return SourceChartLines(outline=parse_outline(child(node, "spPr")))


def _up_down_bars(node: Element | None) -> SourceChartUpDownBars | None:
    if node is None:
        return None
    up = child(node, "upBars")
    down = child(node, "downBars")
    return SourceChartUpDownBars(
        gap_width=num_attr(child(node, "gapWidth"), "val"),
        up_fill=parse_fill(child(up, "spPr")) if up is not None else None,
        up_outline=parse_outline(child(up, "spPr")) if up is not None else None,
        down_fill=parse_fill(child(down, "spPr")) if down is not None else None,
        down_outline=parse_outline(child(down, "spPr")) if down is not None else None,
    )


def _custom_split(node: Element | None) -> list[int]:
    """``c:custSplit`` -- the point indices that belong to the *second* plot."""
    if node is None:
        return []
    out = []
    for point in children(node, "secondPiePt"):
        index = int_attr(point, "val")  # `c:secondPiePt` carries the index on itself
        if index is not None and index >= 0:
            out.append(index)
    return out


def _series(ser: Element, fallback_index: int) -> SourceChartSeries:
    sp_pr = child(ser, "spPr")
    index = int_attr(child(ser, "idx"), "val")
    order = int_attr(child(ser, "order"), "val")

    values, format_code = _numeric_points(child(ser, "val"))
    y_values, y_format = _numeric_points(child(ser, "yVal"))
    if y_values:
        # A scatter series carries its dependent variable in `c:yVal`, not `c:val`.
        values, format_code = y_values, y_format

    x_values, x_format = _numeric_points(child(ser, "xVal"))
    categories, category_format = _string_points(child(ser, "cat"))
    if not categories and x_values:
        # Scatter with a numeric x: the labels a category axis would have shown are the
        # x values themselves, formatted.
        categories, category_format = _string_points(child(ser, "xVal"))

    bubble_sizes, _ = _numeric_points(child(ser, "bubbleSize"))

    return SourceChartSeries(
        index=index if index is not None else fallback_index,
        order=order if order is not None else fallback_index,
        name=_text(child(ser, "tx")),
        categories=_padded(categories, len(values)),
        values=values,
        x_values=x_values or None,
        bubble_sizes=bubble_sizes or None,
        format_code=format_code,
        category_format_code=category_format,
        fill=parse_fill(sp_pr),
        outline=parse_outline(sp_pr),
        invert_if_negative=_optional_flag(child(ser, "invertIfNegative")),
        smooth=_optional_flag(child(ser, "smooth")),
        marker=_marker(child(ser, "marker")),
        explosion=num_attr(child(ser, "explosion"), "val"),
        data_points=[point for point in map(_data_point, children(ser, "dPt")) if point],
        data_labels=_data_labels(child(ser, "dLbls")),
    )


def _data_point(d_pt: Element) -> SourceChartDataPoint | None:
    index = int_attr(child(d_pt, "idx"), "val")
    if index is None or index < 0:
        return None
    sp_pr = child(d_pt, "spPr")
    return SourceChartDataPoint(
        index=index,
        fill=parse_fill(sp_pr),
        outline=parse_outline(sp_pr),
        invert_if_negative=_optional_flag(child(d_pt, "invertIfNegative")),
        explosion=num_attr(child(d_pt, "explosion"), "val"),
    )


def _marker(marker: Element | None) -> SourceChartMarker | None:
    if marker is None:
        return None
    sp_pr = child(marker, "spPr")
    return SourceChartMarker(
        symbol=attr(child(marker, "symbol"), "val"),
        size=num_attr(child(marker, "size"), "val"),
        fill=parse_fill(sp_pr),
        outline=parse_outline(sp_pr),
    )


def _data_labels(d_lbls: Element | None) -> SourceChartDataLabels | None:
    if d_lbls is None:
        return None
    num_fmt = child(d_lbls, "numFmt")
    sp_pr = child(d_lbls, "spPr")
    labels = SourceChartDataLabels(
        delete=_optional_flag(child(d_lbls, "delete")),
        show_value=_optional_flag(child(d_lbls, "showVal")),
        show_category_name=_optional_flag(child(d_lbls, "showCatName")),
        show_series_name=_optional_flag(child(d_lbls, "showSerName")),
        show_percent=_optional_flag(child(d_lbls, "showPercent")),
        show_bubble_size=_optional_flag(child(d_lbls, "showBubbleSize")),
        show_legend_key=_optional_flag(child(d_lbls, "showLegendKey")),
        show_leader_lines=_optional_flag(child(d_lbls, "showLeaderLines")),
        position=attr(child(d_lbls, "dLblPos"), "val"),
        number_format=attr(num_fmt, "formatCode"),
        number_format_source_linked=_bool_attr(num_fmt, "sourceLinked"),
        text_properties=parse_text_body(child(d_lbls, "txPr")),
        fill=parse_fill(sp_pr),
        outline=parse_outline(sp_pr),
    )
    for d_lbl in children(d_lbls, "dLbl"):
        index = int_attr(child(d_lbl, "idx"), "val")
        override = _data_labels(d_lbl)
        if index is not None and index >= 0 and override is not None:
            labels.overrides[index] = override
    return labels


def _axis(node: Element) -> SourceChartAxis:
    scaling = child(node, "scaling")
    num_fmt = child(node, "numFmt")
    major_gridlines = child(node, "majorGridlines")
    minor_gridlines = child(node, "minorGridlines")
    title = child(node, "title")
    return SourceChartAxis(
        kind=local_name(node.tag),  # type: ignore[arg-type]
        axis_id=attr(child(node, "axId"), "val"),
        position=attr(child(node, "axPos"), "val"),
        delete=_flag(child(node, "delete"), default=False),
        orientation=attr(child(scaling, "orientation"), "val") or "minMax",
        minimum=num_attr(child(scaling, "min"), "val"),
        maximum=num_attr(child(scaling, "max"), "val"),
        major_unit=num_attr(child(node, "majorUnit"), "val"),
        minor_unit=num_attr(child(node, "minorUnit"), "val"),
        log_base=num_attr(child(scaling, "logBase"), "val"),
        number_format=attr(num_fmt, "formatCode"),
        number_format_source_linked=_bool_attr(num_fmt, "sourceLinked"),
        major_tick_mark=attr(child(node, "majorTickMark"), "val"),
        minor_tick_mark=attr(child(node, "minorTickMark"), "val"),
        tick_label_position=attr(child(node, "tickLblPos"), "val"),
        major_gridlines=major_gridlines is not None,
        major_gridline_outline=parse_outline(child(major_gridlines, "spPr")),
        minor_gridlines=minor_gridlines is not None,
        minor_gridline_outline=parse_outline(child(minor_gridlines, "spPr")),
        cross_axis_id=attr(child(node, "crossAx"), "val"),
        crosses=attr(child(node, "crosses"), "val"),
        crosses_at=num_attr(child(node, "crossesAt"), "val"),
        cross_between=attr(child(node, "crossBetween"), "val"),
        label_offset=num_attr(child(node, "lblOffset"), "val"),
        label_align=attr(child(node, "lblAlgn"), "val"),
        tick_label_skip=int_attr(child(node, "tickLblSkip"), "val"),
        tick_mark_skip=int_attr(child(node, "tickMarkSkip"), "val"),
        title=_text(title) if title is not None else None,
        text_properties=parse_text_body(child(node, "txPr")),
        outline=parse_outline(child(node, "spPr")),
    )


def _legend(legend: Element | None) -> SourceChartLegend | None:
    if legend is None:
        return None
    sp_pr = child(legend, "spPr")
    deleted: set[int] = set()
    for entry in children(legend, "legendEntry"):
        index = int_attr(child(entry, "idx"), "val")
        if index is not None and _flag(child(entry, "delete"), default=False):
            deleted.add(index)
    return SourceChartLegend(
        position=attr(child(legend, "legendPos"), "val"),
        overlay=_flag(child(legend, "overlay"), default=False),
        text_properties=parse_text_body(child(legend, "txPr")),
        fill=parse_fill(sp_pr),
        outline=parse_outline(sp_pr),
        deleted_entries=deleted,
    )


def _view_3d(view: Element | None) -> SourceChartView3D | None:
    """``c:view3D`` -> :class:`SourceChartView3D`, or ``None`` when the element is absent.

    An element with no children still returns a record, all of whose fields are ``None``:
    the two are different pictures, and only the record can tell them apart.
    """
    if view is None:
        return None
    return SourceChartView3D(
        rot_x=num_attr(child(view, "rotX"), "val"),
        rot_y=num_attr(child(view, "rotY"), "val"),
        depth_percent=num_attr(child(view, "depthPercent"), "val"),
        h_percent=num_attr(child(view, "hPercent"), "val"),
        right_angle_axes=_optional_flag(child(view, "rAngAx")),
        perspective=num_attr(child(view, "perspective"), "val"),
    )


def _color_map_override(clr_map_ovr: Element | None) -> SourceColorMap | None:
    """``c:clrMapOvr`` -- ``a:masterClrMapping`` means "inherit", so it reads as ``None``."""
    if clr_map_ovr is None:
        return None
    override = child(clr_map_ovr, "overrideClrMapping")
    if override is None:
        return None
    mapping = {local_name(key): value for key, value in override.attrib.items()}
    return SourceColorMap(mapping=mapping) if mapping else None


def _rel_id(node: Element | None) -> str | None:
    from ..xmlutil import ns_attr

    return ns_attr(node, "id")


# --------------------------------------------------------------------------------------
# Data caches
# --------------------------------------------------------------------------------------

#: A cache that claims more points than any deck plausibly charts is malformed or hostile;
#: an allocation bound is cheaper than trusting ``c:ptCount``.
MAX_CACHE_POINTS = 100_000


def _cache(reference: Element | None, kind: str) -> Element | None:
    """The populated cache under ``c:numRef``/``c:strRef``/``c:multiLvlStrRef``/``c:*Lit``.

    A reference may carry an empty cache alongside a literal, and a literal may appear
    with no reference at all, so the first candidate holding points wins and an empty
    one is only used when nothing else exists.
    """
    if reference is None:
        return None
    candidates = [
        child(child(reference, f"{kind}Ref"), f"{kind}Cache"),
        child(reference, f"{kind}Lit"),
        child(reference, f"{kind}Cache"),
    ]
    for candidate in candidates:
        if candidate is not None and child(candidate, "pt") is not None:
            return candidate
    for candidate in candidates:
        if candidate is not None:
            return candidate
    return None


def _point_count(cache: Element, points: list[Element] | None = None) -> int:
    """How long the series is.

    ``c:ptCount`` is authoritative but is ``minOccurs="0"`` and not always sane, and the
    ``c:pt`` indices can run past it in files written by other tools, so the length is the
    larger of the two -- clamped, because ``ptCount`` is attacker-controlled.

    ``points`` names where the ``c:pt`` children actually are.  A ``c:multiLvlStrCache``
    holds none directly -- they sit one level down inside ``c:lvl`` -- so without it a
    cache that omits ``c:ptCount`` measures as empty and every category label is lost.
    """
    highest = -1
    for point in children(cache, "pt") if points is None else points:
        index = int_attr(point, "idx")
        if index is not None and 0 <= index < MAX_CACHE_POINTS:
            highest = max(highest, index)
    present = highest + 1

    declared = int_attr(child(cache, "ptCount"), "val")
    if declared is None or declared < 0 or declared > MAX_CACHE_POINTS:
        return present
    return min(max(declared, present), MAX_CACHE_POINTS)


def _numeric_points(reference: Element | None) -> tuple[list[float | None], str | None]:
    """A dense value list with ``None`` for blanks, plus the cache's ``c:formatCode``."""
    cache = _cache(reference, "num")
    if cache is None:
        return [], None

    count = _point_count(cache)
    values: list[float | None] = [None] * count
    for point in children(cache, "pt"):
        index = int_attr(point, "idx")
        if index is None or not 0 <= index < count:
            continue
        raw = (child(point, "v").text or "") if child(point, "v") is not None else ""
        try:
            value = float(raw.strip())
        except (TypeError, ValueError):
            continue
        if value == value and value not in (float("inf"), float("-inf")):
            values[index] = value

    format_node = child(cache, "formatCode")
    format_code = (format_node.text or "").strip() if format_node is not None else ""
    return values, format_code or None


def _string_points(reference: Element | None) -> tuple[list[str], str | None]:
    """Category labels, from whichever of the four cache spellings the file uses.

    ``c:multiLvlStrRef`` holds one ``c:lvl`` per level of a hierarchical category axis,
    **innermost first**.  Only the innermost is used here: drawing the outer levels as
    grouped bands under the axis is a separate piece of work, and taking level 0 is what
    PowerPoint labels each individual bar with.
    """
    if reference is None:
        return [], None

    cache = _cache(reference, "str")
    if cache is None or child(cache, "pt") is None:
        multi = child(child(reference, "multiLvlStrRef"), "multiLvlStrCache")
        if multi is None:
            multi = child(reference, "multiLvlStrLit")
        if multi is not None:
            levels = children(multi, "lvl")
            if levels:
                count = _point_count(multi, children(levels[0], "pt"))
                return _level_strings(levels[0], count), None

    if cache is None:
        values, format_code = _numeric_points(reference)
        return [_default_number_text(value) for value in values], format_code

    if child(cache, "pt") is None and _cache(reference, "num") is not None:
        values, format_code = _numeric_points(reference)
        return [_default_number_text(value) for value in values], format_code

    return _level_strings(cache, _point_count(cache)), None


def _level_strings(container: Element, count: int) -> list[str]:
    values = [""] * count
    for point in children(container, "pt"):
        index = int_attr(point, "idx")
        if index is None or not 0 <= index < count:
            continue
        node = child(point, "v")
        values[index] = "".join(node.itertext()) if node is not None else ""
    return values


def _default_number_text(value: float | None) -> str:
    if value is None:
        return ""
    if value == int(value):
        return str(int(value))
    return repr(round(value, 10))


def _padded(values: list[str], length: int) -> list[str]:
    if len(values) >= length:
        return values
    return values + [""] * (length - len(values))


def _text(node: Element | None) -> SourceChartText | None:
    """``c:tx`` or ``c:title`` -- rich DrawingML, a cached string, or both."""
    if node is None:
        return None
    rich = parse_text_body(child(node, "rich"))
    if rich is None:
        tx = child(node, "tx")
        if tx is not None:
            rich = parse_text_body(child(tx, "rich"))
            node = tx
    cached = _cached_string(node)
    if rich is None and cached is None:
        return None
    return SourceChartText(rich=rich, cached=cached)


def _cached_string(node: Element) -> str | None:
    cache = _cache(node, "str")
    if cache is not None:
        for point in children(cache, "pt"):
            value = child(point, "v")
            if value is not None:
                return "".join(value.itertext())
    literal = child(node, "v")
    if literal is not None:
        return "".join(literal.itertext())
    return None


# --------------------------------------------------------------------------------------
# Chart booleans
# --------------------------------------------------------------------------------------


def _flag(node: Element | None, *, default: bool) -> bool:
    """A ``CT_Boolean`` element: **present with no ``val`` means true**, not false.

    ``<c:delete/>`` deletes the axis and ``<c:overlay/>`` overlays the legend; the
    ``is_true`` used for ordinary DrawingML attributes would read both as false.
    """
    if node is None:
        return default
    value = attr(node, "val")
    if value is None:
        return True
    return value in ("1", "true", "True")


def _optional_flag(node: Element | None) -> bool | None:
    """As :func:`_flag`, but absence stays ``None`` so a default can be applied later."""
    if node is None:
        return None
    return _flag(node, default=True)


def _bool_attr(node: Element | None, name: str) -> bool | None:
    value = attr(node, name)
    if value is None:
        return None
    return value in ("1", "true", "True")
