"""Charts: reading ``c:chartSpace``, and drawing what comes out of it.

Every layout number asserted here was read out of PowerPoint's own PDF export as an exact
vector coordinate, not eyeballed from a raster.  See ``src/pptx2svg/resolve/chart.py`` for
where each constant came from and what its residual against the measurement is.
"""

from __future__ import annotations

import math
import re
import zipfile
from xml.etree.ElementTree import fromstring

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_model
from pptx2svg import model as m
from pptx2svg.parse.chart import flat_chart_kind, parse_chart_space
from pptx2svg.resolve.chart import format_number, nice_axis_scale

C = 'xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart"'
A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'


def chart(body: str):
    return parse_chart_space(fromstring(f"<c:chartSpace {C} {A} {R}>{body}</c:chartSpace>"))


def bar(series: str, extra: str = "") -> str:
    return (
        f"<c:chart><c:plotArea><c:barChart><c:barDir val='col'/>{series}"
        f"</c:barChart>{extra}</c:plotArea></c:chart>"
    )


# -- Reading ---------------------------------------------------------------------------


def test_chart_space_without_a_chart_reads_as_nothing():
    assert chart("<c:date1904 val='0'/>") is None


def test_series_values_come_from_the_cache_not_the_formula():
    parsed = chart(
        bar(
            "<c:ser><c:idx val='0'/><c:order val='0'/><c:val><c:numRef>"
            "<c:f>Sheet1!$B$2:$B$4</c:f><c:numCache><c:formatCode>#,##0</c:formatCode>"
            "<c:ptCount val='3'/>"
            "<c:pt idx='0'><c:v>3</c:v></c:pt>"
            "<c:pt idx='1'><c:v>4</c:v></c:pt>"
            "<c:pt idx='2'><c:v>5</c:v></c:pt>"
            "</c:numCache></c:numRef></c:val></c:ser>"
        )
    )
    series = parsed.plots[0].series[0]
    assert series.values == [3.0, 4.0, 5.0]
    assert series.format_code == "#,##0"


def test_a_missing_point_is_a_blank_not_a_zero():
    parsed = chart(
        bar(
            "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='4'/>"
            "<c:pt idx='0'><c:v>1</c:v></c:pt>"
            "<c:pt idx='3'><c:v>2</c:v></c:pt>"
            "</c:numCache></c:numRef></c:val></c:ser>"
        )
    )
    assert parsed.plots[0].series[0].values == [1.0, None, None, 2.0]


def test_point_count_is_clamped_so_a_hostile_cache_cannot_allocate():
    parsed = chart(
        bar(
            "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='99999999999'/>"
            "<c:pt idx='0'><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        )
    )
    assert parsed.plots[0].series[0].values == [1.0]


def test_categories_read_from_a_multi_level_reference():
    # real-financial-report.pptx writes every category axis this way; a reader that only
    # looks at c:strRef finds no labels at all and silently draws a bare axis.
    parsed = chart(
        bar(
            "<c:ser><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:ptCount val='3'/>"
            "<c:lvl><c:pt idx='0'><c:v>Q1</c:v></c:pt>"
            "<c:pt idx='1'><c:v>Q2</c:v></c:pt>"
            "<c:pt idx='2'><c:v>Q3</c:v></c:pt></c:lvl>"
            "<c:lvl><c:pt idx='0'><c:v>FY24</c:v></c:pt></c:lvl>"
            "</c:multiLvlStrCache></c:multiLvlStrRef></c:cat>"
            "<c:val><c:numRef><c:numCache><c:ptCount val='3'/>"
            "<c:pt idx='0'><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        )
    )
    assert parsed.plots[0].series[0].categories == ["Q1", "Q2", "Q3"]


def test_numeric_categories_become_labels():
    parsed = chart(
        bar(
            "<c:ser><c:cat><c:numRef><c:numCache><c:ptCount val='2'/>"
            "<c:pt idx='0'><c:v>2023</c:v></c:pt>"
            "<c:pt idx='1'><c:v>2024</c:v></c:pt>"
            "</c:numCache></c:numRef></c:cat>"
            "<c:val><c:numRef><c:numCache><c:ptCount val='2'/>"
            "<c:pt idx='0'><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        )
    )
    assert parsed.plots[0].series[0].categories == ["2023", "2024"]


def test_a_chart_boolean_with_no_val_is_true():
    # <c:delete/> deletes the axis.  Reading it as false -- which the ordinary DrawingML
    # attribute rule would -- draws an axis PowerPoint does not.
    parsed = chart(
        "<c:chart><c:plotArea><c:barChart/>"
        "<c:valAx><c:axId val='1'/><c:delete/></c:valAx>"
        "<c:catAx><c:axId val='2'/><c:delete val='0'/></c:catAx>"
        "</c:plotArea></c:chart>"
    )
    assert [axis.delete for axis in parsed.axes] == [True, False]


def test_series_are_ordered_by_order_not_document_position():
    parsed = chart(
        bar(
            "<c:ser><c:idx val='0'/><c:order val='1'/><c:tx><c:v>second</c:v></c:tx></c:ser>"
            "<c:ser><c:idx val='1'/><c:order val='0'/><c:tx><c:v>first</c:v></c:tx></c:ser>"
        )
    )
    assert [s.name.plain for s in parsed.plots[0].series] == ["first", "second"]


def test_axis_ids_tie_a_plot_group_to_its_axes():
    parsed = chart(
        "<c:chart><c:plotArea>"
        "<c:barChart><c:axId val='11'/><c:axId val='22'/></c:barChart>"
        "<c:catAx><c:axId val='11'/></c:catAx><c:valAx><c:axId val='22'/></c:valAx>"
        "</c:plotArea></c:chart>"
    )
    assert parsed.plots[0].axis_ids == ["11", "22"]
    assert [axis.axis_id for axis in parsed.axes] == ["11", "22"]


def test_chart_colour_map_override_is_read_but_master_mapping_is_not():
    inherited = chart("<c:chart><c:plotArea/></c:chart><c:clrMapOvr><a:masterClrMapping/></c:clrMapOvr>")
    assert inherited.color_map_override is None

    overridden = chart(
        "<c:chart><c:plotArea/></c:chart>"
        "<c:clrMapOvr><a:overrideClrMapping bg1='dk1' tx1='lt1'/></c:clrMapOvr>"
    )
    assert overridden.color_map_override.mapping == {"bg1": "dk1", "tx1": "lt1"}


def test_three_d_chart_kinds_fall_back_to_their_flat_equivalent():
    assert flat_chart_kind("bar3DChart") == "barChart"
    assert flat_chart_kind("pie3DChart") == "pieChart"
    assert flat_chart_kind("barChart") == "barChart"


@pytest.mark.parametrize(
    "deck,part,kind,series_count",
    [
        ("authoring-integration.pptx", "ppt/charts/chart1.xml", "barChart", 1),
        ("real-financial-report.pptx", "ppt/charts/chart1.xml", "barChart", 2),
        ("real-financial-report.pptx", "ppt/charts/chart2.xml", "lineChart", 3),
        ("real-financial-report.pptx", "ppt/charts/chart3.xml", "barChart", 2),
        ("real-financial-report.pptx", "ppt/charts/chart4.xml", "doughnutChart", 1),
        ("real-financial-report.pptx", "ppt/charts/chart5.xml", "radarChart", 2),
    ],
)
def test_every_chart_in_the_corpus_reads(deck, part, kind, series_count):
    from tests.conftest import FIXTURE_DIR  # noqa: PLC0415

    with zipfile.ZipFile(FIXTURE_DIR / deck) as archive:
        parsed = parse_chart_space(fromstring(archive.read(part)))
    assert parsed is not None
    assert parsed.plots[0].kind == kind
    assert len(parsed.plots[0].series) == series_count
    assert all(series.values for series in parsed.plots[0].series)


# -- Axis scaling ----------------------------------------------------------------------


@pytest.mark.parametrize(
    "values,expected",
    [
        # authoring-integration.pptx: PowerPoint draws 0..6 by 1 for data topping out at 5.
        ([3, 4, 5], (0.0, 6.0, 1.0)),
        # real-financial-report chart1, series maxing at 4285: PowerPoint drew 0..5000 by 1000.
        ([3980, 4120, 4285, 465, 488, 512], (0.0, 5000.0, 1000.0)),
        # ...and chart3, maxing at 1842: 0..2000 by 500.
        ([1599, 1185, 663, 334, 1842, 1285, 814, 344], (0.0, 2000.0, 500.0)),
    ],
)
def test_axis_scale_matches_what_powerpoint_drew(values, expected):
    """All three come from reading gridline coordinates out of PowerPoint's PDF.

    Five intervals is the only target tick count that reproduces all three; four and six
    each get one of them wrong.
    """
    assert nice_axis_scale(min(values), max(values)) == expected


def test_the_maximum_is_rounded_strictly_up():
    # A series topping out at exactly the axis maximum would touch the frame; PowerPoint
    # adds an interval instead, which is why authoring-integration's axis reaches 6.
    assert nice_axis_scale(0.0, 5.0)[1] == 6.0
    assert nice_axis_scale(0.0, 4.9)[1] == 5.0


def test_a_negative_minimum_extends_the_axis_below_zero():
    minimum, maximum, unit = nice_axis_scale(-30.0, 120.0)
    assert minimum < 0 and maximum >= 120 and unit > 0
    assert minimum % unit == 0 and maximum % unit == 0


def test_a_flat_series_still_gets_a_usable_axis():
    minimum, maximum, unit = nice_axis_scale(0.0, 0.0)
    assert maximum > minimum and unit > 0


# -- Number formatting -----------------------------------------------------------------


@pytest.mark.parametrize(
    "value,code,expected",
    [
        (3.0, None, "3"),
        (3.0, "General", "3"),
        (3.5, "General", "3.5"),
        (4285.0, "#,##0", "4,285"),
        (4285.0, "0", "4285"),
        (0.125, "0.0%", "12.5%"),
        (-0.125, "0.0%", "-12.5%"),
        (1.5, '"$"0.00', "1.50"),
        # A negative value keeps its own sign when the code has no negative section, and
        # takes that section's decoration -- brackets, or an explicit minus -- when it has.
        (-1234.0, "#,##0", "-1,234"),
        (-5.0, "0.0", "-5.0"),
        (-1234.0, "#,##0;(#,##0)", "(1,234)"),
        (-3.0, "#,##0;-#,##0", "-3"),
        (-3.0, "#,##0;#,##0", "3"),
    ],
)
def test_number_formats(value, code, expected):
    assert format_number(value, code) == expected


# -- Drawing ---------------------------------------------------------------------------

#: Points from the frame's top-left, read out of PowerPoint's PDF export of
#: authoring-integration.pptx as exact vector coordinates.  The frame is
#: 299.2126 x 228.3465 pt.  The tolerance is 0.6 pt -- about one pixel at the 1280 px the
#: fidelity harness scores at -- and the residual actually achieved is noted alongside.
POWERPOINT = {
    "plot_left": 21.0725,           # ours 21.1092
    "plot_right": 288.2131,         # ours 288.2126
    "plot_top": 40.8025,            # ours 40.8031
    "plot_bottom": 179.2985,        # ours 179.0754
    "bar1_left": 47.7625,           # ours 47.8195
    "bar_width": 35.6187,           # ours 35.6138
    "bar1_top": 109.9918,           # ours 109.9392
    "title_baseline": 24.5518,      # ours 24.5261
    "category_baseline": 194.4718,  # ours 194.6152
}
TOLERANCE_PT = 0.6

#: Descents as a fraction of the em.  The chart's own labels are Aptos, its title Arial.
APTOS_DESCENT = 577 / 2048
ARIAL_DESCENT = 434 / 2048


def _chart_of(path):
    deck = convert_pptx_to_model(path.read_bytes())
    charts = [e for e in deck.slides[0].elements if isinstance(e, m.ChartElement)]
    assert len(charts) == 1
    return charts[0]


def _pt(emu):
    return emu / 12700.0


def _plot_rect(chart):
    """The plot rectangle, read back off the drawn lines, in points.

    Which line is the value axis and which the category axis depends on `barDir`, and a
    category axis floats away from the plot's edge once values go negative -- so the rect
    is taken from the *longest* line each way instead, which spans the plot whichever axis
    or gridline it happens to be.
    """
    lines = [c for c in chart.children if isinstance(c, m.ConnectorElement)]
    vertical = [line for line in lines if line.transform.extent_width == 0]
    horizontal = [line for line in lines if line.transform.extent_height == 0]
    assert vertical and horizontal
    down = max(vertical, key=lambda line: line.transform.extent_height).transform
    across = max(horizontal, key=lambda line: line.transform.extent_width).transform
    return (
        _pt(across.offset_x),
        _pt(down.offset_y),
        _pt(across.offset_x + across.extent_width),
        _pt(down.offset_y + down.extent_height),
    )


def _baseline(shape, descent_em=APTOS_DESCENT):
    """A one-line label's baseline, from the box the resolver placed it in.

    The renderer hangs the first baseline off the *bottom* of a 1.2 em line box, so the
    baseline is ``top + (1.2 - descent) * size``; see
    ``pptx2svg.text.measure._first_baseline_ratio``.
    """
    size = shape.text_body.paragraphs[0].runs[0].properties.font_size
    return _pt(shape.transform.offset_y) + (1.2 - descent_em) * size


def _labels(chart):
    return {
        "".join(run.text for p in child.text_body.paragraphs for run in p.runs): child
        for child in chart.children
        if isinstance(child, m.ShapeElement) and child.text_body is not None
    }


def _bars(chart):
    """Every drawn bar, left to right.

    The legend swatch is a filled rectangle too, and in the fixture it is even the same
    colour; it is an order of magnitude smaller, which is what separates them.
    """
    bars = [
        child
        for child in chart.children
        if isinstance(child, m.ShapeElement)
        and isinstance(child.fill, m.SolidFill)
        and child.text_body is None
        and child.transform.extent_width > 10 * 12700
    ]
    return sorted(bars, key=lambda bar: bar.transform.offset_x)


def test_the_plot_rectangle_matches_powerpoints(authoring):
    left, top, right, bottom = _plot_rect(_chart_of(authoring))
    assert left == pytest.approx(POWERPOINT["plot_left"], abs=TOLERANCE_PT)
    assert right == pytest.approx(POWERPOINT["plot_right"], abs=TOLERANCE_PT)
    assert top == pytest.approx(POWERPOINT["plot_top"], abs=TOLERANCE_PT)
    assert bottom == pytest.approx(POWERPOINT["plot_bottom"], abs=TOLERANCE_PT)


def test_the_bars_match_powerpoints(authoring):
    bars = _bars(_chart_of(authoring))
    assert len(bars) == 3

    assert _pt(bars[0].transform.offset_x) == pytest.approx(
        POWERPOINT["bar1_left"], abs=TOLERANCE_PT
    )
    assert _pt(bars[0].transform.offset_y) == pytest.approx(
        POWERPOINT["bar1_top"], abs=TOLERANCE_PT
    )
    for bar in bars:
        assert _pt(bar.transform.extent_width) == pytest.approx(
            POWERPOINT["bar_width"], abs=TOLERANCE_PT
        )
    # 3, 4 and 5 on a 6-unit axis, so the heights stand in 1 : 4/3 : 5/3.
    heights = [bar.transform.extent_height for bar in bars]
    assert heights[1] / heights[0] == pytest.approx(4 / 3, abs=0.01)
    assert heights[2] / heights[0] == pytest.approx(5 / 3, abs=0.01)


def test_the_title_and_category_baselines_match_powerpoints(authoring):
    labels = _labels(_chart_of(authoring))
    # The title takes the deck's plain-text-box face at 18 pt, not the chart's 10 pt
    # default: PowerPoint drew the axis labels in Aptos and this title in Arial.
    title = labels["Chart contract"]
    assert title.text_body.paragraphs[0].runs[0].properties.font_size == 18
    assert _baseline(title, ARIAL_DESCENT) == pytest.approx(
        POWERPOINT["title_baseline"], abs=TOLERANCE_PT
    )
    assert _baseline(labels["Reader"]) == pytest.approx(
        POWERPOINT["category_baseline"], abs=TOLERANCE_PT
    )


def test_axis_and_legend_text_takes_the_chart_default_not_the_deck_default(authoring):
    """10 pt in the theme's minor face, which is not what the title gets."""
    labels = _labels(_chart_of(authoring))
    for name in ("0", "6", "Reader", "Coverage"):
        properties = labels[name].text_body.paragraphs[0].runs[0].properties
        assert properties.font_size == 10
        assert properties.font_family == "Aptos"


def test_the_fixture_chart_lowers_to_bars_gridlines_and_labels(authoring):
    element = _chart_of(authoring)

    assert element.chart.kind == "barChart"
    assert element.chart.title == "Chart contract"
    assert element.chart.categories == ["Reader", "Writer", "Renderer"]
    assert [s.values for s in element.chart.series] == [[3.0, 4.0, 5.0]]
    assert element.chart.series[0].color.hex.upper() == "#F97316"
    assert element.chart.value_axis == m.ChartAxisScale(
        minimum=0.0, maximum=6.0, major_unit=1.0
    )
    assert element.chart.legend_position == "b"

    # Six gridlines plus two axis lines: the zero gridline is the category axis, drawn
    # once rather than twice.
    lines = [c for c in element.children if isinstance(c, m.ConnectorElement)]
    assert len(lines) == 8


def test_a_chart_renders_to_svg_rather_than_an_empty_frame(authoring):
    from pptx2svg import convert_pptx_to_svg

    svg = convert_pptx_to_svg(authoring.read_bytes())[0]
    assert "#f97316" in svg.lower()
    assert "Chart contract" in svg
    assert "Renderer" in svg


def test_a_chart_part_that_is_missing_warns_and_draws_an_empty_frame(authoring):
    from tests.deckbuilder import derive_deck

    frame = (
        "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='99' name='Broken chart'/>"
        "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
        "<p:xfrm><a:off x='0' y='0'/><a:ext cx='1000000' cy='1000000'/></p:xfrm>"
        "<a:graphic><a:graphicData "
        "uri='http://schemas.openxmlformats.org/drawingml/2006/chart'>"
        "<c:chart xmlns:c='http://schemas.openxmlformats.org/drawingml/2006/chart' "
        "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' "
        "r:id='rIdMissing'/></a:graphicData></a:graphic></p:graphicFrame>"
    )
    options = ConvertOptions()
    deck = convert_pptx_to_model(derive_deck(authoring, shapes_xml=frame), options)
    assert any(warning.code == "chart-unreadable" for warning in options.warnings)
    assert isinstance(deck.slides[0].elements[-1], m.ShapeElement)


def test_a_chart_type_that_is_not_implemented_says_so(authoring):
    """Types with no renderer must say so rather than draw a wrong picture."""
    from tests.deckbuilder import derive_deck

    chart_xml = (
        "<?xml version='1.0'?>"
        f"<c:chartSpace {C} {A} {R}><c:chart><c:plotArea><c:stockChart>"
        "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:stockChart></c:plotArea></c:chart></c:chartSpace>"
    ).encode()
    frame = (
        "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='97' name='Stock'/>"
        "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
        "<p:xfrm><a:off x='0' y='0'/><a:ext cx='1000000' cy='1000000'/></p:xfrm>"
        "<a:graphic><a:graphicData "
        "uri='http://schemas.openxmlformats.org/drawingml/2006/chart'>"
        "<c:chart xmlns:c='http://schemas.openxmlformats.org/drawingml/2006/chart' "
        "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' "
        "r:id='rIdPie'/></a:graphicData></a:graphic></p:graphicFrame>"
    )
    chart_type = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    deck_bytes = derive_deck(
        authoring,
        parts={"ppt/charts/chartPie.xml": chart_xml},
        shapes_xml=frame,
        slide_relationships=[
            (
                "rIdPie",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                "../charts/chartPie.xml",
            )
        ],
        overrides={"ppt/charts/chartPie.xml": chart_type},
    )
    options = ConvertOptions()
    convert_pptx_to_model(deck_bytes, options)
    warning = next(w for w in options.warnings if w.code == "chart-unsupported-type")
    assert "stockChart" in warning.message


# -- The probe sweep -------------------------------------------------------------------
#
# Twelve charts differing in exactly one input each, all in a 2800000 x 2300000 EMU frame
# (220.4724 x 181.1024 pt), exported by PowerPoint 16.106 and measured out of the PDF as
# exact vector coordinates.  The deck is rebuilt here from the same generator rather than
# committed, so the inputs stay reviewable as XML.
#
# `insets` are points from the frame's four edges to the plot rectangle.  This is the
# table every constant in `resolve/chart.py` was fitted to, and it is what would notice
# if one of them drifted.

PROBE_FRAME = (2800000, 2300000)

PROBE_SWEEP = {
    "bare": (
        {},
        {"left": 21.0725, "right": 10.9995, "top": 11.1025, "bottom": 24.9647},
    ),
    "title": (
        {"title": "Chart contract"},
        {"left": 21.0725, "right": 10.9995, "top": 40.8025, "bottom": 24.9647},
    ),
    "legend-b": (
        {"legend": "b"},
        {"left": 21.0725, "right": 10.9995, "top": 11.1025, "bottom": 49.0480},
    ),
    "legend-r": (
        {"legend": "r"},
        {"left": 21.0725, "right": 74.9942, "top": 11.1025, "bottom": 24.9647},
    ),
    "legend-t": (
        {"legend": "t"},
        {"left": 21.0725, "right": 10.9995, "top": 35.1858, "bottom": 24.9653},
    ),
    "legend-l": (
        {"legend": "l"},
        {"left": 85.0671, "right": 10.9996, "top": 11.1025, "bottom": 24.9647},
    ),
    "font14": (
        {"text_size": 14},
        {"left": 26.9067, "right": 10.9995, "top": 13.5451, "bottom": 32.3530},
    ),
    "font8": (
        {"text_size": 8},
        {"left": 18.1608, "right": 10.9996, "top": 11.0004, "bottom": 21.2719},
    ),
    "two-series": (
        {"legend": "b", "series": 2},
        {"left": 21.0725, "right": 10.9995, "top": 11.1025, "bottom": 49.0480},
    ),
    "gap50": (
        {"gap_width": 50},
        {"left": 21.0725, "right": 10.9995, "top": 11.1025, "bottom": 24.9647},
    ),
    # A drawn tick mark takes no layout space that an absent one does not: PowerPoint
    # produced a byte-identical plot rectangle for this and for `bare`.
    "ticks-out": (
        {"tick_mark": "out"},
        {"left": 21.0725, "right": 10.9995, "top": 11.1025, "bottom": 24.9647},
    ),
    "title-legend-b": (
        {"title": "Both", "legend": "b"},
        {"left": 21.0725, "right": 10.9995, "top": 40.8029, "bottom": 49.0480},
    ),
}

#: Worst residual across the sweep is 0.47 pt; 0.6 pt is about one pixel at the 1280 px
#: the fidelity harness scores at.
SWEEP_TOLERANCE_PT = 0.6


def probe_chart_xml(
    *,
    title=None,
    legend=None,
    series=1,
    text_size=None,
    gap_width=None,
    tick_mark="none",
):
    """One probe chart, in the same shape the exported deck used."""
    colours = ["F97316", "2563EB"]
    values = [[3, 4, 5], [2, 1, 4]]
    names = ["Coverage", "B"]
    categories = ["Reader", "Writer", "Renderer"]

    series_xml = ""
    for index in range(series):
        points = "".join(
            f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(values[index])
        )
        cats = "".join(
            f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(categories)
        )
        series_xml += (
            f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
            f"<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
            f"<c:pt idx='0'><c:v>{names[index]}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
            f"<c:spPr><a:solidFill><a:srgbClr val='{colours[index]}'/></a:solidFill></c:spPr>"
            f"<c:cat><c:strRef><c:strCache><c:ptCount val='3'/>{cats}"
            "</c:strCache></c:strRef></c:cat>"
            "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
            f"<c:ptCount val='3'/>{points}</c:numCache></c:numRef></c:val></c:ser>"
        )

    title_xml = (
        "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>"
        f"<a:rPr lang='en-US'/><a:t>{title}</a:t></a:r></a:p></c:rich></c:tx>"
        "<c:layout/><c:overlay val='0'/></c:title><c:autoTitleDeleted val='0'/>"
        if title
        else "<c:autoTitleDeleted val='1'/>"
    )
    legend_xml = (
        f"<c:legend><c:legendPos val='{legend}'/><c:layout/><c:overlay val='0'/></c:legend>"
        if legend
        else ""
    )
    tx_pr = (
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        f"<a:defRPr sz='{int(text_size * 100)}'/></a:pPr></a:p></c:txPr>"
        if text_size
        else ""
    )
    gap_xml = f"<c:gapWidth val='{gap_width}'/>" if gap_width is not None else ""

    return (
        "<?xml version='1.0'?>"
        f"<c:chartSpace {C} {A} {R}>"
        f"<c:chart>{title_xml}<c:plotArea><c:layout/>"
        "<c:barChart><c:barDir val='col'/><c:grouping val='clustered'/>"
        f"<c:varyColors val='0'/>{series_xml}{gap_xml}"
        "<c:axId val='100002'/><c:axId val='100003'/></c:barChart>"
        "<c:catAx><c:axId val='100002'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='b'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        f"<c:majorTickMark val='{tick_mark}'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100003'/>"
        "<c:crosses val='autoZero'/><c:auto val='1'/><c:lblAlgn val='ctr'/>"
        "<c:lblOffset val='100'/></c:catAx>"
        "<c:valAx><c:axId val='100003'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='l'/><c:majorGridlines/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        f"<c:majorTickMark val='{tick_mark}'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100002'/>"
        "<c:crosses val='autoZero'/><c:crossBetween val='between'/></c:valAx>"
        f"</c:plotArea>{legend_xml}"
        f"<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>{tx_pr}"
        "</c:chartSpace>"
    ).encode()


@pytest.fixture(scope="module")
def probe_deck(authoring):
    """`authoring-integration.pptx` with the twelve probe charts spliced in."""
    from tests.deckbuilder import derive_deck

    chart_type = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    parts, relationships, overrides, shapes = {}, [], {}, ""
    for index, (name, (kwargs, _)) in enumerate(PROBE_SWEEP.items()):
        part = f"ppt/charts/probe{index}.xml"
        parts[part] = probe_chart_xml(**kwargs)
        overrides[part] = chart_type
        relationships.append(
            (
                f"rIdProbe{index}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                f"../charts/probe{index}.xml",
            )
        )
        shapes += (
            f"<p:graphicFrame><p:nvGraphicFramePr>"
            f"<p:cNvPr id='{200 + index}' name='{name}'/>"
            "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
            f"<p:xfrm><a:off x='0' y='0'/>"
            f"<a:ext cx='{PROBE_FRAME[0]}' cy='{PROBE_FRAME[1]}'/></p:xfrm>"
            "<a:graphic><a:graphicData "
            "uri='http://schemas.openxmlformats.org/drawingml/2006/chart'>"
            "<c:chart xmlns:c='http://schemas.openxmlformats.org/drawingml/2006/chart' "
            "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' "
            f"r:id='rIdProbe{index}'/></a:graphicData></a:graphic></p:graphicFrame>"
        )

    deck = convert_pptx_to_model(
        derive_deck(
            authoring,
            parts=parts,
            shapes_xml=shapes,
            slide_relationships=relationships,
            overrides=overrides,
        )
    )
    charts = [e for e in deck.slides[0].elements if isinstance(e, m.ChartElement)]
    # The fixture's own chart comes first in z-order, then the twelve probes.
    return dict(zip(PROBE_SWEEP, charts[1:]))


@pytest.mark.parametrize("name", list(PROBE_SWEEP))
def test_the_probe_sweep_reproduces_powerpoints_plot_rectangle(name, probe_deck):
    chart = probe_deck[name]
    expected = PROBE_SWEEP[name][1]
    width = chart.transform.extent_width / 12700.0
    height = chart.transform.extent_height / 12700.0
    left, top, right, bottom = _plot_rect(chart)
    ours = {"left": left, "right": width - right, "top": top, "bottom": height - bottom}
    for edge, truth in expected.items():
        assert ours[edge] == pytest.approx(truth, abs=SWEEP_TOLERANCE_PT), (
            f"{name} {edge}: PowerPoint {truth:.4f}, ours {ours[edge]:.4f}"
        )


def test_gap_width_widens_the_bars(probe_deck):
    """`c:gapWidth` is the gap between category groups as a percentage of one bar's width.

    At the 150% default one bar fills 1/2.5 of its category band; at 50% it fills 1/1.5.
    Measured in the probe export as 25.12 pt and 41.867 pt against a 62.80 pt band.
    """
    default = _bars(probe_deck["bare"])[0].transform.extent_width / 12700.0
    widened = _bars(probe_deck["gap50"])[0].transform.extent_width / 12700.0
    band = 220.4724 - 21.0725 - 10.9995  # the plot width, over three categories
    assert default == pytest.approx(band / 3 / 2.5, abs=0.1)
    assert widened == pytest.approx(band / 3 / 1.5, abs=0.1)


def test_clustered_series_sit_side_by_side_inside_the_category_band(probe_deck):
    chart = probe_deck["two-series"]
    bars = [
        child
        for child in chart.children
        if isinstance(child, m.ShapeElement)
        and isinstance(child.fill, m.SolidFill)
        and child.text_body is None
        and child.transform.extent_width > 10 * 12700
    ]
    assert len(bars) == 6
    widths = {round(bar.transform.extent_width) for bar in bars}
    assert len(widths) == 1, "clustered bars all share one width"
    # Two series at gapWidth 150 means a band holds 2 bars plus 1.5 of one more.
    band = (220.4724 - 21.0725 - 10.9995) / 3
    assert bars[0].transform.extent_width / 12700.0 == pytest.approx(
        band / 3.5, abs=0.1
    )


# -- The variant sweep -----------------------------------------------------------------
#
# A second six-chart probe covering the barChart variants no deck in the corpus has:
# horizontal bars, stacked and percent-stacked grouping, negative values with and without
# `invertIfNegative`, and `varyColors`.  Same frame as the first sweep, same method --
# exported by PowerPoint and read back out of the PDF as exact vector coordinates.

VARIANT_SWEEP = {
    "horizontal": (
        {"bar_dir": "bar"},
        {"left": 55.413, "right": 13.670, "top": 11.000, "bottom": 24.965},
        (0.0, 6.0, 2.0),
    ),
    "stacked": (
        {"grouping": "stacked", "series": 2, "overlap": 100},
        {"left": 26.413, "right": 10.999, "top": 11.103, "bottom": 24.965},
        (0.0, 10.0, 1.0),
    ),
    "percent-stacked": (
        {"grouping": "percentStacked", "series": 2, "overlap": 100},
        {"left": 40.012, "right": 10.999, "top": 11.103, "bottom": 24.965},
        (0.0, 1.0, 0.1),
    ),
    # A chart with negative values reserves *no* band under the plot: `tickLblPos`
    # defaults to `nextTo`, the category axis floats at zero, and the labels go with it.
    "negative-default": (
        {"values": [3, -2, 5]},
        {"left": 24.478, "right": 10.999, "top": 11.103, "bottom": 11.103},
        (-3.0, 6.0, 1.0),
    ),
    "negative-noinvert": (
        {"values": [3, -2, 5], "invert": False},
        {"left": 24.478, "right": 10.999, "top": 11.103, "bottom": 11.103},
        (-3.0, 6.0, 1.0),
    ),
    "vary-colors": (
        {"vary": True, "colour": None},
        {"left": 21.073, "right": 10.999, "top": 11.103, "bottom": 24.965},
        (0.0, 6.0, 1.0),
    ),
}


def variant_chart_xml(
    *,
    bar_dir="col",
    grouping="clustered",
    series=1,
    values=None,
    colour="F97316",
    vary=False,
    invert=None,
    overlap=0,
):
    colours = [colour, "2563EB"]
    data = [values or [3, 4, 5], [2, 1, 4]]
    categories = ["Reader", "Writer", "Renderer"]

    series_xml = ""
    for index in range(series):
        points = "".join(
            f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(data[index])
        )
        cats = "".join(
            f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(categories)
        )
        fill = (
            f"<c:spPr><a:solidFill><a:srgbClr val='{colours[index]}'/></a:solidFill></c:spPr>"
            if colours[index]
            else ""
        )
        invert_xml = (
            f"<c:invertIfNegative val='{1 if invert else 0}'/>" if invert is not None else ""
        )
        series_xml += (
            f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
            f"<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
            f"<c:pt idx='0'><c:v>S{index}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
            f"{fill}{invert_xml}"
            f"<c:cat><c:strRef><c:strCache><c:ptCount val='3'/>{cats}"
            "</c:strCache></c:strRef></c:cat>"
            "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
            f"<c:ptCount val='3'/>{points}</c:numCache></c:numRef></c:val></c:ser>"
        )

    return (
        "<?xml version='1.0'?>"
        f"<c:chartSpace {C} {A} {R}>"
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        f"<c:barChart><c:barDir val='{bar_dir}'/><c:grouping val='{grouping}'/>"
        f"<c:varyColors val='{1 if vary else 0}'/>{series_xml}"
        f"<c:gapWidth val='150'/><c:overlap val='{overlap}'/>"
        "<c:axId val='100002'/><c:axId val='100003'/></c:barChart>"
        "<c:catAx><c:axId val='100002'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='b'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='none'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100003'/>"
        "<c:crosses val='autoZero'/><c:auto val='1'/><c:lblAlgn val='ctr'/>"
        "<c:lblOffset val='100'/></c:catAx>"
        "<c:valAx><c:axId val='100003'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='l'/><c:majorGridlines/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='none'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100002'/>"
        "<c:crosses val='autoZero'/><c:crossBetween val='between'/></c:valAx>"
        "</c:plotArea>"
        "<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart></c:chartSpace>"
    ).encode()


@pytest.fixture(scope="module")
def variant_deck(authoring):
    from tests.deckbuilder import derive_deck

    chart_type = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
    parts, relationships, overrides, shapes = {}, [], {}, ""
    for index, (name, (kwargs, _, _)) in enumerate(VARIANT_SWEEP.items()):
        part = f"ppt/charts/variant{index}.xml"
        parts[part] = variant_chart_xml(**kwargs)
        overrides[part] = chart_type
        relationships.append(
            (
                f"rIdVar{index}",
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart",
                f"../charts/variant{index}.xml",
            )
        )
        shapes += (
            f"<p:graphicFrame><p:nvGraphicFramePr>"
            f"<p:cNvPr id='{300 + index}' name='{name}'/>"
            "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
            f"<p:xfrm><a:off x='0' y='0'/>"
            f"<a:ext cx='{PROBE_FRAME[0]}' cy='{PROBE_FRAME[1]}'/></p:xfrm>"
            "<a:graphic><a:graphicData "
            "uri='http://schemas.openxmlformats.org/drawingml/2006/chart'>"
            "<c:chart xmlns:c='http://schemas.openxmlformats.org/drawingml/2006/chart' "
            "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' "
            f"r:id='rIdVar{index}'/></a:graphicData></a:graphic></p:graphicFrame>"
        )

    deck = convert_pptx_to_model(
        derive_deck(
            authoring,
            parts=parts,
            shapes_xml=shapes,
            slide_relationships=relationships,
            overrides=overrides,
        )
    )
    charts = [e for e in deck.slides[0].elements if isinstance(e, m.ChartElement)]
    return dict(zip(VARIANT_SWEEP, charts[1:]))


@pytest.mark.parametrize("name", list(VARIANT_SWEEP))
def test_the_variant_sweep_reproduces_powerpoints_plot_rectangle(name, variant_deck):
    chart = variant_deck[name]
    _, expected, _ = VARIANT_SWEEP[name]
    width = chart.transform.extent_width / 12700.0
    height = chart.transform.extent_height / 12700.0
    left, top, right, bottom = _plot_rect(chart)
    ours = {"left": left, "right": width - right, "top": top, "bottom": height - bottom}
    for edge, truth in expected.items():
        assert ours[edge] == pytest.approx(truth, abs=SWEEP_TOLERANCE_PT), (
            f"{name} {edge}: PowerPoint {truth:.4f}, ours {ours[edge]:.4f}"
        )


@pytest.mark.parametrize("name", list(VARIANT_SWEEP))
def test_the_variant_sweep_reproduces_powerpoints_axis(name, variant_deck):
    scale = variant_deck[name].chart.value_axis
    minimum, maximum, unit = VARIANT_SWEEP[name][2]
    assert (scale.minimum, scale.maximum, scale.major_unit) == pytest.approx(
        (minimum, maximum, unit)
    )


def test_a_negative_bar_is_drawn_hollow_only_when_the_file_asks(variant_deck):
    """Measured: white fill, dark 0.75 pt outline -- and *only* at ``val="1"``.

    This assertion used to be the other way round, on the reading that ECMA-376's
    CT_Boolean defaults `val` to 1 and therefore an absent `c:invertIfNegative` means
    inversion.  `real-college-template.pptx` is the first deck in the corpus with a
    negative datum and no such element, and PowerPoint drew that bar solid.

    Probed properly rather than inferred, because the deck differs from the sweep in two
    ways at once -- stacked rather than clustered, and no `c:spPr` on the series rather
    than an explicit fill.  Four variants of that chart were built by rewriting its
    `ppt/charts/chart1.xml` and exported through PowerPoint:

        stacked   + automatic accent fill -> solid #C00000 (the theme's accent1)
        stacked   + explicit #2563EB      -> solid #2563EB
        clustered + automatic accent fill -> solid #C00000
        clustered + explicit #2563EB      -> solid #2563EB

    Neither axis matters; the element's absence is what decides.  The same chart with
    `<c:invertIfNegative val="1"/>` added came out white with a dark outline in both
    fill variants, which is the case this test now pins.
    """
    default = _bars(variant_deck["negative-default"])
    assert [bar.fill.color.hex.upper() for bar in default] == [
        "#F97316",
        "#F97316",
        "#F97316",
    ]

    kept = _bars(variant_deck["negative-noinvert"])
    assert {bar.fill.color.hex.upper() for bar in kept} == {"#F97316"}


def test_vary_colors_cycles_the_theme_accents_undarkened(variant_deck):
    """pptx-renderer darkens these to 88%; PowerPoint's export says it does not."""
    bars = _bars(variant_deck["vary-colors"])
    assert [bar.fill.color.hex.upper() for bar in bars] == ["#4472C4", "#ED7D31", "#A5A5A5"]


def test_a_horizontal_chart_puts_the_first_category_at_the_bottom(variant_deck):
    chart = variant_deck["horizontal"]
    bars = sorted(
        (
            child
            for child in chart.children
            if isinstance(child, m.ShapeElement)
            and isinstance(child.fill, m.SolidFill)
            and child.text_body is None
            and child.transform.extent_width > 10 * 12700
        ),
        key=lambda bar: bar.transform.offset_y,
    )
    assert len(bars) == 3
    # Values 3, 4, 5 with "Reader" lowest, so the bars get *shorter* going down.
    widths = [bar.transform.extent_width for bar in bars]
    assert widths[0] > widths[1] > widths[2]


def test_a_stacked_series_sits_on_top_of_the_one_before_it(variant_deck):
    chart = variant_deck["stacked"]
    bars = [
        child
        for child in chart.children
        if isinstance(child, m.ShapeElement)
        and isinstance(child.fill, m.SolidFill)
        and child.text_body is None
        and child.transform.extent_width > 10 * 12700
    ]
    first = sorted(
        (b for b in bars if b.fill.color.hex.upper() == "#F97316"),
        key=lambda bar: bar.transform.offset_x,
    )
    second = sorted(
        (b for b in bars if b.fill.color.hex.upper() == "#2563EB"),
        key=lambda bar: bar.transform.offset_x,
    )
    assert len(first) == len(second) == 3
    for lower, upper in zip(first, second):
        assert lower.transform.offset_x == pytest.approx(upper.transform.offset_x)
        # The second series' foot is the first series' head.
        assert upper.transform.offset_y + upper.transform.extent_height == pytest.approx(
            lower.transform.offset_y, abs=1.0
        )


# -- Malformed and hostile input -------------------------------------------------------


def _build(body: str, *, width: float = 200.0, height: float = 150.0):
    """Lay out a bare chart with no colour or text resolution, and return its children."""
    from pptx2svg.resolve.chart import ChartBuilder, ChartStyle

    source = chart(body)
    return ChartBuilder(
        source,
        source.plots[0],
        width_pt=width,
        height_pt=height,
        style=ChartStyle(
            font_family="Aptos",
            font_size=10.0,
            color=m.ResolvedColor(hex="#000000"),
            accents=[
                m.ResolvedColor(hex="#4472C4"),
                m.ResolvedColor(hex="#ED7D31"),
                m.ResolvedColor(hex="#A5A5A5"),
            ],
        ),
        resolve_fill=_fake_fill,
        resolve_outline=_fake_outline,
        resolve_text=lambda rich, text, size, align: m.TextBody(),
    ).build()


def _fake_fill(fill):
    """Resolve the one fill spelling these tests use, without a whole ResolveContext."""
    from pptx2svg.parse import source as s

    if isinstance(fill, s.SourceSolidFill) and isinstance(fill.color, s.SrgbColor):
        return m.SolidFill(color=m.ResolvedColor(hex=f"#{fill.color.hex}"))
    if isinstance(fill, s.SourceNoFill):
        return m.NoFill()
    return None


def _fake_outline(outline):
    from pptx2svg.parse import source as s

    if outline is None:
        return None
    fill = _fake_fill(outline.fill)
    if isinstance(fill, m.NoFill) or fill is None:
        return None
    return m.Outline(width=outline.width or 12700, fill=fill)


@pytest.mark.parametrize(
    "name,body",
    [
        ("no series", "<c:chart><c:plotArea><c:barChart/></c:plotArea></c:chart>"),
        (
            "series with no values",
            "<c:chart><c:plotArea><c:barChart><c:ser/></c:barChart></c:plotArea></c:chart>",
        ),
        (
            "every point blank",
            bar(
                "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='3'/>"
                "</c:numCache></c:numRef></c:val></c:ser>"
            ),
        ),
        (
            "every value zero",
            bar(
                "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='2'/>"
                "<c:pt idx='0'><c:v>0</c:v></c:pt><c:pt idx='1'><c:v>0</c:v></c:pt>"
                "</c:numCache></c:numRef></c:val></c:ser>"
            ),
        ),
        (
            "NaN and infinities",
            bar(
                "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='3'/>"
                "<c:pt idx='0'><c:v>NaN</c:v></c:pt>"
                "<c:pt idx='1'><c:v>Infinity</c:v></c:pt>"
                "<c:pt idx='2'><c:v>1e400</c:v></c:pt>"
                "</c:numCache></c:numRef></c:val></c:ser>"
            ),
        ),
        (
            "values at the edge of the float range",
            bar(
                "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='2'/>"
                "<c:pt idx='0'><c:v>1e300</c:v></c:pt>"
                "<c:pt idx='1'><c:v>-1e300</c:v></c:pt>"
                "</c:numCache></c:numRef></c:val></c:ser>"
            ),
        ),
        (
            "zero gap width",
            bar(
                "<c:gapWidth val='0'/><c:ser><c:val><c:numRef><c:numCache>"
                "<c:ptCount val='1'/><c:pt idx='0'><c:v>5</c:v></c:pt>"
                "</c:numCache></c:numRef></c:val></c:ser>"
            ),
        ),
        (
            "percent stacked whose totals are zero",
            bar(
                "<c:grouping val='percentStacked'/><c:ser><c:val><c:numRef><c:numCache>"
                "<c:ptCount val='2'/><c:pt idx='0'><c:v>0</c:v></c:pt>"
                "<c:pt idx='1'><c:v>0</c:v></c:pt>"
                "</c:numCache></c:numRef></c:val></c:ser>"
            ),
        ),
        (
            "series of different lengths",
            bar(
                "<c:grouping val='stacked'/>"
                "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='5'/>"
                "<c:pt idx='4'><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
                "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
                "<c:pt idx='0'><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
            ),
        ),
    ],
)
def test_malformed_charts_lay_out_rather_than_raise(name, body):
    children, data = _build(body)
    assert data.value_axis.maximum > data.value_axis.minimum
    assert data.value_axis.major_unit > 0
    for child in children:
        assert child.transform.extent_width >= 0
        assert child.transform.extent_height >= 0


def test_a_hostile_major_unit_cannot_generate_unbounded_ticks():
    """`c:majorUnit` is attacker-controlled; 1e-7 over a 1.5e12 axis is 10^19 gridlines."""
    children, data = _build(
        "<c:chart><c:plotArea>"
        "<c:barChart><c:axId val='1'/><c:axId val='2'/>"
        "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>1e12</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:barChart><c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/><c:majorGridlines/><c:majorUnit val='0.0000001'/>"
        "</c:valAx></c:plotArea></c:chart>"
    )
    assert len(children) < 50


def test_an_axis_minimum_above_its_maximum_still_spans_something():
    _, data = _build(
        "<c:chart><c:plotArea>"
        "<c:barChart><c:axId val='1'/><c:axId val='2'/>"
        "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:barChart><c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/><c:scaling><c:min val='100'/><c:max val='1'/>"
        "</c:scaling></c:valAx></c:plotArea></c:chart>"
    )
    assert data.value_axis.maximum > data.value_axis.minimum


def test_an_explicit_axis_range_wins_over_the_computed_one():
    _, data = _build(
        "<c:chart><c:plotArea>"
        "<c:barChart><c:axId val='1'/><c:axId val='2'/>"
        "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:barChart><c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/><c:scaling><c:min val='-10'/><c:max val='40'/>"
        "</c:scaling><c:majorUnit val='10'/></c:valAx></c:plotArea></c:chart>"
    )
    assert (data.value_axis.minimum, data.value_axis.maximum, data.value_axis.major_unit) == (
        -10.0,
        40.0,
        10.0,
    )


def test_a_source_linked_axis_takes_the_cells_format():
    """`sourceLinked="1"` means "whatever the cell says", and the cache records that."""
    _, data = _build(
        "<c:chart><c:plotArea>"
        "<c:barChart><c:axId val='1'/><c:axId val='2'/>"
        "<c:ser><c:val><c:numRef><c:numCache><c:formatCode>#,##0</c:formatCode>"
        "<c:ptCount val='1'/><c:pt idx='0'><c:v>4285</c:v></c:pt>"
        "</c:numCache></c:numRef></c:val></c:ser></c:barChart>"
        "<c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/><c:numFmt formatCode='General' sourceLinked='1'/>"
        "</c:valAx></c:plotArea></c:chart>"
    )
    assert data.series[0].format_code == "#,##0"


def test_an_axis_that_chose_general_does_not_inherit_the_cells_format():
    """`sourceLinked="0"` is a choice, including when the choice is General.

    `real-financial-report.pptx` writes exactly this and PowerPoint prints "1000", not
    "1,000" -- even though the same chart's data labels ask for `#,##0`.
    """
    chart_xml = (
        "<c:chart><c:plotArea>"
        "<c:barChart><c:axId val='1'/><c:axId val='2'/>"
        "<c:ser><c:val><c:numRef><c:numCache><c:formatCode>#,##0</c:formatCode>"
        "<c:ptCount val='1'/><c:pt idx='0'><c:v>4285</c:v></c:pt>"
        "</c:numCache></c:numRef></c:val></c:ser></c:barChart>"
        "<c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/><c:numFmt formatCode='General' sourceLinked='0'/>"
        "</c:valAx></c:plotArea></c:chart>"
    )
    children, _ = _build(chart_xml)
    printed = {
        "".join(run.text for p in child.text_body.paragraphs for run in p.runs)
        for child in children
        if isinstance(child, m.ShapeElement) and child.text_body is not None
    }
    assert "5000" in printed
    assert not any("," in label for label in printed)


# -- Findings from review ---------------------------------------------------------------
#
# Each of these reproduced a real defect before its fix.  Several are orientation bugs the
# probe sweep could not see, because it checks the plot *rectangle* and these get the
# rectangle right while drawing the wrong thing inside it.


def _lines(children):
    return [c for c in children if isinstance(c, m.ConnectorElement)]


def test_a_horizontal_charts_gridlines_are_vertical():
    """The value axis runs along the bottom, so its gridlines run up the plot.

    Measured on the horizontal probe: three vertical lines, at the ticks for 2, 4 and 6.
    The one for 0 is left out because the category axis already draws it -- the same
    rule a column chart applies to its zero gridline.
    """
    children, _ = _build(
        "<c:chart><c:plotArea><c:barChart><c:barDir val='bar'/>"
        "<c:axId val='1'/><c:axId val='2'/>"
        "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='3'/>"
        "<c:pt idx='0'><c:v>3</c:v></c:pt><c:pt idx='1'><c:v>4</c:v></c:pt>"
        "<c:pt idx='2'><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:barChart><c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/><c:majorGridlines/></c:valAx></c:plotArea></c:chart>"
    )
    lines = _lines(children)
    gridlines = [line for line in lines if line.transform.extent_width == 0]
    # Three gridlines plus the category axis, all vertical; one horizontal value axis.
    assert len(gridlines) == 4
    assert len([line for line in lines if line.transform.extent_height == 0]) == 1


def test_barDir_decides_which_axis_line_is_which():
    """`c:delete` and `c:spPr` have to follow their own axis, not a fixed side."""
    body = (
        "<c:chart><c:plotArea><c:barChart><c:barDir val='bar'/>"
        "<c:axId val='1'/><c:axId val='2'/>"
        "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:barChart><c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/>{deleted}</c:valAx></c:plotArea></c:chart>"
    )
    both = _lines(_build(body.format(deleted=""))[0])
    assert {line.transform.extent_height == 0 for line in both} == {True, False}

    # Deleting the *value* axis must drop the horizontal line, not the vertical one.
    remaining = _lines(_build(body.format(deleted="<c:delete val='1'/>"))[0])
    assert len(remaining) == 1
    assert remaining[0].transform.extent_width == 0


def test_an_axis_whose_spPr_says_noFill_draws_no_line():
    """`<a:ln><a:noFill/></a:ln>` is "no line", not "no opinion".

    `real-college-template`'s chart says it on the value axis -- ``<a:ln w="25400">
    <a:noFill/></a:ln>`` -- and PowerPoint draws nothing up the left of that plot.  We
    drew the default black axis line, full plot height, because
    ``resolve.view._resolve_outline`` collapses "noFill" and "absent" to the same
    ``None``.  Its category axis, which *does* state a stroke, still draws.
    """
    body = (
        "<c:chart><c:plotArea><c:barChart><c:barDir val='col'/>"
        "<c:axId val='1'/><c:axId val='2'/>"
        "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:barChart>"
        "<c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/>{spPr}</c:valAx></c:plotArea></c:chart>"
    )
    both = _lines(_build(body.format(spPr=""))[0])
    assert len(both) == 2

    quiet = _lines(_build(body.format(spPr="<c:spPr><a:ln w='25400'><a:noFill/></a:ln></c:spPr>"))[0])
    # Only the category axis is left, and a column chart draws that one horizontally.
    assert len(quiet) == 1
    assert quiet[0].transform.extent_height == 0


def test_a_negative_gap_width_does_not_divide_by_zero():
    """`c:gapWidth` is schema-bounded to 0..500 and a file need not obey.

    At -100 on a single series the bar-width divisor is exactly zero, which used to abort
    the whole conversion rather than the one chart.
    """
    children, _ = _build(
        bar(
            "<c:gapWidth val='-100'/><c:ser><c:val><c:numRef><c:numCache>"
            "<c:ptCount val='1'/><c:pt idx='0'><c:v>3</c:v></c:pt>"
            "</c:numCache></c:numRef></c:val></c:ser>"
        )
    )
    assert children


@pytest.mark.parametrize("value", ["1.5e308", "-1.5e308"])
def test_a_datum_near_the_float_ceiling_does_not_overflow(value):
    """The axis maximum is rounded strictly outwards, which takes 1.5e308 to infinity."""
    _, data = _build(
        bar(
            "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
            f"<c:pt idx='0'><c:v>{value}</c:v></c:pt>"
            "</c:numCache></c:numRef></c:val></c:ser>"
        )
    )
    assert data.value_axis.maximum > data.value_axis.minimum


def test_a_denormal_span_does_not_underflow_the_unit_to_zero():
    minimum, maximum, unit = nice_axis_scale(0.0, 5e-324)
    assert unit > 0 and maximum > minimum


def test_each_axis_styles_its_own_labels():
    """`c:catAx/c:txPr` was parsed and then never consulted; both axes have one."""
    children, _ = _build(
        "<c:chart><c:plotArea><c:barChart><c:axId val='1'/><c:axId val='2'/>"
        "<c:ser><c:cat><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>Reader</c:v></c:pt></c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:barChart>"
        "<c:catAx><c:axId val='1'/><c:txPr><a:bodyPr/><a:p><a:pPr>"
        "<a:defRPr sz='2400'/></a:pPr></a:p></c:txPr></c:catAx>"
        "<c:valAx><c:axId val='2'/><c:txPr><a:bodyPr/><a:p><a:pPr>"
        "<a:defRPr sz='800'/></a:pPr></a:p></c:txPr></c:valAx></c:plotArea></c:chart>"
    )
    sizes = {
        "".join(r.text for p in child.text_body.paragraphs for r in p.runs): (
            child.text_body.paragraphs[0].runs[0].properties.font_size
        )
        for child in children
        if isinstance(child, m.ShapeElement) and child.text_body is not None
    }
    assert sizes["Reader"] == 24
    assert sizes["0"] == 8


def test_an_overlaid_legend_does_not_shrink_the_plot():
    """`c:overlay` draws the legend *over* the plot, so it takes no space away."""
    body = (
        "<c:chart><c:plotArea><c:barChart><c:axId val='1'/><c:axId val='2'/>"
        "<c:ser><c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>Coverage</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        "<c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:barChart><c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/></c:valAx></c:plotArea>"
        "<c:legend><c:legendPos val='b'/><c:overlay val='{overlay}'/></c:legend></c:chart>"
    )
    beside = _build(body.format(overlay="0"))[0]
    over = _build(body.format(overlay="1"))[0]
    assert _plot_bottom(over) > _plot_bottom(beside)


def _plot_bottom(children):
    lines = [c for c in children if isinstance(c, m.ConnectorElement)]
    down = max(lines, key=lambda line: line.transform.extent_height).transform
    return down.offset_y + down.extent_height


def test_a_series_longer_than_the_labelled_one_still_draws_every_bar():
    """Bars are indexed by the category list, so a short list used to lose their tails."""
    children, data = _build(
        bar(
            "<c:ser><c:idx val='0'/><c:order val='0'/>"
            "<c:cat><c:strRef><c:strCache><c:ptCount val='2'/>"
            "<c:pt idx='0'><c:v>a</c:v></c:pt><c:pt idx='1'><c:v>b</c:v></c:pt>"
            "</c:strCache></c:strRef></c:cat>"
            "<c:val><c:numRef><c:numCache><c:ptCount val='2'/>"
            "<c:pt idx='0'><c:v>1</c:v></c:pt><c:pt idx='1'><c:v>2</c:v></c:pt>"
            "</c:numCache></c:numRef></c:val></c:ser>"
            "<c:ser><c:idx val='1'/><c:order val='1'/>"
            "<c:val><c:numRef><c:numCache><c:ptCount val='4'/>"
            "<c:pt idx='0'><c:v>3</c:v></c:pt><c:pt idx='1'><c:v>4</c:v></c:pt>"
            "<c:pt idx='2'><c:v>5</c:v></c:pt><c:pt idx='3'><c:v>6</c:v></c:pt>"
            "</c:numCache></c:numRef></c:val></c:ser>"
        )
    )
    assert data.categories == ["a", "b", "3", "4"]
    bars = [
        child
        for child in children
        if isinstance(child, m.ShapeElement)
        and isinstance(child.fill, m.SolidFill)
        and child.text_body is None
    ]
    assert len(bars) == 6


def _filled_bars(series: str):
    """Every filled rectangle a bare `c:barChart` draws, left to right.

    No `c:legend` in these bodies, so nothing but the bars themselves is filled.
    """
    children, _ = _build(bar(series))
    return sorted(
        (
            child
            for child in children
            if isinstance(child, m.ShapeElement)
            and isinstance(child.fill, m.SolidFill)
            and child.text_body is None
        ),
        key=lambda bar: bar.transform.offset_x,
    )


def test_an_explicit_invert_if_negative_draws_the_bar_hollow():
    """The other half of the probe above: `val="1"` really is white-with-an-outline.

    Measured on `probe-invert-on-auto` / `probe-invert-on-fill` -- the deck's own chart
    with `<c:invertIfNegative val="1"/>` added to the series carrying the -1.0.  Both
    came out white, so the hollow drawing survives; only its default changed.
    """
    bars = _filled_bars(
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:spPr><a:solidFill><a:srgbClr val='F97316'/></a:solidFill></c:spPr>"
        "<c:invertIfNegative val='1'/><c:val><c:numRef><c:numCache>"
        "<c:ptCount val='3'/><c:pt idx='0'><c:v>3</c:v></c:pt>"
        "<c:pt idx='1'><c:v>-2</c:v></c:pt><c:pt idx='2'><c:v>5</c:v></c:pt>"
        "</c:numCache></c:numRef></c:val></c:ser>"
    )
    assert [bar.fill.color.hex.upper() for bar in bars] == [
        "#F97316",
        "#FFFFFF",
        "#F97316",
    ]
    inverted = bars[1]
    assert inverted.outline is not None
    assert inverted.outline.fill.color.hex.upper() == "#000000"
    assert inverted.outline.width == 9525


def test_a_series_takes_the_accent_its_c_idx_names_not_its_position():
    """`real-college-template`'s chart: (idx 2, order 0) then (idx 0, order 1).

    The second series states no fill and PowerPoint drew it in accent1, not accent2 --
    so the accent cycle is indexed by `c:idx`.  Every probe before this deck had
    idx == order, which is why the two readings were indistinguishable.
    """
    bars = _filled_bars(
        "<c:ser><c:idx val='2'/><c:order val='0'/><c:val><c:numRef><c:numCache>"
        "<c:ptCount val='1'/><c:pt idx='0'><c:v>3</c:v></c:pt>"
        "</c:numCache></c:numRef></c:val></c:ser>"
        "<c:ser><c:idx val='0'/><c:order val='1'/><c:val><c:numRef><c:numCache>"
        "<c:ptCount val='1'/><c:pt idx='0'><c:v>4</c:v></c:pt>"
        "</c:numCache></c:numRef></c:val></c:ser>"
    )
    # Accents are [#4472C4, #ED7D31, #A5A5A5]: idx 2 is the third, idx 0 the first.
    assert [bar.fill.color.hex.upper() for bar in bars] == ["#A5A5A5", "#4472C4"]


def test_vary_colors_does_not_suppress_the_negative_bar_inversion():
    """A varyColors fill is not a `c:dPt`, so `c:invertIfNegative` still applies to it."""
    children, _ = _build(
        bar(
            "<c:varyColors val='1'/><c:ser><c:invertIfNegative val='1'/>"
            "<c:val><c:numRef><c:numCache>"
            "<c:ptCount val='3'/><c:pt idx='0'><c:v>3</c:v></c:pt>"
            "<c:pt idx='1'><c:v>-2</c:v></c:pt><c:pt idx='2'><c:v>5</c:v></c:pt>"
            "</c:numCache></c:numRef></c:val></c:ser>"
        )
    )
    bars = sorted(
        (
            child
            for child in children
            if isinstance(child, m.ShapeElement)
            and isinstance(child.fill, m.SolidFill)
            and child.text_body is None
        ),
        key=lambda bar: bar.transform.offset_x,
    )
    assert [bar.fill.color.hex.upper() for bar in bars] == [
        "#4472C4",
        "#FFFFFF",
        "#A5A5A5",
    ]


def test_a_deleted_legend_entry_reserves_no_band_width():
    """`c:legendEntry`/`c:delete` removes the label, so it must not shift the plot."""
    body = (
        "<c:chart><c:plotArea><c:barChart><c:axId val='1'/><c:axId val='2'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>ab</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        "<c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>3</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "<c:ser><c:idx val='1'/><c:order val='1'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>a very long series name indeed</c:v></c:pt>"
        "</c:strCache></c:strRef></c:tx>"
        "<c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>2</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:barChart><c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/></c:valAx></c:plotArea>"
        "<c:legend><c:legendPos val='r'/>{entry}</c:legend></c:chart>"
    )
    kept = _build(body.format(entry=""))[0]
    struck = _build(
        body.format(
            entry="<c:legendEntry><c:idx val='1'/><c:delete val='1'/></c:legendEntry>"
        )
    )[0]

    def plot_width(children):
        lines = [c for c in children if isinstance(c, m.ConnectorElement)]
        across = max(lines, key=lambda line: line.transform.extent_width)
        return across.transform.extent_width

    assert plot_width(struck) > plot_width(kept)


def test_a_multi_level_category_cache_without_ptCount_keeps_its_labels():
    """`c:ptCount` is minOccurs=0, and a multi-level cache has no direct `c:pt` children."""
    parsed = chart(
        bar(
            "<c:ser><c:cat><c:multiLvlStrRef><c:multiLvlStrCache><c:lvl>"
            "<c:pt idx='0'><c:v>Q1</c:v></c:pt><c:pt idx='1'><c:v>Q2</c:v></c:pt>"
            "</c:lvl></c:multiLvlStrCache></c:multiLvlStrRef></c:cat>"
            "<c:val><c:numRef><c:numCache><c:ptCount val='2'/>"
            "<c:pt idx='0'><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        )
    )
    assert parsed.plots[0].series[0].categories == ["Q1", "Q2"]


def test_the_metrics_less_fallback_line_box_matches_the_renderers():
    """A face with no metrics table still has to lay out in the 1.2 em the renderer uses.

    A shorter fallback box made an unmetricked title's band come out short of its own
    text, which is the one thing `FontBox` exists to keep in step.
    """
    from pptx2svg.resolve.chart import font_box
    from pptx2svg.text.measure import DEFAULT_LINE_HEIGHT_RATIO

    box = font_box("a face nothing has metrics for", 18.0)
    assert box.line_height == pytest.approx(DEFAULT_LINE_HEIGHT_RATIO * 18.0)
    assert box.first_baseline < box.line_height


# -- Line charts -----------------------------------------------------------------------
#
# Measured on a six-chart probe deck exported by PowerPoint, plus the real line chart in
# `real-financial-report.pptx`.  Coordinates are frame-relative points read out of the
# PDF as exact vectors.

LINE_CATS = ["Reader", "Writer", "Renderer"]


def line_chart_xml(
    *,
    values=(3, 4, 5),
    line="<a:ln w='25400'><a:solidFill><a:srgbClr val='F97316'/></a:solidFill></a:ln>",
    marker="<c:marker><c:symbol val='none'/></c:marker>",
    smooth=None,
    blanks="gap",
):
    points = "".join(
        f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>"
        for i, v in enumerate(values)
        if v is not None
    )
    cats = "".join(
        f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(LINE_CATS)
    )
    smooth_xml = "" if smooth is None else f"<c:smooth val='{1 if smooth else 0}'/>"
    return (
        "<c:chart><c:plotArea><c:layout/>"
        "<c:lineChart><c:grouping val='standard'/><c:varyColors val='0'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>A</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        f"<c:spPr><a:solidFill><a:srgbClr val='F97316'/></a:solidFill>{line}</c:spPr>"
        f"{marker}"
        f"<c:cat><c:strRef><c:strCache><c:ptCount val='3'/>{cats}"
        "</c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef></c:val>"
        f"{smooth_xml}</c:ser>"
        "<c:axId val='1'/><c:axId val='2'/></c:lineChart>"
        "<c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/><c:majorGridlines/>"
        "<c:crossBetween val='between'/></c:valAx>"
        f"</c:plotArea><c:dispBlanksAs val='{blanks}'/></c:chart>"
    )


def _polylines(children):
    return [
        child
        for child in children
        if isinstance(child, m.ShapeElement)
        and isinstance(child.geometry, m.CustomGeometry)
    ]


def _markers(children):
    return sorted(
        (
            child
            for child in children
            if isinstance(child, m.ShapeElement)
            and isinstance(child.geometry, m.PresetGeometry)
            and child.geometry.preset in ("ellipse", "diamond", "triangle", "star5")
        ),
        key=lambda child: child.transform.offset_x,
    )


def _vertices(shape):
    """A polyline's vertices in frame coordinates, from its box and its path data."""
    transform = shape.transform
    path = shape.geometry.paths[0]
    scale_x = transform.extent_width / 12700.0 / path.width if path.width else 1.0
    scale_y = transform.extent_height / 12700.0 / path.height if path.height else 1.0
    out = []
    for token in re.finditer(r"[ML] ([-\d.]+) ([-\d.]+)", path.commands):
        out.append(
            (
                round(transform.offset_x / 12700.0 + float(token.group(1)) * scale_x, 3),
                round(transform.offset_y / 12700.0 + float(token.group(2)) * scale_y, 3),
            )
        )
    return out


def test_a_line_chart_draws_a_polyline_through_the_band_centres():
    """Measured on the probe: vertices at 52.473, 115.273 and 178.05 pt across the frame.

    `c:crossBetween="between"` puts a point in the middle of its category band, exactly
    where a bar would be -- the real line chart in `real-financial-report.pptx` says the
    same and PowerPoint drew it the same way.
    """
    children, data = _build(line_chart_xml(), width=220.4724, height=181.1024)
    assert data.kind == "lineChart"
    polylines = _polylines(children)
    assert len(polylines) == 1
    vertices = _vertices(polylines[0])
    assert [v[0] for v in vertices] == pytest.approx([52.473, 115.273, 178.052], abs=0.6)
    # Values 3, 4 and 5 on a 0..6 axis, so the vertices climb by equal steps.
    steps = [vertices[i][1] - vertices[i + 1][1] for i in range(2)]
    assert steps[0] == pytest.approx(steps[1], abs=0.01)


def test_a_line_takes_its_colour_from_a_ln_and_not_from_a_solid_fill():
    """Measured: a series stating only `a:solidFill` was drawn in accent1, fill ignored.

    A bar chart takes `a:solidFill` as its bar colour, so reusing that rule here would
    paint the line the wrong colour on any deck that sets both.
    """
    explicit, _ = _build(line_chart_xml(), width=220.0, height=181.0)
    assert _polylines(explicit)[0].outline.fill.color.hex.upper() == "#F97316"

    bare, _ = _build(line_chart_xml(line=""), width=220.0, height=181.0)
    assert _polylines(bare)[0].outline.fill.color.hex.upper() == "#4472C4"


def test_a_line_series_that_states_no_width_gets_one_and_a_half_points():
    bare, _ = _build(line_chart_xml(line=""), width=220.0, height=181.0)
    assert _polylines(bare)[0].outline.width == 19050
    # And a stated width is taken literally.
    stated, _ = _build(line_chart_xml(), width=220.0, height=181.0)
    assert _polylines(stated)[0].outline.width == 25400


def test_a_line_stroke_has_round_caps():
    """Measured: PowerPoint emits `1 J` on a line series whether or not `a:ln` says so."""
    children, _ = _build(line_chart_xml(), width=220.0, height=181.0)
    assert _polylines(children)[0].outline.line_cap == "round"


def test_marker_size_is_a_diameter_in_points():
    """`c:size val="7"` measured 6.96 pt across in the probe."""
    children, _ = _build(
        line_chart_xml(
            marker="<c:marker><c:symbol val='circle'/><c:size val='7'/></c:marker>"
        ),
        width=220.0,
        height=181.0,
    )
    markers = _markers(children)
    assert len(markers) == 3
    for marker in markers:
        assert marker.geometry.preset == "ellipse"
        assert marker.transform.extent_width / 12700.0 == pytest.approx(7.0, abs=0.01)


def test_symbol_none_draws_no_marker():
    children, _ = _build(line_chart_xml(), width=220.0, height=181.0)
    assert _markers(children) == []


def test_a_series_with_no_marker_element_still_gets_one():
    """Measured: PowerPoint drew a **diamond** for series 0 of a marker-less line chart.

    The rest of the cycle is from the specification and is not measured.
    """
    children, _ = _build(line_chart_xml(marker=""), width=220.0, height=181.0)
    markers = _markers(children)
    assert len(markers) == 3
    assert {marker.geometry.preset for marker in markers} == {"diamond"}


def test_a_blank_breaks_the_line_in_two():
    """`dispBlanksAs="gap"` is the default and really leaves a gap."""
    children, _ = _build(
        line_chart_xml(values=(3, None, 5)), width=220.0, height=181.0
    )
    assert len(_polylines(children)) == 0  # two isolated points, no run of two

    spanned, _ = _build(
        line_chart_xml(values=(3, None, 5), blanks="span"), width=220.0, height=181.0
    )
    assert len(_vertices(_polylines(spanned)[0])) == 2


def test_smoothing_emits_curves_rather_than_segments():
    """`c:smooth` is drawn as a spline; the probe's control points are not collinear.

    The tension PowerPoint uses was **not** measured, so only the shape of the output is
    asserted here, not its exact curvature.
    """
    straight, _ = _build(line_chart_xml(), width=220.0, height=181.0)
    assert "C " not in _polylines(straight)[0].geometry.paths[0].commands

    curved, _ = _build(
        line_chart_xml(values=(3, 5, 2), smooth=True), width=220.0, height=181.0
    )
    assert "C " in _polylines(curved)[0].geometry.paths[0].commands


def test_the_real_line_chart_renders():
    """`real-financial-report.pptx` slide 2 holds the only line chart in the corpus."""
    from tests.conftest import FIXTURE_DIR

    deck = convert_pptx_to_model(
        (FIXTURE_DIR / "real-financial-report.pptx").read_bytes()
    )
    charts = [
        element
        for slide in deck.slides
        for element in slide.elements
        if isinstance(element, m.ChartElement) and element.chart.kind == "lineChart"
    ]
    assert len(charts) == 1
    chart = charts[0]
    assert len(chart.chart.series) == 3
    assert len(_polylines(chart.children)) == 3
    # Nine data points, plus one more marker per series in the legend: a line chart's
    # legend key is a rule with the series' marker on its midpoint, not a swatch.
    assert len(_markers(chart.children)) == 9 + 3


# -- Data labels -----------------------------------------------------------------------
#
# Positions measured on a six-chart probe exported by PowerPoint.  Values are
# frame-relative points for the *first* bar's label, which is 25.2 pt wide and spans
# y 72.607..145.035 in a 220.4724 x 181.1024 pt frame.

#: (baseline, left edge of the glyph run), frame-relative points.
DATA_LABEL_TRUTH = {
    "outEnd": (76.029, 49.803),   # ours 76.007, 49.832
    "inEnd": (97.149, 49.803),    # ours 97.114, 49.832
    "ctr": (122.829, 49.802),     # ours 122.660, 49.832
    "inBase": (148.413, 49.803),  # ours 148.578, 49.832
}
DATA_LABEL_TOLERANCE_PT = 0.4


def dlbl_chart_xml(*, show=("Val",), pos=None, fmt=None, size=None, values=(3, 4, 5),
                   name="A"):
    flags = "".join(
        f"<c:show{flag} val='{1 if flag in show else 0}'/>"
        for flag in ("LegendKey", "Val", "CatName", "SerName", "Percent", "BubbleSize")
    )
    fmt_xml = f"<c:numFmt formatCode='{fmt}' sourceLinked='0'/>" if fmt else ""
    tx = (
        f"<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz='{size * 100}'/>"
        "</a:pPr></a:p></c:txPr>"
        if size
        else ""
    )
    pos_xml = f"<c:dLblPos val='{pos}'/>" if pos else ""
    points = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(values))
    cats = "".join(
        f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>"
        for i, v in enumerate(["Reader", "Writer", "Renderer"])
    )
    return (
        "<c:chart><c:plotArea><c:layout/>"
        "<c:barChart><c:barDir val='col'/><c:grouping val='clustered'/>"
        "<c:varyColors val='0'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        f"<c:pt idx='0'><c:v>{name}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        "<c:spPr><a:solidFill><a:srgbClr val='F97316'/></a:solidFill></c:spPr>"
        f"<c:dLbls>{fmt_xml}{tx}{pos_xml}{flags}</c:dLbls>"
        f"<c:cat><c:strRef><c:strCache><c:ptCount val='3'/>{cats}"
        "</c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef></c:val></c:ser>"
        "<c:gapWidth val='150'/><c:overlap val='0'/>"
        "<c:axId val='1'/><c:axId val='2'/></c:barChart>"
        "<c:catAx><c:axId val='1'/></c:catAx>"
        "<c:valAx><c:axId val='2'/><c:majorGridlines/></c:valAx>"
        "</c:plotArea></c:chart>"
    )


def _data_labels(children, frame_height=181.1024):
    """Every text that is neither an axis label nor a category label.

    Axis labels are laid out in a box starting at the frame's left edge; category labels
    sit in the band under the plot.
    """
    out = []
    for child in children:
        if not isinstance(child, m.ShapeElement) or child.text_body is None:
            continue
        if child.transform.offset_x / 12700.0 < 1.0:
            continue
        if child.transform.offset_y / 12700.0 > frame_height - 30:
            continue
        out.append(child)
    return sorted(out, key=lambda child: (child.transform.offset_x, child.transform.offset_y))


def _label_baseline(shape):
    size = shape.text_body.paragraphs[0].runs[0].properties.font_size
    return shape.transform.offset_y / 12700.0 + (1.2 - APTOS_DESCENT) * size


def _label_text_left(shape, font_family="Aptos"):
    """Where the glyphs start: the box is wider than the run and the run is centred."""
    from pptx2svg.resolve.chart import text_width

    size = shape.text_body.paragraphs[0].runs[0].properties.font_size
    text = "".join(r.text for p in shape.text_body.paragraphs for r in p.runs)
    box = shape.transform.extent_width / 12700.0
    run = text_width(text, font_family, size)
    align = shape.text_body.paragraphs[0].properties.alignment
    left = shape.transform.offset_x / 12700.0
    if align == "ctr":
        return left + (box - run) / 2
    if align == "r":
        return left + box - run
    return left


@pytest.mark.parametrize("position", list(DATA_LABEL_TRUTH))
def test_each_data_label_position_matches_powerpoints(position):
    children, _ = _build(
        dlbl_chart_xml(pos=None if position == "outEnd" else position),
        width=220.4724,
        height=181.1024,
    )
    labels = _data_labels(children)
    assert labels, f"{position} drew no data label"
    first = labels[0]
    baseline, left = DATA_LABEL_TRUTH[position]
    assert _label_baseline(first) == pytest.approx(baseline, abs=DATA_LABEL_TOLERANCE_PT)
    assert _label_text_left(first) == pytest.approx(left, abs=DATA_LABEL_TOLERANCE_PT)


def test_outEnd_is_the_default_for_a_bar():
    explicit, _ = _build(dlbl_chart_xml(pos="outEnd"), width=220.4724, height=181.1024)
    implicit, _ = _build(dlbl_chart_xml(), width=220.4724, height=181.1024)
    assert _label_baseline(_data_labels(explicit)[0]) == pytest.approx(
        _label_baseline(_data_labels(implicit)[0])
    )


def test_a_label_at_a_larger_size_keeps_the_same_fixed_gap():
    """The gap to the bar is a fixed 4.85 pt, not a multiple of the font size.

    Measured 4.86 pt at 10 pt and 4.70 pt at 14 pt; an em-proportional gap would have
    grown to 6.8 pt and put the 14 pt label two points too high.
    """
    children, _ = _build(
        dlbl_chart_xml(fmt="#,##0", size=14, values=(3000, 4000, 5000)),
        width=220.4724,
        height=181.1024,
    )
    label = _data_labels(children)[0]
    assert "".join(
        r.text for p in label.text_body.paragraphs for r in p.runs
    ) == "3,000"
    assert _label_baseline(label) == pytest.approx(74.973, abs=DATA_LABEL_TOLERANCE_PT)


def test_the_parts_of_a_multi_part_label_stack_on_separate_lines():
    """Measured: series name, category name then value, top to bottom, one line each.

    PowerPoint wraps a long category onto two lines; we do not, which is recorded in the
    roadmap rather than asserted here.
    """
    children, _ = _build(
        dlbl_chart_xml(show=("Val", "CatName", "SerName"), name="Coverage"),
        width=220.4724,
        height=181.1024,
    )
    first_bar = [
        label
        for label in _data_labels(children)
        if abs(label.transform.offset_x / 12700.0 - 27.003) < 1.0
    ]
    texts = [
        "".join(r.text for p in label.text_body.paragraphs for r in p.runs)
        for label in sorted(first_bar, key=lambda label: label.transform.offset_y)
    ]
    assert texts == ["Coverage", "Reader", "3"]


def test_dLbls_that_switch_everything_off_draw_nothing():
    """Four of the five charts in `real-financial-report.pptx` do exactly this.

    Every `c:show*` flag is stated as 0, so the correct output is no label at all -- the
    presence of a `c:dLbls` block says nothing about whether anything is printed.
    """
    children, _ = _build(dlbl_chart_xml(show=()), width=220.4724, height=181.1024)
    assert _data_labels(children) == []


def test_a_point_can_delete_its_own_label():
    body = dlbl_chart_xml().replace(
        "<c:dLbls>",
        "<c:dLbls><c:dLbl><c:idx val='1'/><c:delete val='1'/></c:dLbl>",
        1,
    )
    children, _ = _build(body, width=220.4724, height=181.1024)
    texts = {
        "".join(r.text for p in label.text_body.paragraphs for r in p.runs)
        for label in _data_labels(children)
    }
    assert texts == {"3", "5"}


def test_a_line_chart_puts_its_labels_to_the_right_of_the_point():
    """ECMA's default for a line series is `r`, which is what PowerPoint drew."""
    body = line_chart_xml(
        marker="<c:marker><c:symbol val='circle'/><c:size val='7'/></c:marker>"
    ).replace(
        "<c:cat>",
        "<c:dLbls><c:showLegendKey val='0'/><c:showVal val='1'/>"
        "<c:showCatName val='0'/><c:showSerName val='0'/>"
        "<c:showPercent val='0'/><c:showBubbleSize val='0'/></c:dLbls><c:cat>",
        1,
    )
    children, _ = _build(body, width=220.4724, height=181.1024)
    labels = _data_labels(children)
    assert len(labels) == 3
    # The point is at x 52.473; the label starts one marker radius plus a gap right of it.
    assert _label_text_left(labels[0]) == pytest.approx(
        306.067 - 244.094, abs=DATA_LABEL_TOLERANCE_PT
    )


# -- Pie and doughnut ------------------------------------------------------------------
#
# The core geometry is measured on `real-financial-report.pptx`'s own doughnut; the rest
# on a twelve-chart probe.  Frame-relative points throughout.

#: chart4 of real-financial-report: a 285 x 150 pt frame with a right-hand legend.
REAL_DOUGHNUT = {
    "centre": (91.008, 75.000),
    "outer": 64.000,
    "inner": 32.000,          # c:holeSize = 50, so half the outer radius
    "legend_band": 113.98,    # the same band that deck's bar charts reserve
}

PIE_TOLERANCE_PT = 0.4


def pie_chart_xml(
    *,
    kind="pieChart",
    values=(43, 30, 19, 8),
    cats=("Alpha", "Beta", "Gamma", "Delta"),
    first_angle=None,
    hole=None,
    explosion=None,
    vary=None,
    dlbl_show=(),
    dlbl_pos=None,
    series=1,
    legend=None,
):
    body = ""
    for index in range(series):
        points = "".join(
            f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(values)
        )
        cpts = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(cats))
        exp = "" if explosion is None else f"<c:explosion val='{explosion}'/>"
        labels = ""
        if dlbl_show:
            flags = "".join(
                f"<c:show{flag} val='{1 if flag in dlbl_show else 0}'/>"
                for flag in ("LegendKey", "Val", "CatName", "SerName", "Percent",
                             "BubbleSize")
            )
            pos = "" if dlbl_pos is None else f"<c:dLblPos val='{dlbl_pos}'/>"
            labels = f"<c:dLbls>{pos}{flags}</c:dLbls>"
        body += (
            f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
            "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
            f"<c:pt idx='0'><c:v>S{index}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
            f"{exp}{labels}"
            f"<c:cat><c:strRef><c:strCache><c:ptCount val='{len(cats)}'/>{cpts}"
            "</c:strCache></c:strRef></c:cat>"
            "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
            f"<c:ptCount val='{len(values)}'/>{points}"
            "</c:numCache></c:numRef></c:val></c:ser>"
        )
    vary_xml = "" if vary is None else f"<c:varyColors val='{1 if vary else 0}'/>"
    ang = "" if first_angle is None else f"<c:firstSliceAng val='{first_angle}'/>"
    hole_xml = "" if hole is None else f"<c:holeSize val='{hole}'/>"
    legend_xml = (
        f"<c:legend><c:legendPos val='{legend}'/><c:overlay val='0'/></c:legend>"
        if legend
        else ""
    )
    return (
        f"<c:chart><c:plotArea><c:layout/><c:{kind}>{vary_xml}{body}{ang}{hole_xml}"
        f"</c:{kind}></c:plotArea>{legend_xml}</c:chart>"
    )


def _slices(children):
    return [
        child
        for child in children
        if isinstance(child, m.ShapeElement)
        and isinstance(child.geometry, m.CustomGeometry)
    ]


def _path_points(shape):
    transform = shape.transform
    path = shape.geometry.paths[0]
    sx = transform.extent_width / 12700.0 / path.width if path.width else 1.0
    sy = transform.extent_height / 12700.0 / path.height if path.height else 1.0
    points = []
    for token in re.finditer(r"([MLC])((?: -?[\d.]+)+)", path.commands):
        numbers = [float(v) for v in token.group(2).split()]
        for i in range(0, len(numbers) - 1, 2):
            points.append(
                (
                    transform.offset_x / 12700.0 + numbers[i] * sx,
                    transform.offset_y / 12700.0 + numbers[i + 1] * sy,
                )
            )
    return points


def _radii(shape, centre):
    return sorted({round(math.hypot(x - centre[0], y - centre[1]), 2)
                   for x, y in _path_points(shape)})


def test_the_real_doughnut_matches_powerpoints_geometry():
    """chart4 of `real-financial-report.pptx`, whose export pins every polar rule."""
    from tests.conftest import FIXTURE_DIR

    deck = convert_pptx_to_model((FIXTURE_DIR / "real-financial-report.pptx").read_bytes())
    charts = [
        element
        for slide in deck.slides
        for element in slide.elements
        if isinstance(element, m.ChartElement)
        and element.chart.kind == "doughnutChart"
    ]
    assert len(charts) == 1
    parts = _slices(charts[0].children)
    assert len(parts) == 4
    centre = REAL_DOUGHNUT["centre"]
    for part in parts:
        radii = _radii(part, centre)
        # The path visits both arcs; its Bezier control points sit off them either way, so
        # the test is that each radius is present rather than that it bounds the set.
        for wanted in (REAL_DOUGHNUT["inner"], REAL_DOUGHNUT["outer"]):
            assert any(abs(r - wanted) < PIE_TOLERANCE_PT for r in radii), (wanted, radii)


def test_angle_zero_is_twelve_oclock_and_slices_run_clockwise():
    """The real doughnut's first slice, a 43% share, ends at 154.80 deg = 43% of 360."""
    children, _ = _build(
        pie_chart_xml(values=(43, 30, 19, 8)), width=220.4724, height=181.1024
    )
    first = _slices(children)[0]
    centre = (110.236, 90.551)
    start = _path_points(first)[0]
    angle = math.degrees(
        math.atan2(start[0] - centre[0], -(start[1] - centre[1]))
    ) % 360
    assert angle == pytest.approx(0.0, abs=0.5) or angle == pytest.approx(360.0, abs=0.5)


def test_first_slice_angle_rotates_clockwise_from_twelve():
    children, _ = _build(
        pie_chart_xml(first_angle=90), width=220.4724, height=181.1024
    )
    start = _path_points(_slices(children)[0])[0]
    centre = (110.236, 90.551)
    angle = math.degrees(
        math.atan2(start[0] - centre[0], -(start[1] - centre[1]))
    ) % 360
    assert angle == pytest.approx(90.0, abs=0.5)


def test_hole_size_is_a_percentage_of_the_outer_radius():
    children, _ = _build(
        pie_chart_xml(kind="doughnutChart", hole=25), width=220.4724, height=181.1024
    )
    centre = (110.236, 90.551)
    radii = _radii(_slices(children)[0], centre)
    for wanted in (79.551, 0.25 * 79.551):
        assert any(abs(r - wanted) < PIE_TOLERANCE_PT for r in radii), (wanted, radii)


def test_a_doughnut_with_no_hole_size_draws_as_a_pie():
    """ECMA-376 documents a default of 10; PowerPoint's export has no inner arc at all."""
    children, _ = _build(
        pie_chart_xml(kind="doughnutChart"), width=220.4724, height=181.1024
    )
    centre = (110.236, 90.551)
    assert min(_radii(_slices(children)[0], centre)) == pytest.approx(0.0, abs=0.05)


def test_several_series_make_concentric_rings():
    children, _ = _build(
        pie_chart_xml(kind="doughnutChart", series=2), width=220.4724, height=181.1024
    )
    parts = _slices(children)
    assert len(parts) == 8
    centre = (110.236, 90.551)
    band = 79.551 / 2
    assert any(abs(r - band) < PIE_TOLERANCE_PT for r in _radii(parts[0], centre))
    outer_ring = _radii(parts[4], centre)
    for wanted in (band, 79.551):
        assert any(abs(r - wanted) < PIE_TOLERANCE_PT for r in outer_ring), (
            wanted, outer_ring
        )


def test_explosion_shrinks_the_radius_and_pushes_each_slice_out():
    """Measured: radius x 1/(1+e), offset e x the *shrunk* radius along the bisector."""
    children, _ = _build(pie_chart_xml(explosion=20), width=220.4724, height=181.1024)
    centre = (110.236, 90.551)
    radius = 79.551 / 1.2
    first = _slices(children)[0]
    points = _path_points(first)
    # The wedge's apex is the exploded centre.
    apex = min(points, key=lambda p: math.hypot(p[0] - centre[0], p[1] - centre[1]))
    offset = math.hypot(apex[0] - centre[0], apex[1] - centre[1])
    assert offset == pytest.approx(radius * 0.2, abs=PIE_TOLERANCE_PT)
    far = max(math.hypot(x - apex[0], y - apex[1]) for x, y in points)
    assert far >= radius - PIE_TOLERANCE_PT


def test_a_pie_varies_its_colours_by_point_without_being_asked():
    """A bar chart with no `c:varyColors` draws one colour; a pie cycles the accents."""
    children, _ = _build(pie_chart_xml(), width=220.4724, height=181.1024)
    fills = [
        part.fill.color.hex.upper()
        for part in _slices(children)
        if isinstance(part.fill, m.SolidFill)
    ]
    assert fills == ["#4472C4", "#ED7D31", "#A5A5A5", "#4472C4"]


def test_percentages_add_up_to_one_hundred():
    """Measured: three equal values are labelled 34%, 33%, 33%, not 33% three times."""
    from pptx2svg.resolve.chart import _percent_shares

    assert _percent_shares([1, 1, 1]) == [34, 33, 33]
    assert sum(_percent_shares([1, 1, 1, 1, 1, 1])) == 100
    assert _percent_shares([43, 30, 19, 8]) == [43, 30, 19, 8]
    assert _percent_shares([0, 0]) == [0, 0]


def test_a_slice_label_sits_at_the_measured_fraction_of_the_radius():
    centre = (110.236, 90.551)
    radius = 79.551
    for position, fraction in (("ctr", 0.500), ("inEnd", 0.856),
                               ("outEnd", 1.020), ("bestFit", 0.710)):
        children, _ = _build(
            pie_chart_xml(dlbl_show=("Val",), dlbl_pos=position),
            width=220.4724,
            height=181.1024,
        )
        labels = [
            child
            for child in children
            if isinstance(child, m.ShapeElement) and child.text_body is not None
        ]
        assert len(labels) == 4, position
        # The first slice spans 0..154.8 deg, so its bisector is at 77.4.
        first = labels[0]
        size = first.text_body.paragraphs[0].runs[0].properties.font_size
        cx = first.transform.offset_x / 12700.0 + first.transform.extent_width / 25400.0
        cy = first.transform.offset_y / 12700.0 + (1.2 - APTOS_DESCENT) * size - size * 0.3
        distance = math.hypot(cx - centre[0], cy - centre[1])
        assert distance / radius == pytest.approx(fraction, abs=0.06), position


def test_a_pie_legends_its_categories_not_its_series():
    children, _ = _build(
        pie_chart_xml(legend="r"), width=220.4724, height=181.1024
    )
    texts = {
        "".join(r.text for p in child.text_body.paragraphs for r in p.runs)
        for child in children
        if isinstance(child, m.ShapeElement) and child.text_body is not None
    }
    assert {"Alpha", "Beta", "Gamma", "Delta"} <= texts
    assert "S0" not in texts
# -- Radar ------------------------------------------------------------------------------
#
# Measured on eighteen probe charts across three decks, all in the same
# 220.4724 x 181.1024 pt frame, exported by PowerPoint 16.106 and read back out of the PDF
# as exact vector coordinates -- plus `real-financial-report.pptx`'s own radar, whose
# export is in the corpus and whose web, spokes, rings and polygons are all measurable
# there despite the deck being skipped by the fidelity harness for want of Noto Sans JP.
# Vector geometry does not move when a font is substituted, which is why the corpus can
# pin the polar conventions that the brief for this work assumed only probes could.
#
# Frame-relative points throughout.  The frame centre is (110.2362, 90.5512).

RADAR_CENTRE = (110.2362, 90.5512)
RADAR_TOLERANCE_PT = 0.35

#: PowerPoint's own radius for each probe, in points, against a plot region whose half
#: height is 79.5512 pt and half width 99.2362 pt.  Five of these are bound by the
#: vertical label reserve and three by the horizontal one; `no-cat-axis` has no labels at
#: all and is the anchor that says the reserve goes to zero.
RADAR_RADIUS = {
    "standard": ({}, 69.47),
    "aptos-8": ({"size": 8.0}, 72.48),
    "aptos-14": ({"size": 14.0}, 63.36),
    "arial-10": ({"size": 10.0, "face": "Arial"}, 70.80),
    "arial-14": ({"size": 14.0, "face": "Arial"}, 65.04),
    "no-cat-axis": ({"delete_cat": True}, 79.44),
    "long-labels": (
        {
            "cats": (
                "Category One", "Category Two", "Category Three",
                "Category Four", "Category Five",
            )
        },
        60.24,
    ),
    "two-line": (
        {"cats": ("One Two", "Two Three", "Red Blue", "Six Ten", "Sun Moon")},
        55.44,
    ),
    "four-cats": ({"cats": ("A", "B", "C", "D"), "values": ((3, 4, 5, 2),)}, 69.47),
    "eight-cats": (
        {"cats": tuple("ABCDEFGH"), "values": ((3, 4, 5, 2, 1, 3, 4, 2),)},
        69.47,
    ),
}


def radar_chart_xml(
    *,
    style="standard",
    cats=("A", "B", "C", "D", "E"),
    values=((3, 4, 5, 2, 1),),
    legend=None,
    size=None,
    face=None,
    delete_cat=False,
    delete_val=False,
    gridlines=True,
    cat_axis_line=False,
    series_line=None,
    dlbl_show=(),
):
    """One radar probe, in the same shape the exported decks used."""
    body = ""
    for index, row in enumerate(values):
        points = "".join(
            f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(row)
        )
        cpts = "".join(
            f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(cats)
        )
        sp = (
            f"<c:spPr><a:solidFill><a:srgbClr val='2563EB'/></a:solidFill>"
            f"<a:ln w='{series_line}'>"
            "<a:solidFill><a:srgbClr val='2563EB'/></a:solidFill></a:ln></c:spPr>"
            if series_line
            else ""
        )
        labels = ""
        if dlbl_show:
            flags = "".join(
                f"<c:show{flag} val='{1 if flag in dlbl_show else 0}'/>"
                for flag in ("LegendKey", "Val", "CatName", "SerName", "Percent",
                             "BubbleSize")
            )
            labels = f"<c:dLbls>{flags}</c:dLbls>"
        # CT_RadarSer's sequence is strict: idx, order, tx, spPr, marker, dPt, dLbls,
        # cat, val.  Out of order is one of the four known PowerPoint hang signatures.
        body += (
            f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
            "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
            f"<c:pt idx='0'><c:v>Series {index + 1}</c:v></c:pt>"
            "</c:strCache></c:strRef></c:tx>"
            f"{sp}{labels}"
            f"<c:cat><c:strRef><c:strCache><c:ptCount val='{len(cats)}'/>{cpts}"
            "</c:strCache></c:strRef></c:cat>"
            "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
            f"<c:ptCount val='{len(row)}'/>{points}"
            "</c:numCache></c:numRef></c:val></c:ser>"
        )
    legend_xml = (
        f"<c:legend><c:legendPos val='{legend}'/><c:overlay val='0'/></c:legend>"
        if legend
        else ""
    )
    tx_pr = (
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        f"<a:defRPr sz='{int(size * 100)}'>"
        + (f"<a:latin typeface='{face}'/>" if face else "")
        + "</a:defRPr></a:pPr><a:endParaRPr lang='en-US'/></a:p></c:txPr>"
        if size
        else ""
    )
    cat_line = (
        "<c:spPr><a:ln w='12700'><a:solidFill><a:srgbClr val='888888'/></a:solidFill>"
        "</a:ln></c:spPr>"
        if cat_axis_line
        else ""
    )
    grid = "<c:majorGridlines/>" if gridlines else ""
    return (
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        f"<c:radarChart><c:radarStyle val='{style}'/><c:varyColors val='0'/>{body}"
        "<c:axId val='100002'/><c:axId val='100003'/></c:radarChart>"
        "<c:catAx><c:axId val='100002'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        f"<c:delete val='{1 if delete_cat else 0}'/><c:axPos val='b'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        f"<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        f"<c:tickLblPos val='nextTo'/>{cat_line}<c:crossAx val='100003'/>"
        "<c:crosses val='autoZero'/><c:auto val='1'/><c:lblAlgn val='ctr'/>"
        "<c:lblOffset val='100'/></c:catAx>"
        "<c:valAx><c:axId val='100003'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        f"<c:delete val='{1 if delete_val else 0}'/><c:axPos val='l'/>{grid}"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='out'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100002'/>"
        "<c:crosses val='autoZero'/><c:crossBetween val='between'/></c:valAx>"
        f"</c:plotArea>{legend_xml}<c:plotVisOnly val='1'/>"
        f"<c:dispBlanksAs val='gap'/></c:chart>{tx_pr}"
    )


def _radar(**kwargs):
    children, _ = _build(
        radar_chart_xml(**kwargs), width=220.4724, height=181.1024
    )
    return children


def _paths(children):
    """Every custom-geometry shape, with its points in frame coordinates."""
    out = []
    for child in children:
        if not isinstance(child, m.ShapeElement):
            continue
        if not isinstance(child.geometry, m.CustomGeometry):
            continue
        out.append((child, _path_points(child)))
    return out


def _radius_of(children, centre=RADAR_CENTRE):
    points = [point for _, points in _paths(children) for point in points]
    return max(math.hypot(x - centre[0], y - centre[1]) for x, y in points)


@pytest.mark.parametrize("name", list(RADAR_RADIUS))
def test_the_radar_radius_matches_powerpoints(name):
    """The whole layout in one number, across two faces, three sizes and four shapes."""
    kwargs, expected = RADAR_RADIUS[name]
    ours = _radius_of(_radar(**kwargs))
    assert ours == pytest.approx(expected, abs=RADAR_TOLERANCE_PT), (
        f"{name}: PowerPoint {expected:.2f}, ours {ours:.2f}"
    )


def test_the_web_is_polygonal_and_follows_the_category_count():
    """Three, five, six and eight categories gave triangles, pentagons, hexagons, octagons."""
    for count, cats in ((3, "ABC"), (5, "ABCDE"), (6, "ABCDEF"), (8, "ABCDEFGH")):
        values = tuple((index % 5) + 1 for index in range(count))
        rings = [
            points
            for shape, points in _paths(_radar(cats=tuple(cats), values=(values,)))
            if isinstance(shape.fill, m.NoFill)
        ]
        assert rings, count
        # A ring closes, so it visits one more point than it has corners.
        assert all(len(points) == count + 1 for points in rings), (
            count, [len(p) for p in rings]
        )


def test_angle_zero_is_twelve_oclock_and_categories_run_clockwise():
    """Measured on the four-category probe: due north, east, south then west."""
    children = _radar(cats=("A", "B", "C", "D"), values=((5, 5, 5, 5),))
    ring = max(
        (
            points
            for shape, points in _paths(children)
            if isinstance(shape.fill, m.NoFill)
        ),
        key=lambda points: max(
            math.hypot(x - RADAR_CENTRE[0], y - RADAR_CENTRE[1]) for x, y in points
        ),
    )
    radius = max(
        math.hypot(x - RADAR_CENTRE[0], y - RADAR_CENTRE[1]) for x, y in ring
    )
    expected = [
        (RADAR_CENTRE[0], RADAR_CENTRE[1] - radius),
        (RADAR_CENTRE[0] + radius, RADAR_CENTRE[1]),
        (RADAR_CENTRE[0], RADAR_CENTRE[1] + radius),
        (RADAR_CENTRE[0] - radius, RADAR_CENTRE[1]),
    ]
    for wanted in expected:
        assert any(
            abs(x - wanted[0]) < 0.05 and abs(y - wanted[1]) < 0.05 for x, y in ring
        ), (wanted, ring)


def test_a_point_sits_at_its_fraction_of_the_radius():
    """Values 3, 4, 5, 2, 1 on a 0..5 axis: 0.6, 0.8, 1.0, 0.4 and 0.2 of the radius."""
    children = _radar()
    radius = _radius_of(children)
    series = [
        points
        for shape, points in _paths(children)
        if isinstance(shape.fill, m.NoFill) and len(points) == 6
    ]
    # Rings and the series polygon are all five-sided; the series is the one whose
    # vertices are at different radii.
    ragged = [
        points
        for points in series
        if len({
            round(math.hypot(x - RADAR_CENTRE[0], y - RADAR_CENTRE[1]), 1)
            for x, y in points
        }) > 1
    ]
    assert len(ragged) == 1
    distances = [
        math.hypot(x - RADAR_CENTRE[0], y - RADAR_CENTRE[1]) for x, y in ragged[0][:5]
    ]
    assert [d / radius for d in distances] == pytest.approx(
        [0.6, 0.8, 1.0, 0.4, 0.2], abs=0.01
    )


def test_a_radars_axis_stops_at_the_data_where_a_bars_goes_past_it():
    """The one discriminating observation: 0..5 of data draws five rings, not six."""
    from pptx2svg.resolve.chart import nice_axis_scale

    assert nice_axis_scale(0, 5) == (0.0, 6.0, 1.0)
    assert nice_axis_scale(0, 5, strict=False) == (0.0, 5.0, 1.0)

    children = _radar()
    rings = [
        points
        for shape, points in _paths(children)
        if isinstance(shape.fill, m.NoFill) and len(points) == 6
    ]
    # Five rings plus the series polygon.
    assert len(rings) == 6


def test_standard_and_marker_draw_the_same_picture():
    """ECMA-376 says `standard` has no markers; PowerPoint's two exports are identical."""
    def summary(style):
        children = _radar(style=style)
        return [
            (
                type(child).__name__,
                round(child.transform.offset_x / 12700.0, 3),
                round(child.transform.offset_y / 12700.0, 3),
            )
            for child in children
        ]

    assert len(summary("standard")) > 10, "nothing drawn, so the comparison is vacuous"
    assert summary("standard") == summary("marker")


def test_a_filled_radar_fills_and_does_not_stroke_or_mark():
    """Measured: the probe, which states no `a:ln`, emits a bare `f` and no markers."""
    filled = [
        shape
        for shape, _ in _paths(_radar(style="filled"))
        if isinstance(shape.fill, m.SolidFill)
    ]
    assert len(filled) == 1
    assert filled[0].outline is None
    # No markers either: every remaining shape is a web ring.
    markers = [
        child
        for child in _radar(style="filled")
        if isinstance(child, m.ShapeElement)
        and isinstance(child.geometry, m.PresetGeometry)
        and child.geometry.preset in ("diamond", "ellipse", "rect")
        and child.text_body is None
    ]
    assert markers == []


def test_a_filled_radar_strokes_when_the_series_states_a_line():
    """The other half of the pair: the corpus radar says `w="25400"` and is stroked."""
    filled = [
        shape
        for shape, _ in _paths(_radar(style="filled", series_line=25400))
        if isinstance(shape.fill, m.SolidFill)
    ]
    assert len(filled) == 1
    assert filled[0].outline is not None
    assert filled[0].outline.width == 25400


def test_the_web_is_drawn_without_major_gridlines_and_with_the_value_axis_deleted():
    """Both probes still drew every ring, so the web is not the value axis' gridlines."""
    def rings(**kwargs):
        return [
            points
            for shape, points in _paths(_radar(**kwargs))
            if isinstance(shape.fill, m.NoFill) and len(points) == 6
        ]

    assert len(rings()) == len(rings(gridlines=False)) == len(rings(delete_val=True))


def test_spokes_are_drawn_only_when_the_category_axis_states_a_line():
    """No probe without a `c:spPr` drew any; the corpus radar's #888888 line drew six."""
    assert [c for c in _radar() if isinstance(c, m.ConnectorElement)] == []
    spokes = [c for c in _radar(cat_axis_line=True) if isinstance(c, m.ConnectorElement)]
    assert len(spokes) == 5
    assert all(spoke.outline.width == 12700 for spoke in spokes)


def test_the_value_labels_are_right_aligned_two_digits_left_of_the_spoke():
    """Measured on six charts across two faces and three sizes, worst residual 0.14 pt."""
    from pptx2svg.resolve.chart import text_width

    for size, face, pen in ((10.0, None, 94.216), (8.0, None, 97.411),
                            (14.0, None, 87.796), (10.0, "Arial", 93.556),
                            (14.0, "Arial", 86.881)):
        family = face or "Aptos"
        children = _radar(size=size, face=face)
        labels = [
            child
            for child in children
            if isinstance(child, m.ShapeElement) and child.text_body is not None
            and "".join(
                r.text for p in child.text_body.paragraphs for r in p.runs
            ) == "0"
        ]
        assert len(labels) == 1, (size, face)
        box = labels[0].transform
        # The text is right-aligned in its box with no right inset, so the box's right
        # edge is the text's.  PowerPoint's pen x plus the digit's advance is where that
        # edge landed.
        right = (box.offset_x + box.extent_width) / 12700.0
        expected = pen + text_width("0", family, size)
        assert right == pytest.approx(expected, abs=0.2), (size, face)
        # And that edge is two digit widths left of the spoke, which is the rule itself.
        assert right == pytest.approx(
            RADAR_CENTRE[0] - 2 * text_width("0", family, size), abs=0.01
        ), (size, face)


def test_a_category_label_wraps_rather_than_rotating():
    """Every `rot` in the long-label probe is zero, and "Category Three" is two lines."""
    from pptx2svg.resolve.chart import _wrap_to_width, font_box, ChartFont

    font = ChartFont(family="Aptos", box=font_box("Aptos", 10.0))
    cap = 198.4724 * 0.25
    assert _wrap_to_width("Category Three", font, cap) == ["Category", "Three"]
    # The bracket: 43.72 pt stayed on one line, 59.10 pt wrapped.
    assert _wrap_to_width("Two Three", font, cap) == ["Two Three"]
    assert _wrap_to_width("Category One", font, cap) == ["Category", "One"]
    # A single word is never split.
    assert _wrap_to_width("Supercalifragilistic", font, 10.0) == [
        "Supercalifragilistic"
    ]


def test_a_side_category_label_is_anchored_at_the_vertex():
    """Measured pen positions: B at 178.945, C at 152.701, D right edge 68.132."""
    children = _radar()
    found = {}
    for child in children:
        if not isinstance(child, m.ShapeElement) or child.text_body is None:
            continue
        text = "".join(r.text for p in child.text_body.paragraphs for r in p.runs)
        if text in ("B", "C", "D", "E"):
            size = child.text_body.paragraphs[0].runs[0].properties.font_size
            left = child.transform.offset_x / 12700.0 + size / 2
            found[text] = (left, left + child.transform.extent_width / 12700.0 - size)
    assert found["B"][0] == pytest.approx(178.945, abs=0.4)
    assert found["C"][0] == pytest.approx(152.701, abs=0.4)
    assert found["D"][1] == pytest.approx(68.132, abs=0.5)
    assert found["E"][1] == pytest.approx(42.637, abs=1.2)


def test_a_line_style_radar_legends_with_a_line_and_marker_not_a_swatch():
    """Measured: a 19.200 pt rule with the marker on it, then 2.025 pt before the text."""
    children = _radar(legend="r")
    keys = [c for c in children if isinstance(c, m.ConnectorElement)]
    assert len(keys) == 1
    assert keys[0].transform.extent_width / 12700.0 == pytest.approx(19.2, abs=0.01)
    assert keys[0].transform.offset_x / 12700.0 == pytest.approx(154.957, abs=0.4)
    # A filled radar takes the ordinary swatch instead.
    assert [
        c for c in _radar(style="filled", legend="r")
        if isinstance(c, m.ConnectorElement)
    ] == []


def test_a_radar_legends_its_series_not_its_categories():
    texts = {
        "".join(r.text for p in child.text_body.paragraphs for r in p.runs)
        for child in _radar(legend="r")
        if isinstance(child, m.ShapeElement) and child.text_body is not None
    }
    assert "Series 1" in texts


def test_a_radar_series_with_no_marker_size_gets_six_points_not_seven():
    """Measured 6.0 pt square on the probe's second series; ECMA-376's default is 7."""
    markers = [
        child
        for child in _radar(values=((3, 4, 5, 2, 1), (1, 2, 3, 4, 5)))
        if isinstance(child, m.ShapeElement)
        and isinstance(child.geometry, m.PresetGeometry)
        and child.geometry.preset in ("diamond", "rect")
        and child.text_body is None
    ]
    assert markers, "no markers drawn"
    assert all(
        marker.transform.extent_width / 12700.0 == pytest.approx(6.0, abs=0.01)
        for marker in markers
    )
    # The cycle is diamond then square, the same one a line chart uses.
    presets = [marker.geometry.preset for marker in markers]
    assert presets[:5] == ["diamond"] * 5
    assert presets[5:] == ["rect"] * 5


def test_a_blank_leaves_the_radar_ring_open():
    """Not measured -- mirrors the line chart, whose blank behaviour was.

    A whole ring visits six points (five corners and the close).  A blank at one corner
    leaves **one** open run of four, not two runs of two and three: the ring is a cycle,
    so the stretch that passes through index 0 is a single run.  No marker is drawn for
    the missing point.
    """
    children = _radar(values=((3, 4, None, 2, 1),))
    runs = sorted(
        len(points)
        for shape, points in _paths(children)
        if isinstance(shape.fill, m.NoFill) and len(points) != 6
    )
    assert runs == [4]
    markers = [
        child
        for child in children
        if isinstance(child, m.ShapeElement)
        and isinstance(child.geometry, m.PresetGeometry)
        and child.geometry.preset == "diamond"
    ]
    assert len(markers) == 4


def test_the_real_radar_matches_powerpoints_geometry():
    """chart5 of `real-financial-report.pptx`, the only radar in the corpus.

    PowerPoint's own export puts the centre at (132.72, 75.60) with a radius of 45.56 pt,
    so its label reserve is 18.44 pt of the 64 pt half-region.

    **The radius was 6.83 pt out and most of that is now gone.**  The category labels are
    Japanese; the chart names `<a:latin typeface="Arial"/>` and no `<a:ea>` at all, and
    `font_box` used to read the Latin face for a label the Latin face cannot draw a single
    glyph of.  Arial's line box is 1.117 em, which the fitted reserve turns into 11.61 pt.
    The face PowerPoint actually drew those labels in is the theme's
    `<a:font script="Jpan" typeface="游ゴシック"/>` -- its export embeds YuGothic-Regular
    for them -- and our table for it has a 1.448 em line box, giving 16.61 pt and a radius
    of 47.39.

    The 1.83 pt left is the fitted reserve itself, not the face: `RADAR_LABEL_RESERVE_*`
    was fitted to five Latin probes and reproducing PowerPoint's 18.44 exactly would want
    a 1.5696 em line box, which is neither Noto Sans JP's 1.448 nor Yu Gothic's own hhea
    figure.  Whether the reserve has a term that only CJK exercises is unmeasured; there
    is one CJK radar in the corpus and no probe deck for it.
    """
    from tests.conftest import FIXTURE_DIR

    deck = convert_pptx_to_model((FIXTURE_DIR / "real-financial-report.pptx").read_bytes())
    charts = [
        element
        for slide in deck.slides
        for element in slide.elements
        if isinstance(element, m.ChartElement) and element.chart.kind == "radarChart"
    ]
    assert len(charts) == 1
    chart = charts[0]
    assert chart.chart.value_axis.minimum == 0.0
    assert chart.chart.value_axis.maximum == 100.0
    assert chart.chart.value_axis.major_unit == 50.0

    points = [point for _, points in _paths(chart.children) for point in points]
    xs = [x for x, _ in points]
    ys = [y for _, y in points]
    assert (min(xs) + max(xs)) / 2 == pytest.approx(132.720, abs=0.2)
    assert (min(ys) + max(ys)) / 2 == pytest.approx(75.600, abs=0.7)
    # The outer ring is the 100 % series, so its half-height is the drawn radius.
    radius = (max(ys) - min(ys)) / 2
    assert radius == pytest.approx(47.39, abs=0.1)
    assert abs(radius - 45.56) < abs(52.39 - 45.56)

    # Two filled series, both stroked because the file states `a:ln w="25400"`, and six
    # spokes because its category axis states a #888888 line.
    filled = [
        shape for shape, _ in _paths(chart.children)
        if isinstance(shape.fill, m.SolidFill)
    ]
    assert [shape.fill.color.hex.upper() for shape in filled] == ["#2563EB", "#94A3B8"]
    assert all(shape.outline is not None for shape in filled)
    assert len([c for c in chart.children if isinstance(c, m.ConnectorElement)]) == 6


# -- Rotated category labels ------------------------------------------------------------
#
# Two probe decks of six bar charts, all in the same 220.4724 x 181.1024 pt frame with
# five categories and 10 pt Aptos labels, exported by PowerPoint 16.106 and read back out
# of the PDF as exact vector coordinates.  The first sweeps the label from a fifth of its
# band to four times it; the second straddles exactly one band width, which is what turns
# the threshold from a guess into a 5% window.
#
# `real-financial-report.pptx`'s own chart3 is the third source: its export rotates too,
# at 12 pt, which is where the 45 degrees is confirmed off a real deck rather than a probe.

#: The band every chart in both decks has: (209.47 - 21.07) / 5, measured off the
#: gridlines.
ROTATION_BAND_PT = 37.68

#: name -> (categories, widest label's width in pt, does PowerPoint turn them,
#: PowerPoint's bottom inset in pt).  The width is the widest of the five, which is what
#: sets the band -- the deck whose five labels differ only by a trailing letter showed
#: that to 1.01 pt, and using the first label instead left a 0.73 pt residual everywhere.
ROTATION_SWEEP = {
    # The straddle: 36.62 pt on a 37.68 pt band stays level, 38.59 pt turns.
    "r085": (("000000",) * 5, 32.05, False, 24.965),
    "r092": (("CCCCC",) * 5, 34.62, False, 24.965),
    "r097": (("OOOOO",) * 5, 36.62, False, 24.965),
    "r102": (("hhhhhhh",) * 5, 38.59, True, 48.688),
    "r108": (("vvvvvvvvv",) * 5, 40.69, True, 50.155),
    "r115": (("wwwwww",) * 5, 43.24, True, 51.959),
    # The range.
    "w1": (tuple("ABCDE"), 6.90, False, 24.965),
    "w2": (tuple(f"Cat{n}" for n in "ABCDE"), 22.37, False, 24.965),
    "w3": (tuple(f"Category{n}" for n in "ABCDE"), 45.87, True, 53.843),
    "w4": (tuple(f"CategoryLong{n}" for n in "ABCDE"), 66.74, True, 68.604),
    "w5": (tuple(f"CategoryLongerStill{n}" for n in "ABCDE"), 91.84, True, 86.356),
}

#: The unrotated band is a different formula and was fitted separately, so it keeps the
#: older sweep's tolerance; the rotated one lands an order of magnitude closer.
ROTATION_TOLERANCE_PT = 0.15
LEVEL_TOLERANCE_PT = 0.2


def rotation_chart_xml(cats, values=None):
    """A plain clustered column chart whose only variable is its category labels."""
    values = values or tuple(3 + (index % 3) for index in range(len(cats)))
    points = "".join(
        f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(values)
    )
    cpts = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(cats))
    return (
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        "<c:barChart><c:barDir val='col'/><c:grouping val='clustered'/>"
        "<c:varyColors val='0'/>"
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>Coverage</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        "<c:spPr><a:solidFill><a:srgbClr val='2563EB'/></a:solidFill></c:spPr>"
        f"<c:cat><c:strRef><c:strCache><c:ptCount val='{len(cats)}'/>{cpts}"
        "</c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{points}"
        "</c:numCache></c:numRef></c:val></c:ser>"
        "<c:axId val='100002'/><c:axId val='100003'/></c:barChart>"
        "<c:catAx><c:axId val='100002'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='b'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='none'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100003'/>"
        "<c:crosses val='autoZero'/><c:auto val='1'/><c:lblAlgn val='ctr'/>"
        "<c:lblOffset val='100'/></c:catAx>"
        "<c:valAx><c:axId val='100003'/>"
        "<c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='l'/><c:majorGridlines/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='none'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100002'/>"
        "<c:crosses val='autoZero'/><c:crossBetween val='between'/></c:valAx>"
        "</c:plotArea><c:plotVisOnly val='1'/>"
        "<c:dispBlanksAs val='gap'/></c:chart>"
    )


def _rotation_probe(cats, **kwargs):
    return _build(
        rotation_chart_xml(cats, **kwargs), width=220.4724, height=181.1024
    )[0]


def _turned(children):
    """Every label the resolver rotated, left to right."""
    turned = [
        child
        for child in children
        if isinstance(child, m.ShapeElement)
        and child.text_body is not None
        and child.transform.rotation
    ]
    return sorted(turned, key=lambda shape: shape.transform.offset_x)


def _bottom_inset(children, frame_height=181.1024):
    return frame_height - _pt(_plot_bottom(children))


@pytest.mark.parametrize("name", list(ROTATION_SWEEP))
def test_a_category_label_turns_when_it_is_wider_than_its_band(name):
    """The threshold, and the 5% window the probe pins it inside.

    ``OOOOO`` is 36.62 pt on a 37.68 pt band and PowerPoint left it level; ``hhhhhhh`` is
    38.59 pt on the same band and PowerPoint turned it.  One band width is not a round
    guess -- it is the middle of (0.972, 1.024].
    """
    cats, width, turns, _ = ROTATION_SWEEP[name]
    children = _rotation_probe(cats)
    assert bool(_turned(children)) is turns, (
        f"{name}: widest label {width:.2f} pt on a {ROTATION_BAND_PT:.2f} pt band, "
        f"{width / ROTATION_BAND_PT:.3f} of it"
    )


@pytest.mark.parametrize("name", list(ROTATION_SWEEP))
def test_the_band_under_the_plot_matches_powerpoints(name):
    """Both formulas in one number: the level band and the turned one.

    Turned, it is ``21.39 + width * sin 45``, and the six rotated rows land within
    0.015 pt of PowerPoint.  Level, it is the older ``6.5 + lineHeight + 0.615 em``,
    which is 0.11 pt light and was fitted before this sweep existed.
    """
    cats, _, turns, expected = ROTATION_SWEEP[name]
    ours = _bottom_inset(_rotation_probe(cats))
    tolerance = ROTATION_TOLERANCE_PT if turns else LEVEL_TOLERANCE_PT
    assert ours == pytest.approx(expected, abs=tolerance), (
        f"{name}: PowerPoint {expected:.3f}, ours {ours:.3f}"
    )


def test_a_turned_label_snaps_to_forty_five_degrees():
    """Twelve probes from 1.02 band widths to 4.18 all came out at exactly 45.

    No intermediate angle appeared anywhere in that range and nothing went to 90, so the
    angle is a constant rather than a function of the crowding.
    """
    for name, (cats, _, turns, _inset) in ROTATION_SWEEP.items():
        if not turns:
            continue
        turned = _turned(_rotation_probe(cats))
        assert len(turned) == 5, name
        assert {shape.transform.rotation for shape in turned} == {-45.0}, name


def test_a_turned_label_lands_where_powerpoint_put_it():
    """The anchor is the far end of the rotated baseline, not the box.

    PowerPoint's pens for ``CategoryA``..``CategoryE`` are 10.781, 48.238, 85.171,
    122.776 and 161.254, all on frame row 171.600.  Recovering ours means undoing the
    rotation about the box centre, which is the same arithmetic the resolver does forwards.
    """
    from pptx2svg.resolve.chart import font_box

    children = _rotation_probe(tuple(f"Category{n}" for n in "ABCDE"))
    box = font_box("Aptos", 10.0)
    radians = math.radians(-45.0)
    cos, sin = math.cos(radians), math.sin(radians)
    pens = []
    for shape in _turned(children):
        transform = shape.transform
        width = _pt(transform.extent_width)
        height = _pt(transform.extent_height)
        centre_x = _pt(transform.offset_x) + width / 2
        centre_y = _pt(transform.offset_y) + height / 2
        # The *left* end of the right-aligned baseline, inside the unrotated box.
        offset_x = -width / 2 + box.size / 2
        offset_y = box.first_baseline - height / 2
        pens.append(
            (
                centre_x + offset_x * cos - offset_y * sin,
                centre_y + offset_x * sin + offset_y * cos,
            )
        )
    expected = [10.781, 48.238, 85.171, 122.776, 161.254]
    assert len(pens) == len(expected)
    for (x, y), truth in zip(pens, expected):
        assert x == pytest.approx(truth, abs=0.6), (x, truth)
        assert y == pytest.approx(171.600, abs=0.9), y


def test_a_label_four_times_its_band_is_not_capped_the_way_powerpoints_is():
    """The one measurement the shipped formula does **not** reproduce, kept visible.

    PowerPoint reserved 85.628 pt for a 130.88 pt label -- *less* than the 86.356 pt it
    gave the 91.84 pt label one step below it -- so no clamp on the width produces both,
    and whatever it does past about 90 pt of label was never identified.  Ours keeps
    going up the fitted line.  This test asserts today's behaviour so the divergence is
    a recorded number rather than a surprise.
    """
    cats = tuple(f"CategoryLongerStillAndMore{n}" for n in "ABCDE")
    ours = _bottom_inset(_rotation_probe(cats))
    assert ours == pytest.approx(113.9, abs=0.6)
    assert ours > 85.628 + 20


def test_a_horizontal_chart_does_not_turn_the_labels_under_it():
    """Its bottom band holds *value* ticks, and no probe measured those rotating.

    The same labels on a ``col`` chart do turn, which is what makes this a statement
    about orientation rather than a test that passes because nothing rotates anywhere.
    """
    cats = tuple(f"CategoryLongerStill{n}" for n in "ABCDE")
    assert len(_turned(_rotation_probe(cats))) == 5
    body = rotation_chart_xml(cats).replace(
        "<c:barDir val='col'/>", "<c:barDir val='bar'/>"
    )
    children, _ = _build(body, width=220.4724, height=181.1024)
    assert _turned(children) == []


def test_the_real_bar_chart_with_long_labels_turns_them():
    """chart3 of `real-financial-report.pptx`, the corpus deck this was worth doing for.

    PowerPoint turns its Japanese category labels 45 degrees -- its export's text matrix
    is 8.4853 at 12 pt, which is ``12 * cos 45`` -- and reserves **69.538 pt** under the
    plot.  We turn them too, which we did not before, and reserve 89.3.

    **The remaining 19.8 pt is not the width, and that was the standing diagnosis until
    PowerPoint's own export was read.**  This file used to say ``プラットフォーム`` came
    out 96 pt where PowerPoint laid it out "at about 68"; the 68 was this formula inverted
    through PowerPoint's inset, never a measurement.  The export settles it: every glyph
    of every Japanese label on this chart advances exactly 12.000 pt -- one em of the
    YuGothic-Regular it embeds for them -- so the string is 96 pt drawn and 96 pt
    measured, and ``デジタル`` (48), ``グローバル`` (60) and ``その他`` (36) match to
    0.000 pt as well.

    What the export does show is something nothing here models: **PowerPoint truncated the
    label.**  It drew ``プラット…`` -- four katakana and an ellipsis -- where the category
    is eight characters, and it did the same to two of the radar's six.  The reserve it
    kept, 69.538 pt, implies a label width of 68.09 through this formula, which is 1.06 of
    the 64.39 pt band; the real label is 1.49 bands.  That is the same unidentified cap
    ``_bottom_label_band`` already records for a 4.18-band probe, seen from the other side.
    Before the rotation work the inset was 27.29 pt, so the error more than halved.
    """
    from tests.conftest import FIXTURE_DIR

    deck = convert_pptx_to_model(
        (FIXTURE_DIR / "real-financial-report.pptx").read_bytes()
    )
    charts = [
        element
        for element in deck.slides[2].elements
        if isinstance(element, m.ChartElement) and element.chart.kind == "barChart"
    ]
    assert len(charts) == 1
    children = charts[0].children
    turned = _turned(children)
    assert len(turned) == 4
    assert {shape.transform.rotation for shape in turned} == {-45.0}
    inset = _bottom_inset(children, frame_height=135.0)
    assert inset == pytest.approx(89.3, abs=0.5)
    # Still wrong, but nearer than the 27.29 pt it reserved when it drew them level.
    assert abs(inset - 69.538) < abs(27.29 - 69.538)


def test_labels_that_float_inside_the_plot_do_not_turn():
    """Negative values lift the category axis off the plot's floor.

    `_plot_rect` then reserves no band under the plot at all -- measured, on the negative
    probe -- and prints the labels beside the zero line instead.  Turning them there would
    draw into space nothing set aside, and no probe measured that corner, so the two stay
    in step: the labels turn only where a band was reserved for them.
    """
    cats = tuple(f"CategoryLongerStill{n}" for n in "ABCDE")
    # The same labels over positive data do turn, so this is about the axis floating.
    assert len(_turned(_rotation_probe(cats))) == 5
    children = _rotation_probe(cats, values=(3, -4, 5, -2, 1))
    assert _turned(children) == []
    # And they are still all drawn, level, rather than dropped.
    labels = {
        "".join(run.text for p in child.text_body.paragraphs for run in p.runs)
        for child in children
        if isinstance(child, m.ShapeElement) and child.text_body is not None
    }
    assert set(cats) <= labels


# -- Wrapped category labels -------------------------------------------------------------
#
# Three probe decks of 57 bar charts in the same 220.4724 x 181.1024 pt frame as the
# rotation sweep, five categories each, exported by PowerPoint 16.106 and read back out of
# the PDF as exact vector coordinates.  The category labels are pinned to Arial except
# where a row names another face, because the question these ask is what PowerPoint does
# with a *face's* line box and Aptos alone cannot answer that.
#
# What they establish, in order of how much they change:
#
# * **Wrapping comes before rotation.**  `MMM MM` is 44.43 pt on a 37.761 pt band and came
#   back level on two lines; `MMMMM` is 41.65 pt on the same band and turned.  The rule is
#   the widest *unbreakable token*, not the widest label.
# * **A label breaks at a space and nowhere else.**  Hyphen, slash, comma, underscore, en
#   dash and CJK are all not break opportunities; U+00A0 is.
# * **The band grows by one line box per extra line**, measured on a four-rung ladder in
#   five faces.
# * **One label that must turn turns all of them**, wrappable neighbours included.

#: The band every chart in these decks has, read off the axis rule: 188.807 / 5.
WRAP_BAND_PT = 37.761


def wrap_chart_xml(cats, face="Arial", size=10.0, values=None):
    """The rotation sweep's chart with its label face and size pinned."""
    text = (
        f"<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz='{int(size * 100)}'>"
        f"<a:latin typeface='{face}'/><a:cs typeface='{face}'/></a:defRPr></a:pPr>"
        "</a:p></c:txPr>"
    )
    value_text = text.replace(f"sz='{int(size * 100)}'", "sz='1000'")
    body = rotation_chart_xml(cats, values=values)
    body = body.replace("<c:lblOffset val='100'/>", f"{text}<c:lblOffset val='100'/>")
    return body.replace(
        "<c:crossBetween val='between'/>", f"{value_text}<c:crossBetween val='between'/>"
    )


def _wrap_probe(cats, **kwargs):
    return _build(wrap_chart_xml(cats, **kwargs), width=220.4724, height=181.1024)[0]


def _label_rows(children):
    """Every level category label's text, grouped by baseline, top row first.

    Taken as "below the plot", which is what separates a category label from a value one
    without having to know either font.
    """
    floor = _pt(_plot_bottom(children))
    rows: dict[float, list[str]] = {}
    for child in children:
        if not isinstance(child, m.ShapeElement) or child.text_body is None:
            continue
        if child.transform.rotation or _pt(child.transform.offset_y) < floor:
            continue
        text = "".join(run.text for p in child.text_body.paragraphs for run in p.runs)
        rows.setdefault(round(_pt(child.transform.offset_y), 3), []).append(text)
    return [rows[key] for key in sorted(rows)]


#: name -> (categories, the widest unbreakable token in pt, does PowerPoint turn them).
#: Every row is one probe chart.  The two that bracket the threshold are `token-under` and
#: `token-over`: 37.22 pt wrapped and 38.33 pt turned on a 37.761 pt band, which puts the
#: ratio inside (0.9857, 1.0151] -- and one band width is the only value that is also
#: inside the unbroken sweep's (0.972, 1.024].
WRAP_DECISION = {
    # Wrapping wins whenever breaking the label saves it.
    "space-break": (("MMM MM",) * 5, 24.99, False),
    "three-lines": (("MM MM MM MM MM MM",) * 5, 16.66, False),
    "just-over": (("MMMM m",) * 5, 33.32, False),
    "just-under": (("MMM m",) * 5, 24.99, False),
    "one-wraps": (("Aa", "Bb", "MMM MM", "Cc", "Dd"), 24.99, False),
    "widest-wraps": (("MMM MM",) * 4 + ("MMMM MMMM",), 33.32, False),
    "token-under": (("xxxxxxxi M",) * 5, 37.22, False),
    # A no-break space breaks.  It is the one character here that surprises.
    "nbsp": (("MMM\u00a0MM",) * 5, 24.99, False),
    # Rotation is what is left when breaking does not help.
    "single-token": (("MMMMM",) * 5, 41.65, True),
    "token-over": (("HHHHHi M",) * 5, 38.33, True),
    "all-tokens-wide": (("MMMMM MMMMM",) * 5, 41.65, True),
    "one-rotates": (("Aa", "Bb", "MMMMM", "Cc", "Dd"), 41.65, True),
    # The case that separates "widest token" from every other candidate rule: two spaces
    # in it, and PowerPoint turned it anyway because the middle token is 99.96 pt.
    "wide-token-and-space": (("Fiscal MMMMMMMMMMMM 2012",) * 5, 99.96, True),
    # Not break opportunities.
    "hyphen": (("MMM-MM",) * 5, 44.98, True),
    "slash": (("MMM/MM",) * 5, 44.43, True),
    "comma": (("MMM,MM",) * 5, 44.43, True),
    "underscore": (("MMM_MM",) * 5, 47.21, True),
    "en-dash": (("MMM–MM",) * 5, 47.21, True),
    "cjk": (("プラットフォーム",) * 5, 80.00, True),
}


@pytest.mark.parametrize("name", list(WRAP_DECISION))
def test_a_label_is_broken_before_it_is_turned(name):
    """Twenty probes: PowerPoint wraps where it can and turns only where it cannot."""
    cats, token, turns = WRAP_DECISION[name]
    children = _wrap_probe(cats)
    assert bool(_turned(children)) is turns, (
        f"{name}: widest token {token:.2f} pt on a {WRAP_BAND_PT:.2f} pt band, "
        f"{token / WRAP_BAND_PT:.4f} of it"
    )


def test_one_label_that_must_turn_turns_the_ones_that_could_have_wrapped():
    """Rotation is a decision for the axis, not for each label.

    Four ``MMM MM`` -- which wrap on their own -- beside one ``MMMMM``, and PowerPoint
    turned all five, none of them broken.
    """
    children = _wrap_probe(("MMM MM", "MMM MM", "MMMMM", "MMM MM", "MMM MM"))
    turned = _turned(children)
    assert len(turned) == 5
    assert {shape.transform.rotation for shape in turned} == {-45.0}
    assert _label_rows(children) == []


#: name -> (face, categories, lines PowerPoint used, the band it reserved in pt).  One
#: token per line by construction, so the token count is the line count.  Read off the
#: axis rule against the frame's foot.
WRAP_LADDER = {
    "arial-1": ("Arial", ("MMM",) * 5, 1, 23.712),
    "arial-2": ("Arial", ("MMM MMM",) * 5, 2, 35.212),
    "arial-3": ("Arial", ("MMM MMM MMM",) * 5, 3, 46.712),
    "arial-4": ("Arial", ("MMM MMM MMM MMM",) * 5, 4, 58.212),
    "calibri-1": ("Calibri", ("MMM",) * 5, 1, 25.052),
    "calibri-2": ("Calibri", ("MMM MMM",) * 5, 2, 37.257),
    "calibri-3": ("Calibri", ("MMM MMM MMM",) * 5, 3, 49.462),
    "calibri-4": ("Calibri", ("MMM MMM MMM MMM",) * 5, 4, 61.667),
    "aptos-1": ("Aptos", ("MMM",) * 5, 1, 24.965),
    "aptos-2": ("Aptos", ("MMM MMM",) * 5, 2, 37.170),
    "aptos-3": ("Aptos", ("MMM MMM MMM",) * 5, 3, 49.375),
    "aptos-4": ("Aptos", ("MMM MMM MMM MMM",) * 5, 4, 61.580),
    "times-1": ("Times New Roman", ("MMM",) * 5, 1, 23.515),
    "times-2": ("Times New Roman", ("MMM MMM",) * 5, 2, 34.590),
    "times-3": ("Times New Roman", ("MMM MMM MMM",) * 5, 3, 45.665),
    "times-4": ("Times New Roman", ("MMM MMM MMM MMM",) * 5, 4, 56.740),
    "courier-1": ("Courier New", ("MMMMM",) * 5, 1, 23.380),
    "courier-2": ("Courier New", ("MMMMM MMMMM",) * 5, 2, 34.710),
    "courier-3": ("Courier New", ("MMMMM MMMMM MMMMM",) * 5, 3, 46.040),
    "courier-4": ("Courier New", ("MMMMM MMMMM MMMMM MMMMM",) * 5, 4, 57.370),
}

#: Every face here but Arial lands inside this at every rung.  What it allows is the level
#: band's own pre-existing bias, which runs from 0.11 pt light on Arial to 0.60 pt heavy on
#: Courier New and was fitted long before this sweep; what the test is really asserting is
#: that the *slope* is right, which is why the residual is flat across all four rungs
#: instead of fanning out.
WRAP_BAND_TOLERANCE_PT = 0.65

#: How far short of PowerPoint our band falls for each line after the first, in Arial only.
#: This is Arial's ``hhea`` lineGap -- 67 units of 2048, 0.0327 em -- which PowerPoint adds
#: to the line box and :mod:`pptx2svg.text.metrics` does not carry.  See
#: ``_bottom_label_band`` for why adding it is not the improvement it looks like.
ARIAL_LINE_GAP_PT_PER_LINE = 0.328


@pytest.mark.parametrize("name", list(WRAP_LADDER))
def test_the_wrapped_band_grows_by_one_line_box_a_line(name):
    """Four rungs in five faces, and the slope is the face's own line box.

    PowerPoint's band grew by exactly the same amount from one line to two, two to three
    and three to four in every face, so the residual is flat rather than fanning -- which
    is what says the per-line term is right and the constant is the level band's older fit.
    """
    face, cats, lines, expected = WRAP_LADDER[name]
    children = _wrap_probe(cats, face=face)
    assert not _turned(children), f"{name} turned; it should have wrapped"
    assert len(_label_rows(children)) == lines
    allowance = WRAP_BAND_TOLERANCE_PT + (
        ARIAL_LINE_GAP_PT_PER_LINE * (lines - 1) if face == "Arial" else 0.0
    )
    assert _bottom_inset(children) == pytest.approx(expected, abs=allowance)


def test_arial_is_short_by_its_line_gap_and_nothing_else():
    """The one face the line box alone does not fit, pinned so it cannot drift.

    Arial is the only one of the five probe faces whose ``hhea`` lineGap is not zero, and
    PowerPoint adds it: 67 units of 2048 is 0.328 pt at 10 pt, which is exactly how far our
    band falls short per extra line.  Recorded rather than corrected -- the substitute we
    would read a lineGap from carries 87 where Office's own ``times.ttf`` carries 0, so
    adding it would trade this error for a bigger one on Times New Roman.
    """
    for name in ("arial-1", "arial-2", "arial-3", "arial-4"):
        _, cats, lines, expected = WRAP_LADDER[name]
        residual = _bottom_inset(_wrap_probe(cats)) - expected
        predicted = 0.110 - ARIAL_LINE_GAP_PT_PER_LINE * (lines - 1)
        assert residual == pytest.approx(predicted, abs=0.05), name


def test_the_label_needing_the_most_lines_sets_the_band():
    """One three-line label among four that fit lifts the whole band to three lines.

    Measured: the band came out 46.712 pt, the same as a chart where every label needed
    three lines, and the four short ones sat on the *first* row with nothing under them.
    """
    cats = ("MM MM", "MM MM", "MM MM MM MM MM MM", "MM MM", "MM MM")
    children = _wrap_probe(cats)
    assert _bottom_inset(children) == pytest.approx(
        46.712, abs=WRAP_BAND_TOLERANCE_PT + 2 * ARIAL_LINE_GAP_PT_PER_LINE
    )
    rows = _label_rows(children)
    assert [len(row) for row in rows] == [5, 1, 1]


def test_a_wrapped_block_hangs_from_the_top_like_a_one_line_label():
    """Adding lines does not move the first baseline; the lines are added below it.

    The Arial ladder put the first baseline 15.03, 15.05 and 15.07 pt under the axis at
    one, two and three lines -- one number inside PowerPoint's 0.12 pt output grid.
    """
    tops = [
        _first_label_top(_wrap_probe(cats))
        for cats in (("MMM",) * 5, ("MMM MMM",) * 5, ("MMM MMM MMM",) * 5)
    ]
    assert tops[1] == pytest.approx(tops[0], abs=0.01)
    assert tops[2] == pytest.approx(tops[0], abs=0.01)


def _first_label_top(children):
    """Where the first row of the category-label block sits, relative to the plot."""
    floor = _pt(_plot_bottom(children))
    return min(
        _pt(child.transform.offset_y) - floor
        for child in children
        if isinstance(child, m.ShapeElement)
        and child.text_body is not None
        and _pt(child.transform.offset_y) >= floor
    )


def test_the_band_stops_growing_where_the_measurements_stop():
    """Six lines is the last rung PowerPoint laid out; past it, it stops wrapping.

    An eight-token label came back on **two** lines, each four band widths wide and
    overlapping its neighbours, in a 35.212 pt band -- and a twelve- and a twenty-four-token
    label did exactly the same.  No rule reproduces the linear part and that collapse, so
    the band stops where the ladder stops, and this test records how far the disagreement
    goes rather than leaving it latent.
    """
    six = _bottom_inset(_wrap_probe((" ".join(["MMM"] * 6),) * 5))
    assert six == pytest.approx(
        81.212, abs=WRAP_BAND_TOLERANCE_PT + 5 * ARIAL_LINE_GAP_PT_PER_LINE
    )
    for tokens in (8, 12, 24):
        children = _wrap_probe((" ".join(["MMM"] * tokens),) * 5)
        assert len(_label_rows(children)) == 6
        # PowerPoint reserved 35.212 pt for all three of these.  We are 44 pt over, and
        # that is the recorded price of not extrapolating a rule it refutes.
        assert _bottom_inset(children) == pytest.approx(six, abs=0.01)


def test_a_label_breaks_only_where_powerpoint_breaks_it():
    """The break set, straight off the probes, without going through the band."""
    from pptx2svg.resolve.chart import ChartFont, font_box, wrap_label

    font = ChartFont(family="Arial", box=font_box("Arial", 10.0))
    assert wrap_label("MMM MM", font, WRAP_BAND_PT) == ["MMM", "MM"]
    # The no-break space breaks, which is the one character here that surprises.
    assert wrap_label("MMM\u00a0MM", font, WRAP_BAND_PT) == ["MMM", "MM"]
    for unbroken in ("MMM-MM", "MMM/MM", "MMM,MM", "MMM_MM", "MMM\u2013MM"):
        assert wrap_label(unbroken, font, WRAP_BAND_PT) == [unbroken]
    # Greedy, and a token too wide to sit alone still gets its own line rather than
    # being split -- the state `_labels_rotate` has already turned the axis for.
    assert wrap_label("MM MM MM", font, WRAP_BAND_PT) == ["MM MM", "MM"]
    assert wrap_label("Fiscal MMMMMMMMMMMM 2012", font, WRAP_BAND_PT) == [
        "Fiscal",
        "MMMMMMMMMMMM",
        "2012",
    ]


# -- A line chart's legend key -----------------------------------------------------------


def test_a_line_chart_legends_with_a_rule_and_its_marker():
    """Confirmed on four line-chart legends after the radar measured it once.

    19.200 pt of rule with the series' marker on its midpoint, then 2.025 pt before the
    text -- the same on a right-hand legend and a bottom one, the same for a series whose
    marker is ``none``, and **the same at 14 pt**, which is what says these are points and
    not the ems the radar's single 10 pt measurement was carried as.
    """
    from pptx2svg.resolve.chart import (
        LINE_LEGEND_KEY_GAP_PT,
        LINE_LEGEND_KEY_PT,
        ChartBuilder,
        ChartFont,
        font_box,
    )

    for size in (10.0, 14.0):
        font = ChartFont(family="Arial", box=font_box("Arial", size))
        key, gap = ChartBuilder._legend_key_size(_FakeLine(), font)
        assert (key, gap) == (LINE_LEGEND_KEY_PT, LINE_LEGEND_KEY_GAP_PT)
        assert (key, gap) == (19.200, 2.025)


class _FakeLine:
    """Just enough of a builder for `_legend_key_size` to answer "a line chart"."""

    _is_line = True
    _is_radar = False
    _radar_style = "marker"



def test_a_chart_measures_and_draws_its_japanese_in_the_themes_script_face():
    """The defect this file used to attribute to the label width, and where it lived.

    `real-financial-report.pptx`'s charts name `<a:latin typeface="Arial"/>` and nothing
    else; its theme writes `<a:ea typeface=""/>` in both collections and
    `<a:font script="Jpan" typeface="游ゴシック"/>` beside it.  PowerPoint's own export
    embeds **YuGothic-Regular** for every Japanese category label and **ArialMT** for the
    Latin runs inside the same labels, so the script list beats the Latin face and the two
    split within one label.

    `ChartFont` used to carry one face, which left a label measured through Arial -- which
    has no Japanese glyph at all -- and drawn in whatever the rasteriser happened to find.
    Both halves are asserted here: the resolved run names the East Asian face, and the
    widths are the ones PowerPoint drew, 48.000, 96.000, 60.000 and 36.000 pt.
    """
    from tests.conftest import FIXTURE_DIR

    deck = convert_pptx_to_model(
        (FIXTURE_DIR / "real-financial-report.pptx").read_bytes()
    )
    charts = [
        element
        for element in deck.slides[2].elements
        if isinstance(element, m.ChartElement) and element.chart.kind == "barChart"
    ]
    assert len(charts) == 1
    labels = {
        run.text: run.properties
        for shape in _turned(charts[0].children)
        for paragraph in shape.text_body.paragraphs
        for run in paragraph.runs
    }
    assert set(labels) == {"デジタル", "プラットフォーム", "グローバル", "その他"}
    for properties in labels.values():
        assert properties.font_family == "Arial"
        assert properties.font_family_ea == "游ゴシック"

    from pptx2svg.resolve.chart import ChartFont, font_box

    font = ChartFont(
        family="Arial",
        box=font_box("Arial", 12.0),
        family_ea="游ゴシック",
        box_ea=font_box("游ゴシック", 12.0),
    )
    for text, drawn in (
        ("デジタル", 48.0),
        ("プラットフォーム", 96.0),
        ("グローバル", 60.0),
        ("その他", 36.0),
    ):
        assert font.width(text) == pytest.approx(drawn, abs=0.001), text
    # The radar's two mixed labels, from the same export: Arial for the Latin half.
    assert font.width("DX投資額") == pytest.approx(52.669, abs=0.01)
    assert font.width("CO2削減") == pytest.approx(48.674, abs=0.01)
    # A Latin-only label is untouched by any of it.
    assert font.width("Q1") == pytest.approx(16.008, abs=0.001)
    # And the line box the reserves read is the face that draws the label, per label.
    assert font.box_for("Q1").line_height == pytest.approx(13.406, abs=0.001)
    assert font.box_for("その他").line_height == pytest.approx(17.376, abs=0.001)


def test_a_chart_label_that_mixes_scripts_is_drawn_in_two_chunks():
    """resvg falls back per text chunk, not per glyph, so the split has to be in the SVG.

    One `<tspan>` naming both faces would be drawn entirely in whichever one the
    rasteriser settles on, at a width nothing computed.  `ChartFont.width` measures
    `DX投資額` as Arial plus 游ゴシック, so the emitted markup has to say the same thing.
    """
    from pptx2svg import convert_pptx_to_svg
    from tests.conftest import FIXTURE_DIR

    svg = convert_pptx_to_svg(
        str(FIXTURE_DIR / "real-financial-report.pptx"),
        ConvertOptions(warn_on_font_substitution=False),
    )[3]
    chunks = re.findall(
        r'<tspan[^>]*font-family="([^"]*)"[^>]*>(DX|投資額)</tspan>', svg
    )
    assert [text for _, text in chunks] == ["DX", "投資額"]
    latin, east_asian = (family for family, _ in chunks)
    assert latin.startswith("Arial")
    assert east_asian.startswith("游ゴシック")
