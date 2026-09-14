"""Charts: reading ``c:chartSpace``, and drawing what comes out of it.

Every layout number asserted here was read out of PowerPoint's own PDF export as an exact
vector coordinate, not eyeballed from a raster.  See ``src/pptx2svg/resolve/chart.py`` for
where each constant came from and what its residual against the measurement is.
"""

from __future__ import annotations

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
        (-1234.0, "#,##0;(#,##0)", "(1,234)"),
        (1.5, '"$"0.00', "1.50"),
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
    """The plot rectangle, read back off the two axis lines, in points."""
    lines = [c for c in chart.children if isinstance(c, m.ConnectorElement)]
    vertical = [line for line in lines if line.transform.extent_width == 0]
    horizontal = [line for line in lines if line.transform.extent_height == 0]
    assert vertical and horizontal
    value = vertical[0].transform
    # The category axis is the lowest horizontal line; the rest are gridlines.
    category = max(horizontal, key=lambda line: line.transform.offset_y).transform
    return (
        _pt(value.offset_x),
        _pt(value.offset_y),
        _pt(category.offset_x + category.extent_width),
        _pt(category.offset_y),
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
    bars = [
        child
        for child in chart.children
        if isinstance(child, m.ShapeElement)
        and isinstance(child.fill, m.SolidFill)
        and child.fill.color.hex.upper() == "#F97316"
        and child.text_body is None
        # The legend swatch is the same colour and an order of magnitude smaller.
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
    """Only barChart is drawn so far; the rest must say so, not draw a wrong picture."""
    from tests.deckbuilder import derive_deck

    chart_xml = (
        "<?xml version='1.0'?>"
        f"<c:chartSpace {C} {A} {R}><c:chart><c:plotArea><c:pieChart>"
        "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:pieChart></c:plotArea></c:chart></c:chartSpace>"
    ).encode()
    frame = (
        "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='97' name='Pie'/>"
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
        overrides={"/ppt/charts/chartPie.xml": chart_type},
    )
    options = ConvertOptions()
    convert_pptx_to_model(deck_bytes, options)
    warning = next(w for w in options.warnings if w.code == "chart-unsupported-type")
    assert "pieChart" in warning.message


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
        overrides[f"/{part}"] = chart_type
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
    ours = {
        "left": left,
        "right": width - right,
        "top": top,
        "bottom": height - bottom,
    }
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
