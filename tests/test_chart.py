"""Charts: reading ``c:chartSpace``, and drawing what comes out of it.

Every layout number asserted here was read out of PowerPoint's own PDF export as an exact
vector coordinate, not eyeballed from a raster.  See ``src/pptx2svg/resolve/chart.py`` for
where each constant came from and what its residual against the measurement is.
"""

from __future__ import annotations

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
        f"<c:chartSpace {C} {A} {R}><c:chart><c:plotArea><c:radarChart>"
        "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:radarChart></c:plotArea></c:chart></c:chartSpace>"
    ).encode()
    frame = (
        "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='97' name='Radar'/>"
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
    assert "radarChart" in warning.message


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
        overrides[f"/{part}"] = chart_type
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


def test_a_negative_bar_is_drawn_hollow_unless_the_file_opts_out(variant_deck):
    """Measured: white fill, black 0.75 pt outline -- and the series colour at val="0"."""
    default = _bars(variant_deck["negative-default"])
    assert [bar.fill.color.hex.upper() for bar in default] == [
        "#F97316",
        "#FFFFFF",
        "#F97316",
    ]
    inverted = next(bar for bar in default if bar.fill.color.hex.upper() == "#FFFFFF")
    assert inverted.outline is not None
    assert inverted.outline.fill.color.hex.upper() == "#000000"
    assert inverted.outline.width == 9525

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


def test_vary_colors_does_not_suppress_the_negative_bar_inversion():
    """A varyColors fill is not a `c:dPt`, so `c:invertIfNegative` still applies to it."""
    children, _ = _build(
        bar(
            "<c:varyColors val='1'/><c:ser><c:val><c:numRef><c:numCache>"
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
    assert len(_markers(chart.children)) == 9


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
