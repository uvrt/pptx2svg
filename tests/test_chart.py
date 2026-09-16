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
from pptx2svg.resolve.chart import (
    _decade,
    _nice_unit,
    format_number,
    nice_axis_scale,
)

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
    "values,intervals,expected",
    [
        # authoring-integration.pptx: PowerPoint draws 0..6 by 1 for data topping out at
        # 5, on a 228.3 pt frame with 10 pt labels -- ten intervals of room, so the base
        # rule answers on its own.
        ([3, 4, 5], 10, (0.0, 6.0, 1.0)),
        # real-financial-report chart1, maxing at 4285: PowerPoint drew 0..5000 by 1000 on
        # a 142.5 pt frame at 12 pt, which is six intervals of room.
        ([3980, 4120, 4285, 465, 488, 512], 6, (0.0, 5000.0, 1000.0)),
        # ...and chart3, maxing at 1842: 0..2000 by 500 on a 135 pt frame, also six.
        ([1599, 1185, 663, 334, 1842, 1285, 814, 344], 6, (0.0, 2000.0, 500.0)),
        # The same 1842 on a frame with room for ten draws 0..2000 by **200**, which is
        # what says the by-500 axis above is a coarsened one rather than the base rule.
        # Measured: `tools/make_axis_probe.py` deck `axis-bar`, probe `bar1842t`.
        ([1599, 1185, 663, 334, 1842, 1285, 814, 344], 10, (0.0, 2000.0, 200.0)),
    ],
)
def test_axis_scale_matches_what_powerpoint_drew(values, intervals, expected):
    """All four come from reading gridline coordinates out of PowerPoint's PDF.

    The unit is the finest 1-2-5 step that divides the padded range into *intervals*, and
    the count is what the frame has room for; see :func:`side_axis_intervals`.
    """
    assert nice_axis_scale(min(values), max(values), intervals=intervals) == expected


def test_the_side_axis_rung_is_the_faces_full_hhea_pitch():
    """Arial is the face that separates the pitch from the line box, and it does it twice.

    ``axis-wider``'s Arial cell walks a 10 pt axis over frames of 160, 165, 169, 172 and
    175 pt; PowerPoint steps from nine intervals to ten somewhere in (160, 165].  The rung
    decides where we step, and the two candidates put it in different places:

        line box  22 + 12 * 11.1719 = 156.06    outside PowerPoint's window
        pitch     22 + 12 * 11.4990 = 159.99    inside it, by 0.01 pt

    So the line box is not loose here, it is refuted.  What the pitch leaves is the one
    reading the sweep still misses -- 160 pt itself, by 0.012 pt of rung -- and
    :data:`~pptx2svg.resolve.chart.AXIS_EDGE_RESERVE_PT` carries why that is not patched.
    """
    from pptx2svg.resolve.chart import font_box, side_axis_intervals

    box = font_box("Arial", 10.0)
    assert box.line_height == pytest.approx(11.1719, abs=5e-4)
    assert box.pitch == pytest.approx(11.4990, abs=5e-4)
    assert box.gap == pytest.approx(0.3271, abs=5e-4)

    # Where each rung puts the nine-to-ten step, walked a point at a time.
    def step(rung):
        return next(h for h in range(140, 200) if side_axis_intervals(h, rung) == 10)

    assert step(box.line_height) == 157
    assert step(box.pitch) == 160

    # A zero-gap face is untouched: every number measured with Aptos still stands.
    aptos = font_box("Aptos", 10.0)
    assert aptos.gap == 0.0
    assert aptos.pitch == aptos.line_height


def test_the_maximum_clears_the_data_by_five_per_cent_not_by_a_whole_unit():
    """A series topping out at the axis maximum would touch the frame, and PowerPoint
    clears it -- but by padding the range 5% before rounding, not by adding a unit.

    The two are the same answer when the data lands on a unit boundary and different when
    it does not, and **4.9 is where they part**: a whole-unit bump leaves 0..5, the pad
    carries 5.145 past 5 and the axis goes to 6.  PowerPoint draws 0..6, measured on the
    `axis-pad` deck, and the same deck brackets the pad itself -- 0.3..4.8 takes unit 1 and
    0.3..4.76 takes 0.5, which is (4.17%, 5.04%] of the maximum.
    """
    assert nice_axis_scale(0.0, 5.0)[1] == 6.0
    assert nice_axis_scale(0.0, 4.9)[1] == 6.0
    assert nice_axis_scale(0.0, 4.7)[1] == 5.0


def test_a_negative_minimum_extends_the_axis_below_zero():
    minimum, maximum, unit = nice_axis_scale(-30.0, 120.0)
    assert minimum < 0 and maximum >= 120 and unit > 0
    assert minimum % unit == 0 and maximum % unit == 0


def test_a_flat_series_still_gets_a_usable_axis():
    minimum, maximum, unit = nice_axis_scale(0.0, 0.0)
    assert maximum > minimum and unit > 0


def _log10_off_by_one_ulp(direction: float):
    """``math.log10`` with every answer nudged one ulp, the way a faithful libm may."""
    correctly_rounded = math.log10

    def perturbed(value: float) -> float:
        return math.nextafter(correctly_rounded(value), direction)

    return perturbed


@pytest.mark.parametrize("direction", [-math.inf, math.inf])
def test_a_one_ulp_error_in_log10_does_not_move_the_axis(monkeypatch, direction):
    """The property a tripwire in `tests/test_vrt.py` used to stand in for.

    That one asserted this *platform* rounds `log10(100.0)` to 2.0.  This one asserts our
    own code does not care, which is the half that survives being run somewhere else:
    `log10` is perturbed by an ulp in both directions -- harsher than any real libm
    disagreement, since a faithful implementation is only allowed to be wrong one way at a
    time -- and every axis has to come out where it was.

    0..100 is `real-financial-report.pptx` slide 4's radar, the one place in the corpus
    where the ulp used to show: with four intervals of radius the unit is a hundred over
    four rounded up the ladder, which is 50, and it draws two rings.  One ulp low under the
    old expression and `floor` gave a decade of 10, the unit became 10, and the slide grew
    by five thousand characters of SVG.
    """
    monkeypatch.setattr(math, "log10", _log10_off_by_one_ulp(direction))

    assert _decade(100.0) == 100.0
    assert nice_axis_scale(0.0, 100.0, intervals=4, strict=False) == (0.0, 100.0, 50.0)
    # The 1-2-5 ladder reads a magnitude the same way, and failed the same way: an ulp low
    # made the magnitude of 100 be 10 and the rung above it 20, so an axis wanting a
    # hundred-unit step got a twenty.
    assert _nice_unit(100.0) == 100.0
    assert _nice_unit(100.1) == 200.0
    # And the measured axes are unmoved, ulp or no ulp.
    assert nice_axis_scale(0.0, 5.0) == (0.0, 6.0, 1.0)
    assert nice_axis_scale(465.0, 4285.0, intervals=6) == (0.0, 5000.0, 1000.0)
    assert nice_axis_scale(0.0, 0.07) == pytest.approx((0.0, 0.08, 0.01))


def test_every_decade_of_the_double_range_is_found_exactly():
    """The contract, stated on the helper rather than on a chart.

    Ten to the something, at or below the value, with the next one up above it: that is
    the whole of it, and it is what the axis rules are written against.  Swept over the
    double range rather than sampled near one, because the failure being guarded against
    is not a small one -- an exponent out by one is an axis out by a factor of ten, and
    `log10` is no more exact at 1e-300 than it is at 100.
    """
    for exponent in range(-300, 301):
        power = float(f"1e{exponent}")
        for value in (power, power * 1.000001, power * 5, power * 9.999):
            assert _decade(value) == power, (value, exponent)


def test_a_span_a_hair_under_a_power_of_ten_is_still_that_power_of_ten():
    """Spans are subtractions, and subtractions of decimals land just under round numbers.

    ``0.24 - 0.14`` is ``0.09999999999999998``.  Nothing about that chart is a hundredth;
    the axis wanted is the one for a span of 0.1, and :data:`_DECADE_SLACK` is what keeps
    it.  `log10`'s rounding used to supply a window like this by accident, but a narrower
    and less even one: it put ``0.24 - 0.14`` (two ulps under a tenth) in the right decade
    and ``1.13 - 1.03`` (ten ulps under the same tenth) in the one below, an axis stepping
    by 0.005 for data that spans a tenth.  Data like that is authored every day.

    The slack is not a licence to round.  A span that genuinely falls short of a power of
    ten still falls short of it: 9.9999999999 misses by a hundred times more than this
    window reaches, and nothing anyone would author comes close to being that near.

    Its **width** is not measured and cannot be on this oracle -- PowerPoint draws a span
    of 0.095 and a span one ulp under a tenth identically, so nothing observable moves when
    the window does.  See :data:`_DECADE_SLACK` and the test below.
    """
    assert _decade(0.24 - 0.14) == 0.1
    assert _decade(1.13 - 1.03) == 0.1
    assert _decade(1.13 - 0.13) == 1.0
    assert _decade(9.99) == 1.0
    assert _decade(9.9999999999) == 1.0
    # Two ulps apart, and `floor(log10(...))` put them in different decades: the first
    # rounds to exactly 2.0 and the second to 1.9999999999999998.  Both are a hundred, and
    # agreeing that they are is also what makes an ulp *up* in `log10` a no-op here --
    # every value an upward ulp could promote is inside this window already.
    assert _decade(math.nextafter(100.0, 0.0)) == 100.0
    assert _decade(math.nextafter(math.nextafter(100.0, 0.0), 0.0)) == 100.0


#: Every pair of two-decimal numbers a tenth apart whose ratio leaves a scatter's value
#: axis unanchored leaves one of exactly four residues in the subtraction: none, ten ulps,
#: twenty-six, ninety.  PowerPoint was given all four and drew all four **alike**, every
#: one of them stepping by 0.02, so whatever we do with a residue, the one thing it may
#: not do is depend on its size.
AUTHORED_TENTH_PAIRS = ((1.00, 1.10), (1.03, 1.13), (2.16, 2.26), (7.94, 8.04))


def test_the_residue_of_an_authored_subtraction_never_changes_the_axis():
    """The evenness :data:`_DECADE_SLACK` exists for, and PowerPoint agrees it is right.

    ``1.10 - 1.00`` lands a hair *above* a tenth, ``1.13 - 1.03`` ten ulps under it,
    ``2.26 - 2.16`` twenty-six and ``8.04 - 7.94`` ninety.  Nothing about those four charts
    differs to a reader, and PowerPoint draws all four with the same unit -- measured, one
    probe each, ``tools/make_axis_probe.py`` deck ``axis-decade``.  Under
    ``floor(log10(span))`` the last three fall into the decade of a hundredth and the first
    does not, so a chart of 1.00..1.10 got an axis by 0.05 and its neighbour 1.03..1.13 one
    by 0.01: five times the gridlines for a tenth of data either way.

    The unit we pick is PowerPoint's own 0.02 on all four, which it was not while the rule
    was the power of ten below the span; what a residue is allowed to do is nothing, and it
    does nothing either way.
    """
    for low, high in AUTHORED_TENTH_PAIRS:
        minimum, maximum, unit = nice_axis_scale(low, high, anchor_zero=False)
        assert unit == pytest.approx(0.02), (low, high)
        assert minimum <= low and maximum >= high


def test_powerpoint_ignores_the_decade_of_a_span_on_an_unanchored_axis():
    """A measurement kept as a test because it refutes the shape, not a constant.

    :data:`_DECADE_SLACK` decides which decade a span that falls just under a power of ten
    belongs to, and the only axis where that decision survives to be drawn is an
    unanchored one -- a scatter's.  Thirty such probes, with the span under a tenth by a
    relative 1e-15 through 5e-2, came back from PowerPoint with **one** axis between them:

    ========================  =====================  ======================
    data                      PowerPoint             ours
    ========================  =====================  ======================
    1.03..1.13 (10 ulps)      1.02..1.14 by 0.02     1.0..1.15 by 0.05
    1.03..1.1299999999        1.02..1.14 by 0.02     1.02..1.13 by 0.01
    1.03..1.125 (a flat .095) 1.02..1.14 by 0.02     1.02..1.13 by 0.01
    ========================  =====================  ======================

    A five-per-cent shortfall and a one-ulp shortfall drawn identically is not a wide
    forgiveness window; it is a rule that never asks which decade the span is in.  So the
    slack's width is unmeasurable here and no number is fitted to these -- and since the
    unit is a count of the padded range rather than a decade, all three now come out where
    PowerPoint drew them.
    """
    assert nice_axis_scale(1.03, 1.13, anchor_zero=False) == pytest.approx((1.02, 1.14, 0.02))
    assert nice_axis_scale(1.03, 1.1299999999, anchor_zero=False) == pytest.approx(
        (1.02, 1.14, 0.02)
    )
    assert nice_axis_scale(1.03, 1.125, anchor_zero=False) == pytest.approx((1.02, 1.14, 0.02))
    # The extent comes out of the same padded range: 0.3..4.9 is already clear of a 0..5
    # axis and PowerPoint draws 0..6, which is the pad carrying 4.9 to 5.145.
    assert nice_axis_scale(0.3, 4.9) == pytest.approx((0.0, 6.0, 1.0))


def test_the_decade_helper_survives_the_ends_of_the_double_range():
    """Where there is no power of ten to return, and what the callers do about it.

    A denormal has none below it, so the helper says 0.0 and `nice_axis_scale` falls back
    to its 0..1 axis -- the same answer, by the same guard, as before the helper existed.
    Nothing finite reaches 1e309, so nothing comes back infinite.
    """
    assert _decade(5e-324) == 0.0
    assert nice_axis_scale(0.0, 5e-324) == (0.0, 1.0, 1.0)
    assert _decade(1e-320) == 1e-320
    assert _decade(1.7976931348623157e308) == 1e308


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
        f"<c:chartSpace {C} {A} {R}><c:chart><c:plotArea><c:surfaceChart>"
        "<c:ser><c:val><c:numRef><c:numCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>1</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
        "</c:surfaceChart></c:plotArea></c:chart></c:chartSpace>"
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
    assert "surfaceChart" in warning.message


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
        (
            "area with no series",
            "<c:chart><c:plotArea><c:areaChart/></c:plotArea></c:chart>",
        ),
        (
            "area whose only value is blank",
            "<c:chart><c:plotArea><c:areaChart><c:ser><c:val><c:numRef><c:numCache>"
            "<c:ptCount val='3'/></c:numCache></c:numRef></c:val></c:ser>"
            "</c:areaChart></c:plotArea></c:chart>",
        ),
        (
            "scatter with no series",
            "<c:chart><c:plotArea><c:scatterChart/></c:plotArea></c:chart>",
        ),
        (
            "scatter with y values and no x",
            "<c:chart><c:plotArea><c:scatterChart><c:ser><c:yVal><c:numRef><c:numCache>"
            "<c:ptCount val='3'/><c:pt idx='0'><c:v>1</c:v></c:pt>"
            "<c:pt idx='2'><c:v>4</c:v></c:pt>"
            "</c:numCache></c:numRef></c:yVal></c:ser></c:scatterChart>"
            "</c:plotArea></c:chart>",
        ),
        (
            "scatter whose x list is shorter than its y list",
            "<c:chart><c:plotArea><c:scatterChart><c:ser>"
            "<c:xVal><c:numRef><c:numCache><c:ptCount val='1'/>"
            "<c:pt idx='0'><c:v>7</c:v></c:pt></c:numCache></c:numRef></c:xVal>"
            "<c:yVal><c:numRef><c:numCache><c:ptCount val='4'/>"
            "<c:pt idx='0'><c:v>1</c:v></c:pt><c:pt idx='3'><c:v>4</c:v></c:pt>"
            "</c:numCache></c:numRef></c:yVal></c:ser></c:scatterChart>"
            "</c:plotArea></c:chart>",
        ),
        (
            "scatter whose x values are all the same",
            "<c:chart><c:plotArea><c:scatterChart><c:ser>"
            "<c:xVal><c:numRef><c:numCache><c:ptCount val='2'/>"
            "<c:pt idx='0'><c:v>2</c:v></c:pt><c:pt idx='1'><c:v>2</c:v></c:pt>"
            "</c:numCache></c:numRef></c:xVal>"
            "<c:yVal><c:numRef><c:numCache><c:ptCount val='2'/>"
            "<c:pt idx='0'><c:v>1</c:v></c:pt><c:pt idx='1'><c:v>4</c:v></c:pt>"
            "</c:numCache></c:numRef></c:yVal></c:ser></c:scatterChart>"
            "</c:plotArea></c:chart>",
        ),
        (
            "scatter whose x values are NaN and infinite",
            "<c:chart><c:plotArea><c:scatterChart><c:ser>"
            "<c:xVal><c:numRef><c:numCache><c:ptCount val='2'/>"
            "<c:pt idx='0'><c:v>NaN</c:v></c:pt><c:pt idx='1'><c:v>1e400</c:v></c:pt>"
            "</c:numCache></c:numRef></c:xVal>"
            "<c:yVal><c:numRef><c:numCache><c:ptCount val='2'/>"
            "<c:pt idx='0'><c:v>1</c:v></c:pt><c:pt idx='1'><c:v>4</c:v></c:pt>"
            "</c:numCache></c:numRef></c:yVal></c:ser></c:scatterChart>"
            "</c:plotArea></c:chart>",
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
    # `<c:smooth val="0"/>`, because **an absent element means smooth** -- see
    # `test_an_absent_smooth_element_smooths`.  Every test below that asserts a straight
    # line wants the element, and PowerPoint's own writer emits it on every series.
    smooth=False,
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
    """`c:smooth` is drawn as a spline; the probe's control points are not collinear."""
    straight, _ = _build(line_chart_xml(), width=220.0, height=181.0)
    assert "C " not in _polylines(straight)[0].geometry.paths[0].commands

    curved, _ = _build(
        line_chart_xml(values=(3, 5, 2), smooth=True), width=220.0, height=181.0
    )
    assert "C " in _polylines(curved)[0].geometry.paths[0].commands


def test_an_absent_smooth_element_smooths():
    """A chart boolean: the element missing is **not** the same as `val="0"`.

    Three probes at one frame size: a line chart with `<c:smooth val="1"/>`, one with no
    `c:smooth` at all and one with `val="0"` came back from PowerPoint as four cubics, the
    *same* four cubics, and a four-segment polyline.  This reader used `bool(None)` and so
    drew the middle case straight.
    """
    absent, _ = _build(
        line_chart_xml(values=(3, 5, 2), smooth=None), width=220.0, height=181.0
    )
    stated, _ = _build(
        line_chart_xml(values=(3, 5, 2), smooth=True), width=220.0, height=181.0
    )
    assert (
        _polylines(absent)[0].geometry.paths[0].commands
        == _polylines(stated)[0].geometry.paths[0].commands
    )


def test_the_spline_tension_is_powerpoints():
    """The control points, not just the shape: PowerPoint's Catmull-Rom is the plain 1/6.

    A five-point series over 3, 4, 5, 2, 6 exports as four cubics.  PowerPoint's thirteen
    ordinates, read out of the PDF and put back through its own 0..7 axis, are the exact
    thirds and sixths below -- which pin both halves of the rule: the interior controls are
    a sixth of the *neighbours'* chord from their vertex, and the two terminal ones are a
    **third** of their own chord, which duplicating the end point gets wrong by a factor of
    two.  Read in value space so the assertion does not depend on the frame.
    """
    children, _ = _build(
        line_chart_xml(values=(3, 4, 5, 2, 6), smooth=True),
        width=220.4724,
        height=181.1024,
    )
    commands = _polylines(children)[0].geometry.paths[0].commands
    numbers = [float(value) for value in commands.replace("M", "").replace("C", "").split()]
    ys = numbers[1::2]
    # The first and last vertices are the data's own 3 and 6, which calibrates the axis.
    values = [3.0 + (ys[0] - y) * 3.0 / (ys[0] - ys[-1]) for y in ys]
    assert values == pytest.approx(
        [
            3.0, 3 + 1 / 3, 3 + 2 / 3,
            4.0, 4 + 1 / 3, 5 + 1 / 3,
            5.0, 4 + 2 / 3, 1 + 5 / 6,
            2.0, 2 + 1 / 6, 4 + 2 / 3,
            6.0,
        ],
        abs=0.01,
    )


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
    """0..5 of data draws five rings where a bar chart's axis goes to six.

    A radar pads nothing -- that is the same measurement as its extent stopping at the
    data -- so both halves of `strict=False` are exercised here.  The unit then follows the
    radius: five rings on a radius with room for five or six intervals, and ten rings on
    one with room for ten, which is what the `axis-ring` deck drew at 0..5 on its widest
    frame.
    """
    from pptx2svg.resolve.chart import nice_axis_scale

    assert nice_axis_scale(0, 5) == (0.0, 6.0, 1.0)
    assert nice_axis_scale(0, 5, intervals=5, strict=False) == (0.0, 5.0, 1.0)
    assert nice_axis_scale(0, 5, intervals=10, strict=False) == (0.0, 5.0, 0.5)

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

#: Every face lands inside this at every rung.  What it allows is the level band's own
#: pre-existing bias, which runs from 0.11 pt heavy on Arial to 0.60 pt heavy on Courier
#: New and was fitted long before this sweep; what the test is really asserting is that the
#: *slope* is right, which is why the residual is flat across all four rungs instead of
#: fanning out.
WRAP_BAND_TOLERANCE_PT = 0.65

#: Arial's ``hhea`` lineGap at 10 pt: 67 units of 2048, 0.0327 em.  It used to be how far
#: short of PowerPoint our band fell for each line after the first -- Arial is the only one
#: of the five probe faces whose gap is not zero -- and
#: :class:`~pptx2svg.text.metrics.FontMetrics` now carries it, so the tests below assert
#: its *absence* from the residual rather than its presence.
ARIAL_LINE_GAP_PT_PER_LINE = 0.328


@pytest.mark.parametrize("name", list(WRAP_LADDER))
def test_the_wrapped_band_grows_by_one_line_pitch_a_line(name):
    """Four rungs in five faces, and the slope is the face's own baseline-to-baseline pitch.

    PowerPoint's band grew by exactly the same amount from one line to two, two to three
    and three to four in every face, so the residual is flat rather than fanning -- which
    is what says the per-line term is right and the constant is the level band's older fit.
    """
    face, cats, lines, expected = WRAP_LADDER[name]
    children = _wrap_probe(cats, face=face)
    assert not _turned(children), f"{name} turned; it should have wrapped"
    assert len(_label_rows(children)) == lines
    assert _bottom_inset(children) == pytest.approx(expected, abs=WRAP_BAND_TOLERANCE_PT)


def test_arial_carries_its_line_gap_and_the_residual_goes_flat():
    """The face that identified the term, pinned so the term cannot be dropped again.

    Arial is the only one of the five probe faces whose ``hhea`` lineGap is not zero, and
    PowerPoint adds it: 67 units of 2048 is 0.328 pt at 10 pt, which is exactly how far our
    band used to fall short per extra line.  With the gap carried -- Arial's own, read from
    ``arial.ttf`` rather than from the Arimo we draw with -- the residual stops growing and
    settles on the level band's own +0.11 pt bias, the same constant offset every other
    face in the ladder shows.

    The assertion is the *flatness*, not the value: a slope would mean the per-line term is
    wrong again, and it is the slope that the substitute's 87-unit gap would have got wrong
    in the other direction.
    """
    residuals = []
    for name in ("arial-1", "arial-2", "arial-3", "arial-4"):
        _, cats, lines, expected = WRAP_LADDER[name]
        residuals.append(_bottom_inset(_wrap_probe(cats)) - expected)
    for residual in residuals:
        assert residual == pytest.approx(0.110, abs=0.05)
    # A dropped gap would fan these out by 0.328 pt a rung; Tinos' 87-unit gap would fan
    # them the other way by 0.098.  Either is an order of magnitude outside this.
    assert max(residuals) - min(residuals) < 0.01


def test_the_label_needing_the_most_lines_sets_the_band():
    """One three-line label among four that fit lifts the whole band to three lines.

    Measured: the band came out 46.712 pt, the same as a chart where every label needed
    three lines, and the four short ones sat on the *first* row with nothing under them.
    """
    cats = ("MM MM", "MM MM", "MM MM MM MM MM MM", "MM MM", "MM MM")
    children = _wrap_probe(cats)
    assert _bottom_inset(children) == pytest.approx(
        46.712, abs=WRAP_BAND_TOLERANCE_PT
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
        81.212, abs=WRAP_BAND_TOLERANCE_PT
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
    _is_scatter = False
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
# -- Area charts -----------------------------------------------------------------------
#
# Eighteen charts differing in one input each, in the same 2800000 x 2300000 EMU frame the
# bar sweep uses, exported by PowerPoint 16.106 and read back out of the PDF as exact path
# vertices.  The generator below is the deck; nothing binary is committed.

AREA_CATEGORIES = ("Reader", "Writer", "Renderer")
FRAME_W = 220.4724
FRAME_H = 181.1024


def area_chart_xml(
    *,
    values=((3, 4, 5),),
    names=("Alpha", "Beta"),
    categories=AREA_CATEGORIES,
    grouping="standard",
    fills=(None, None),
    outline="",
    legend=None,
    title=None,
    text_size=None,
    cross_between="between",
    blanks="gap",
    dlbls="",
):
    """One area chart, in the shape the exported probe deck used.

    ``cross_between=None`` leaves the element out, which is **not** the same as ``between``
    -- see :func:`test_an_area_with_no_cross_between_draws_as_midcat`.
    """
    series_xml = ""
    for index, column in enumerate(values):
        points = "".join(
            f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>"
            for i, v in enumerate(column)
            if v is not None
        )
        cats = "".join(
            f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(categories)
        )
        fill = fills[index] if index < len(fills) else None
        sp_pr = (
            f"<c:spPr><a:solidFill><a:srgbClr val='{fill}'/></a:solidFill>{outline}</c:spPr>"
            if fill
            else (f"<c:spPr>{outline}</c:spPr>" if outline else "")
        )
        series_xml += (
            f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
            f"<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
            f"<c:pt idx='0'><c:v>{names[index]}</c:v></c:pt>"
            "</c:strCache></c:strRef></c:tx>"
            f"{sp_pr}"
            f"<c:cat><c:strRef><c:strCache><c:ptCount val='{len(categories)}'/>{cats}"
            "</c:strCache></c:strRef></c:cat>"
            "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
            f"<c:ptCount val='{len(column)}'/>{points}</c:numCache></c:numRef></c:val>"
            "</c:ser>"
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
    cross = f"<c:crossBetween val='{cross_between}'/>" if cross_between else ""
    tx_pr = (
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        f"<a:defRPr sz='{int(text_size * 100)}'/></a:pPr></a:p></c:txPr>"
        if text_size
        else ""
    )
    return (
        f"<c:chart>{title_xml}<c:plotArea><c:layout/>"
        f"<c:areaChart><c:grouping val='{grouping}'/><c:varyColors val='0'/>"
        f"{series_xml}{dlbls}"
        "<c:axId val='100002'/><c:axId val='100003'/></c:areaChart>"
        "<c:catAx><c:axId val='100002'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='b'/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='none'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100003'/>"
        "<c:crosses val='autoZero'/><c:auto val='1'/><c:lblAlgn val='ctr'/>"
        "<c:lblOffset val='100'/><c:noMultiLvlLbl val='0'/></c:catAx>"
        "<c:valAx><c:axId val='100003'/><c:scaling><c:orientation val='minMax'/></c:scaling>"
        "<c:delete val='0'/><c:axPos val='l'/><c:majorGridlines/>"
        "<c:numFmt formatCode='General' sourceLinked='1'/>"
        "<c:majorTickMark val='none'/><c:minorTickMark val='none'/>"
        "<c:tickLblPos val='nextTo'/><c:crossAx val='100002'/>"
        f"<c:crosses val='autoZero'/>{cross}</c:valAx>"
        f"</c:plotArea>{legend_xml}"
        f"<c:plotVisOnly val='1'/><c:dispBlanksAs val='{blanks}'/></c:chart>{tx_pr}"
    )


def area_dlbls(*flags):
    """``c:dLbls`` with the named ``c:show*`` flags on.

    Deliberately **no** ``c:dLblPos``: PowerPoint opens a deck whose area chart carries one
    as ``[Repaired]`` and will not export it, for ``ctr`` -- the only value ECMA-376 lists
    for an area series -- as much as for anything else.  So an area label has exactly one
    placement and the file cannot ask for another.
    """
    body = "".join(
        f"<c:show{flag} val='{1 if flag in flags else 0}'/>"
        for flag in ("LegendKey", "Val", "CatName", "SerName", "Percent", "BubbleSize")
    )
    return f"<c:dLbls>{body}</c:dLbls>"


#: Points from the frame's four edges to the plot rectangle, measured off PowerPoint's
#: export of the probe deck.  Worst residual across the eighteen is 0.23 pt.
AREA_SWEEP = {
    "bare": ({}, {"left": 21.073, "right": 11.000, "top": 11.102, "bottom": 24.965}),
    "midcat": (
        {"cross_between": "midCat"},
        {"left": 26.433, "right": 30.840, "top": 11.102, "bottom": 24.965},
    ),
    "no-crossbetween": (
        {"cross_between": None},
        {"left": 26.433, "right": 30.840, "top": 11.102, "bottom": 24.965},
    ),
    "two": (
        {"values": ((3, 4, 5), (2, 5, 1)), "legend": "b"},
        {"left": 21.072, "right": 11.000, "top": 11.102, "bottom": 49.048},
    ),
    "stacked": (
        {"values": ((3, 4, 5), (2, 5, 1)), "grouping": "stacked", "legend": "b"},
        {"left": 26.413, "right": 11.000, "top": 11.103, "bottom": 49.048},
    ),
    "percent": (
        {"values": ((3, 4, 5), (2, 5, 1)), "grouping": "percentStacked", "legend": "b"},
        {"left": 40.012, "right": 11.000, "top": 11.103, "bottom": 49.048},
    ),
    "negative": (
        {"values": ((3, -2, 5),)},
        {"left": 24.478, "right": 11.000, "top": 11.102, "bottom": 11.102},
    ),
    "five-cats": (
        {"values": ((3, 4, 5, 2, 6),), "categories": ("A", "B", "C", "D", "E")},
        {"left": 21.073, "right": 11.000, "top": 11.103, "bottom": 24.965},
    ),
    "legend-r": (
        {"values": ((3, 4, 5), (2, 5, 1)), "legend": "r"},
        {"left": 21.073, "right": 58.925, "top": 11.102, "bottom": 24.965},
    ),
    "font14": (
        {"text_size": 14},
        {"left": 26.907, "right": 11.000, "top": 13.545, "bottom": 32.353},
    ),
    # The `title` probe measured a 40.803 pt top inset, exactly what a bar chart's title
    # band gives -- but the title's *face* comes from the deck's own cascade, which
    # `_build` stubs out, so the number is only reproducible through the full pipeline.
    # It is the bar sweep's constant and the bar sweep asserts it.
}


def _rect_of(children):
    """The plot rectangle, off the longest drawn line each way; see `_plot_rect`."""
    lines = [c for c in children if isinstance(c, m.ConnectorElement)]
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


def _insets(children, width=FRAME_W, height=FRAME_H):
    left, top, right, bottom = _rect_of(children)
    return {
        "left": left,
        "right": width - right,
        "top": top,
        "bottom": height - bottom,
    }


def _custom_paths(children):
    """Every custom-geometry shape, in draw order, with absolute vertices."""
    out = []
    for child in children:
        if not isinstance(child, m.ShapeElement):
            continue
        if not isinstance(child.geometry, m.CustomGeometry):
            continue
        ox = _pt(child.transform.offset_x)
        oy = _pt(child.transform.offset_y)
        points = []
        for token in re.finditer(
            r"([MLC])((?:\s+-?[\d.]+)+)", child.geometry.paths[0].commands
        ):
            nums = [float(value) for value in token.group(2).split()]
            for index in range(0, len(nums), 2):
                points.append((ox + nums[index], oy + nums[index + 1]))
        out.append((child, points))
    return out


def _areas(children):
    return [pair for pair in _custom_paths(children) if not isinstance(pair[0].fill, m.NoFill)]


def _labels_by_text(children):
    return {
        "".join(run.text for p in child.text_body.paragraphs for run in p.runs): child
        for child in children
        if isinstance(child, m.ShapeElement) and child.text_body is not None
    }


@pytest.mark.parametrize("name", list(AREA_SWEEP))
def test_the_area_sweep_reproduces_powerpoints_plot_rectangle(name):
    kwargs, expected = AREA_SWEEP[name]
    children, data = _build(area_chart_xml(**kwargs), width=FRAME_W, height=FRAME_H)
    assert data.kind == "areaChart"
    ours = _insets(children)
    for edge, truth in expected.items():
        assert ours[edge] == pytest.approx(truth, abs=SWEEP_TOLERANCE_PT), (
            f"{name} {edge}: PowerPoint {truth:.4f}, ours {ours[edge]:.4f}"
        )


def test_an_area_closes_to_the_zero_line_and_not_to_the_plot_floor():
    """Measured on the negative probe: the polygon's return edge is the value axis' zero.

    Its three vertices are at 64.029, 152.345 and 28.758 pt down the frame and its bottom
    edge runs along 117.069 -- which is where a -3..6 axis puts 0, with the plot's own
    floor 53 pt lower.  The path therefore crosses itself where the line crosses zero, and
    PowerPoint leaves the bow tie exactly as it falls: it fills nonzero-winding, which is
    the renderer's default.
    """
    children, _ = _build(
        area_chart_xml(values=((3, -2, 5),)), width=FRAME_W, height=FRAME_H
    )
    ((_, points),) = _areas(children)
    assert [y for _, y in points[:3]] == pytest.approx([64.029, 152.345, 28.758], abs=0.6)
    assert [y for _, y in points[3:]] == pytest.approx([117.069] * 3, abs=0.6)


def test_area_vertices_sit_at_the_band_centres():
    """`c:crossBetween="between"`: 52.473, 115.292 and 178.072 pt across the frame."""
    children, _ = _build(area_chart_xml(), width=FRAME_W, height=FRAME_H)
    ((_, points),) = _areas(children)
    assert [x for x, _ in points[:3]] == pytest.approx(
        [52.473, 115.292, 178.072], abs=0.6
    )


def test_an_area_with_no_cross_between_draws_as_midcat():
    """**An absent `c:crossBetween` is `midCat` on an area chart**, not `between`.

    Three probes settle it: `between` put the three vertices at the band centres, while
    `midCat` and **no element at all** were byte-identical to each other at 26.433,
    108.015 and 189.632 -- the plot's left edge, its midpoint and its right edge.  The
    plot itself narrows to make room for the half of the first and last category label
    that now hangs outside it, which the sweep above asserts.

    A *line* chart's absent case is not measured, and the default there is left alone.
    """
    absent, _ = _build(area_chart_xml(cross_between=None), width=FRAME_W, height=FRAME_H)
    stated, _ = _build(
        area_chart_xml(cross_between="midCat"), width=FRAME_W, height=FRAME_H
    )
    ((_, theirs),) = _areas(stated)
    ((_, points),) = _areas(absent)
    assert points == theirs
    assert [x for x, _ in points[:3]] == pytest.approx(
        [26.433, 108.015, 189.632], abs=0.6
    )


def test_midcat_centres_each_category_label_on_its_tick():
    """Measured: `Reader` centred on the plot's left edge, 26.433 pt across the frame.

    `Writer` lands on the plot's midpoint and `Renderer` on its right edge.  A `between`
    label is centred in a *band* instead, which is what every other test here exercises.
    """
    children, _ = _build(
        area_chart_xml(cross_between="midCat"), width=FRAME_W, height=FRAME_H
    )
    labels = _labels_by_text(children)
    for text, expected in (
        ("Reader", 26.433),
        ("Writer", 108.015),
        ("Renderer", 189.632),
    ):
        box = labels[text].transform
        centre = _pt(box.offset_x) + _pt(box.extent_width) / 2
        assert centre == pytest.approx(expected, abs=0.6), text


def test_areas_are_painted_in_series_order_with_the_first_at_the_back():
    """Two probes say it is order and not size, and the fills are fully opaque.

    Swapping the two series' values swapped which one ended up hidden, and a probe whose
    first series covers the second entirely still emitted the first path first.  Every one
    of the eighteen area fills came back at alpha 255, so the front series really does
    hide what is behind it -- that is PowerPoint's picture, not an omission here.
    """
    children, _ = _build(
        area_chart_xml(values=((6, 6, 6), (2, 3, 1)), legend="b"),
        width=FRAME_W,
        height=FRAME_H,
    )
    shapes = [shape for shape, _ in _areas(children)]
    assert [shape.fill.color.hex for shape in shapes] == ["#4472C4", "#ED7D31"]


def test_a_stacked_area_rides_on_the_running_total():
    """The band's lower edge is the total *without* this series, traced backwards.

    Measured on the stacked probe, whose second series' polygon returns along the first
    series' own line rather than along the axis.
    """
    children, _ = _build(
        area_chart_xml(values=((3, 4, 5), (2, 5, 1)), grouping="stacked"),
        width=FRAME_W,
        height=FRAME_H,
    )
    first, second = _areas(children)
    tops = [y for _, y in first[1][:3]]
    returns = [y for _, y in second[1][3:]]
    assert returns == pytest.approx(list(reversed(tops)), abs=0.01)


def test_a_percent_stacked_area_fills_the_whole_plot():
    """The last series reaches 100% in every category, so its top edge is the plot's."""
    children, _ = _build(
        area_chart_xml(values=((3, 4, 5), (2, 5, 1)), grouping="percentStacked"),
        width=FRAME_W,
        height=FRAME_H,
    )
    _, top, _, _ = _rect_of(children)
    _, second = _areas(children)
    assert [y for _, y in second[1][:3]] == pytest.approx([top] * 3, abs=0.01)


def test_a_blank_area_run_of_one_point_draws_nothing():
    """The `dispBlanksAs="gap"` probe drew **no area at all**.

    Its middle value is missing, which leaves two runs of one point each, and a single
    point has no area.  That is the whole rule -- a blank splits an area exactly as it
    splits a line -- but it is worth an assertion because "draw nothing" is also what a
    silent failure looks like.  With `zero` the same file draws one unbroken area dipping
    to the axis.
    """
    children, _ = _build(
        area_chart_xml(values=((3, None, 5),)), width=FRAME_W, height=FRAME_H
    )
    assert _areas(children) == []

    zeroed, _ = _build(
        area_chart_xml(values=((3, None, 5),), blanks="zero"),
        width=FRAME_W,
        height=FRAME_H,
    )
    ((_, points),) = _areas(zeroed)
    _, _, _, bottom = _rect_of(zeroed)
    assert points[1][1] == pytest.approx(bottom, abs=0.6)


def test_an_area_takes_no_outline_unless_the_file_states_one():
    """Eight probes stating none got a bare fill; one stating `w="25400"` got 2 pt."""
    plain, _ = _build(area_chart_xml(), width=FRAME_W, height=FRAME_H)
    ((shape, _),) = _areas(plain)
    assert shape.outline is None

    stroked, _ = _build(
        area_chart_xml(
            fills=("F97316",),
            outline=(
                "<a:ln w='25400'><a:solidFill><a:srgbClr val='111827'/>"
                "</a:solidFill></a:ln>"
            ),
        ),
        width=FRAME_W,
        height=FRAME_H,
    )
    ((shape, _),) = _areas(stroked)
    assert shape.fill.color.hex == "#F97316"
    assert shape.outline.width == 25400
    assert shape.outline.fill.color.hex == "#111827"


def test_an_area_label_sits_at_the_centre_of_its_own_band():
    """Measured on two probes, to within 0.21 pt.

    An unstacked 3/4/5 put its labels at 1.5, 2.0 and 2.5 on the value axis -- the
    midpoints between the line and the zero it fills to -- and a stacked pair put the
    second series' at 4, 6.5 and 5.5, the midpoints of its *segments* and not of the
    stack.
    """
    children, _ = _build(
        area_chart_xml(dlbls=area_dlbls("Val")), width=FRAME_W, height=FRAME_H
    )
    labels = _labels_by_text(children)
    assert _baseline(labels["3"]) == pytest.approx(122.762, abs=0.6)
    assert _baseline(labels["4"]) == pytest.approx(110.762, abs=0.6)
    assert _baseline(labels["5"]) == pytest.approx(98.522, abs=0.6)


def test_an_area_legends_with_a_swatch_and_not_a_rule():
    """Measured: a 5.492 pt square -- the bar chart's key, not the line chart's rule."""
    from pptx2svg.resolve.chart import LEGEND_SWATCH_EM

    children, _ = _build(
        area_chart_xml(values=((3, 4, 5), (2, 5, 1)), legend="b"),
        width=FRAME_W,
        height=FRAME_H,
    )
    swatches = [
        child
        for child in children
        if isinstance(child, m.ShapeElement)
        and child.text_body is None
        and isinstance(child.geometry, m.PresetGeometry)
        and child.geometry.preset == "rect"
    ]
    assert len(swatches) == 2
    for swatch in swatches:
        assert _pt(swatch.transform.extent_width) == pytest.approx(
            LEGEND_SWATCH_EM * 10.0, abs=0.05
        )


# -- Scatter charts --------------------------------------------------------------------
#
# Twenty-four charts across two decks, same frame, exported by PowerPoint 16.106.  A
# scatter is the one Cartesian kind with **two value axes and no category axis**, so
# nothing about a category band -- the wrap, the 45 degree turn, the reserve those ask for
# -- applies to it; what its bottom band actually is, measured, is the horizontal bar
# chart's row of value labels.

SCATTER_XS = (1, 2, 3, 4, 5)
SCATTER_YS = (3, 4, 5, 2, 6)
SCATTER_YS2 = (1, 5, 2, 6, 3)


def scatter_chart_xml(
    *,
    series=((SCATTER_XS, SCATTER_YS),),
    names=("Alpha", "Beta"),
    style="lineMarker",
    markers=(None, None),
    lines=(None, None),
    smooth=(0, 0),
    legend=None,
    text_size=None,
    x_gridlines=False,
    x_delete=False,
    dlbls=("",),
):
    """One scatter chart, in the shape the exported probe deck used."""
    series_xml = ""
    for index, (xs, ys) in enumerate(series):

        def _cache(values):
            points = "".join(
                f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>"
                for i, v in enumerate(values)
                if v is not None
            )
            return (
                "<c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
                f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef>"
            )

        line = lines[index] if index < len(lines) else None
        marker = markers[index] if index < len(markers) else None
        smooth_value = smooth[index] if index < len(smooth) else 0
        labels = dlbls[index] if index < len(dlbls) else ""
        series_xml += (
            f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
            f"<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
            f"<c:pt idx='0'><c:v>{names[index]}</c:v></c:pt>"
            "</c:strCache></c:strRef></c:tx>"
            + (f"<c:spPr>{line}</c:spPr>" if line else "")
            + (marker or "")
            + labels
            + f"<c:xVal>{_cache(xs)}</c:xVal><c:yVal>{_cache(ys)}</c:yVal>"
            + ("" if smooth_value is None else f"<c:smooth val='{smooth_value}'/>")
            + "</c:ser>"
        )

    def _axis(axis_id, cross_id, position, *, gridlines, delete=False):
        grid = "<c:majorGridlines/>" if gridlines else ""
        return (
            f"<c:valAx><c:axId val='{axis_id}'/>"
            "<c:scaling><c:orientation val='minMax'/></c:scaling>"
            f"<c:delete val='{1 if delete else 0}'/><c:axPos val='{position}'/>{grid}"
            "<c:numFmt formatCode='General' sourceLinked='1'/>"
            "<c:majorTickMark val='none'/><c:minorTickMark val='none'/>"
            "<c:tickLblPos val='nextTo'/>"
            f"<c:crossAx val='{cross_id}'/><c:crosses val='autoZero'/>"
            "<c:crossBetween val='midCat'/></c:valAx>"
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
    return (
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        f"<c:scatterChart><c:scatterStyle val='{style}'/><c:varyColors val='0'/>"
        f"{series_xml}"
        "<c:axId val='100002'/><c:axId val='100003'/></c:scatterChart>"
        + _axis("100002", "100003", "b", gridlines=x_gridlines, delete=x_delete)
        + _axis("100003", "100002", "l", gridlines=True)
        + f"</c:plotArea>{legend_xml}"
        f"<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>{tx_pr}"
    )


def scatter_dlbls(*flags, position=None):
    pos = f"<c:dLblPos val='{position}'/>" if position else ""
    body = pos + "".join(
        f"<c:show{flag} val='{1 if flag in flags else 0}'/>"
        for flag in ("LegendKey", "Val", "CatName", "SerName", "Percent", "BubbleSize")
    )
    return f"<c:dLbls>{body}</c:dLbls>"


#: Insets from the frame's four edges, measured off PowerPoint's export.  Worst residual
#: across the eighteen scatter probes is 0.32 pt.
SCATTER_SWEEP = {
    "bare": ({}, {"left": 21.073, "right": 13.670, "top": 11.102, "bottom": 24.965}),
    "x-wide": (
        {"series": (((10, 20, 30, 40, 50), SCATTER_YS),)},
        {"left": 21.073, "right": 16.340, "top": 11.102, "bottom": 24.965},
    ),
    # A negative x range floats the value axis into the plot, so the y labels go with it
    # and the left inset is the *first x label's* overhang instead of a label column.
    "x-neg": (
        {"series": (((-2, -1, 0, 1, 2), SCATTER_YS),)},
        {"left": 15.373, "right": 13.670, "top": 11.102, "bottom": 24.965},
    ),
    # A negative y range does the mirror image: no band under the plot at all.
    "y-neg": (
        {"series": ((SCATTER_XS, (3, -2, 5, -1, 4)),)},
        {"left": 24.478, "right": 13.670, "top": 11.102, "bottom": 11.102},
    ),
    "x-deleted": (
        {"x_delete": True},
        {"left": 21.072, "right": 11.000, "top": 11.103, "bottom": 11.102},
    ),
    "two": (
        {"series": ((SCATTER_XS, SCATTER_YS), (SCATTER_XS, SCATTER_YS2)), "legend": "b"},
        {"left": 21.072, "right": 13.670, "top": 11.103, "bottom": 49.048},
    ),
    "font14": (
        {"text_size": 14},
        {"left": 26.907, "right": 14.740, "top": 13.545, "bottom": 32.353},
    ),
}


def _polyline_points(children):
    return [points for shape, points in _custom_paths(children)
            if isinstance(shape.fill, m.NoFill)]


def _marker_boxes(children):
    """Every marker, left to right: a preset shape with no text in it."""
    markers = [
        child
        for child in children
        if isinstance(child, m.ShapeElement)
        and child.text_body is None
        and isinstance(child.geometry, m.PresetGeometry)
    ]
    return sorted(markers, key=lambda shape: shape.transform.offset_x)


@pytest.mark.parametrize("name", list(SCATTER_SWEEP))
def test_the_scatter_sweep_reproduces_powerpoints_plot_rectangle(name):
    kwargs, expected = SCATTER_SWEEP[name]
    children, data = _build(scatter_chart_xml(**kwargs), width=FRAME_W, height=FRAME_H)
    assert data.kind == "scatterChart"
    ours = _insets(children)
    for edge, truth in expected.items():
        assert ours[edge] == pytest.approx(truth, abs=SWEEP_TOLERANCE_PT), (
            f"{name} {edge}: PowerPoint {truth:.4f}, ours {ours[edge]:.4f}"
        )


def test_a_scatter_has_no_category_axis_and_takes_the_coarse_bottom_axis():
    """Both axes are value axes, and **the one along the bottom comes out coarser**.

    The same 1..5 of data that the y axis draws as 0..7 by 1 the x axis draws as 0..6 by
    2, which is the rule a horizontal bar chart's value axis already follows.  Four x-axis
    probes bracket the cap at four intervals: `x-float`, whose 0.5..4.5 rounds to a *five*
    interval 0..5 at unit 1, was coarsened to 0..6 by 2 anyway, and `x-neg`'s four-interval
    -4..4 by 2 was kept.
    """
    children, _ = _build(scatter_chart_xml(), width=FRAME_W, height=FRAME_H)
    labels = _labels_by_text(children)
    assert set("01234567") <= set(labels)
    _, top, _, bottom = _rect_of(children)
    # Four x labels, 0/2/4/6, along the bottom band.
    along_bottom = [
        text
        for text, shape in labels.items()
        if _pt(shape.transform.offset_y) > bottom
    ]
    assert sorted(along_bottom) == ["0", "2", "4", "6"]

    floats, _ = _build(
        scatter_chart_xml(series=(((0.5, 1.5, 2.5, 3.5, 4.5), SCATTER_YS),)),
        width=FRAME_W,
        height=FRAME_H,
    )
    _, _, _, floats_bottom = _rect_of(floats)
    assert sorted(
        text
        for text, shape in _labels_by_text(floats).items()
        if _pt(shape.transform.offset_y) > floats_bottom
    ) == ["0", "2", "4", "6"]


def test_a_scatter_point_maps_through_both_value_axes():
    """Measured: x = 1..5 on a 0..6 axis and y = 3, 4, 5, 2, 6 on a 0..7 one."""
    children, _ = _build(
        scatter_chart_xml(), width=FRAME_W, height=FRAME_H
    )
    (points,) = _polyline_points(children)
    assert points[0] == pytest.approx((52.027, 93.980), abs=0.6)
    assert points[-1] == pytest.approx((175.847, 31.822), abs=0.6)


def test_scatter_style_decides_nothing_and_the_series_markup_decides_everything():
    """`marker`, `line` and `lineMarker` came back **byte-identical**: line and markers.

    So `c:scatterStyle` is not what turns either off.  What does is the series' own
    markup, each measured: `<a:ln><a:noFill/></a:ln>` leaves the markers alone on the
    plot, and `<c:symbol val="none"/>` leaves the line alone.
    """
    shapes = {}
    for style in ("marker", "line", "lineMarker"):
        children, _ = _build(
            scatter_chart_xml(style=style), width=FRAME_W, height=FRAME_H
        )
        shapes[style] = (len(_polyline_points(children)), len(_marker_boxes(children)))
    assert shapes["marker"] == shapes["line"] == shapes["lineMarker"] == (1, 5)

    no_line, _ = _build(
        scatter_chart_xml(lines=("<a:ln><a:noFill/></a:ln>",)),
        width=FRAME_W,
        height=FRAME_H,
    )
    assert _polyline_points(no_line) == []
    assert len(_marker_boxes(no_line)) == 5

    no_marker, _ = _build(
        scatter_chart_xml(markers=("<c:marker><c:symbol val='none'/></c:marker>",)),
        width=FRAME_W,
        height=FRAME_H,
    )
    assert len(_polyline_points(no_marker)) == 1
    assert _marker_boxes(no_marker) == []


def test_a_scatter_marker_with_no_size_is_six_points():
    """Not ECMA-376's seven.  The second series' **square** measured 6.000 x 6.000.

    A square is axis-aligned, so unlike the diamond the radar measured -- 5.76 pt across,
    which is a 6 pt box with its tips inside PowerPoint's 0.24 pt output grid -- there is
    nothing to argue about.  A line chart with an explicit `c:size val="6"` square measured
    the same 6.000 on the same export.
    """
    from pptx2svg.resolve.chart import DEFAULT_MARKER_SIZE_PT

    assert DEFAULT_MARKER_SIZE_PT == 6.0
    children, _ = _build(scatter_chart_xml(), width=FRAME_W, height=FRAME_H)
    for marker in _marker_boxes(children):
        assert _pt(marker.transform.extent_width) == pytest.approx(6.0, abs=0.01)
        assert _pt(marker.transform.extent_height) == pytest.approx(6.0, abs=0.01)

    sized, _ = _build(
        scatter_chart_xml(
            markers=("<c:marker><c:symbol val='circle'/><c:size val='12'/></c:marker>",)
        ),
        width=FRAME_W,
        height=FRAME_H,
    )
    assert _pt(_marker_boxes(sized)[0].transform.extent_width) == pytest.approx(
        12.0, abs=0.01
    )


def test_a_scatter_joins_its_points_in_the_order_the_file_lists_them():
    """The unsorted-x probe's path runs 3, 1, 5, 2, 4 -- data order, not sorted by x."""
    children, _ = _build(
        scatter_chart_xml(series=(((3, 1, 5, 2, 4), SCATTER_YS),)),
        width=FRAME_W,
        height=FRAME_H,
    )
    (points,) = _polyline_points(children)
    assert [x for x, _ in points] == pytest.approx(
        [113.937, 52.027, 175.847, 82.982, 144.892], abs=0.6
    )


def test_a_blank_breaks_a_scatter_into_two_strokes():
    """And the vertex list alone does not say so -- only the path *operators* do.

    The probe with a missing middle y, under `dispBlanksAs="gap"`, exports as one path
    object with four points, which reads as unbroken; its segment kinds are move, line,
    move, line.  Two disjoint strokes with a gap where the blank is, and five markers
    become four.  Reading coordinates without reading the operators is how that gets
    missed -- it was, until the render was put beside PowerPoint's.
    """
    children, _ = _build(
        scatter_chart_xml(series=((SCATTER_XS, (3, 4, None, 2, 6)),)),
        width=FRAME_W,
        height=FRAME_H,
    )
    runs = _polyline_points(children)
    assert [len(run) for run in runs] == [2, 2]
    assert [x for x, _ in runs[0]] == pytest.approx([52.027, 82.978], abs=0.6)
    assert [x for x, _ in runs[1]] == pytest.approx([144.898, 175.847], abs=0.6)
    assert len(_marker_boxes(children)) == 4


SCATTER_LABEL_BASELINES = {
    # position -> (left edge, baseline) of the label on the first point, measured.
    None: (61.027, 96.842),
    "l": (37.687, 96.842),
    "t": (49.357, 83.162),
    "b": (49.357, 110.522),
    "ctr": (49.357, 96.909),
}


@pytest.mark.parametrize("position", list(SCATTER_LABEL_BASELINES))
def test_each_scatter_label_position_matches_powerpoints(position):
    """All five, one probe each, against a point at (52.027, 93.980) with a 6 pt marker.

    Every one is the line chart's own geometry read off a different edge of the marker:
    `r` -- the default when the file states none -- and `l` put the label's near edge a
    marker radius plus 0.6 em from the point, which is 9.000 pt measured and 9.000 pt
    predicted on both sides; `t` and `b` put its line box a radius plus the bar chart's
    4.85 pt gap away; `ctr` is the ink centre.

    `b` is the one loose number, 0.9 pt low, and it is the slack this file records
    elsewhere: PowerPoint's line box runs about a point taller than our metrics give, so a
    placement hung off the *ascent* inherits all of the difference where one hung off the
    descent inherits none.
    """
    expected_left, expected_baseline = SCATTER_LABEL_BASELINES[position]
    children, _ = _build(
        scatter_chart_xml(dlbls=(scatter_dlbls("Val", position=position),)),
        width=FRAME_W,
        height=FRAME_H,
    )
    # The first point's label is the "3"; the y axis' own "3" is the one at the far left.
    threes = [
        shape
        for text, shape in _labels_by_text(children).items()
        if text == "3"
    ]
    assert threes, "no data label drawn"
    label = max(threes, key=lambda shape: shape.transform.offset_x)
    body = label.text_body
    align = body.paragraphs[0].properties.alignment
    box_left = _pt(label.transform.offset_x)
    width = _pt(label.transform.extent_width)
    from pptx2svg.resolve.chart import text_width

    ink = text_width("3", "Aptos", 10.0)
    if align == "ctr":
        box_left += (width - ink) / 2
    elif align == "r":
        box_left += width - ink
    assert box_left == pytest.approx(expected_left, abs=0.6)
    assert _baseline(label) == pytest.approx(expected_baseline, abs=1.0)


def test_a_scatter_legends_with_a_rule_and_its_marker():
    """The line chart's key, not the bar's swatch: 19.200 pt of rule, marker on its middle."""
    from pptx2svg.resolve.chart import (
        LINE_LEGEND_KEY_GAP_PT,
        LINE_LEGEND_KEY_PT,
        ChartBuilder,
        ChartFont,
        font_box,
    )

    class _FakeScatter:
        _is_line = False
        _is_scatter = True
        _is_bubble = False
        _is_radar = False
        _radar_style = "marker"

    font = ChartFont(family="Aptos", box=font_box("Aptos", 10.0))
    assert ChartBuilder._legend_key_size(_FakeScatter(), font) == (
        LINE_LEGEND_KEY_PT,
        LINE_LEGEND_KEY_GAP_PT,
    )

    children, _ = _build(
        scatter_chart_xml(
            series=((SCATTER_XS, SCATTER_YS), (SCATTER_XS, SCATTER_YS2)), legend="b"
        ),
        width=FRAME_W,
        height=FRAME_H,
    )
    keys = [
        child
        for child in children
        if isinstance(child, m.ConnectorElement)
        and _pt(child.transform.extent_width) == pytest.approx(19.2, abs=0.01)
    ]
    assert len(keys) == 2


def test_a_scatter_with_negative_x_puts_its_value_labels_beside_the_axis():
    """`tickLblPos="nextTo"` means next to the *axis*, and the axis is at x = 0.

    Measured: the value axis is drawn at 111.088 pt -- the zero tick -- with the plot's own
    left edge 95.7 pt away, and the y labels are right-aligned 9.23 pt to its left, which
    is the same `descent + 0.645 em` the column always uses.
    """
    children, _ = _build(
        scatter_chart_xml(series=(((-2, -1, 0, 1, 2), SCATTER_YS),)),
        width=FRAME_W,
        height=FRAME_H,
    )
    verticals = [
        child
        for child in children
        if isinstance(child, m.ConnectorElement) and child.transform.extent_width == 0
    ]
    axis = max(verticals, key=lambda line: line.transform.extent_height)
    assert _pt(axis.transform.offset_x) == pytest.approx(111.088, abs=0.6)
    # Two labels read "0" -- the y axis' own and the x axis' -- so pick the one that is
    # not in the band under the plot.
    _, _, _, bottom = _rect_of(children)
    zero = next(
        child
        for child in children
        if isinstance(child, m.ShapeElement)
        and child.text_body is not None
        and "".join(
            run.text for p in child.text_body.paragraphs for run in p.runs
        ) == "0"
        and _pt(child.transform.offset_y) < bottom
    )
    right = _pt(zero.transform.offset_x) + _pt(zero.transform.extent_width)
    assert right == pytest.approx(111.088 - 9.267, abs=0.6)


#: ``(x data, PowerPoint's x axis)`` for the six probes that ask whether a scatter's axis
#: has to start at zero.  The bracket on the threshold is [0.80, 0.84): 40/50 keeps zero
#: and 42/50 does not, with Excel's folklore 5/6 = 0.8333 inside it.
SCATTER_ZERO_ANCHOR = {
    "years": ((2010, 2012, 2014, 2016, 2020), (2005.0, 2025.0, 5.0)),
    "40-50-keeps-zero": ((40, 42, 44, 46, 50), (0.0, 60.0, 20.0)),
    "42-50-drops-zero": ((42, 44, 46, 48, 50), (40.0, 55.0, 5.0)),
    "tight-and-far": ((100, 101, 102, 103, 104), (98.0, 106.0, 2.0)),
    "all-negative": ((-50, -40, -30, -20, -10), (-60.0, 0.0, 20.0)),
    "near-zero-keeps-it": ((10, 20, 30, 40, 50), (0.0, 60.0, 20.0)),
}


@pytest.mark.parametrize("name", list(SCATTER_ZERO_ANCHOR))
def test_a_scatters_axis_leaves_zero_out_when_the_data_sits_far_from_it(name):
    """**A scatter is the one type whose axis is not anchored at zero.**

    A bar must start at its axis -- a bar that does not is a different picture -- and
    `nice_axis_scale` anchors every other type there.  Applied to a scatter it destroys the
    chart: years 2010..2020 against a measurement collapse into a 1% sliver of a 0..3000
    axis.  PowerPoint agrees, and the six probes here bracket when it lets go: the near end
    has to be past 5/6 of the far one, measured to [0.80, 0.84) by the two that straddle
    it.

    The unanchored extent rounds strictly outward at **both** ends, where an anchored one
    holds its low end at zero: 100..104 came back 98..106 and 2010..2020 came back
    2005..2025, each a whole unit clear of the data.
    """
    from pptx2svg.resolve.chart import nice_axis_scale

    xs, expected = SCATTER_ZERO_ANCHOR[name]
    # Four intervals is what a 220.47 pt frame gives a bottom axis with 10 pt labels --
    # `bottom_axis_intervals` -- which is the frame every one of these probes was drawn on.
    assert nice_axis_scale(
        min(xs), max(xs), intervals=4, anchor_zero=False
    ) == pytest.approx(expected)

    # And it reaches the drawing: the first and last x labels are the axis' own ends.
    children, _ = _build(
        scatter_chart_xml(series=((xs, SCATTER_YS),)), width=FRAME_W, height=FRAME_H
    )
    _, _, _, bottom = _rect_of(children)
    along_bottom = sorted(
        (
            _pt(shape.transform.offset_x),
            "".join(r.text for p in shape.text_body.paragraphs for r in p.runs),
        )
        for shape in children
        if isinstance(shape, m.ShapeElement)
        and shape.text_body is not None
        and _pt(shape.transform.offset_y) > bottom
    )
    assert float(along_bottom[0][1]) == pytest.approx(expected[0])
    assert float(along_bottom[-1][1]) == pytest.approx(expected[1])


def test_the_zero_anchor_still_holds_for_every_other_chart_type():
    """Only a scatter was measured, so only a scatter lets go.

    `real-financial-report.pptx`'s radar has 65..100 of data and PowerPoint draws it from
    zero; so does every bar in the corpus.  A line chart of temperatures has the same
    problem a scatter does and **no probe has ever shown what PowerPoint does with one**.
    """
    from pptx2svg.resolve.chart import nice_axis_scale

    assert nice_axis_scale(2010, 2020) == pytest.approx((0.0, 2500.0, 500.0))
    assert nice_axis_scale(2010, 2020, intervals=4, anchor_zero=False) == pytest.approx(
        (2005.0, 2025.0, 5.0)
    )


def test_the_two_observations_that_refuted_every_ladder_now_come_out_together():
    """The pair that no ratio rule could order, and both are the tick count.

    A scatter probe of 120..160 came back 0..180 **by 20** on a 145 pt axis where halving
    the power of ten gave 0..200 by 50, and `real-financial-report.pptx`'s radar came back
    0..100 by 50 -- two rings at radii 22.78 and 45.56 -- where the ladder that reproduced
    the scatter gave five.  A ratio threshold has to accept 2.0 and refuse 3.2 at once,
    which is why three of them were written and reverted.

    Neither is about the ratio.  The scatter's y axis has a 181.1 pt frame at 10 pt labels,
    which is room for ten intervals, and 168 padded over ten rounds up to 20.  The radar's
    radius is 45.56 pt at 13.406 pt of line box, which is room for four, and a hundred over
    four rounds up the 1-2-5 ladder to 50.
    """
    from pptx2svg.resolve.chart import nice_axis_scale

    assert nice_axis_scale(120, 160, intervals=10) == pytest.approx((0.0, 180.0, 20.0))
    assert nice_axis_scale(65, 100, intervals=4, strict=False) == pytest.approx(
        (0.0, 100.0, 50.0)
    )


def test_a_midcat_bar_chart_keeps_its_labels_in_the_bands_with_its_bars():
    """Excel writes `midCat` on a column chart's value axis, and the bars do not move.

    It is the "Axis position: on tick marks" checkbox, so a perfectly ordinary column chart
    carries it.  Only `_draw_lines` and `_draw_areas` put their marks on the ticks, so
    reading it on a bar chart would place every label up to 41 pt from the bar it names and
    push the first one outside the plot.  What PowerPoint does with such a file is **not
    measured**; keeping the bars and their labels in the bands together is the one reading
    that cannot contradict itself.
    """
    body = bar(
        "<c:ser><c:cat><c:strRef><c:strCache><c:ptCount val='3'/>"
        "<c:pt idx='0'><c:v>Alpha</c:v></c:pt><c:pt idx='1'><c:v>Beta</c:v></c:pt>"
        "<c:pt idx='2'><c:v>Gamma</c:v></c:pt></c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef><c:numCache><c:ptCount val='3'/>"
        "<c:pt idx='0'><c:v>3</c:v></c:pt><c:pt idx='1'><c:v>4</c:v></c:pt>"
        "<c:pt idx='2'><c:v>5</c:v></c:pt></c:numCache></c:numRef></c:val></c:ser>"
    ).replace("<c:valAx>", "<c:valAx><c:crossBetween val='midCat'/>")
    children, _ = _build(body, width=FRAME_W, height=FRAME_H)
    labels = _labels_by_text(children)
    bars = sorted(
        (
            child
            for child in children
            if isinstance(child, m.ShapeElement)
            and isinstance(child.fill, m.SolidFill)
            and child.text_body is None
        ),
        key=lambda shape: shape.transform.offset_x,
    )
    assert len(bars) == 3
    for shape, text in zip(bars, ("Alpha", "Beta", "Gamma")):
        bar_centre = _pt(shape.transform.offset_x) + _pt(shape.transform.extent_width) / 2
        box = labels[text].transform
        label_centre = _pt(box.offset_x) + _pt(box.extent_width) / 2
        assert label_centre == pytest.approx(bar_centre, abs=0.5), text
    # And nothing is pushed outside the frame the way the tick placement would.
    assert min(_pt(labels[t].transform.offset_x) for t in ("Alpha", "Beta", "Gamma")) > 10.0


def test_an_axis_states_where_it_itself_crosses_the_other_one():
    """`c:crosses` belongs to the axis it is written on, on a scatter as on a bar.

    `<c:crosses val="max"/>` on the **bottom** axis moves the *horizontal* line to the top
    of the plot, not the vertical line to the right of it -- which is how the sibling
    `_category_axis_position` has always read it.  Reading each axis' own `c:crosses` for
    the perpendicular line put the value axis at the plot's right edge and dragged the y
    tick labels across the plot with it.  Only `autoZero` is measured; this pins the
    reading rather than a second one invented for scatters.
    """
    xml = scatter_chart_xml()
    at_max = xml.replace(
        "<c:crossAx val='100003'/><c:crosses val='autoZero'/>",
        "<c:crossAx val='100003'/><c:crosses val='max'/>",
        1,
    )
    assert at_max != xml
    children, _ = _build(at_max, width=FRAME_W, height=FRAME_H)
    left, top, right, bottom = _rect_of(children)
    vertical = [
        c for c in children
        if isinstance(c, m.ConnectorElement) and c.transform.extent_width == 0
    ]
    horizontal = [
        c for c in children
        if isinstance(c, m.ConnectorElement) and c.transform.extent_height == 0
    ]
    # The bottom axis said `max`, so the *horizontal* line went to the plot's top edge.
    axis = min(horizontal, key=lambda c: c.transform.offset_y)
    assert _pt(axis.transform.offset_y) == pytest.approx(top, abs=0.01)
    # The vertical line stayed where the y axis' own `autoZero` puts it.
    assert _pt(
        min(vertical, key=lambda c: c.transform.offset_x).transform.offset_x
    ) == pytest.approx(left, abs=0.01)


def test_a_scatter_honours_crosses_at():
    """`c:crossesAt` is a value on the axis being crossed.

    `_category_axis_position` honours it and the scatter's own crossing used to drop it
    silently.  Unmeasured -- no probe states it -- but implemented from the same schema
    reading its sibling uses rather than from a second one.
    """
    xml = scatter_chart_xml().replace(
        "<c:crossAx val='100003'/><c:crosses val='autoZero'/>",
        "<c:crossAx val='100003'/><c:crosses val='val'/><c:crossesAt val='4'/>",
        1,
    )
    children, _ = _build(xml, width=FRAME_W, height=FRAME_H)
    left, top, right, bottom = _rect_of(children)
    crossing = bottom - (bottom - top) * 4 / 7
    # A gridline and the axis line are both black rules across the whole plot, so the axis
    # cannot be told apart by geometry -- but its tick labels hang off it, and `nextTo`
    # means next to the axis.  They move up with it.
    assert any(
        abs(_pt(c.transform.offset_y) - crossing) < 0.2
        for c in children
        if isinstance(c, m.ConnectorElement) and c.transform.extent_height == 0
    )
    zero = min(
        (
            shape
            for shape in children
            if isinstance(shape, m.ShapeElement)
            and shape.text_body is not None
            and "".join(r.text for p in shape.text_body.paragraphs for r in p.runs) == "0"
            # The y axis has a "0" too, right-aligned in a column that starts at the
            # frame's own left edge; the x axis' is centred on its tick.
            and shape.transform.offset_x > 1.0
        ),
        key=lambda shape: shape.transform.offset_x,
    )
    assert _baseline(zero) == pytest.approx(crossing + 15.17, abs=0.6)


def test_a_scatter_data_label_can_print_its_x_value_as_the_category_name():
    """`parse/chart` already caches a scatter's `c:xVal` as its categories; this dropped it.

    `<c:showCatName val="1"/>` printed nothing at all, because the label text was built
    against an empty category list.
    """
    children, _ = _build(
        scatter_chart_xml(dlbls=(scatter_dlbls("CatName"),)),
        width=FRAME_W,
        height=FRAME_H,
    )
    _, _, _, bottom = _rect_of(children)
    inside = {
        "".join(r.text for p in shape.text_body.paragraphs for r in p.runs)
        for shape in children
        if isinstance(shape, m.ShapeElement)
        and shape.text_body is not None
        and _pt(shape.transform.offset_x) > 30.0
        and _pt(shape.transform.offset_y) < bottom
    }
    assert {"1", "2", "3", "4", "5"} <= inside


def test_a_label_below_its_point_stacks_downward():
    """`b` is the one placement that hangs *under* its anchor, so its block grows down.

    Every other placement stacks upward, which for `b` put the first line back on the
    marker it was supposed to clear.
    """
    children, _ = _build(
        scatter_chart_xml(dlbls=(scatter_dlbls("SerName", "Val", position="b"),)),
        width=FRAME_W,
        height=FRAME_H,
    )
    def leftmost(text):
        return min(
            (
                shape
                for shape in children
                if isinstance(shape, m.ShapeElement)
                and shape.text_body is not None
                and "".join(r.text for p in shape.text_body.paragraphs for r in p.runs)
                == text
                and _pt(shape.transform.offset_x) > 30.0
            ),
            key=lambda shape: shape.transform.offset_x,
        )

    # Point one is (52.027, 93.980); "Alpha" is the first line and "3" the second, so the
    # block starts where a one-line `b` label would and grows *away* from the marker.
    name, value = leftmost("Alpha"), leftmost("3")
    assert _baseline(name) == pytest.approx(110.522, abs=1.0)
    assert _baseline(value) == pytest.approx(_baseline(name) + 12.207, abs=0.1)
    assert _baseline(name) > 93.980 + 3


def test_area_and_scatter_draw_rather_than_warning():
    """The gate in `resolve/view.py`, which is what turns a warning into a picture."""
    from pptx2svg.resolve.view import DRAWABLE_CHART_KINDS

    assert {"areaChart", "scatterChart"} <= DRAWABLE_CHART_KINDS
    # `area3DChart` degrades to `areaChart` through `flat_chart_kind`, so it draws flat.
    assert flat_chart_kind("area3DChart") == "areaChart"



# -- bubbleChart -----------------------------------------------------------------------
#
# Thirty probe charts across three decks in the same 220.4724 x 181.1024 pt frame as the
# scatter sweep, exported by PowerPoint 16.106 and read back as exact path vertices.  Every
# drawn circle was axis-aligned and square to 0.001 pt, so its bounding box *is* its
# diameter.  The tables below are PowerPoint's own numbers.

BUBBLE_XS = (1, 2, 3)
BUBBLE_YS = (3, 4, 5)


def bubble_chart_xml(
    *,
    series=((BUBBLE_XS, BUBBLE_YS, (1, 4, 9)),),
    names=("Alpha", "Beta"),
    scale=None,
    represents=None,
    show_neg=None,
    legend=None,
    title=None,
    dlbls=("",),
    text_size=None,
    fixed=True,
):
    """One bubble chart, in the shape the exported probe deck used."""

    def cache(values):
        points = "".join(
            f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>"
            for i, v in enumerate(values)
            if v is not None
        )
        return (
            "<c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
            f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef>"
        )

    body = ""
    for index, (xs, ys, sizes) in enumerate(series):
        body += (
            f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
            "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
            f"<c:pt idx='0'><c:v>{names[index]}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
            + (dlbls[index] if index < len(dlbls) else "")
            + f"<c:xVal>{cache(xs)}</c:xVal><c:yVal>{cache(ys)}</c:yVal>"
            + f"<c:bubbleSize>{cache(sizes)}</c:bubbleSize></c:ser>"
        )

    tail = ""
    if scale is not None:
        tail += f"<c:bubbleScale val='{scale}'/>"
    if show_neg is not None:
        tail += f"<c:showNegBubbles val='{show_neg}'/>"
    if represents is not None:
        tail += f"<c:sizeRepresents val='{represents}'/>"

    def axis(axis_id, cross_id, position, low, high, gridlines):
        scaling = "<c:orientation val='minMax'/>"
        if fixed:
            scaling += f"<c:max val='{high}'/><c:min val='{low}'/>"
        return (
            f"<c:valAx><c:axId val='{axis_id}'/><c:scaling>{scaling}</c:scaling>"
            f"<c:delete val='0'/><c:axPos val='{position}'/>"
            + ("<c:majorGridlines/>" if gridlines else "")
            + "<c:numFmt formatCode='General' sourceLinked='1'/>"
            "<c:majorTickMark val='none'/><c:minorTickMark val='none'/>"
            "<c:tickLblPos val='nextTo'/>"
            f"<c:crossAx val='{cross_id}'/><c:crosses val='autoZero'/>"
            "<c:crossBetween val='midCat'/></c:valAx>"
        )

    legend_xml = (
        f"<c:legend><c:legendPos val='{legend}'/><c:layout/><c:overlay val='0'/></c:legend>"
        if legend
        else ""
    )
    title_xml = (
        "<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>"
        f"<a:rPr lang='en-US'/><a:t>{title}</a:t></a:r></a:p></c:rich></c:tx>"
        "<c:layout/><c:overlay val='0'/></c:title><c:autoTitleDeleted val='0'/>"
        if title
        else "<c:autoTitleDeleted val='1'/>"
    )
    tx_pr = (
        "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>"
        f"<a:defRPr sz='{int(text_size * 100)}'/></a:pPr></a:p></c:txPr>"
        if text_size
        else ""
    )
    return (
        f"<c:chart>{title_xml}<c:plotArea><c:layout/>"
        f"<c:bubbleChart><c:varyColors val='0'/>{body}{tail}"
        "<c:axId val='100002'/><c:axId val='100003'/></c:bubbleChart>"
        + axis("100002", "100003", "b", 0, 4, False)
        + axis("100003", "100002", "l", 0, 6, True)
        + f"</c:plotArea>{legend_xml}"
        f"<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>{tx_pr}"
    )


def _discs(children):
    """Every drawn bubble, smallest first: ``(diameter, centre x, centre y, fill)``."""
    out = []
    for child in children:
        if not isinstance(child, m.ShapeElement):
            continue
        if not isinstance(child.geometry, m.PresetGeometry):
            continue
        if child.geometry.preset != "ellipse":
            continue
        t = child.transform
        out.append((
            _pt(t.extent_width),
            _pt(t.offset_x) + _pt(t.extent_width) / 2,
            _pt(t.offset_y) + _pt(t.extent_height) / 2,
            child.fill,
        ))
    return sorted(out, key=lambda disc: disc[0])


#: ``c:bubbleScale`` against the diameter PowerPoint drew, on an identical chart whose
#: sizing region is the 181.1024 pt frame less 5 pt a side.  A *linear* reading is refuted
#: at both ends: it predicts 19.74 at 50 where PowerPoint drew 22.318, and 78.97 at 200
#: where it drew 64.163.
BUBBLE_SCALE_SWEEP = {
    1: 0.512,
    10: 4.983,
    25: 11.937,
    50: 22.318,
    75: 31.427,
    100: 39.485,
    150: 53.101,
    200: 64.163,
    300: 81.048,
}


@pytest.mark.parametrize("scale", list(BUBBLE_SCALE_SWEEP))
def test_bubble_scale_is_a_soft_clamp_not_a_multiplier(scale):
    kwargs = {} if scale == 100 else {"scale": scale}
    children, data = _build(bubble_chart_xml(**kwargs), width=FRAME_W, height=FRAME_H)
    assert data.kind == "bubbleChart"
    largest = _discs(children)[-1][0]
    assert largest == pytest.approx(BUBBLE_SCALE_SWEEP[scale], abs=0.01)


def test_a_bubbles_area_is_proportional_to_its_size():
    """Sizes 1, 4, 9 drew 13.162, 26.323 and 39.485 -- exactly 1:2:3."""
    children, _ = _build(bubble_chart_xml(), width=FRAME_W, height=FRAME_H)
    diameters = [disc[0] for disc in _discs(children)]
    assert diameters == pytest.approx([13.162, 26.323, 39.485], abs=0.01)


def test_size_represents_w_makes_the_diameter_proportional_instead():
    """The same sizes with ``<c:sizeRepresents val="w"/>`` drew 4.387, 17.549, 39.485."""
    children, _ = _build(bubble_chart_xml(represents="w"), width=FRAME_W, height=FRAME_H)
    diameters = [disc[0] for disc in _discs(children)]
    assert diameters == pytest.approx([4.387, 17.549, 39.485], abs=0.01)


def test_the_largest_bubble_is_the_reference_and_it_is_global():
    """Four size distributions all drew 39.485 for their biggest, so the reference is the
    maximum and not the sum; and a two-series probe shares one maximum across both -- a
    series topping out at 9 drew 27.920 beside one topping out at 18."""
    for sizes in ((1, 2, 3), (1, 4, 9), (5, 5, 5), (1, 2, 100)):
        children, _ = _build(
            bubble_chart_xml(series=((BUBBLE_XS, BUBBLE_YS, sizes),)),
            width=FRAME_W,
            height=FRAME_H,
        )
        assert _discs(children)[-1][0] == pytest.approx(39.485, abs=0.01)

    children, _ = _build(
        bubble_chart_xml(
            series=(
                (BUBBLE_XS, BUBBLE_YS, (1, 4, 9)),
                (BUBBLE_XS, (2, 3, 4), (2, 8, 18)),
            )
        ),
        width=FRAME_W,
        height=FRAME_H,
    )
    diameters = [disc[0] for disc in _discs(children)]
    assert diameters == pytest.approx(
        [9.307, 13.162, 18.613, 26.323, 27.920, 39.485], abs=0.02
    )


def test_a_zero_size_draws_nothing_and_a_negative_one_draws_hollow():
    """Sizes -4, 0, 9 drew **two** circles: the 9 in the series colour and the -4 at its
    magnitude, white with a black 0.75 pt outline -- a negative bar's drawing.
    ``<c:showNegBubbles val="0"/>`` removes it; the element absent or 1 keeps it."""
    negative = ((BUBBLE_XS, BUBBLE_YS, (-4, 0, 9)),)
    for kwargs in ({}, {"show_neg": 1}):
        children, _ = _build(
            bubble_chart_xml(series=negative, **kwargs), width=FRAME_W, height=FRAME_H
        )
        discs = _discs(children)
        assert [disc[0] for disc in discs] == pytest.approx([26.323, 39.485], abs=0.01)
        assert isinstance(discs[0][3], m.SolidFill)
        assert discs[0][3].color.hex.upper() == "#FFFFFF"

    children, _ = _build(
        bubble_chart_xml(series=negative, show_neg=0), width=FRAME_W, height=FRAME_H
    )
    assert [disc[0] for disc in _discs(children)] == pytest.approx([39.485], abs=0.01)


def test_a_bubbles_plot_rectangle_is_a_scatters():
    """``bare`` in the scatter sweep is 21.073 / 13.670 / 11.102 / 24.965, and a bubble
    with automatic axes over the same data reproduces it."""
    children, _ = _build(
        bubble_chart_xml(series=((BUBBLE_XS, BUBBLE_YS, (1, 2, 3)),), fixed=False),
        width=FRAME_W,
        height=FRAME_H,
    )
    ours = _insets(children)
    for edge, truth in (
        ("left", 21.073), ("right", 13.670), ("top", 11.102), ("bottom", 24.965)
    ):
        assert ours[edge] == pytest.approx(truth, abs=SWEEP_TOLERANCE_PT)


def test_the_region_a_bubble_is_sized_against_is_not_the_plot():
    """Two probes move every plot edge and draw the **same** bubble, and two move only the
    legend band and change it.  ``font14`` and ``font8`` both drew 39.485; a right legend
    drew 37.511 and a bottom one 33.928.

    The title probe is measured too -- 32.631 pt, which the full pipeline reproduces to
    0.001 -- and is **not** asserted here: this harness stubs text resolution, so its title
    font is not the 18 pt fallback PowerPoint gave an unstyled ``c:title`` and its band is
    2.2 pt out.  That is a property of the stub, not of the rule.
    """
    for kwargs, expected in (
        ({"text_size": 14}, 39.485),
        ({"text_size": 8}, 39.485),
        ({"legend": "r"}, 37.511),
        ({"legend": "b"}, 33.928),
    ):
        children, _ = _build(bubble_chart_xml(**kwargs), width=FRAME_W, height=FRAME_H)
        assert _discs(children)[-1][0] == pytest.approx(expected, abs=0.15), kwargs


def test_a_bubble_legends_with_a_swatch_not_a_rule():
    """The one place a bubble's legend parts company with the scatter it copies.  Both
    legend probes drew a 5.492 pt key, and the reserve feeds the sizing region too: with
    the line key the drawn bubble came out 3.1 pt small and the plot 13.4 pt narrow."""
    from pptx2svg.resolve.chart import (
        ChartBuilder,
        ChartFont,
        LEGEND_SWATCH_EM,
        font_box,
    )

    class _FakeBubble:
        _is_line = False
        _is_scatter = True
        _is_bubble = True
        _is_radar = False
        _radar_style = "marker"

    font = ChartFont(family="Aptos", box=font_box("Aptos", 10.0))
    key, _ = ChartBuilder._legend_key_size(_FakeBubble(), font)
    assert key == pytest.approx(LEGEND_SWATCH_EM * 10.0, abs=0.01)


def test_a_bubble_draws_no_line_and_no_marker():
    """Every one of the thirty probes emitted bare discs -- no connecting stroke, no
    marker, and no outline unless the series states one."""
    children, _ = _build(bubble_chart_xml(), width=FRAME_W, height=FRAME_H)
    assert _polyline_points(children) == []
    presets = [
        child.geometry.preset
        for child in children
        if isinstance(child, m.ShapeElement)
        and isinstance(child.geometry, m.PresetGeometry)
        and child.text_body is None
    ]
    assert presets.count("ellipse") == 3
    assert set(presets) <= {"ellipse"}


def test_a_bubbles_label_stands_further_off_than_a_markers():
    """``r`` puts the label box's near edge 8.494 pt past the **disc's** edge at 10 pt, on
    three bubbles of radius 6.581, 13.162 and 19.742 -- not the scatter's 6.0 off a 3 pt
    marker."""
    children, _ = _build(
        bubble_chart_xml(dlbls=(scatter_dlbls("Val", position="r"),)),
        width=FRAME_W,
        height=FRAME_H,
    )
    # The axis labels are text too, so keep only what sits beside a bubble: the y labels
    # are at x < 20 and the x labels below y = 160.
    boxes = [
        _pt(shape.transform.offset_x)
        for shape in children
        if isinstance(shape, m.ShapeElement)
        and shape.text_body is not None
        and _pt(shape.transform.offset_x) > 40.0
        and _pt(shape.transform.offset_y) < 150.0
    ]
    assert max(boxes) == pytest.approx(160.370 + 39.485 / 2 + 8.494, abs=0.4)


# -- ofPieChart ------------------------------------------------------------------------
#
# Twenty-four probe charts across two decks, same frame, exported by PowerPoint 16.106.
# The numbers below are read off its own path vertices: a wedge closes through the pie's
# centre, which is the second-to-last point of the emitted path, so the radius and both
# edge angles come out exactly.

OF_PIE_VALUES = (40, 25, 15, 10, 6, 4)


def of_pie_chart_xml(
    *,
    values=OF_PIE_VALUES,
    of_pie_type="pie",
    split_type=None,
    split_pos=None,
    cust_split=None,
    second_size=None,
    gap_width=None,
    ser_lines=None,
    legend=None,
    dlbls="",
):
    """One ofPie chart, in the shape the exported probe deck used."""
    cats = "".join(
        f"<c:pt idx='{i}'><c:v>Cat{i}</c:v></c:pt>" for i in range(len(values))
    )
    points = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(values))
    series = (
        "<c:ser><c:idx val='0'/><c:order val='0'/>"
        "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
        "<c:pt idx='0'><c:v>Share</c:v></c:pt></c:strCache></c:strRef></c:tx>"
        f"<c:cat><c:strRef><c:strCache><c:ptCount val='{len(values)}'/>{cats}"
        "</c:strCache></c:strRef></c:cat>"
        "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
        f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef></c:val></c:ser>"
    )
    tail = "" if gap_width is None else f"<c:gapWidth val='{gap_width}'/>"
    if split_type is not None:
        tail += f"<c:splitType val='{split_type}'/>"
    if split_pos is not None:
        tail += f"<c:splitPos val='{split_pos}'/>"
    if cust_split is not None:
        body = "".join(f"<c:secondPiePt val='{i}'/>" for i in cust_split)
        tail += f"<c:custSplit>{body}</c:custSplit>"
    if second_size is not None:
        tail += f"<c:secondPieSize val='{second_size}'/>"
    if ser_lines is not None:
        tail += f"<c:serLines>{ser_lines}</c:serLines>"
    legend_xml = (
        f"<c:legend><c:legendPos val='{legend}'/><c:layout/><c:overlay val='0'/></c:legend>"
        if legend
        else ""
    )
    return (
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        f"<c:ofPieChart><c:ofPieType val='{of_pie_type}'/>{series}{dlbls}{tail}"
        f"</c:ofPieChart></c:plotArea>{legend_xml}"
        "<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
    )


def _wedges(children):
    """Every drawn wedge as ``(centre, radius, start, sweep)``, in draw order.

    The path is ``M`` at the arc's start, cubics round it, then ``L`` to the centre; the
    centre is therefore the last emitted point, which is what makes the two plots
    separable without knowing the layout rule.
    """
    out = []
    for shape, points in _custom_paths(children):
        if len(points) < 4:
            continue
        centre = points[-1]
        radius = math.hypot(points[0][0] - centre[0], points[0][1] - centre[1])
        end = points[-2]

        def clockwise(point):
            return math.degrees(
                math.atan2(point[0] - centre[0], centre[1] - point[1])
            ) % 360

        start = clockwise(points[0])
        sweep = (clockwise(end) - start) % 360
        out.append((centre, radius, start, sweep or 360.0, shape))
    return out


def _plots(children):
    """The wedges grouped by the centre they share, left plot first."""
    groups = {}
    for centre, radius, start, sweep, shape in _wedges(children):
        groups.setdefault((round(centre[0], 1), round(centre[1], 1)), []).append(
            (radius, start, sweep, shape)
        )
    return [groups[key] for key in sorted(groups)]


#: How many points ``auto`` moves to the second plot, at five point counts.  ``round(n/3)``
#: is refuted twice: it predicts 1 at n=4 and 2 at n=7.
OF_PIE_AUTO_SPLIT = {3: 1, 4: 2, 6: 2, 7: 3, 8: 3}


@pytest.mark.parametrize("count", list(OF_PIE_AUTO_SPLIT))
def test_an_of_pie_with_no_split_type_moves_the_last_third(count):
    values = tuple(range(count, 0, -1))
    children, data = _build(
        of_pie_chart_xml(values=values), width=FRAME_W, height=FRAME_H
    )
    assert data.kind == "ofPieChart"
    moved = OF_PIE_AUTO_SPLIT[count]
    first, second = _plots(children)
    # The first plot keeps the rest and gains one aggregated slice.
    assert len(first) == count - moved + 1
    assert len(second) == moved


def test_the_four_stated_split_rules():
    """40/25/15/10/6/4, one probe each.

    * ``pos`` moves the **last** ``c:splitPos`` points;
    * ``val`` moves every point **below** ``c:splitPos`` -- 12 moved 10, 6 and 4;
    * ``percent`` is the same test on the share, and it is strict: 15 kept the 15%;
    * ``cust`` moves exactly the listed indices.
    """
    for kwargs, kept, moved in (
        ({"split_type": "pos", "split_pos": 4}, 2, 4),
        ({"split_type": "val", "split_pos": 12}, 3, 3),
        ({"split_type": "percent", "split_pos": 15}, 3, 3),
        ({"split_type": "cust", "cust_split": (0, 3)}, 4, 2),
    ):
        children, _ = _build(of_pie_chart_xml(**kwargs), width=FRAME_W, height=FRAME_H)
        first, second = _plots(children)
        assert (len(first), len(second)) == (kept + 1, moved), kwargs


def test_the_aggregated_slice_is_centred_at_three_oclock():
    """Five probes, every one of them centred on 90 degrees clockwise from twelve: the
    slice runs 72..108 here, and 27..153, 54..126, 0..180 and 54..126 on the others."""
    children, _ = _build(of_pie_chart_xml(), width=FRAME_W, height=FRAME_H)
    first, second = _plots(children)
    # The aggregated slice is the last one drawn on the first plot.
    _, start, sweep, _ = first[-1]
    assert (start + sweep / 2) % 360 == pytest.approx(90.0, abs=0.01)
    assert start == pytest.approx(72.0, abs=0.01)
    assert sweep == pytest.approx(36.0, abs=0.01)
    # Both plots start at the same angle.
    assert first[0][1] == pytest.approx(108.0, abs=0.01)
    assert second[0][1] == pytest.approx(108.0, abs=0.01)


#: ``(kwargs, first radius, first centre x, second radius, second centre x)`` -- the
#: packing law, read off six probes.  ``r = W / (2 + 2s + g/100)`` on a 198.472 pt region.
OF_PIE_LAYOUT = {
    "default": ({}, 44.105, 55.105, 33.079, 176.394),
    "size50": ({"second_size": 50}, 49.618, 60.618, 24.809, 184.663),
    "size100": ({"second_size": 100}, 39.694, 50.694, 39.694, 169.778),
    "size25": ({"second_size": 25}, 56.706, 67.706, 14.177, 195.295),
    "gap300": ({"gap_width": 300}, 30.534, 41.534, 22.901, 186.571),
    "gap0": ({"gap_width": 0}, 56.706, 67.706, 42.530, 166.942),
}


@pytest.mark.parametrize("name", list(OF_PIE_LAYOUT))
def test_the_two_plots_are_packed_across_the_polar_region(name):
    kwargs, radius, centre_x, second_radius, second_x = OF_PIE_LAYOUT[name]
    children, _ = _build(of_pie_chart_xml(**kwargs), width=FRAME_W, height=FRAME_H)
    first, second = _plots(children)
    assert first[0][0] == pytest.approx(radius, abs=0.02)
    assert second[0][0] == pytest.approx(second_radius, abs=0.02)
    centres = sorted({round(_wedge[0][0], 1) for _wedge in _wedges(children)})
    assert centres[0] == pytest.approx(centre_x, abs=0.05)
    assert centres[-1] == pytest.approx(second_x, abs=0.05)


def test_the_bar_form_packs_by_a_different_divisor():
    """``r = W / (2 + s + g/200)``: 61.068 at the default 75% and 66.157 at 50%, and the
    bar itself is ``s*r`` wide by ``2*s*r`` tall with the first moved point on top."""
    for kwargs, radius, width, height in (
        ({}, 61.068, 45.801, 91.602),
        ({"second_size": 50}, 66.157, 33.078, 66.157),
    ):
        children, _ = _build(
            of_pie_chart_xml(of_pie_type="bar", **kwargs), width=FRAME_W, height=FRAME_H
        )
        first = _plots(children)[0]
        assert first[0][0] == pytest.approx(radius, abs=0.02), kwargs
        bars = sorted(
            (
                child
                for child in children
                if isinstance(child, m.ShapeElement)
                and isinstance(child.geometry, m.PresetGeometry)
                and child.geometry.preset == "rect"
            ),
            key=lambda shape: shape.transform.offset_y,
        )
        assert len(bars) == 2
        assert _pt(bars[0].transform.extent_width) == pytest.approx(width, abs=0.02)
        total = sum(_pt(bar.transform.extent_height) for bar in bars)
        assert total == pytest.approx(height, abs=0.02)
        # 6 and 4 of 10: the 60% segment sits above the 40% one.
        assert _pt(bars[0].transform.extent_height) == pytest.approx(
            height * 0.6, abs=0.02
        )
        assert _pt(bars[0].transform.offset_x) + width == pytest.approx(
            FRAME_W - 11.0, abs=0.05
        )


def test_ser_lines_are_tangent_to_the_second_pie():
    """Measured: each connector leaves a **corner of the aggregated slice** and touches
    the second pie.  The upper one ran (97.051, 76.922) to (168.104, 58.528), where the
    radius and the line are perpendicular to 0.000 and the length is sqrt(d^2 - r^2).

    Presence is the switch -- a probe with no ``c:serLines`` drew none -- and the default
    is black at 0.5 pt, the axis default.
    """
    children, _ = _build(of_pie_chart_xml(), width=FRAME_W, height=FRAME_H)
    assert _lines(children) == []

    children, _ = _build(of_pie_chart_xml(ser_lines=""), width=FRAME_W, height=FRAME_H)
    lines = _lines(children)
    assert len(lines) == 2
    boxes = sorted(
        (
            (
                _pt(line.transform.offset_x),
                _pt(line.transform.offset_y),
                _pt(line.transform.offset_x + line.transform.extent_width),
                _pt(line.transform.offset_y + line.transform.extent_height),
            )
            for line in lines
        ),
        key=lambda box: box[1],
    )
    assert boxes[0] == pytest.approx((97.051, 58.528, 168.104, 76.922), abs=0.05)
    assert boxes[1] == pytest.approx((97.051, 104.180, 168.104, 122.574), abs=0.05)


def test_past_six_colours_the_accent_cycle_is_shaded_then_tinted():
    """Found on the ofPie probes, which always need one colour more than they have points.

    Seven slices came back as accent1..accent6 *darkened* and a light accent1; nine
    repeated the same darkened six and then three light ones, so the variation is per cycle
    of six and not a function of the count.  Both factors are exact against 27 measured
    channels; the round numbers either side are off by up to 1 and 5.
    """
    from pptx2svg.resolve.chart import _cycle_shift

    accents = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"]
    shaded = ["#3b64ad", "#d26e2a", "#929292", "#e2aa00", "#5089bc", "#62993e"]
    tinted = ["#8fa2d4", "#f1a78a", "#bfbfbf"]
    assert [_cycle_shift(value, 0) for value in accents] == shaded
    assert [_cycle_shift(value, 1) for value in accents[:3]] == tinted


def test_an_of_pie_that_moves_nothing_or_everything_still_draws():
    """``val`` with a threshold under the smallest value leaves a plain pie, and one over
    the largest leaves the first plot a single whole-circle slice.  PowerPoint draws a dark
    filled disc where the empty second plot would be; this draws nothing there, which is
    the divergence and is recorded rather than reproduced."""
    children, _ = _build(
        of_pie_chart_xml(split_type="val", split_pos=1), width=FRAME_W, height=FRAME_H
    )
    plots = _plots(children)
    assert len(plots) == 1
    assert len(plots[0]) == len(OF_PIE_VALUES)

    children, _ = _build(
        of_pie_chart_xml(split_type="val", split_pos=100), width=FRAME_W, height=FRAME_H
    )
    first, second = _plots(children)
    assert len(first) == 1
    assert first[0][2] == pytest.approx(360.0, abs=0.01)
    assert len(second) == len(OF_PIE_VALUES)


# -- stockChart ------------------------------------------------------------------------
#
# Twelve probe charts.  The plot rectangle, the axis, the bands and the legend key came
# back a line chart's in every one.

STOCK_CATS = ("Mon", "Tue", "Wed", "Thu", "Fri")
STOCK_OPEN = (10, 12, 11, 14, 13)
STOCK_HIGH = (15, 16, 14, 18, 17)
STOCK_LOW = (8, 9, 10, 12, 11)
STOCK_CLOSE = (12, 11, 13, 13, 16)
STOCK_OHLC = (
    ("Open", STOCK_OPEN), ("High", STOCK_HIGH), ("Low", STOCK_LOW), ("Close", STOCK_CLOSE)
)
STOCK_HLC = (("High", STOCK_HIGH), ("Low", STOCK_LOW), ("Close", STOCK_CLOSE))


def stock_chart_xml(
    *,
    series=STOCK_OHLC,
    hi_low=None,
    up_down=None,
    up_down_gap=None,
    up_fill=None,
    down_fill=None,
    legend=None,
):
    """One stock chart, in the shape the exported probe deck used."""
    cats = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(STOCK_CATS))
    body = ""
    for index, (name, values) in enumerate(series):
        points = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(values))
        body += (
            f"<c:ser><c:idx val='{index}'/><c:order val='{index}'/>"
            "<c:tx><c:strRef><c:strCache><c:ptCount val='1'/>"
            f"<c:pt idx='0'><c:v>{name}</c:v></c:pt></c:strCache></c:strRef></c:tx>"
            f"<c:cat><c:strRef><c:strCache><c:ptCount val='{len(STOCK_CATS)}'/>{cats}"
            "</c:strCache></c:strRef></c:cat>"
            "<c:val><c:numRef><c:numCache><c:formatCode>General</c:formatCode>"
            f"<c:ptCount val='{len(values)}'/>{points}</c:numCache></c:numRef></c:val></c:ser>"
        )
    tail = "" if hi_low is None else f"<c:hiLowLines>{hi_low}</c:hiLowLines>"
    if up_down is not None:
        inner = "" if up_down_gap is None else f"<c:gapWidth val='{up_down_gap}'/>"
        if up_fill is not None:
            inner += f"<c:upBars>{up_fill}</c:upBars>"
        if down_fill is not None:
            inner += f"<c:downBars>{down_fill}</c:downBars>"
        tail += f"<c:upDownBars>{inner}</c:upDownBars>"
    legend_xml = (
        f"<c:legend><c:legendPos val='{legend}'/><c:layout/><c:overlay val='0'/></c:legend>"
        if legend
        else ""
    )
    return (
        "<c:chart><c:autoTitleDeleted val='1'/><c:plotArea><c:layout/>"
        f"<c:stockChart>{body}{tail}"
        "<c:axId val='100002'/><c:axId val='100003'/></c:stockChart>"
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
        f"</c:plotArea>{legend_xml}"
        "<c:plotVisOnly val='1'/><c:dispBlanksAs val='gap'/></c:chart>"
    )


def _up_down_bars(children):
    """The open-to-close bodies, left to right, as ``(box, fill hex)``."""
    out = []
    for child in children:
        if not isinstance(child, m.ShapeElement) or child.text_body is not None:
            continue
        if not isinstance(child.geometry, m.PresetGeometry):
            continue
        if child.geometry.preset != "rect" or not isinstance(child.fill, m.SolidFill):
            continue
        t = child.transform
        # A square marker is a 6 pt rect with a solid fill too; the narrowest bar these
        # probes draw is 9.154 pt, at `gapWidth=300`.
        if _pt(t.extent_width) < 8.0:
            continue
        out.append((
            (_pt(t.offset_x), _pt(t.offset_y), _pt(t.extent_width), _pt(t.extent_height)),
            child.fill.color.hex.upper(),
        ))
    return sorted(out)


def test_a_stock_chart_with_no_decorations_is_a_line_chart():
    """Measured, and it is the opposite of what the schema's name suggests: PowerPoint
    drew one 1.5 pt polyline per series with the ordinary marker cycle on it.  What a real
    stock chart hides is hidden by the *file*, with `<a:ln><a:noFill/></a:ln>`."""
    children, data = _build(stock_chart_xml(), width=FRAME_W, height=FRAME_H)
    assert data.kind == "stockChart"
    assert len(_polyline_points(children)) == 4
    assert _up_down_bars(children) == []


def test_hi_low_lines_span_the_category_and_only_when_asked():
    """The probe's line runs from the largest value in the category to the smallest -- 15
    down to 8 in the first -- at the band centre, in black at 0.5 pt."""
    children, _ = _build(stock_chart_xml(), width=FRAME_W, height=FRAME_H)
    plot = _rect_of(children)
    before = len(_lines(children))

    children, _ = _build(stock_chart_xml(hi_low=""), width=FRAME_W, height=FRAME_H)
    added = [
        line
        for line in _lines(children)
        if line.transform.extent_width == 0
        and _pt(line.transform.offset_x) > plot[0] + 1.0
    ]
    assert len(added) == 5
    assert len(_lines(children)) == before + 5
    # 0..20 on this plot, so one unit is height/20; the first category spans 15 to 8.
    left, top, right, bottom = plot
    unit = (bottom - top) / 20.0
    first = min(added, key=lambda line: line.transform.offset_x)
    assert _pt(first.transform.offset_x) == pytest.approx(left + (right - left) / 10, abs=0.2)
    assert _pt(first.transform.offset_y) == pytest.approx(bottom - 15 * unit, abs=0.2)
    assert _pt(first.transform.extent_height) == pytest.approx(7 * unit, abs=0.2)


#: ``c:upDownBars/c:gapWidth`` against the width PowerPoint drew on a 36.612 pt band.
STOCK_GAP_SWEEP = {None: 14.646, 50: 24.410, 150: 14.646, 300: 9.154}


@pytest.mark.parametrize("gap", list(STOCK_GAP_SWEEP))
def test_up_down_bar_width_follows_the_gap_width(gap):
    children, _ = _build(
        stock_chart_xml(up_down="", up_down_gap=gap), width=FRAME_W, height=FRAME_H
    )
    bars = _up_down_bars(children)
    assert len(bars) == 5
    assert bars[0][0][2] == pytest.approx(STOCK_GAP_SWEEP[gap], abs=0.05)


def test_up_down_bars_take_the_first_and_last_series_whatever_they_are_called():
    """**The series order carries the meaning and the labels carry none.**  A three-series
    High/Low/Close chart with ``c:upDownBars`` drew all five bars *down*, from each
    category's High to its Close; four series drew three up and two down.  The defaults are
    #F9F9F9 and #3F3F3F, which are not theme accents."""
    children, _ = _build(stock_chart_xml(up_down=""), width=FRAME_W, height=FRAME_H)
    fills = [fill for _, fill in _up_down_bars(children)]
    assert fills == ["#F9F9F9", "#3F3F3F", "#F9F9F9", "#3F3F3F", "#F9F9F9"]

    children, _ = _build(
        stock_chart_xml(series=STOCK_HLC, up_down=""), width=FRAME_W, height=FRAME_H
    )
    bars = _up_down_bars(children)
    assert [fill for _, fill in bars] == ["#3F3F3F"] * 5
    left, top, right, bottom = _rect_of(children)
    unit = (bottom - top) / 20.0
    # The first bar runs High 15 down to Close 12.
    assert bars[0][0][1] == pytest.approx(bottom - 15 * unit, abs=0.2)
    assert bars[0][0][3] == pytest.approx(3 * unit, abs=0.2)


def test_up_and_down_bar_fills_come_from_the_file_when_it_states_them():
    red = "<c:spPr><a:solidFill><a:srgbClr val='FF0000'/></a:solidFill></c:spPr>"
    blue = "<c:spPr><a:solidFill><a:srgbClr val='0000FF'/></a:solidFill></c:spPr>"
    children, _ = _build(
        stock_chart_xml(up_down="", up_fill=red, down_fill=blue),
        width=FRAME_W,
        height=FRAME_H,
    )
    fills = {fill for _, fill in _up_down_bars(children)}
    assert fills == {"#FF0000", "#0000FF"}


def test_the_new_four_types_and_what_still_warns():
    """The gate in `resolve/view.py`, which is what turns a warning into a picture."""
    from pptx2svg.resolve.view import DRAWABLE_CHART_KINDS

    assert {"bubbleChart", "ofPieChart", "stockChart"} <= DRAWABLE_CHART_KINDS
    # `surfaceChart` is deliberately still out: PowerPoint draws it as a lit 3-D mesh --
    # both spellings, with and without `c:view3D` -- and none of that is built.  See the
    # roadmap.
    assert "surfaceChart" not in DRAWABLE_CHART_KINDS
    assert flat_chart_kind("surface3DChart") == "surfaceChart"


#: Numbers a file controls, at three frame sizes including a 1 x 1 pt one.  Every one of
#: these was run before the rule it exercises was written; the lesson the last review left
#: is that a measurement pins the frame and says nothing about what a hand-written number
#: does inside it.
DEGENERATE_CHARTS = {
    "bubble with no sizes": lambda: bubble_chart_xml(
        series=((BUBBLE_XS, BUBBLE_YS, (None, None, None)),)
    ),
    "bubble all zero": lambda: bubble_chart_xml(
        series=((BUBBLE_XS, BUBBLE_YS, (0, 0, 0)),)
    ),
    "bubble scale 0": lambda: bubble_chart_xml(scale=0),
    "of pie all zero": lambda: of_pie_chart_xml(values=(0, 0, 0)),
    "of pie gap 5000": lambda: of_pie_chart_xml(gap_width=5000, ser_lines=""),
    "of pie gap -500": lambda: of_pie_chart_xml(gap_width=-500, ser_lines=""),
    "of pie second size 0": lambda: of_pie_chart_xml(
        of_pie_type="bar", second_size=0, ser_lines=""
    ),
    "of pie custom split out of range": lambda: of_pie_chart_xml(
        split_type="cust", cust_split=(99, -3)
    ),
    "of pie split position past the end": lambda: of_pie_chart_xml(
        split_type="pos", split_pos=900
    ),
    "of pie unknown split type": lambda: of_pie_chart_xml(split_type="nonsense"),
    "stock with one series": lambda: stock_chart_xml(
        series=(("Only", STOCK_CLOSE),), hi_low="", up_down=""
    ),
    "stock gap -100": lambda: stock_chart_xml(up_down="", up_down_gap=-100),
}


@pytest.mark.parametrize("name", list(DEGENERATE_CHARTS))
@pytest.mark.parametrize("size", [(FRAME_W, FRAME_H), (1.0, 1.0), (4000.0, 3.0)])
def test_a_file_controlled_number_does_not_crash_the_new_types(name, size):
    children, data = _build(
        DEGENERATE_CHARTS[name](), width=size[0], height=size[1]
    )
    assert data.kind in ("bubbleChart", "ofPieChart", "stockChart")


def test_a_split_that_moves_nothing_leaves_no_second_plot_and_no_connector():
    """``cust`` with only out-of-range indices is the shape of an empty split, and the two
    connector lines have nothing to point at."""
    children, _ = _build(
        of_pie_chart_xml(split_type="cust", cust_split=(99,), ser_lines=""),
        width=FRAME_W,
        height=FRAME_H,
    )
    assert len(_plots(children)) == 1
    assert _lines(children) == []


def test_an_office_2016_chartex_frame_says_what_it_is(authoring):
    """A ``cx:chartSpace`` frame used to warn ``chart-unreadable``: "names no chart part".

    That is true of the `c:` relationship and false about the file -- the frame's own first
    child is *also* called ``chart``, so it fell through to the ordinary chart path with no
    relationship id.  The family (treemap, sunburst, histogram, box-and-whisker, waterfall,
    funnel, map) shares no markup with ``c:chartSpace``: different namespace, different
    data model, no ``c:*Chart`` group anywhere in it.  The picture was always right -- an
    empty positioned frame -- and only the diagnosis was wrong.
    """
    from tests.deckbuilder import derive_deck

    cx = "http://schemas.microsoft.com/office/drawing/2014/chartex"
    part = (
        "<?xml version='1.0'?>"
        f"<cx:chartSpace xmlns:cx='{cx}' "
        "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships'>"
        "<cx:chart><cx:plotArea><cx:plotAreaRegion>"
        "<cx:series layoutId='treemap'/>"
        "</cx:plotAreaRegion></cx:plotArea></cx:chart></cx:chartSpace>"
    ).encode()
    frame = (
        "<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id='91' name='Treemap'/>"
        "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr>"
        "<p:xfrm><a:off x='0' y='0'/><a:ext cx='2800000' cy='2300000'/></p:xfrm>"
        f"<a:graphic><a:graphicData uri='{cx}'>"
        f"<cx:chart xmlns:cx='{cx}' "
        "xmlns:r='http://schemas.openxmlformats.org/officeDocument/2006/relationships' "
        "r:id='rIdCx'/></a:graphicData></a:graphic></p:graphicFrame>"
    )
    deck = derive_deck(
        authoring,
        parts={"ppt/charts/chartEx1.xml": part},
        shapes_xml=frame,
        slide_relationships=[
            (
                "rIdCx",
                "http://schemas.microsoft.com/office/2014/relationships/chartEx",
                "../charts/chartEx1.xml",
            )
        ],
        overrides={"ppt/charts/chartEx1.xml": "application/vnd.ms-office.chartex+xml"},
    )
    options = ConvertOptions()
    model = convert_pptx_to_model(deck, options)
    warning = next(w for w in options.warnings if w.code == "chart-unsupported-type")
    assert "Office 2016" in warning.message
    assert not any(w.code == "chart-unreadable" for w in options.warnings)
    assert isinstance(model.slides[0].elements[-1], m.ShapeElement)


# --------------------------------------------------------------------------------------
# The chart gallery fixture
# --------------------------------------------------------------------------------------

#: Slide number -> the ``c:*Chart`` group element it leads with, in the order
#: `chart-gallery.pptx` puts them.  Written down here rather than read back out of the
#: deck so that a slide reordered in `tools/make_chart_gallery.py` fails this test instead
#: of silently renumbering what every other assertion about that deck means.
GALLERY_SLIDES = (
    "barChart",
    "barChart",
    "lineChart",
    "areaChart",
    "scatterChart",
    "bubbleChart",
    "pieChart",
    "doughnutChart",
    "ofPieChart",
    "radarChart",
    "stockChart",
    "surfaceChart",
    "bar3DChart",
    "line3DChart",
    "pie3DChart",
    "area3DChart",
    "barChart",  # the combo's first group; its `lineChart` is the one we drop
)


def test_the_chart_gallery_holds_one_of_every_group_element(chart_gallery):
    """The fixture's coverage claim, checked against the file rather than its README.

    Every ``c:*Chart`` element `parse/chart.CHART_GROUP_ELEMENTS` recognises has to appear
    in this deck, because that is the whole reason it exists: the chart renderer is the
    largest body of measured behaviour here and no other committed deck exercises more
    than four of its types.  A group element added to the reader without a slide here is
    a type nothing renders end to end.
    """
    from pptx2svg.parse.chart import CHART_GROUP_ELEMENTS

    with zipfile.ZipFile(chart_gallery) as archive:
        names = sorted(
            (name for name in archive.namelist() if re.fullmatch(r"ppt/charts/chart\d+\.xml", name)),
            key=lambda name: int(re.search(r"(\d+)", name).group(1)),
        )
        parts = [archive.read(name).decode("utf-8") for name in names]
    assert len(parts) == len(GALLERY_SLIDES)
    for number, (part, expected) in enumerate(zip(parts, GALLERY_SLIDES), start=1):
        groups = re.findall(r"<c:(\w+Chart)[ >]", part)
        assert groups[0] == expected, f"slide {number} leads with {groups[0]}"

    present = {group for part in parts for group in re.findall(r"<c:(\w+Chart)[ >]", part)}
    # `surface3DChart` is the one group element with no slide of its own: `surfaceChart`
    # already covers the deferral, and `flat_chart_kind` maps the two to one answer.
    assert set(CHART_GROUP_ELEMENTS) - present == {"surface3DChart"}


def test_only_the_gallerys_surface_slide_refuses_to_draw(chart_gallery):
    """Sixteen slides draw; the seventeenth says so and draws an empty frame.

    `surfaceChart` is measured and deliberately deferred, and this is the committed
    end-to-end evidence that the refusal is the *honest* one -- a warning naming the type,
    plus a positioned but undrawn frame -- rather than a blank slide nobody notices.  It
    is also the tripwire for the opposite mistake: if some future change starts drawing a
    surface, this test and `tests/vrt/chart-gallery/slide-12.svg` both fail, which is the
    right amount of noise for a chart type going from refused to drawn.
    """
    options = ConvertOptions()
    model = convert_pptx_to_model(chart_gallery, options)

    unsupported = [w for w in options.warnings if w.code == "chart-unsupported-type"]
    assert len(unsupported) == 1
    assert "surfaceChart" in unsupported[0].message
    assert not any(w.code == "chart-unreadable" for w in options.warnings)

    # An undrawn chart is a `ShapeElement` with neither fill nor outline, which is what
    # `_empty_graphic_frame` produces; every other slide ends in a real chart.
    frame = model.slides[11].elements[-1]
    assert isinstance(frame, m.ShapeElement)
    assert frame.fill is None and frame.outline is None
    for number, slide in enumerate(model.slides, start=1):
        if number == 12:
            continue
        assert isinstance(slide.elements[-1], m.ChartElement), f"slide {number}"
