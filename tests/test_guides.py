"""The DrawingML guide-formula language (ECMA-376 §20.1.9.11)."""

from __future__ import annotations

import math

import pytest

from pptx2svg.guides import (
    DEGREE,
    arc_endpoint,
    arc_segments,
    builtin_variables,
    evaluate_formula,
    evaluate_guides,
    resolve_value,
)


def variables(width=200.0, height=100.0):
    return builtin_variables(width, height)


# --------------------------------------------------------------------------------------
# Built-in names
# --------------------------------------------------------------------------------------


def test_the_basic_box_names():
    v = variables()
    assert (v["l"], v["t"], v["r"], v["b"]) == (0.0, 0.0, 200.0, 100.0)
    assert (v["hc"], v["vc"]) == (100.0, 50.0)
    assert (v["ss"], v["ls"]) == (100.0, 200.0)


@pytest.mark.parametrize("divisor", [2, 3, 4, 5, 6, 8, 10, 12, 16, 20, 24, 32])
def test_every_divisor_shorthand_exists(divisor):
    """A missing name resolves to zero and silently drops path commands, so the set
    must be complete -- `cloud` and `wave` reference wd16 and wd32."""
    v = variables()
    assert v[f"wd{divisor}"] == 200.0 / divisor
    assert v[f"hd{divisor}"] == 100.0 / divisor
    assert v[f"ssd{divisor}"] == 100.0 / divisor


def test_named_angles_are_in_sixtythousandths_of_a_degree():
    v = variables()
    assert v["cd4"] == 90 * DEGREE
    assert v["cd2"] == 180 * DEGREE
    assert v["3cd4"] == 270 * DEGREE
    assert v["cd8"] == 45 * DEGREE
    assert v["5cd8"] == 225 * DEGREE
    assert v["7cd8"] == 315 * DEGREE


def test_an_unknown_name_resolves_to_zero_rather_than_raising():
    assert resolve_value("nosuchguide", variables()) == 0.0
    assert resolve_value("-8333", variables()) == -8333.0


# --------------------------------------------------------------------------------------
# Operators
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "formula,expected",
    [
        ("val 18750", 18750),
        ("*/ w 1 2", 100),          # 200 * 1 / 2
        ("+- w 0 h", 100),          # 200 + 0 - 100
        ("+/ w h 2", 150),          # (200 + 100) / 2
        ("min w h", 100),
        ("max w h", 200),
        ("abs -5", 5),
        ("sqrt 144", 12),
        ("?: 1 w h", 200),          # positive -> second operand
        ("?: -1 w h", 100),         # non-positive -> third
        ("?: 0 w h", 100),          # zero counts as non-positive
    ],
)
def test_operators(formula, expected):
    assert evaluate_formula(formula, variables()) == expected


def test_pin_clamps_its_middle_operand():
    """`pin x y z` limits y to [x, z] -- the operand order is easy to get backwards."""
    v = variables()
    assert evaluate_formula("pin 0 500 100", v) == 100
    assert evaluate_formula("pin 0 -500 100", v) == 0
    assert evaluate_formula("pin 0 50 100", v) == 50


def test_trigonometric_operators_scale_their_first_operand():
    v = variables()
    assert evaluate_formula("sin w cd4", v) == 200        # 200 * sin(90°)
    assert evaluate_formula("cos w cd2", v) == -200       # 200 * cos(180°)
    assert evaluate_formula("tan h cd8", v) == 100        # 100 * tan(45°)


def test_at2_returns_an_angle_in_drawingml_units():
    assert evaluate_formula("at2 1 1", variables()) == round(45 * DEGREE)


def test_mod_is_the_three_dimensional_vector_length():
    assert evaluate_formula("mod 3 4 0", variables()) == 5


def test_an_unknown_operator_yields_zero_rather_than_raising():
    """Geometry is cosmetic; a shape drawn wrong beats a deck that will not convert."""
    assert evaluate_formula("frobnicate w h", variables()) == 0.0
    assert evaluate_formula("", variables()) == 0.0


# --------------------------------------------------------------------------------------
# Rounding
# --------------------------------------------------------------------------------------


def test_guides_are_integers_by_default():
    """ECMA-376 defines every guide as an integer; Office evaluates them in EMU."""
    assert evaluate_formula("*/ 100 1 3", variables()) == 33


def test_precise_mode_keeps_the_fraction():
    """Presets are evaluated in pixels, where rounding is a visible half-pixel error."""
    assert evaluate_formula("*/ 100 1 3", variables(), precise=True) == pytest.approx(100 / 3)


def test_precise_mode_propagates_through_a_guide_list():
    guides = [[("a", "*/ w 1 3")], [("b", "*/ a 1 1")]]
    rounded = evaluate_guides(guides, 200.0, 100.0)
    exact = evaluate_guides(guides, 200.0, 100.0, precise=True)
    assert rounded["b"] == 67
    assert exact["b"] == pytest.approx(200 / 3)


# --------------------------------------------------------------------------------------
# Guide lists
# --------------------------------------------------------------------------------------


def test_later_guides_see_earlier_ones():
    """Evaluation is in document order, not alphabetical: gdLst may use avLst."""
    result = evaluate_guides(
        [[("adj1", "val 25000")], [("dx", "*/ w adj1 100000"), ("x", "+- hc 0 dx")]],
        200.0,
        100.0,
    )
    assert result["dx"] == 50
    assert result["x"] == 50


def test_builtins_are_available_to_every_list():
    result = evaluate_guides([[("half", "*/ ss 1 2")]], 200.0, 100.0)
    assert result["half"] == 50


# --------------------------------------------------------------------------------------
# arcTo conversion
# --------------------------------------------------------------------------------------


def test_a_quarter_arc_ends_where_the_geometry_says():
    """Start at 3 o'clock on a circle of radius 10 centred at (0,0); sweep 90° down."""
    end_x, end_y, large, sweep = arc_endpoint(10.0, 0.0, 10.0, 10.0, 0.0, 90 * DEGREE)
    assert end_x == pytest.approx(0.0, abs=1e-9)
    assert end_y == pytest.approx(10.0)
    assert (large, sweep) == (0, 1)


def test_a_sweep_over_half_a_turn_is_cut_into_representable_pieces():
    """Rather than lean on SVG's large-arc flag, a long sweep is split into pieces of at
    most half a turn each -- which is what makes the full-circle case work at all."""
    segments = arc_segments(10.0, 0.0, 10.0, 10.0, 0.0, 270 * DEGREE)
    assert len(segments) == 2
    assert all(large == 0 for _, _, large, _ in segments)
    assert all(sweep == 1 for _, _, _, sweep in segments)
    # Centre is (0, 0); three quarters clockwise from 3 o'clock lands at 12 o'clock.
    end_x, end_y, _, _ = segments[-1]
    assert end_x == pytest.approx(0.0, abs=1e-9)
    assert end_y == pytest.approx(-10.0)


def test_a_negative_sweep_reverses_the_direction_flag():
    _, _, _, sweep = arc_endpoint(10.0, 0.0, 10.0, 10.0, 0.0, -90 * DEGREE)
    assert sweep == 0


def test_the_centre_is_reconstructed_from_the_start_angle():
    """DrawingML implies the centre; SVG needs the end point. Start at 12 o'clock."""
    end_x, end_y, _, _ = arc_endpoint(0.0, 0.0, 10.0, 10.0, 270 * DEGREE, 90 * DEGREE)
    # Centre is (0, 10); sweeping 90° clockwise lands at 3 o'clock: (10, 10).
    assert end_x == pytest.approx(10.0)
    assert end_y == pytest.approx(10.0, abs=1e-9)


def test_an_elliptical_arc_uses_both_radii():
    end_x, end_y, _, _ = arc_endpoint(20.0, 0.0, 20.0, 5.0, 0.0, 90 * DEGREE)
    assert end_x == pytest.approx(0.0, abs=1e-9)
    assert end_y == pytest.approx(5.0)


def test_angles_agree_with_the_degree_constant():
    assert DEGREE == 60000.0
    assert math.isclose(evaluate_formula("cos 1 0", variables(), precise=True), 1.0)


# --------------------------------------------------------------------------------------
# Geometric vs parametric angles
#
# DrawingML's stAng/swAng are *geometric* angles -- the direction of a ray from the
# ellipse centre.  The parametric angle that generates a point, (wR·cos t, hR·sin t), is
# a different number unless the ellipse is a circle.  Conflating them is silent for
# circles and grows with the aspect ratio, so these tests pin the elliptical case.
# --------------------------------------------------------------------------------------


def test_a_circular_arc_is_unaffected_by_the_conversion():
    """wR == hR means geometric and parametric angles coincide, which is why the bug
    this guards against stayed invisible in every square shape."""
    end_x, end_y, _, _ = arc_endpoint(10.0, 0.0, 10.0, 10.0, 0.0, 45 * DEGREE)
    assert end_x == pytest.approx(10.0 * math.cos(math.radians(45)))
    assert end_y == pytest.approx(10.0 * math.sin(math.radians(45)))


def test_an_elliptical_arc_reconstructs_the_centre_the_spec_intended():
    """Taken from `curvedUpArrow` at 200x100, where the spec names the answer.

    Its guides put the arc's start at (x7, y1) = (184.92, 25) with wR = 81.25, hR = h =
    100 and stAng3 = 17.629 degrees, about a centre at (x3, t) = (106.25, 0).  Reading
    stAng as a parametric angle instead puts the centre at (107.5, -5.3), and the band
    then fails to meet the arrowhead -- which is exactly how the bug was spotted.
    """
    start_x, start_y = 184.92, 25.0
    width_radius, height_radius = 81.25, 100.0
    start_angle = 17.629 * DEGREE

    # A zero-length sweep is refused, so sweep a little and work the centre back out.
    end_x, end_y, _, _ = arc_endpoint(
        start_x, start_y, width_radius, height_radius, start_angle, 0.001 * DEGREE
    )
    # The start point lies on the ellipse centred at (106.25, 0):
    assert ((start_x - 106.25) / width_radius) ** 2 + (
        (start_y - 0.0) / height_radius
    ) ** 2 == pytest.approx(1.0, abs=1e-3)
    # ... and so does the point an instant later.
    assert ((end_x - 106.25) / width_radius) ** 2 + (
        (end_y - 0.0) / height_radius
    ) ** 2 == pytest.approx(1.0, abs=1e-3)


def test_a_quarter_sweep_on_an_ellipse_lands_on_the_axis():
    """From the rightmost point of an ellipse, sweeping 90 degrees must reach the
    bottom -- (cx, cy + hR) -- whatever the radii."""
    end_x, end_y, _, _ = arc_endpoint(30.0, 0.0, 30.0, 5.0, 0.0, 90 * DEGREE)
    assert end_x == pytest.approx(0.0, abs=1e-9)   # centre is (0, 0)
    assert end_y == pytest.approx(5.0)


def test_the_conversion_preserves_the_revolution():
    """`atan2` returns (-pi, pi], so an arc crossing 12 o'clock would jump a whole turn
    and reverse itself if the revolution were not carried through."""
    segments = arc_segments(30.0, 0.0, 30.0, 5.0, 0.0, 300 * DEGREE)
    assert len(segments) == 2
    assert all(sweep == 1 for _, _, _, sweep in segments)


def test_a_full_turn_is_split_rather_than_silently_dropped():
    """SVG draws nothing when an arc's endpoints coincide, and DrawingML writes a circle
    as one arcTo sweeping 360 degrees.  Taken literally that erases every circle in the
    preset catalogue -- smileyFace's face and eyes, the middle of sun, donut's hole."""
    segments = arc_segments(10.0, 0.0, 10.0, 10.0, 0.0, 360 * DEGREE)
    assert len(segments) == 2, "one segment would start and end at the same point"
    midpoint = segments[0][:2]
    assert midpoint[0] == pytest.approx(-10.0)
    assert segments[-1][0] == pytest.approx(10.0)
    assert segments[-1][1] == pytest.approx(0.0, abs=1e-9)


def test_a_zero_sweep_produces_no_segments():
    assert arc_segments(10.0, 0.0, 10.0, 10.0, 0.0, 0.0) == []


def test_a_full_sweep_returns_to_the_starting_point():
    end_x, end_y, _, _ = arc_endpoint(30.0, 0.0, 30.0, 5.0, 0.0, 360 * DEGREE)
    assert end_x == pytest.approx(30.0)
    assert end_y == pytest.approx(0.0, abs=1e-9)


@pytest.mark.parametrize("geometric_degrees", [0, 30, 45, 60, 90, 135, 180, 270, 359])
def test_the_start_point_always_lies_on_the_reconstructed_ellipse(geometric_degrees):
    """Whatever the start angle, the current point must end up on the ellipse the arc is
    drawn around -- otherwise the arc starts somewhere other than where the pen is."""
    width_radius, height_radius = 40.0, 12.0
    start_x, start_y = 7.0, 3.0
    end_x, end_y, _, _ = arc_endpoint(
        start_x,
        start_y,
        width_radius,
        height_radius,
        geometric_degrees * DEGREE,
        20 * DEGREE,
    )
    # Recover the centre from the start, then confirm the end is on the same ellipse.
    start_parametric = math.atan2(
        width_radius * math.sin(math.radians(geometric_degrees)),
        height_radius * math.cos(math.radians(geometric_degrees)),
    )
    center_x = start_x - width_radius * math.cos(start_parametric)
    center_y = start_y - height_radius * math.sin(start_parametric)
    assert ((end_x - center_x) / width_radius) ** 2 + (
        (end_y - center_y) / height_radius
    ) ** 2 == pytest.approx(1.0)
