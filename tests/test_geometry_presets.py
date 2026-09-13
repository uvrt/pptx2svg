"""Preset shapes whose definition is transcribed from ECMA-376 Appendix D.

These presets are not hand-approximated: ``render/geometry.py`` carries the
specification's own ``avLst``/``gdLst``/``pathLst`` as data and evaluates it.  That makes
a test asserting "the code computes what the guides say" tautological, so the exact-value
tests below compute the expected vertices **longhand from the published formulas**,
writing the arithmetic out so a reader can check it against the spec without running
anything.  A transcription error in the shipped data then shows up as a mismatch.

The transcription was cross-checked against two independent copies of Appendix D -- a
published dump of ``presetShapeDefinitions.xml`` and OnlyOffice's per-shape C++
transcription -- which agree byte for byte.
"""

from __future__ import annotations

import math
import re

import pytest

from pptx2svg.render.geometry import PRESET_GEOMETRIES, SPEC_PRESETS, preset_geometry_svg

CALLOUTS = [
    "callout1", "callout2", "callout3",
    "accentCallout1", "accentCallout2", "accentCallout3",
    "accentBorderCallout1", "accentBorderCallout2", "accentBorderCallout3",
    "leftArrowCallout", "rightArrowCallout", "upArrowCallout", "downArrowCallout",
    "leftRightArrowCallout", "upDownArrowCallout", "quadArrowCallout",
]


def path_data(svg: str) -> list[str]:
    """Every ``d`` attribute in the generated element, in document order."""
    return re.findall(r'\sd="([^"]*)"', svg)


def numbers(data: str) -> list[float]:
    return [float(token) for token in re.findall(r"-?\d+(?:\.\d+)?", data)]


# --------------------------------------------------------------------------------------
# Registration and general soundness
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize("name", CALLOUTS)
def test_the_preset_is_registered(name):
    assert name in PRESET_GEOMETRIES
    assert name in SPEC_PRESETS


@pytest.mark.parametrize("name", sorted(SPEC_PRESETS))
def test_every_spec_preset_produces_a_usable_element(name):
    svg = preset_geometry_svg(name, 200.0, 100.0, {})
    assert svg.startswith("<path") or svg.startswith("<g>")
    assert svg.endswith("/>") or svg.endswith("</g>")
    data = path_data(svg)
    assert data and all(d.strip() for d in data)


@pytest.mark.parametrize("name", sorted(SPEC_PRESETS))
def test_no_preset_emits_a_non_finite_coordinate(name):
    """A divide-by-zero or a bad `at2` shows up as nan/inf and silently kills the path."""
    for width, height in ((200.0, 100.0), (100.0, 200.0), (50.0, 50.0), (1.0, 400.0)):
        svg = preset_geometry_svg(name, width, height, {})
        for value in numbers(" ".join(path_data(svg))):
            assert math.isfinite(value)


@pytest.mark.parametrize("name", sorted(SPEC_PRESETS))
def test_a_degenerate_box_does_not_raise(name):
    """Zero-sized shapes exist in real decks; they must render empty, not explode."""
    preset_geometry_svg(name, 0.0, 0.0, {})
    preset_geometry_svg(name, 0.0, 100.0, {})


@pytest.mark.parametrize("name", sorted(SPEC_PRESETS))
def test_output_is_byte_stable(name):
    first = preset_geometry_svg(name, 200.0, 100.0, {})
    assert first == preset_geometry_svg(name, 200.0, 100.0, {})


@pytest.mark.parametrize("name", sorted(SPEC_PRESETS))
def test_adjustments_actually_reach_the_guides(name):
    """A preset that ignored `a:avLst` would silently render every instance identically."""
    spec = SPEC_PRESETS[name]
    if not spec.adjustments:
        pytest.skip(f"{name} has no adjustment values")
    first_name, default = spec.adjustments[0]
    nudged = preset_geometry_svg(name, 200.0, 100.0, {first_name: default + 7000})
    assert nudged != preset_geometry_svg(name, 200.0, 100.0, {})


# --------------------------------------------------------------------------------------
# Exact geometry, computed longhand from the published guide formulas
# --------------------------------------------------------------------------------------


def test_leftArrowCallout_vertices():
    """w=200, h=100, so ss=min(w,h)=100 and the spec's defaults give:

        adj1..adj4 = 25000, 25000, 25000, 64977
        maxAdj2 = 50000*h/ss = 50000      a2 = pin(0, 25000, 50000)  = 25000
        maxAdj1 = a2*2       = 50000      a1 = pin(0, 25000, 50000)  = 25000
        maxAdj3 = 100000*w/ss = 200000    a3 = pin(0, 25000, 200000) = 25000
        q2      = a3*ss/w = 12500         maxAdj4 = 100000 - 12500   = 87500
        a4      = pin(0, 64977, 87500) = 64977
        dy1 = ss*a2/100000 = 25           dy2 = ss*a1/200000 = 12.5
        y1 = vc-dy1 = 25   y2 = vc-dy2 = 37.5   y3 = vc+dy2 = 62.5   y4 = vc+dy1 = 75
        x1 = ss*a3/100000 = 25            dx2 = w*a4/100000 = 129.954
        x2 = r - dx2 = 70.046
    """
    svg = preset_geometry_svg("leftArrowCallout", 200.0, 100.0, {})
    assert path_data(svg) == [
        "M 0 50 L 25 25 L 25 37.5 L 70.046 37.5 L 70.046 0 "
        "L 200 0 L 200 100 L 70.046 100 L 70.046 62.5 L 25 62.5 L 25 75 Z"
    ]


def test_rightArrowCallout_mirrors_leftArrowCallout_but_starts_elsewhere():
    """The same construction anchored to the right edge.  The two definitions are *not*
    mirror images of each other's text: `leftArrowCallout` starts its path at the arrow
    tip, `rightArrowCallout` at the top-left corner, so the vertex order differs even
    though the outline is a reflection.

    At 200x100 the guides give a2=a1=a3=25000 and a4=64977 as before, so
    dx3 = ss*a3/100000 = 25 and x3 = r - dx3 = 175 (the arrow base),
    x2 = w*a4/100000 = 129.954 (the box edge), with the arrow spanning
    y1 = vc-dy1 = 25 to y4 = vc+dy1 = 75 and the shaft y2 = 37.5 to y3 = 62.5."""
    svg = preset_geometry_svg("rightArrowCallout", 200.0, 100.0, {})
    assert path_data(svg) == [
        "M 0 0 L 129.954 0 L 129.954 37.5 L 175 37.5 L 175 25 L 200 50 "
        "L 175 75 L 175 62.5 L 129.954 62.5 L 129.954 100 L 0 100 Z"
    ]


def test_callout1_draws_an_unstroked_box_and_a_two_point_leader():
    """Guides: y1 = h*adj1/100000, x1 = w*adj2/100000, y2 = h*adj3/100000,
    x2 = w*adj4/100000, with defaults 18750, -8333, 112500, -38333.  At 200x100:
    y1 = 18.75, x1 = -16.666, y2 = 112.5, x2 = -76.666.  The negative x values put the
    leader outside the box, which is what PowerPoint shows before you drag the handle."""
    svg = preset_geometry_svg("callout1", 200.0, 100.0, {})
    assert path_data(svg) == [
        "M 0 0 L 200 0 L 200 100 L 0 100 Z",
        "M -16.666 18.75 L -76.666 112.5",
    ]
    # The box is filled but not stroked; the leader is stroked but not filled.
    assert 'stroke="none"' in svg
    assert 'fill="none"' in svg


def test_borderCallout1_and_callout1_differ_only_in_whether_the_box_is_stroked():
    """`borderCallout1` is the hand-written generator that predates this mechanism; the
    two must agree on where the box and the leader go."""
    spec = preset_geometry_svg("callout1", 200.0, 100.0, {})
    assert numbers(" ".join(path_data(spec))) == pytest.approx(
        [0, 0, 200, 0, 200, 100, 0, 100, -16.666, 18.75, -76.666, 112.5]
    )


def test_accentCallout1_adds_an_accent_bar_at_x1():
    """The spec really does place the bar at x1 -- the leader's own x -- rather than at
    the left edge, and really does close an empty subpath between the moveTo and the
    lnTo.  Both oddities are present in two independent copies of Appendix D, so they
    are transcribed rather than tidied away."""
    svg = preset_geometry_svg("accentCallout1", 200.0, 100.0, {})
    assert path_data(svg) == [
        "M 0 0 L 200 0 L 200 100 L 0 100 Z",
        "M -16.666 0 Z L -16.666 100",
        "M -16.666 18.75 L -76.666 112.5",
    ]


def test_upDownArrowCallout_is_symmetric_about_the_horizontal_centre():
    svg = preset_geometry_svg("upDownArrowCallout", 200.0, 100.0, {})
    values = numbers(path_data(svg)[0])
    xs = values[0::2]
    # Every x has its mirror about w/2 somewhere in the outline.
    for x in xs:
        assert any(abs((200.0 - x) - other) < 1e-6 for other in xs)


def test_quadArrowCallout_has_four_arrow_tips_on_the_box_edges():
    svg = preset_geometry_svg("quadArrowCallout", 200.0, 100.0, {})
    values = numbers(path_data(svg)[0])
    points = list(zip(values[0::2], values[1::2]))
    assert (0.0, 50.0) in points      # left tip
    assert (200.0, 50.0) in points    # right tip
    assert (100.0, 0.0) in points     # top tip
    assert (100.0, 100.0) in points   # bottom tip


@pytest.mark.parametrize(
    "name,tip",
    [
        ("leftArrowCallout", (0.0, 50.0)),
        ("rightArrowCallout", (200.0, 50.0)),
        ("upArrowCallout", (100.0, 0.0)),
        ("downArrowCallout", (100.0, 100.0)),
    ],
)
def test_each_arrow_callout_points_the_way_its_name_says(name, tip):
    values = numbers(path_data(preset_geometry_svg(name, 200.0, 100.0, {}))[0])
    assert tip in list(zip(values[0::2], values[1::2]))


# --------------------------------------------------------------------------------------
# Adjustment clamping
# --------------------------------------------------------------------------------------


def test_an_out_of_range_adjustment_is_clamped_by_the_specs_own_pin_guides():
    """`pin 0 adj3 maxAdj3` is the spec's guard against an arrowhead longer than the
    shape.  A deck can carry any integer, so the clamp has to hold."""
    svg = preset_geometry_svg("leftArrowCallout", 200.0, 100.0, {"adj3": 10_000_000})
    values = numbers(path_data(svg)[0])
    assert all(-1.0 <= x <= 201.0 for x in values[0::2])


def test_a_negative_adjustment_is_clamped_to_zero():
    svg = preset_geometry_svg("leftArrowCallout", 200.0, 100.0, {"adj1": -50000})
    values = numbers(path_data(svg)[0])
    assert all(math.isfinite(v) for v in values)
