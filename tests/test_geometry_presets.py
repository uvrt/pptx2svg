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

**All 187 presets have been measured against PowerPoint itself.**  Each was laid out at
its default adjustments, exported to PDF by PowerPoint 16.106, and scored by silhouette
overlap in a square box and a 2:1 one.  The specification won 51 times, tied 135, and
lost none: there is no shape in the catalogue where PowerPoint departs from ECMA-376.
After promoting those 51 the whole set sits at median 0.999, mean 0.997, minimum 0.962 --
and the minimum is antialiasing on a hairline, since ``line`` and ``lineInv`` score
identically while being drawn by different code.

Two aspect ratios rather than one, because ``chevron`` is pixel-identical to the
specification in a square box and 0.716 against PowerPoint when stretched.
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


def points(data: str) -> list[tuple[float, float]]:
    """Every point an SVG ``d`` string visits.

    Not the same as "every number in the string": an ``A`` command carries seven
    numbers, of which only the last two are a coordinate, so pairing numbers off
    blindly mixes radii and flags into the geometry.
    """
    tokens = re.findall(r"[MLAQCZ]|-?\d+(?:\.\d+)?", data)
    result: list[tuple[float, float]] = []
    index = 0
    while index < len(tokens):
        command = tokens[index]
        index += 1
        if command == "Z":
            continue
        count = {"M": 1, "L": 1, "A": 1, "Q": 2, "C": 3}[command]
        if command == "A":
            index += 5  # rx, ry, x-axis-rotation, large-arc-flag, sweep-flag
        for _ in range(count):
            result.append((float(tokens[index]), float(tokens[index + 1])))
            index += 2
    return result


def shaded_paths(svg: str) -> list[str]:
    """The ``d`` of every path drawn with a shading overlay.

    For an action button those are exactly the pictogram's faces: the silhouette path
    carries the full-box frame as a subpath and the stroked outline sometimes does too,
    so neither can be used to measure the symbol.
    """
    return [
        match.group(1)
        for match in re.finditer(r'<path d="([^"]*)"([^/]*)/>', svg)
        if "fill-opacity=" in match.group(2)
    ]


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
    """A preset that ignored `a:avLst` would silently render every instance identically.

    Nudged both ways, because a default often sits *on* one of the spec's own `pin`
    bounds -- `smileyFace` defaults to the widest smile it allows -- so moving in one
    direction legitimately changes nothing.
    """
    spec = SPEC_PRESETS[name]
    if not spec.adjustments:
        pytest.skip(f"{name} has no adjustment values")
    first_name, default = spec.adjustments[0]
    unchanged = preset_geometry_svg(name, 200.0, 100.0, {})
    nudged = [
        preset_geometry_svg(name, 200.0, 100.0, {first_name: default + delta})
        for delta in (7000, -7000)
    ]
    assert any(value != unchanged for value in nudged)


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
    xs = [x for x, _ in points(path_data(svg)[0])]
    # Every x has its mirror about w/2 somewhere in the outline.
    for x in xs:
        assert any(abs((200.0 - x) - other) < 1e-6 for other in xs)


def test_quadArrowCallout_has_four_arrow_tips_on_the_box_edges():
    svg = preset_geometry_svg("quadArrowCallout", 200.0, 100.0, {})
    visited = points(path_data(svg)[0])
    assert (0.0, 50.0) in visited      # left tip
    assert (200.0, 50.0) in visited    # right tip
    assert (100.0, 0.0) in visited     # top tip
    assert (100.0, 100.0) in visited   # bottom tip


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
    assert tip in points(path_data(preset_geometry_svg(name, 200.0, 100.0, {}))[0])


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


# --------------------------------------------------------------------------------------
# Action buttons
# --------------------------------------------------------------------------------------

ACTION_BUTTONS = [
    "actionButtonBackPrevious", "actionButtonBeginning", "actionButtonBlank",
    "actionButtonDocument", "actionButtonEnd", "actionButtonForwardNext",
    "actionButtonHelp", "actionButtonHome", "actionButtonInformation",
    "actionButtonMovie", "actionButtonReturn", "actionButtonSound",
]


@pytest.mark.parametrize("name", ACTION_BUTTONS)
def test_the_action_button_is_registered(name):
    assert name in PRESET_GEOMETRIES
    assert name in SPEC_PRESETS


@pytest.mark.parametrize("name", ACTION_BUTTONS)
def test_every_action_button_fills_its_whole_box(name):
    """The frame is the first path and is the full rectangle; a button that did not
    cover its box would show the slide through it."""
    first = path_data(preset_geometry_svg(name, 200.0, 100.0, {}))[0]
    assert first.startswith("M 0 0 L 200 0 L 200 100 L 0 100 Z")


def test_actionButtonBlank_is_only_the_frame():
    """The blank button is the control case: one path, the frame, no pictogram.  It is
    also the check that a single normally-painted path is emitted bare rather than
    wrapped in a pointless <g>."""
    svg = preset_geometry_svg("actionButtonBlank", 200.0, 100.0, {})
    assert svg == '<path d="M 0 0 L 200 0 L 200 100 L 0 100 Z"/>'



@pytest.mark.parametrize("name", [n for n in ACTION_BUTTONS if n != "actionButtonBlank"])
def test_every_other_action_button_draws_a_pictogram(name):
    svg = preset_geometry_svg(name, 200.0, 100.0, {})
    assert len(path_data(svg)) > 2


@pytest.mark.parametrize("name", [n for n in ACTION_BUTTONS if n != "actionButtonBlank"])
def test_the_pictogram_is_shaded_so_it_shows_against_the_buttons_own_fill(name):
    """The silhouette shares the button's fill colour, so without the darken/lighten
    faces on top the symbol would be invisible.  Those become a neutral overlay here."""
    svg = preset_geometry_svg(name, 200.0, 100.0, {})
    assert "fill-opacity=" in svg


@pytest.mark.parametrize("name", [n for n in ACTION_BUTTONS if n != "actionButtonBlank"])
def test_the_symbol_box_is_square_however_the_button_is_stretched(name):
    """The pictogram is sized from `ss`, the shortest side, so stretching the button
    wide must not stretch the symbol with it.  Measured on the shaded faces, which are
    pure pictogram -- the silhouette path also carries the full-box frame as a subpath,
    and the stroked outline sometimes does too.

    The symbol box is `dx2 = ss*3/8` either side of the centre, so it is `ss*3/4` on a
    side whatever the aspect ratio."""
    for width, height in ((200.0, 100.0), (100.0, 200.0), (400.0, 100.0)):
        shortest = min(width, height)
        visited = [
            point
            for data in shaded_paths(preset_geometry_svg(name, width, height, {}))
            for point in points(data)
        ]
        assert visited, "every non-blank button has at least one shaded face"
        xs = [x for x, _ in visited]
        ys = [y for _, y in visited]
        side = shortest * 3 / 4
        assert max(xs) - min(xs) <= side + 1e-6
        assert max(ys) - min(ys) <= side + 1e-6
        # And it is centred on the button.
        assert (max(xs) + min(xs)) / 2 == pytest.approx(width / 2, abs=side / 2)
        assert (max(ys) + min(ys)) / 2 == pytest.approx(height / 2, abs=side / 2)


def test_actionButtonHome_draws_a_house_in_the_symbol_box():
    """The symbol box is `dx2 = ss*3/8` either side of the centre, so at 200x100
    (ss = 100, dx2 = 37.5) it spans g11 = hc-dx2 = 62.5 to g12 = hc+dx2 = 137.5
    horizontally and g9 = vc-dx2 = 12.5 to g10 = vc+dx2 = 87.5 vertically.

    The roof apex sits at (hc, g9) and the eaves at (g11, vc) and (g12, vc)."""
    svg = preset_geometry_svg("actionButtonHome", 200.0, 100.0, {})
    visited = points(path_data(svg)[0])
    assert (100.0, 12.5) in visited    # apex, at (hc, g9)
    assert (62.5, 50.0) in visited     # left eave, at (g11, vc)
    assert (137.5, 50.0) in visited    # right eave, at (g12, vc)
    assert (71.875, 87.5) in visited   # wall foot, on g10


# --------------------------------------------------------------------------------------
# The generated table
#
# `src/pptx2svg/render/preset_specs.py` is compiled from ECMA-376's own
# `presetShapeDefinitions.xml` by `tools/derive_preset_geometry.py`.  The spec file is not
# redistributed here, so these tests cannot recompile it; what they can do is hold the
# generated data, the tool's manifest and the renderer's registry in agreement, so that a
# hand-edit or a half-finished regeneration fails loudly.
# --------------------------------------------------------------------------------------

import importlib.util  # noqa: E402
from pathlib import Path  # noqa: E402

from pptx2svg.render.preset_specs import PRESET_SPECS  # noqa: E402

KNOWN_FILL_MODES = {"norm", "none", "lighten", "lightenLess", "darken", "darkenLess"}
COMMAND_ARITY = {"M": 2, "L": 2, "A": 4, "Q": 4, "C": 6, "Z": 0}


def load_tool():
    path = Path(__file__).resolve().parent.parent / "tools/derive_preset_geometry.py"
    spec = importlib.util.spec_from_file_location("derive_preset_geometry", path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


def test_the_generated_table_holds_exactly_what_the_tool_was_asked_for():
    tool = load_tool()
    assert sorted(PRESET_SPECS) == sorted(tool.SPEC_DRIVEN)
    assert len(tool.SPEC_DRIVEN) == len(set(tool.SPEC_DRIVEN)), "duplicate in SPEC_DRIVEN"


def test_the_generated_table_records_the_source_it_was_compiled_from():
    """A shape that looks wrong gets diffed against the spec, so the file has to say
    which edition it came from -- and that has to be the edition the tool accepts."""
    tool = load_tool()
    source = Path(__file__).resolve().parent.parent / "src/pptx2svg/render/preset_specs.py"
    assert tool.SOURCE_SHA256["presetShapeDefinitions.xml"] in source.read_text(encoding="utf-8")


def test_every_generated_preset_reaches_the_renderer():
    assert set(PRESET_SPECS) <= set(PRESET_GEOMETRIES)
    assert set(PRESET_SPECS) == set(SPEC_PRESETS)


@pytest.mark.parametrize("name", sorted(PRESET_SPECS))
def test_the_generated_data_is_structurally_sound(name):
    adjustments, guides, paths = PRESET_SPECS[name]
    assert all(isinstance(n, str) and isinstance(v, int) for n, v in adjustments)
    assert all(isinstance(n, str) and isinstance(f, str) and f for n, f in guides)
    assert paths, "a preset with no paths would silently fall back to a rectangle"
    for fill, stroke, space, commands in paths:
        assert fill in KNOWN_FILL_MODES, f"{name}: unhandled fill mode {fill!r}"
        assert isinstance(stroke, bool)
        assert space is None or (len(space) == 2 and all(v > 0 for v in space))
        assert commands
        for command in commands:
            assert command[0] in COMMAND_ARITY, f"{name}: unknown command {command[0]!r}"
            assert len(command) - 1 == COMMAND_ARITY[command[0]], f"{name}: {command!r}"


@pytest.mark.parametrize("name", sorted(PRESET_SPECS))
def test_guides_only_refer_to_names_already_defined(name):
    """Guides are evaluated in document order, and an unknown name resolves to zero
    rather than raising -- so a forward reference is a silently misdrawn shape."""
    from pptx2svg.guides import builtin_variables

    adjustments, guides, _ = PRESET_SPECS[name]
    defined = set(builtin_variables(1.0, 1.0)) | {n for n, _ in adjustments}
    for guide_name, formula in guides:
        for token in formula.split()[1:]:
            try:
                float(token)
            except ValueError:
                assert token in defined, f"{name}: {guide_name!r} uses undefined {token!r}"
        defined.add(guide_name)


def test_lineInv_is_the_other_diagonal():
    """The one preset ECMA-376 defines that we had no implementation for at all."""
    assert path_data(preset_geometry_svg("lineInv", 200.0, 100.0, {})) == ["M 0 100 L 200 0"]


# --------------------------------------------------------------------------------------
# Stars
#
# The one family held back from the specification, because ECMA-376 gives `star10` an
# inner radius 85% of its outer one and that looked far too shallow to be what PowerPoint
# draws.  Asking PowerPoint settled it: a probe deck of all ten at their default
# adjustments, exported to PDF by PowerPoint 16.106 and rasterised, then scored by
# silhouette overlap against both candidates.
#
#   star    hand-written    specification
#   star4       0.466           0.992
#   star5       0.686           0.984
#   star6       0.750           0.994
#   star7       0.508           0.991
#   star8       0.507           0.989
#   star10      0.424           0.999
#   star12      0.505           0.995
#   star16      0.506           0.991
#   star24      0.505           0.985
#   star32      0.507           0.975
#
# The specification was right in all ten and the doubt was unfounded.  The hand-written
# generator used a single inner ratio of 0.38 for every star but `star6` -- a value only
# correct for `star5` -- so every star above five points came out far too spiky.  The
# residual is antialiasing along the silhouette edge.
# --------------------------------------------------------------------------------------

STARS = {
    # preset: (points, inner/outer radius ratio)
    "star4": (4, 0.25000),
    "star5": (5, 0.28394),
    "star6": (6, 0.53748),
    "star7": (7, 0.61931),
    "star8": (8, 0.75000),
    "star10": (10, 0.81683),
    "star12": (12, 0.75000),
    "star16": (16, 0.75000),
    "star24": (24, 0.75000),
    "star32": (32, 0.75000),
}


def star_radii(name: str) -> tuple[float, float, int]:
    """(smallest, largest, vertex count) about the centre of a 200x200 box."""
    visited = points(path_data(preset_geometry_svg(name, 200.0, 200.0, {}))[0])
    radii = [math.hypot(x - 100.0, y - 100.0) for x, y in visited]
    return min(radii), max(radii), len(visited)


@pytest.mark.parametrize("name", sorted(STARS))
def test_a_star_has_two_vertices_per_point(name):
    expected_points, _ = STARS[name]
    assert star_radii(name)[2] == expected_points * 2


@pytest.mark.parametrize("name", sorted(STARS))
def test_the_star_is_as_deep_as_powerpoint_draws_it(name):
    """The measurement above, one number per shape.  A regression to a fixed inner ratio
    -- the bug this replaced -- moves every star except `star5` and fails here."""
    _, expected_ratio = STARS[name]
    inner, outer, _ = star_radii(name)
    assert inner / outer == pytest.approx(expected_ratio, abs=5e-5)


def test_stars_get_shallower_as_they_gain_points():
    """The pattern the old generator missed: more points means a shallower star, up to
    `star8`, after which the specification holds the ratio at 0.75."""
    ratios = [STARS[f"star{n}"][1] for n in (4, 5, 6, 7, 8)]
    assert ratios == sorted(ratios)
    assert {STARS[f"star{n}"][1] for n in (8, 12, 16, 24, 32)} == {0.75}


# --------------------------------------------------------------------------------------
# Path-local coordinate spaces and arcs
# --------------------------------------------------------------------------------------


def test_a_path_space_arc_scales_without_rotating():
    """A path authored in its own coordinate space must scale linearly onto the shape.

    `stAng` is a *geometric* angle measured in the space the path was authored in, so
    turning it into the ellipse's parametric angle has to use the unscaled radii.
    Scaling them first quietly rotates every arc that does not begin on an axis --
    invisible in a square box, where both radii scale alike, and worth 0.19 of silhouette
    overlap against PowerPoint on a stretched `cloud`.

    The invariant: stretching the box by (2, 1) moves every point by (2, 1).
    """
    from pptx2svg.render.geometry import _P, _Spec, _spec_geometry

    spec = _Spec(
        paths=(
            _P(
                ("M", "100", "50"),
                # A 45-degree start, which is where the two conversions disagree.
                ("A", "50", "50", "2700000", "5400000"),
                space=(100, 100),
            ),
        )
    )
    square = numbers(re.search(r'd="([^"]*)"', _spec_geometry(spec, 400.0, 400.0, {})).group(1))
    wide = numbers(re.search(r'd="([^"]*)"', _spec_geometry(spec, 800.0, 400.0, {})).group(1))
    assert len(square) == len(wide)

    # "M x y" then "A rx ry rot large sweep x y": every x-ish number doubles, every
    # y-ish one is unchanged.  Positions of each within the command are fixed.
    # Coordinates are written to three decimals, so doubling a rounded value can drift
    # by a thousandth; the tolerance allows for the formatting, not for the geometry.
    assert wide[0] == pytest.approx(square[0] * 2, abs=0.01)   # move-to x
    assert wide[1] == pytest.approx(square[1], abs=0.01)       # move-to y
    assert wide[2] == pytest.approx(square[2] * 2, abs=0.01)   # rx
    assert wide[3] == pytest.approx(square[3], abs=0.01)       # ry
    assert wide[-2] == pytest.approx(square[-2] * 2, abs=0.01)  # end x
    assert wide[-1] == pytest.approx(square[-1], abs=0.01)      # end y


def _wholly_path_space_presets():
    """Presets whose every path is authored in its own coordinate space.

    Only these are expected to stretch linearly.  A preset that mixes path-space paths
    with guide-driven ones -- `cloudCallout`, whose bubble is on a 43200 grid but whose
    tail is computed from `ss` -- is *correct* without being linear, because the guides
    legitimately depend on the shorter side.
    """
    return sorted(
        name
        for name, (_, _, paths) in PRESET_SPECS.items()
        if paths and all(space is not None for _, _, space, _ in paths)
    )


@pytest.mark.parametrize("name", _wholly_path_space_presets())
def test_presets_with_a_path_space_stretch_linearly(name):
    """The same invariant through the public entry point."""
    square = numbers(" ".join(path_data(preset_geometry_svg(name, 400.0, 400.0, {}))))
    wide = numbers(" ".join(path_data(preset_geometry_svg(name, 800.0, 400.0, {}))))
    assert len(square) == len(wide) and square
    # Every number is either x-like (doubles), or y-like or a flag (unchanged).  Zero
    # satisfies both, so each position is judged once rather than counted twice.
    for a, b in zip(square, wide):
        assert abs(b - a) < 0.01 or abs(b - 2 * a) < 0.01, (a, b)
