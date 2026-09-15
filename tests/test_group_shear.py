"""A rotated child inside a non-uniformly scaled group is a rectangle, never a shear.

A group's ``ext``/``chExt`` ratio looks like a matrix, and composing it with a rotated
child's own matrix produces a shear: a square would come out a parallelogram.  PowerPoint
does not do that.  It hands each of the two factors to one of the child's *own* axes --
choosing by which slide axis that axis currently lies nearer to -- and rotates the grown
box afterwards, so the corners stay square at every angle.

Everything asserted here was read out of PowerPoint 16.x's own PDF export of a probe deck
(``tools/make_group_shear_probe.py``, three decks, 82 probes), by walking the PDF's path
objects for the drawn vertices rather than rasterising and guessing.  The discriminating
case is a square rotated 45 degrees inside a 4:1 group: scale-after-rotate gives a rhombus
with diagonals in the ratio 4:1, and PowerPoint drew a 4:1 *rectangle* with an interior
angle of 90.000 degrees.  See :func:`~pptx2svg.render.svg.swaps_group_axes` for the full
measurement, including the reading it refuted.
"""

from __future__ import annotations

import math
import re

import pytest

from pptx2svg import model as m
from pptx2svg.render.context import RenderContext
from pptx2svg.render.svg import render_slide_to_svg

SLIDE = m.SlideSize(9144000, 5143500)
EMU_PER_POINT = 12700
#: The probe square, 400000 EMU on a side, centred in a 1000000 EMU child space.
BOX = 400000
CHILD = 1000000


def render(*elements: m.SlideElement) -> str:
    return render_slide_to_svg(
        m.Slide(slide_number=1, elements=list(elements)), SLIDE, RenderContext()
    )


def square(rotation: float = 0.0, flip_h: bool = False, flip_v: bool = False,
           width: int = BOX, height: int = BOX) -> m.ShapeElement:
    return m.ShapeElement(
        transform=m.Transform(
            offset_x=(CHILD - width) // 2,
            offset_y=(CHILD - height) // 2,
            extent_width=width,
            extent_height=height,
            rotation=rotation,
            flip_h=flip_h,
            flip_v=flip_v,
        ),
        geometry=m.PresetGeometry(preset="rect"),
    )


def group(scale_x: float, scale_y: float, *children: m.SlideElement,
          rotation: float = 0.0, child: int = CHILD) -> m.GroupElement:
    return m.GroupElement(
        transform=m.Transform(
            offset_x=0,
            offset_y=0,
            extent_width=int(child * scale_x),
            extent_height=int(child * scale_y),
            rotation=rotation,
        ),
        child_transform=m.Transform(
            offset_x=0, offset_y=0, extent_width=child, extent_height=child
        ),
        children=list(children),
    )


# ------------------------------------------------------------------------------------
# Reading the drawing back out
# ------------------------------------------------------------------------------------


def _multiply(outer, inner):
    a0, a1, a2, a3, a4, a5 = outer
    b0, b1, b2, b3, b4, b5 = inner
    return (
        a0 * b0 + a2 * b1, a1 * b0 + a3 * b1,
        a0 * b2 + a2 * b3, a1 * b2 + a3 * b3,
        a0 * b4 + a2 * b5 + a4, a1 * b4 + a3 * b5 + a5,
    )


def _parse(text: str):
    matrix = (1.0, 0.0, 0.0, 1.0, 0.0, 0.0)
    for name, args in re.findall(r"(\w+)\(([^)]*)\)", text):
        values = [float(v) for v in re.split(r"[\s,]+", args.strip()) if v]
        if name == "translate":
            part = (1, 0, 0, 1, values[0], values[1] if len(values) > 1 else 0)
        elif name == "scale":
            part = (values[0], 0, 0, values[1] if len(values) > 1 else values[0], 0, 0)
        elif name == "rotate":
            radians = math.radians(values[0])
            cos, sin = math.cos(radians), math.sin(radians)
            part = (cos, sin, -sin, cos, 0, 0)
            if len(values) == 3:
                shift = (1, 0, 0, 1, values[1], values[2])
                part = _multiply(_multiply(shift, part), (1, 0, 0, 1, -values[1], -values[2]))
        else:  # pragma: no cover - the renderer emits nothing else
            raise AssertionError(f"unhandled SVG transform {name}")
        matrix = _multiply(matrix, part)
    return matrix


def drawn(svg: str) -> dict:
    """What the single ``<rect>`` in ``svg`` actually draws, in points.

    Composes every enclosing ``<g transform>`` so the answer is about the picture and not
    about which of the nested transforms happens to carry which term -- the same thing
    the PDF readout does to PowerPoint's output.
    """
    stack = [(1.0, 0.0, 0.0, 1.0, 0.0, 0.0)]
    found = None
    for token in re.finditer(r"<g\b([^>]*)>|</g>|<rect\b([^>]*)/>", svg):
        if token.group(0) == "</g>":
            stack.pop()
        elif token.group(1) is not None:
            attr = re.search(r'transform="([^"]*)"', token.group(1))
            stack.append(_multiply(stack[-1], _parse(attr.group(1))) if attr else stack[-1])
        elif "width=" in token.group(2) and "height=" in token.group(2):
            width = float(re.search(r'(?<!-)\bwidth="([\d.]+)"', token.group(2)).group(1))
            height = float(re.search(r'(?<!-)\bheight="([\d.]+)"', token.group(2)).group(1))
            found = (stack[-1], width, height)
    assert found is not None, "no <rect> in the rendered slide"
    matrix, width, height = found
    scale = 72 / 96  # the renderer works in CSS pixels; PowerPoint was measured in points

    def at(x, y):
        return (
            (matrix[0] * x + matrix[2] * y + matrix[4]) * scale,
            (matrix[1] * x + matrix[3] * y + matrix[5]) * scale,
        )

    origin = at(0, 0)
    across = at(width, 0)
    down = at(0, height)
    x_axis = (across[0] - origin[0], across[1] - origin[1])
    y_axis = (down[0] - origin[0], down[1] - origin[1])
    centre = at(width / 2, height / 2)
    return {
        "width": round(math.hypot(*x_axis), 3),
        "height": round(math.hypot(*y_axis), 3),
        # Clockwise from east, as OOXML measures rotation, modulo 180 because a side and
        # its reverse are one axis.
        "angle": round(math.degrees(math.atan2(x_axis[1], x_axis[0])) % 180, 3),
        "interior": round(
            math.degrees(
                math.acos(
                    (x_axis[0] * y_axis[0] + x_axis[1] * y_axis[1])
                    / (math.hypot(*x_axis) * math.hypot(*y_axis))
                )
            ),
            3,
        ),
        "centre": (round(centre[0], 3), round(centre[1], 3)),
    }


#: The probe square's side, in points.  400000 EMU.
SIDE = BOX / EMU_PER_POINT


def side(multiple: float = 1.0):
    """``multiple`` times the probe square's side, to within the formatter's resolution.

    SVG coordinates are written to three decimals, so a length that the group arithmetic
    puts at 62.9925 pt comes back as 62.993 where the authored figure is 62.992.  A
    hundredth of a point is two orders of magnitude finer than the 1x-against-4x the
    assertions are actually about.
    """
    return pytest.approx(SIDE * multiple, abs=0.01)


# ------------------------------------------------------------------------------------
# The discriminator
# ------------------------------------------------------------------------------------


def test_a_rotated_square_in_a_four_to_one_group_is_a_rectangle_not_a_rhombus():
    """The case that settles the order of operations.

    Scale-after-rotate would give a rhombus: diagonals 178.2 and 44.5 pt, corners at 28
    and 152 degrees.  PowerPoint drew a 31.496 x 125.984 pt rectangle with a 90.000
    degree corner, so the scale reaches the child's own axes and the rotation is applied
    to the result.
    """
    box = drawn(render(group(4, 1, square(rotation=45))))
    assert box["interior"] == 90.0
    assert (box["width"], box["height"]) == (side(), side(4))
    assert box["angle"] == 45.0


def test_the_centre_moves_by_the_plain_scale():
    """Only the extents are reassigned; the child's centre is mapped by ``S`` as usual.

    Measured: the centre of every one of the 43 sweep probes landed on
    ``origin + S * (child centre - chOff)`` to the hundredth of a point, at every angle
    and with every flip.
    """
    box = drawn(render(group(4, 1, square(rotation=45))))
    # To the hundredth of a point, which is what was measured -- not to the thousandth,
    # which is only an artefact of how `drawn` rounds.  The y centre here lands on
    # 39.3705, right on the 3-decimal boundary, and the rotation puts `math.sin`/`cos`
    # in front of it: Windows composed 39.371 where macOS composed 39.370, and an
    # exact comparison turned a half-thousandth of a point into a red build.
    assert box["centre"] == pytest.approx(
        (CHILD / 2 * 4 / EMU_PER_POINT, CHILD / 2 / EMU_PER_POINT), abs=0.01
    )


def test_an_unrotated_child_is_scaled_the_obvious_way():
    """The control: with nothing rotated there is no reassignment to make."""
    box = drawn(render(group(4, 1, square())))
    assert (box["width"], box["height"], box["angle"]) == (side(4), side(), 0.0)


# ------------------------------------------------------------------------------------
# Where the two factors change places
# ------------------------------------------------------------------------------------


def test_the_factors_change_places_between_44_9_and_45_degrees():
    """Measured on consecutive probes: 44.9 put the 4 on the width, 45.0 on the height.

    Both drew exactly 1x and 4x the authored side -- nothing in between -- so this is a
    switch and not a blend of the two factors.
    """
    below = drawn(render(group(4, 1, square(rotation=44.9))))
    at_45 = drawn(render(group(4, 1, square(rotation=45.0))))
    assert (below["width"], below["height"]) == (side(4), side())
    assert (at_45["width"], at_45["height"]) == (side(), side(4))


def test_the_switch_repeats_every_ninety_degrees_in_both_directions():
    """Measured: 134 swapped, 135 did not; 180 behaved as 0, 225 as 45, 270 as 90,
    315 as 135, 405 as 45, -30 as 150 and -60 as 120."""
    swapped = [45, 50, 89, 91, 100, 134, 225, 270, 405, -60]
    straight = [0, 15, 30, 40, 44, 135, 136, 180, 315, -30, -45]
    for angle in swapped:
        box = drawn(render(group(4, 1, square(rotation=angle))))
        assert (box["width"], box["height"]) == (side(), side(4)), angle
    for angle in straight:
        box = drawn(render(group(4, 1, square(rotation=angle))))
        assert (box["width"], box["height"]) == (side(4), side()), angle


def test_the_switch_does_not_move_with_the_scale_ratio():
    """2:1 and 10:1 both flipped between an authored 44 and an authored 45.

    That is what makes it a property of the angle rather than of the stretch: a rule
    derived from the factors would put the crossing somewhere else for each ratio.
    """
    for ratio in (2, 10):
        assert drawn(render(group(ratio, 1, square(rotation=44))))["width"] == round(SIDE * ratio, 3)
        assert drawn(render(group(ratio, 1, square(rotation=45))))["height"] == round(SIDE * ratio, 3)


def test_neither_factor_has_to_be_one():
    """A 4:2 group put 4 and 2 on the child's own axes, and swapped them past 45.

    With a 4:1 group a "swap" could be read as "one axis is left alone"; 4:2 cannot.
    """
    thirty = drawn(render(group(4, 2, square(rotation=30))))
    sixty = drawn(render(group(4, 2, square(rotation=60))))
    assert (thirty["width"], thirty["height"]) == (side(4), side(2))
    assert (sixty["width"], sixty["height"]) == (side(2), side(4))


def test_the_child_does_not_have_to_be_square():
    """A 400000 x 200000 child was treated the same way, so the rule is about the axes
    and not about a square's symmetry."""
    box = drawn(render(group(4, 1, square(rotation=60, height=BOX // 2))))
    assert (box["width"], box["height"]) == (side(), side(2))
    assert box["interior"] == 90.0


# ------------------------------------------------------------------------------------
# Flips
# ------------------------------------------------------------------------------------


def test_a_flip_does_not_matter_away_from_the_crossing():
    """flipH, flipV and both were measured at 30 and 60 degrees; all matched the
    unflipped probe, because a flip only changes the angle's sign and the test is
    otherwise symmetric."""
    for flip_h, flip_v in ((True, False), (False, True), (True, True)):
        at_30 = drawn(render(group(4, 1, square(rotation=30, flip_h=flip_h, flip_v=flip_v))))
        at_60 = drawn(render(group(4, 1, square(rotation=60, flip_h=flip_h, flip_v=flip_v))))
        assert at_30["width"] == side(4), (flip_h, flip_v)
        assert at_60["height"] == side(4), (flip_h, flip_v)


def test_one_flip_turns_the_crossing_round_at_the_tie():
    """Exactly at 45 and 135 degrees a single flip reverses the answer.

    Measured, and the reason the rule is stated on a signed angle rather than on
    ``|sin| > |cos|``: at an authored 45 the unflipped square swapped and the flipped one
    did not; at an authored 135 the unflipped square did not swap and the flipped one
    did.  Two flips cancel and behave like none.
    """
    assert drawn(render(group(4, 1, square(rotation=45))))["height"] == side(4)
    assert drawn(render(group(4, 1, square(rotation=45, flip_h=True))))["width"] == side(4)
    assert drawn(render(group(4, 1, square(rotation=45, flip_v=True))))["width"] == side(4)

    assert drawn(render(group(4, 1, square(rotation=135))))["width"] == side(4)
    assert drawn(render(group(4, 1, square(rotation=135, flip_h=True))))["height"] == side(4)

    both = square(rotation=45, flip_h=True, flip_v=True)
    assert drawn(render(group(4, 1, both)))["height"] == side(4)


# ------------------------------------------------------------------------------------
# Nesting and the group's own rotation
# ------------------------------------------------------------------------------------


def test_a_rotated_nested_group_is_treated_like_a_rotated_shape():
    """Measured: an unrotated square inside a group rotated 45 inside a 4:1 group drew
    exactly what a square rotated 45 directly inside the 4:1 group drew."""
    inner = m.GroupElement(
        transform=m.Transform(offset_x=0, offset_y=0, extent_width=CHILD, extent_height=CHILD,
                              rotation=45),
        child_transform=m.Transform(offset_x=0, offset_y=0, extent_width=CHILD,
                                    extent_height=CHILD),
        children=[square()],
    )
    assert drawn(render(group(4, 1, inner))) == drawn(render(group(4, 1, square(rotation=45))))


def test_nested_scales_multiply_before_the_axes_are_assigned():
    """A 2:1 group inside a 4:1 group scaled a rotated child by 8, not by 2 then 4.

    Measured at 251.968 pt against the authored 31.496, which is 8x -- so the assignment
    sees the product and cannot be applied twice with two different answers.
    """
    inner = m.GroupElement(
        transform=m.Transform(offset_x=0, offset_y=0, extent_width=CHILD * 2,
                              extent_height=CHILD),
        child_transform=m.Transform(offset_x=0, offset_y=0, extent_width=CHILD,
                                    extent_height=CHILD),
        children=[square(rotation=45)],
    )
    box = drawn(render(group(4, 1, inner)))
    assert (box["width"], box["height"]) == (side(), side(8))
    assert box["interior"] == 90.0


def test_the_groups_own_rotation_is_applied_after_its_scale():
    """A group rotated 30 with a 4:1 scale and a child rotated 45 drew the child's long
    side at 165 degrees -- the 135 the group gives it, turned another 30."""
    box = drawn(render(group(4, 1, square(rotation=45), rotation=30)))
    assert box["interior"] == 90.0
    assert (box["width"], box["height"]) == (side(), side(4))
    assert box["angle"] == 75.0  # the *short* side; the long one is at 165


# ------------------------------------------------------------------------------------
# What the change must not disturb
# ------------------------------------------------------------------------------------


def test_a_uniform_group_still_draws_its_children_through_one_scale():
    """A uniform scale commutes with rotation, so the established ``scale()`` around the
    whole subtree is already exact and is kept -- along with the counter-transform that
    keeps the glyphs at their authored size."""
    svg = render(group(4, 4, square(rotation=45)))
    assert 'scale(4, 4)' in svg


def test_a_scale_uniform_only_to_rounding_is_still_uniform():
    """Google Slides rounds a uniform scale into two EMU pairs that no longer divide to
    the same number -- 3.796875 across against 3.7968797 down in the meal planning
    template.  Those are one scale, and must not take the folding path on float noise.
    """
    almost = m.GroupElement(
        transform=m.Transform(offset_x=0, offset_y=0, extent_width=5024671, extent_height=7367088),
        child_transform=m.Transform(offset_x=0, offset_y=0, extent_width=1323370,
                                    extent_height=1940303),
        children=[square(rotation=45)],
    )
    assert "scale(3.797, 3.797)" in render(almost)


# ------------------------------------------------------------------------------------
# Text inside the same composite
# ------------------------------------------------------------------------------------


def text_box(rotation: float = 0.0, width: int = BOX, height: int = BOX // 2):
    return m.ShapeElement(
        transform=m.Transform(
            offset_x=(CHILD - width) // 2,
            offset_y=(CHILD - height) // 2,
            extent_width=width,
            extent_height=height,
            rotation=rotation,
        ),
        geometry=m.PresetGeometry(preset="rect"),
        text_body=m.TextBody(
            paragraphs=[
                m.Paragraph(
                    runs=[
                        m.TextRun("HOHOHO", m.RunProperties(font_size=18.0,
                                                            font_family="Arial"))
                    ]
                )
            ],
            body_properties=m.BodyProperties(wrap="none", margin_left=0, margin_right=0,
                                             margin_top=0, margin_bottom=0),
        ),
    )


def text_frame(svg: str) -> dict:
    """Where and how the glyphs are drawn: baseline origin, angle, scale and skew."""
    stack = [(1.0, 0.0, 0.0, 1.0, 0.0, 0.0)]
    found = None
    for token in re.finditer(r"<g\b([^>]*)>|</g>|<text\b([^>]*)>", svg):
        if token.group(0) == "</g>":
            stack.pop()
        elif token.group(1) is not None:
            attr = re.search(r'transform="([^"]*)"', token.group(1))
            stack.append(_multiply(stack[-1], _parse(attr.group(1))) if attr else stack[-1])
        else:
            x = float(re.search(r'\bx="([-\d.]+)"', token.group(2)).group(1))
            y = float(re.search(r'\by="([-\d.]+)"', token.group(2)).group(1))
            found = (stack[-1], x, y)
    assert found is not None, "no <text> in the rendered slide"
    matrix, x, y = found
    a, b, c, d, e, f = matrix
    scale = 72 / 96
    return {
        "origin": (round((a * x + c * y + e) * scale, 2),
                   round((b * x + d * y + f) * scale, 2)),
        "angle": round(math.degrees(math.atan2(b, a)) % 360, 3),
        "scale": round(math.hypot(a, b), 6),
        # Zero for a pure rotation whatever the scale, and non-zero the moment the
        # composite stops being a similarity -- which is what a shear would do.
        "skew": round(a * c + b * d, 6),
    }


def test_text_in_a_rotated_child_of_a_stretched_group_is_neither_skewed_nor_scaled():
    """PowerPoint writes text with a rotation, not a general matrix.

    Measured on ten text probes: every character of an 18 pt Arial run inside a 4:1, a
    1:4 and a 2:1 group came out with a text matrix of exactly ``18 * R(theta)`` -- the
    same matrix as the ungrouped control, with no skew term at any angle and the size
    untouched.  The angle is the one the run was *authored* at, not the effective angle
    the surrounding geometry takes.
    """
    for rotation in (0, 20, 45):
        for scale_x, scale_y in ((4, 1), (1, 4), (2, 1)):
            frame = text_frame(render(group(scale_x, scale_y, text_box(rotation=rotation))))
            assert frame["skew"] == 0.0, (rotation, scale_x, scale_y)
            assert frame["scale"] == 1.0, (rotation, scale_x, scale_y)
            assert frame["angle"] == rotation, (rotation, scale_x, scale_y)


def test_the_text_frame_follows_the_reassigned_box():
    """The glyphs are absolute but the frame is not: it is the box the geometry rule
    gives, so the run wraps and anchors as it would in an ungrouped shape of that size.

    Measured by predicting the first baseline from the swapped 31.496 x 62.992 pt frame
    and finding PowerPoint's within 0.07 pt of it; the unswapped 125.984 x 15.748 frame
    predicts a point 50 pt away.
    """
    def bare(width, height):
        return m.ShapeElement(
            transform=m.Transform(
                offset_x=CHILD * 4 // 2 - width // 2,
                offset_y=CHILD // 2 - height // 2,
                extent_width=width,
                extent_height=height,
                rotation=45,
            ),
            geometry=m.PresetGeometry(preset="rect"),
            text_body=text_box().text_body,
        )

    inside = text_frame(render(group(4, 1, text_box(rotation=45))))
    # 400000 x 200000 authored at 45 degrees, so the 4 lands on the height.
    swapped = text_frame(render(bare(BOX, BOX // 2 * 4)))
    unswapped = text_frame(render(bare(BOX * 4, BOX // 2)))
    assert inside == swapped
    assert inside["origin"] != unswapped["origin"]
