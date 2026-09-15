#!/usr/bin/env python3
"""Build the deck that settles how PowerPoint composes a group's scale with rotation.

A group maps its children's authored coordinate space onto its on-slide box.  When that
map is non-uniform *and* something inside it is rotated, two compositions are possible
and they draw differently:

* **scale-after-rotate** (``S . R``) -- the child is rotated in its own space and the
  whole space is then stretched, so a square comes out a *parallelogram*: the two axes
  are stretched by different amounts and stop being perpendicular.
* **rotate-after-scale** (``R . S``) -- the child's box is stretched first and the result
  rotated as a rigid body, so a square comes out an unskewed *rectangle* turned by the
  authored angle.

A square rotated 45 degrees inside a 4:1 group separates them completely: ``S . R`` gives
a rhombus with diagonals in the ratio 4:1, ``R . S`` a 4:1 rectangle at 45 degrees whose
diagonals are equal.  Everything else on the deck -- other angles, the other axis, flips,
group rotation, two nesting orders, and the same cases with text in them -- exists to
check that whatever the discriminator says also holds away from it.

There are three decks, built in the order they were needed: ``probe`` asks the question,
``sweep`` finds exactly where the two factors change places and whether that place moves,
and ``tie`` settles the two angles where a flip changes the answer (and, on the same
pages, whether a group's scale reaches the pen).  The reader picks the probe table by the
PDF's file name, so a deck and its readout cannot drift apart.

Usage::

    python3 tools/make_group_shear_probe.py ~/pptx2svg-oracle/group-shear.pptx probe
    osascript tools/powerpoint_export_pdf.applescript \
        ~/pptx2svg-oracle/group-shear.pptx ~/pptx2svg-oracle/group-shear.pdf
    python3 tools/read_group_shear_probe.py ~/pptx2svg-oracle/group-shear.pdf

and the same three lines for ``group-sweep`` / ``sweep`` and ``group-tie`` / ``tie``.

PowerPoint is sandboxed: both paths must be in a directory it has already been granted
access to, which in this repository means ``~/pptx2svg-oracle`` and nothing else.  A fresh
directory, even under ``$HOME``, fails with -9074 *after* opening the deck, so it reads as
a broken deck rather than an unapproved path.  See the header of
``tools/powerpoint_export_pdf.applescript``.  Delete the decks and their PDFs afterwards:
that directory is meant to hold only the committed fixture pairs.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from derive_table_styles import slide_document, write_deck  # noqa: E402

HERE = Path(__file__).resolve().parent
SOURCE = HERE.parent / "tests/fixtures/authoring-integration.pptx"
BACKDROP = "FFFFFF"

#: The child coordinate space every group on this deck declares.  Square, so the group's
#: own ``ext`` is the only thing that makes the mapping non-uniform.
CHILD = 1000000
#: The probe shape, centred in that space.  Square, so a rotation cannot be mistaken for
#: an extent swap and the two compositions differ in shape and not only in angle.
BOX = 400000
BOX_OFF = (CHILD - BOX) // 2

#: Distinct fills so a shape can be identified in the exported PDF without relying on
#: drawing order.  PowerPoint emits one path object per shape; the fill colour survives.
FILL = "C00000"
CONTROL_FILL = "0070C0"

#: Where the group's on-slide box sits.  Placed so a 4x stretch of the child space still
#: lands on the 720x405 pt slide at every angle on the deck.
ORIGIN = (2072000, 1571750)


def _xfrm(x, y, cx, cy, rot=0, flip_h=False, flip_v=False, child=None):
    attrs = ""
    if rot:
        attrs += f' rot="{int(round(rot * 60000))}"'
    if flip_h:
        attrs += ' flipH="1"'
    if flip_v:
        attrs += ' flipV="1"'
    inner = f'<a:off x="{x}" y="{y}"/><a:ext cx="{cx}" cy="{cy}"/>'
    if child is not None:
        chx, chy, chcx, chcy = child
        inner += f'<a:chOff x="{chx}" y="{chy}"/><a:chExt cx="{chcx}" cy="{chcy}"/>'
    return f"<a:xfrm{attrs}>{inner}</a:xfrm>"


def rect(shape_id, name, x, y, cx, cy, *, rot=0, flip_h=False, flip_v=False, fill=FILL,
         line_emu=0):
    line = (
        f'<a:ln w="{line_emu}"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>'
        if line_emu else "<a:ln><a:noFill/></a:ln>"
    )
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/>'
        "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>"
        + _xfrm(x, y, cx, cy, rot, flip_h, flip_v)
        + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        f'<a:solidFill><a:srgbClr val="{fill}"/></a:solidFill>'
        + line
        + "</p:spPr>"
        "<p:txBody><a:bodyPr/><a:lstStyle/><a:p/></p:txBody></p:sp>"
    )


def textbox(shape_id, name, x, y, cx, cy, text, *, rot=0, flip_h=False, flip_v=False,
            size=1800):
    """A left-and-top anchored, non-wrapping run of Arial.

    Every knob that could move the glyphs for a reason other than the group transform is
    pinned: no wrap so the line cannot break at a scaled width, zero insets so the
    frame's own padding does not enter the reading, and Arial because it is the face the
    metrics in this repository are real for.
    """
    return (
        f'<p:sp><p:nvSpPr><p:cNvPr id="{shape_id}" name="{name}"/>'
        '<p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr><p:spPr>'
        + _xfrm(x, y, cx, cy, rot, flip_h, flip_v)
        + '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/></p:spPr>'
        '<p:txBody><a:bodyPr wrap="none" lIns="0" tIns="0" rIns="0" bIns="0" anchor="t"/>'
        "<a:lstStyle/>"
        '<a:p><a:pPr algn="l"><a:buNone/></a:pPr>'
        f'<a:r><a:rPr lang="en-US" sz="{size}" dirty="0">'
        '<a:solidFill><a:srgbClr val="000000"/></a:solidFill>'
        '<a:latin typeface="Arial"/></a:rPr>'
        f"<a:t>{text}</a:t></a:r></a:p></p:txBody></p:sp>"
    )


def group(shape_id, name, x, y, cx, cy, children, *, rot=0, flip_h=False, flip_v=False,
          child_space=(0, 0, CHILD, CHILD)):
    return (
        f'<p:grpSp><p:nvGrpSpPr><p:cNvPr id="{shape_id}" name="{name}"/>'
        "<p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr>"
        + _xfrm(x, y, cx, cy, rot, flip_h, flip_v, child=child_space)
        + "</p:grpSpPr>"
        + children
        + "</p:grpSp>"
    )


def label(shape_id, text):
    """A caption, so a human flipping through the PDF can tell the slides apart."""
    return textbox(shape_id, "caption", 200000, 150000, 8700000, 400000, text, size=1200)


#: Every probe, in slide order.  The reader imports this list, so the deck and the
#: readout cannot drift apart.
GEOMETRY_PROBES: list[dict] = [
    # The controls: the same square at the on-slide size the 4:1 group would give it,
    # rotated by the same authored angle and *not* grouped.  Under rotate-after-scale a
    # grouped probe must land exactly on its control; under scale-after-rotate it cannot.
    dict(key="control-rot0", scale=(4, 1), rot=0, grouped=False),
    dict(key="control-rot45", scale=(4, 1), rot=45, grouped=False),
    # The discriminator and its neighbours.
    dict(key="x4-rot0", scale=(4, 1), rot=0),
    dict(key="x4-rot45", scale=(4, 1), rot=45),
    dict(key="x4-rot30", scale=(4, 1), rot=30),
    dict(key="x4-rot15", scale=(4, 1), rot=15),
    dict(key="x4-rot60", scale=(4, 1), rot=60),
    dict(key="x4-rot90", scale=(4, 1), rot=90),
    dict(key="y4-rot45", scale=(1, 4), rot=45),
    dict(key="x2-rot45", scale=(2, 1), rot=45),
    dict(key="x4-rot20", scale=(4, 1), rot=20),
    # A flip is a negative scale, so it should compose the same way.
    dict(key="x4-rot0-flipH", scale=(4, 1), rot=0, flip_h=True),
    dict(key="x4-rot45-flipH", scale=(4, 1), rot=45, flip_h=True),
    dict(key="x4-rot45-flipV", scale=(4, 1), rot=45, flip_v=True),
    # The group itself rotated, with a non-uniform scale under it.
    dict(key="grot30-x4-rot0", scale=(4, 1), rot=0, group_rot=30),
    dict(key="grot30-x4-rot45", scale=(4, 1), rot=45, group_rot=30),
    # Nesting, all three orders.
    dict(key="nest-x2-x2-rot45", scale=(4, 1), rot=45, nest="split"),
    dict(key="nest-outer-rot-inner-x4", scale=(4, 1), rot=0, nest="outer-rot45"),
    dict(key="nest-outer-x4-inner-rot", scale=(4, 1), rot=0, nest="inner-rot45"),
]

TEXT_PROBES: list[dict] = [
    dict(key="text-control-rot0", scale=(4, 1), rot=0, grouped=False),
    dict(key="text-control-rot45", scale=(4, 1), rot=45, grouped=False),
    dict(key="text-x4-rot0", scale=(4, 1), rot=0),
    dict(key="text-x4-rot45", scale=(4, 1), rot=45),
    dict(key="text-x4-rot20", scale=(4, 1), rot=20),
    dict(key="text-y4-rot45", scale=(1, 4), rot=45),
    dict(key="text-x2-rot45", scale=(2, 1), rot=45),
    dict(key="text-x4-rot45-flipH", scale=(4, 1), rot=45, flip_h=True),
    dict(key="text-grot45-x4-rot0", scale=(4, 1), rot=0, group_rot=45),
    dict(key="text-nest-x2-x2-rot45", scale=(4, 1), rot=45, nest="split"),
]

PROBES = GEOMETRY_PROBES + TEXT_PROBES


def _sweep(prefix, scale, angles, **extra):
    return [dict(key=f"{prefix}-rot{a:g}".replace(".", "_"), scale=scale, rot=a, **extra)
            for a in angles]


#: The follow-up deck.  The first deck showed the drawn rectangle's long axis jumping by
#: 90 degrees somewhere between an authored 30 and an authored 45; these probes find
#: where, check that the place does not move with the scale ratio, and check that the two
#: factors really are ``sx`` and ``sy`` rather than anything derived from them.
SWEEP_PROBES: list[dict] = (
    # Where the jump is, and whether it repeats every 90 degrees, in both directions.
    _sweep("t4", (4, 1), [40, 44, 44.9, 45, 45.1, 46, 50, 89, 91, 100,
                          134, 135, 136, 180, 225, 270, 315, -30, -45, -60])
    # Does the threshold move with the ratio?  A rule that snaps each of the shape's own
    # axes to the nearer slide axis puts it at 45 for every ratio; anything derived from
    # the stretch would not.
    + _sweep("t2", (2, 1), [40, 44, 45, 46, 50])
    + _sweep("t10", (10, 1), [40, 44, 45, 46, 50])
    # A non-square child: separates "the factors are assigned to the shape's own axes"
    # from every model that works only because the child happened to be square.
    + _sweep("wide", (4, 1), [0, 30, 45, 60, 90], cy=BOX // 2)
    # Neither factor equal to 1, so a swap cannot hide inside an identity.
    + _sweep("s42", (4, 2), [30, 60])
    # Flips, at an angle on each side of the threshold.
    + _sweep("fh", (4, 1), [30, 60], flip_h=True)
    + _sweep("fv", (4, 1), [30, 60], flip_v=True)
    + _sweep("fhv", (4, 1), [30, 60], flip_h=True, flip_v=True)
)

#: The third deck.  Two things the sweep left open.
#:
#: *The tie.*  The sweep put the swap exactly on ``[45, 135)`` of the rotation reduced
#: modulo 180 -- but the first deck had shown a flip changing the answer at exactly 45,
#: which that interval cannot explain.  These probes ask whether a flip negates the angle
#: the test is made on (which would move the tie to the other end of the interval) or
#: whether something else is going on.
#:
#: *Line width.*  Whether a group's scale reaches the pen, which decides whether the
#: scale can be folded into each child's extents or has to stay an SVG ``scale()``.
TIE_PROBES: list[dict] = (
    _sweep("tie-fh", (4, 1), [135, 315], flip_h=True)
    + _sweep("tie-fv", (4, 1), [135], flip_v=True)
    + _sweep("tie-fhv", (4, 1), [45, 135], flip_h=True, flip_v=True)
    + _sweep("tie", (4, 1), [405])
    + [
        dict(key="pen-control", scale=(4, 1), rot=0, grouped=False, line_emu=76200),
        dict(key="pen-x4", scale=(4, 1), rot=0, line_emu=76200),
        dict(key="pen-x4y4", scale=(4, 4), rot=0, line_emu=76200),
        dict(key="pen-x4-rot30", scale=(4, 1), rot=30, line_emu=76200),
    ]
)

DECKS = {"probe": PROBES, "sweep": SWEEP_PROBES, "tie": TIE_PROBES}


def _content(probe, shape_id, is_text):
    """The probe shape, in child coordinates."""
    cx = probe.get("cx", BOX)
    cy = probe.get("cy", BOX)
    if is_text:
        return textbox(
            shape_id, probe["key"], BOX_OFF, BOX_OFF, cx, cy // 2, "HOHOHO",
            rot=probe["rot"], flip_h=probe.get("flip_h", False),
            flip_v=probe.get("flip_v", False),
        )
    return rect(
        shape_id, probe["key"], BOX_OFF, BOX_OFF, cx, cy,
        rot=probe["rot"], flip_h=probe.get("flip_h", False),
        flip_v=probe.get("flip_v", False), line_emu=probe.get("line_emu", 0),
    )


def probe_shapes(probe, is_text):
    sx, sy = probe["scale"]
    gx, gy = ORIGIN
    gw, gh = CHILD * sx, CHILD * sy
    shape_id = 100

    if not probe.get("grouped", True):
        if is_text:
            return textbox(
                shape_id, probe["key"], gx + BOX_OFF * sx, gy + BOX_OFF * sy,
                BOX * sx, BOX * sy // 2, "HOHOHO", rot=probe["rot"],
                flip_h=probe.get("flip_h", False), flip_v=probe.get("flip_v", False),
            )
        return rect(
            shape_id, probe["key"], gx + BOX_OFF * sx, gy + BOX_OFF * sy,
            BOX * sx, BOX * sy, rot=probe["rot"],
            flip_h=probe.get("flip_h", False), flip_v=probe.get("flip_v", False),
            fill=CONTROL_FILL, line_emu=probe.get("line_emu", 0),
        )

    nest = probe.get("nest")
    if nest == "split":
        # Two 2:1 groups, one inside the other.  If the composite is associative the
        # result must equal the single 4:1 group at the same angle.
        inner = group(
            shape_id + 2, "inner", 0, 0, CHILD * 2, CHILD,
            _content(probe, shape_id, is_text),
        )
        return group(shape_id + 1, "outer", gx, gy, gw, gh, inner)
    if nest == "outer-rot45":
        # Rotation on the outside, non-uniform scale on the inside.
        inner = group(shape_id + 2, "inner", 0, 0, CHILD, CHILD,
                      _content(probe, shape_id, is_text))
        return group(shape_id + 1, "outer", gx, gy, gw, gh, inner, rot=45)
    if nest == "inner-rot45":
        # Non-uniform scale on the outside, rotation on the inside.
        inner = group(shape_id + 2, "inner", 0, 0, CHILD, CHILD,
                      _content(probe, shape_id, is_text), rot=45)
        return group(shape_id + 1, "outer", gx, gy, gw, gh, inner)

    return group(
        shape_id + 1, "grp", gx, gy, gw, gh, _content(probe, shape_id, is_text),
        rot=probe.get("group_rot", 0),
    )


def slides(probes) -> list[str]:
    out = []
    for probe in probes:
        is_text = probe["key"].startswith("text-")
        body = label(9000, probe["key"]) + probe_shapes(probe, is_text)
        if is_text and probe.get("grouped", True):
            # Keep an ungrouped, unrotated 18 pt control on the same page as every
            # grouped text probe, so cap height and advance width are read off one page.
            body += textbox(9001, "ref", 400000, 4400000, 4000000, 500000, "HOHOHO")
        out.append(slide_document(BACKDROP, body))
    return out


def main() -> int:
    which = sys.argv[2] if len(sys.argv) > 2 else "probe"
    probes = DECKS[which]
    target = Path(
        sys.argv[1] if len(sys.argv) > 1
        else os.path.expanduser(f"~/pptx2svg-oracle/group-{which}.pptx")
    )
    write_deck(SOURCE, target, slides(probes), BACKDROP)
    print(f"wrote {target} ({len(probes)} probes)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
