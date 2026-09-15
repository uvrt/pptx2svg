"""Text inside a group keeps the point size it was authored at.

A group maps its children's coordinate space onto its on-slide box with
``ext``/``chExt``, and PowerPoint applies that mapping to geometry only -- the frame
grows, the type inside it does not.  Google Slides exports lean on this hard: a title
slide authored in a 2387101-unit-wide space and placed in a 9063527 EMU box scales 3.80x,
so scaling the text with it drew a 24 pt subtitle at 91 pt, off the bottom of the slide.

The numbers asserted here are the ones read out of PowerPoint's own PDF export of a probe
deck built the same way these fixtures are; :attr:`RenderContext.group_scale` records the
measurements in full.  These tests assert the *consequence*: the markup inside a scaled
group is the markup of the same text in an equally sized ungrouped box, wrapped in the
transform that cancels the group's scale.
"""

from __future__ import annotations

import re

from pptx2svg import model as m
from pptx2svg.render.context import RenderContext
from pptx2svg.render.svg import render_slide_to_svg
from pptx2svg.render.text import compute_sp_autofit_height

SLIDE = m.SlideSize(9144000, 5143500)

#: The group's on-slide box throughout: 2000000 x 400000 EMU.
OUTER = m.Transform(offset_x=0, offset_y=0, extent_width=2000000, extent_height=400000)
#: The child space a 4x group maps onto it.
QUARTER = m.Transform(offset_x=0, offset_y=0, extent_width=500000, extent_height=100000)


def text_shape(transform: m.Transform, text: str = "Group scale probe", **body) -> m.ShapeElement:
    return m.ShapeElement(
        transform=transform,
        geometry=m.PresetGeometry(preset="rect"),
        text_body=m.TextBody(
            paragraphs=[
                m.Paragraph(
                    runs=[
                        m.TextRun(
                            text,
                            m.RunProperties(font_size=18.0, font_family="Arial"),
                        )
                    ]
                )
            ],
            body_properties=m.BodyProperties(**body),
        ),
    )


def render(*elements: m.SlideElement) -> str:
    return render_slide_to_svg(
        m.Slide(slide_number=1, elements=list(elements)), SLIDE, RenderContext()
    )


def text_markup(svg: str) -> str:
    """Just the ``<text>`` elements, so the surrounding transforms do not mask a match."""
    return "".join(re.findall(r"<text\b.*?</text>", svg, re.S))


def grouped(child_space: m.Transform, *children: m.SlideElement) -> m.GroupElement:
    return m.GroupElement(transform=OUTER, child_transform=child_space, children=list(children))


def test_a_scaled_group_draws_its_text_at_the_authored_size():
    """4x group, and the run still has to come out the size an ungrouped one would."""
    inside = render(grouped(QUARTER, text_shape(QUARTER)))
    bare = render(text_shape(OUTER))
    assert text_markup(inside) == text_markup(bare)
    # ... which only holds because the group's scale(4, 4) is cancelled on the way in.
    assert 'transform="scale(0.25, 0.25)"' in inside


def test_the_frame_scales_even_though_the_text_does_not():
    """The wrap width is the *on-slide* width, so both break into the same lines.

    This is the half of the rule that is easy to lose: undoing the scale by shrinking the
    font instead would wrap at the child-space width and produce different lines.
    """
    long_text = "alpha bravo charlie delta echo foxtrot golf hotel india juliett"
    inside = render(grouped(QUARTER, text_shape(QUARTER, long_text)))
    bare = render(text_shape(OUTER, long_text))
    assert text_markup(inside) == text_markup(bare)
    assert inside.count("<tspan") > 1


def test_nested_group_scales_compound():
    """Outer 2x and inner 3x is a 6x child space, and the text is still 18 pt."""
    half = m.Transform(offset_x=0, offset_y=0, extent_width=1000000, extent_height=200000)
    inner = m.GroupElement(
        transform=half, child_transform=QUARTER, children=[text_shape(QUARTER)]
    )
    svg = render(grouped(half, inner))
    assert text_markup(svg) == text_markup(render(text_shape(OUTER)))
    assert 'transform="scale(0.25, 0.25)"' in svg


def test_a_non_uniform_group_does_not_stretch_the_glyphs():
    """``ext``/``chExt`` differing per axis leaves the glyphs alone on both axes.

    Measured: in a group scaling 4x horizontally and 1x vertically, PowerPoint drew the
    probe string 58.67 pt wide -- exactly what the ungrouped control drew.

    A non-uniform scale is no longer emitted as an SVG ``scale()`` at all, because one
    around a rotated child composes to a shear that PowerPoint never draws (see
    :func:`~pptx2svg.render.svg.swaps_group_axes`); it is folded into each child's own
    box instead.  So there is no counter-scale left to assert, and the consequence --
    the same markup as an ungrouped box of the on-slide size -- is asserted directly,
    which is the stronger check anyway.
    """
    tall = m.Transform(offset_x=0, offset_y=0, extent_width=500000, extent_height=400000)
    svg = render(grouped(tall, text_shape(tall)))
    assert text_markup(svg) == text_markup(render(text_shape(OUTER)))
    assert "scale(" not in svg


def test_a_group_that_shrinks_its_children_leaves_the_text_alone():
    """The rule is not "never grow"; a 0.25x group keeps 18 pt too, overflowing the box.

    Note which box the comparison uses: the group shrinks a 2000000 EMU child frame onto
    a 500000 EMU slot, so the match is against an ungrouped box of *that* size -- 18 pt
    text wrapping inside 500000 EMU, not inside 2000000.
    """
    outer = m.Transform(offset_x=0, offset_y=0, extent_width=500000, extent_height=100000)
    group = m.GroupElement(
        transform=outer, child_transform=OUTER, children=[text_shape(OUTER)]
    )
    svg = render(group)
    assert text_markup(svg) == text_markup(render(text_shape(outer)))
    assert 'transform="scale(4, 4)"' in svg


def test_an_unscaled_group_wraps_the_text_in_nothing():
    """The control.  A group that only positions must leave the text markup untouched --
    including emitting no counter-scale, so the common case costs no extra element."""
    svg = render(grouped(OUTER, text_shape(OUTER)))
    assert text_markup(svg) == text_markup(render(text_shape(OUTER)))
    # The group's own `scale(1, 1)` is there; a second one around the text is not.
    assert svg.count("scale(") == 1


def test_the_group_scale_is_restored_for_the_next_sibling():
    """A shape drawn after a scaled group must not inherit the group's scale."""
    svg = render(grouped(QUARTER, text_shape(QUARTER)), text_shape(OUTER))
    assert svg.count('transform="scale(0.25, 0.25)"') == 1


def test_sp_autofit_height_is_measured_in_slide_units_and_returned_in_child_units():
    """``spAutofit`` grows the shape, whose extent is in the group's child space."""
    body = text_shape(QUARTER, "alpha bravo charlie delta echo foxtrot", auto_fit="spAutofit")
    context = RenderContext()
    ungrouped = compute_sp_autofit_height(body.text_body, OUTER, context)
    context.group_scale = (4.0, 4.0)
    inside = compute_sp_autofit_height(body.text_body, QUARTER, context)
    assert ungrouped is not None and inside is not None
    assert inside == ungrouped / 4.0


# -- End to end, through the markup a Google Slides export actually writes ---------------

#: Slide 1 of the sales template that reported this, reduced to the shapes that matter
#: and moved onto the fixture's 9144000 x 5143500 EMU slide; the group's own numbers are
#: untouched.  It maps a 2387101 x 161700 child space onto 9063527 x 613953 EMU --
#: 3.7969x on both axes -- and the subtitle inside it is authored at 24 pt.  Drawn at
#: 24 x 3.7969 = 91.1 pt it wrapped onto three lines and ran off the bottom of the slide,
#: which is the reported bug.
GOOGLE_SLIDES_GROUP = """
<p:grpSp>
  <p:nvGrpSpPr><p:cNvPr id="79" name="Google Shape;79;p13"/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr>
  <p:grpSpPr><a:xfrm>
    <a:off x="40236" y="2200000"/><a:ext cx="9063527" cy="613953"/>
    <a:chOff x="0" y="-3"/><a:chExt cx="2387101" cy="161700"/>
  </a:xfrm></p:grpSpPr>
  <p:sp>
    <p:nvSpPr><p:cNvPr id="81" name="Google Shape;81;p13"/>
      <p:cNvSpPr txBox="1"/><p:nvPr/></p:nvSpPr>
    <p:spPr>
      <a:xfrm><a:off x="1" y="-3"/><a:ext cx="2387100" cy="161700"/></a:xfrm>
      <a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln>
    </p:spPr>
    <p:txBody>
      <a:bodyPr anchor="ctr" lIns="50800" rIns="50800" tIns="50800" bIns="50800"
                wrap="square"><a:noAutofit/></a:bodyPr>
      <a:lstStyle/>
      <a:p>
        <a:pPr algn="ctr"><a:lnSpc><a:spcPct val="150000"/></a:lnSpc></a:pPr>
        <a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Arial"/></a:rPr>
          <a:t>Modern productivity for secure, high-performing teams</a:t></a:r>
      </a:p>
    </p:txBody>
  </p:sp>
</p:grpSp>
"""


def test_a_google_slides_subtitle_keeps_its_authored_size(authoring):
    """The reported case, end to end: parse, resolve and render the real idiom."""
    from deckbuilder import derive_deck

    from pptx2svg import ConvertOptions, convert_pptx_to_svg

    deck = derive_deck(authoring, shapes_xml=GOOGLE_SLIDES_GROUP)
    svg = convert_pptx_to_svg(deck, ConvertOptions(slide_numbers=[1]))[0]
    subtitle = re.search(r"<text\b(?:(?!</text>).)*?Modern.*?</text>", svg, re.S)
    assert subtitle is not None
    # 24 pt is 32 SVG px.  Scaled with the group it would have been 121.5.
    assert set(re.findall(r'font-size="([\d.]+)"', subtitle.group())) == {"32"}
    # And on one line, which is the whole point.
    assert subtitle.group().count("<tspan") == 1
