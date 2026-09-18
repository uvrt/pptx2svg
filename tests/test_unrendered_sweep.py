"""Features the parser used to read and the renderer used to throw away.

Each of these was found by a *coverage* sweep rather than by anything looking wrong:
nothing in the corpus exercises them, so there was no fidelity signal and no snapshot to
move.  That is precisely the shape of defect this file exists to stop coming back --
``a:clrChange`` sat parsed, resolved onto ``BlipEffects.clr_change`` and read by nothing
for as long as the field existed, wearing an excuse in ``UNRENDERED_FIELDS`` that said it
had no SVG equivalent.  It has one.

The decks here are *derived* with :mod:`tests.deckbuilder` rather than committed, so the
markup under test is visible in the test that uses it.  ``tests/fixtures/feature-sweep``
covers the same ground as a real deck PowerPoint can open and export; these are the unit
half, which runs without PowerPoint and pins the exact output.
"""

from __future__ import annotations

import base64
import re
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).parent))

from deckbuilder import derive_deck  # noqa: E402

from pptx2svg import ConvertOptions, convert_pptx_to_model, convert_pptx_to_svg  # noqa: E402
from pptx2svg import model as m  # noqa: E402
from pptx2svg.parse.drawing import parse_blip_effects, parse_line, parse_svg_blip_rel_id  # noqa: E402
from pptx2svg.parse.shapes import parse_alternate_content  # noqa: E402
from pptx2svg.parse.source import SourceBlipBullet, SourceImage, SourceUnsupported  # noqa: E402
from pptx2svg.parse.text import parse_bullet, parse_paragraph_properties  # noqa: E402
from pptx2svg.xmlutil import parse_xml  # noqa: E402

A = 'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"'
R = 'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"'
P = 'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"'
MC = 'xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
ASVG = 'xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG"'

#: A one-pixel PNG, and the smallest SVG that draws something.  Both are inline so the
#: test says what it embedded rather than pointing at a binary nobody reads.
PNG_1PX = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=="
)
SVG_MARK = (
    b'<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 10 10">'
    b'<circle cx="5" cy="5" r="4" fill="#ff8800"/></svg>'
)


# --------------------------------------------------------------------------------------
# a:clrChange -- replacing one colour with another
# --------------------------------------------------------------------------------------


def test_clr_change_parses_both_ends():
    blip = parse_xml(
        f"<a:blip {A}><a:clrChange>"
        '<a:clrFrom><a:srgbClr val="FF0000"/></a:clrFrom>'
        '<a:clrTo><a:srgbClr val="0000FF"/></a:clrTo>'
        "</a:clrChange></a:blip>"
    )
    effects = parse_blip_effects(blip)
    assert effects is not None
    assert effects.clr_change is not None
    assert effects.clr_change[0].hex == "FF0000"
    assert effects.clr_change[1].hex == "0000FF"


def test_clr_change_renders_as_a_keyed_colour_matrix_and_composite():
    """The recipe, pinned.

    ``feComponentTransfer`` cannot express it: whether a pixel is the source colour is a
    fact about three channels at once and a transfer function sees one at a time.  So the
    filter subtracts the key colour both ways round (a filter clamps at 0, so the two
    clamped halves add up to the absolute difference a single linear matrix cannot
    produce), sums them with an arithmetic composite, and turns that distance into an
    alpha mask over a flat fill of the replacement colour.
    """
    from pptx2svg.render.context import RenderContext
    from pptx2svg.render.effect import render_blip_effects
    from pptx2svg.text.measure import DefaultTextMeasurer

    context = RenderContext(
        measurer=DefaultTextMeasurer(), font_mapping={}, jpan_fallback_font=None
    )
    effects = m.BlipEffects(
        clr_change=m.ClrChangeEffect(
            clr_from=m.ResolvedColor(hex="#ff0000"), clr_to=m.ResolvedColor(hex="#0000ff")
        )
    )
    assert render_blip_effects(effects, context).startswith('filter="url(#blip-')
    filters = "".join(context.defs)

    # Both halves of the absolute difference, keyed on full red.
    assert 'values="1 0 0 0 -1  0 1 0 0 0  0 0 1 0 0  0 0 0 0 1"' in filters
    assert 'values="-1 0 0 0 1  0 -1 0 0 0  0 0 -1 0 0  0 0 0 0 1"' in filters
    # Summed, not multiplied: k2 and k3 both 1, k1 and k4 zero.
    assert 'operator="arithmetic" k1="0" k2="1" k3="1" k4="0"' in filters
    # The replacement colour comes out of the constant column, and the alpha row
    # subtracts the summed distance so only an exact match survives.
    assert 'values="0 0 0 0 0  0 0 0 0 0  0 0 0 0 1  -64 -64 -64 0 1"' in filters
    assert '<feComposite in="clrChangeKeyed" in2="SourceGraphic" operator="over"/>' in filters


def test_resvg_actually_runs_the_clr_change_recipe():
    """Empirical, not assumed -- the same way ``feDiffuseLighting`` was checked.

    A filter this project cannot verify is worse than one it does not emit: a filter
    primitive resvg ignores fails *silently*, leaving the picture unrecoloured with
    nothing to say so.
    """
    from PIL import Image

    from pptx2svg.png import svg_to_png
    from pptx2svg.render.context import RenderContext
    from pptx2svg.render.effect import render_blip_effects
    from pptx2svg.text.measure import DefaultTextMeasurer

    context = RenderContext(
        measurer=DefaultTextMeasurer(), font_mapping={}, jpan_fallback_font=None
    )
    attr = render_blip_effects(
        m.BlipEffects(
            clr_change=m.ClrChangeEffect(
                clr_from=m.ResolvedColor(hex="#ff0000"), clr_to=m.ResolvedColor(hex="#0000ff")
            )
        ),
        context,
    )
    svg = (
        '<svg xmlns="http://www.w3.org/2000/svg" width="120" height="40" '
        'viewBox="0 0 120 40">'
        f'<defs>{"".join(context.defs)}</defs>'
        f"<g {attr}>"
        '<rect x="0" y="0" width="40" height="40" fill="#ff0000"/>'
        '<rect x="40" y="0" width="40" height="40" fill="#00ff00"/>'
        '<rect x="80" y="0" width="40" height="40" fill="#fa0505"/>'
        "</g></svg>"
    )
    image = Image.open(__import__("io").BytesIO(svg_to_png(svg, backend="resvg"))).convert("RGB")
    # The exact match is replaced; an unrelated colour and a near miss are not.
    assert image.getpixel((20, 20)) == (0, 0, 255)
    assert image.getpixel((60, 20)) == (0, 255, 0)
    assert image.getpixel((100, 20)) == (250, 5, 5)


# --------------------------------------------------------------------------------------
# a:alphaModFix -- picture opacity
# --------------------------------------------------------------------------------------


def test_alpha_mod_fix_is_read_as_an_opacity():
    blip = parse_xml(f'<a:blip {A}><a:alphaModFix amt="40000"/></a:blip>')
    effects = parse_blip_effects(blip)
    assert effects is not None and effects.alpha == pytest.approx(0.4)


def test_a_bare_alpha_mod_fix_is_not_an_effect_at_all():
    """``<a:alphaModFix/>`` defaults to 100% -- it asks for nothing.

    Google Slides writes one on every exported picture.  Recording it would put a no-op
    ``feFuncA slope="1"`` into the filter chain of each, which is output churn standing
    for no instruction.
    """
    assert parse_blip_effects(parse_xml(f"<a:blip {A}><a:alphaModFix/></a:blip>")) is None
    assert parse_blip_effects(parse_xml(f'<a:blip {A}><a:alphaModFix amt="100000"/></a:blip>')) is None


def test_alpha_mod_fix_scales_the_existing_alpha_rather_than_replacing_it():
    from pptx2svg.render.context import RenderContext
    from pptx2svg.render.effect import render_blip_effects
    from pptx2svg.text.measure import DefaultTextMeasurer

    context = RenderContext(
        measurer=DefaultTextMeasurer(), font_mapping={}, jpan_fallback_font=None
    )
    render_blip_effects(m.BlipEffects(alpha=0.4), context)
    # `linear` on the alpha channel: a PNG's own transparency survives, which an
    # `opacity` attribute on the <image> would also do but a `table` transfer would not.
    assert '<feFuncA type="linear" slope="0.4" intercept="0"/>' in "".join(context.defs)


# --------------------------------------------------------------------------------------
# a:ln@cmpd -- compound lines
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize("value", ["dbl", "thickThin", "thinThick", "tri"])
def test_compound_line_spellings_survive_the_parse(value):
    outline = parse_line(parse_xml(f'<a:ln {A} w="38100" cmpd="{value}"><a:solidFill>'
                                   '<a:srgbClr val="000000"/></a:solidFill></a:ln>'))
    assert outline is not None and outline.compound == value


def test_a_single_compound_line_is_not_worth_recording_as_a_deviation():
    outline = parse_line(parse_xml(f'<a:ln {A} cmpd="sng"/>'))
    assert outline is not None and outline.compound == "sng"


def test_a_compound_line_warns_that_it_was_flattened(basic_theme):
    """The ``chart-3d-flattened`` shape of warning: drawn, and simplified.

    SVG gives a path one stroke, centred on it, so ``dbl`` comes out as a single stroke of
    the full width -- right colour, right weight, right place, missing split.  Saying so is
    this project's standing position; the alternative is a picture that is quietly wrong.
    """
    shape = (
        f'<p:sp {P} {A}><p:nvSpPr><p:cNvPr id="900" name="Compound"/>'
        "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>"
        '<a:xfrm><a:off x="500000" y="500000"/><a:ext cx="2000000" cy="1000000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:ln w="57150" cmpd="dbl"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:ln>'
        "</p:spPr></p:sp>"
    )
    options = ConvertOptions(slide_numbers=[1])
    convert_pptx_to_model(derive_deck(basic_theme, shapes_xml=shape), options)
    warnings = [w for w in options.warnings if w.code == "line-compound-flattened"]
    assert len(warnings) == 1
    assert "dbl" in warnings[0].message


def test_the_compound_warning_fires_once_per_spelling_not_once_per_line(basic_theme):
    """A themed table puts ``cmpd`` on all four borders of every cell."""
    shapes = "".join(
        f'<p:sp {P} {A}><p:nvSpPr><p:cNvPr id="{910 + n}" name="C{n}"/>'
        "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>"
        f'<a:xfrm><a:off x="{300000 * n}" y="300000"/><a:ext cx="200000" cy="200000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
        '<a:ln w="38100" cmpd="dbl"><a:solidFill><a:srgbClr val="1F4E79"/></a:solidFill></a:ln>'
        "</p:spPr></p:sp>"
        for n in range(6)
    )
    options = ConvertOptions(slide_numbers=[1])
    convert_pptx_to_model(derive_deck(basic_theme, shapes_xml=shapes), options)
    assert len([w for w in options.warnings if w.code == "line-compound-flattened"]) == 1


# --------------------------------------------------------------------------------------
# asvg:svgBlip -- the vector original of a picture
# --------------------------------------------------------------------------------------


def test_the_svg_extension_on_a_blip_is_read():
    blip = parse_xml(
        f'<a:blip {A} {R} {ASVG} r:embed="rId2"><a:extLst>'
        '<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
        '<asvg:svgBlip r:embed="rId3"/></a:ext></a:extLst></a:blip>'
    )
    assert parse_svg_blip_rel_id(blip) == "rId3"


def test_a_blip_with_no_extension_has_no_vector():
    assert parse_svg_blip_rel_id(parse_xml(f'<a:blip {A} {R} r:embed="rId2"/>')) is None


def _deck_with_svg_picture(basic_theme: Path, *, svg_part: bytes | None = SVG_MARK) -> bytes:
    """A picture whose blip names a PNG and whose extension names the SVG it came from."""
    parts = {"ppt/media/sweep.png": PNG_1PX}
    rels = [
        ("rIdSweepPng", "http://schemas.openxmlformats.org/officeDocument/2006/"
                        "relationships/image", "../media/sweep.png"),
    ]
    if svg_part is not None:
        parts["ppt/media/sweep.svg"] = svg_part
        rels.append(
            ("rIdSweepSvg", "http://schemas.openxmlformats.org/officeDocument/2006/"
                            "relationships/image", "../media/sweep.svg")
        )
    extension = (
        '<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
        f'<asvg:svgBlip {ASVG} r:embed="rIdSweepSvg"/></a:ext></a:extLst>'
        if svg_part is not None
        else ""
    )
    shape = (
        f'<p:pic {P} {A} {R}><p:nvPicPr><p:cNvPr id="940" name="SVG picture"/>'
        "<p:cNvPicPr/><p:nvPr/></p:nvPicPr><p:blipFill>"
        f'<a:blip r:embed="rIdSweepPng">{extension}</a:blip>'
        "<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>"
        '<a:xfrm><a:off x="1000000" y="1000000"/><a:ext cx="1000000" cy="1000000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )
    return derive_deck(
        basic_theme,
        parts=parts,
        shapes_xml=shape,
        slide_relationships=rels,
        defaults={"svg": "image/svg+xml"},
    )


def _swept_picture(slide) -> m.ImageElement | None:
    for element in slide.elements:
        if isinstance(element, m.ImageElement) and element.alt_text == "SVG picture":
            return element
    return None


def test_a_picture_that_ships_an_svg_embeds_the_vector_not_the_raster(basic_theme):
    resolved = convert_pptx_to_model(
        _deck_with_svg_picture(basic_theme), ConvertOptions(slide_numbers=[1])
    )
    picture = _swept_picture(resolved.slides[0])
    assert picture is not None
    assert picture.mime_type == "image/svg+xml"
    assert base64.b64decode(picture.image_data) == SVG_MARK


def test_a_picture_with_no_svg_still_embeds_its_raster(basic_theme):
    resolved = convert_pptx_to_model(
        _deck_with_svg_picture(basic_theme, svg_part=None), ConvertOptions(slide_numbers=[1])
    )
    picture = _swept_picture(resolved.slides[0])
    assert picture is not None
    assert picture.mime_type == "image/png"


# --------------------------------------------------------------------------------------
# mc:AlternateContent -- which branch to believe
# --------------------------------------------------------------------------------------


def _alternate(choice_inner: str, fallback_inner: str) -> str:
    return (
        f"<mc:AlternateContent {MC} {P} {A} {R}>"
        f'<mc:Choice Requires="asvg">{choice_inner}</mc:Choice>'
        f"<mc:Fallback>{fallback_inner}</mc:Fallback>"
        "</mc:AlternateContent>"
    )


PIC_CHOICE = (
    '<p:pic><p:nvPicPr><p:cNvPr id="1" name="vector"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
    f'<p:blipFill><a:blip r:embed="rIdRaster"><a:extLst><a:ext uri="{{96DAC541}}">'
    f'<asvg:svgBlip {ASVG} r:embed="rIdVector"/></a:ext></a:extLst></a:blip>'
    "</p:blipFill><p:spPr/></p:pic>"
)
PIC_FALLBACK = (
    '<p:pic><p:nvPicPr><p:cNvPr id="2" name="raster"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
    '<p:blipFill><a:blip r:embed="rIdRaster"/></p:blipFill><p:spPr/></p:pic>'
)


def test_a_choice_we_can_draw_beats_the_fallback():
    """The Fallback is lossy by construction -- it is what the Choice degrades *to*."""
    node = parse_alternate_content(parse_xml(_alternate(PIC_CHOICE, PIC_FALLBACK)))
    assert isinstance(node, SourceImage)
    assert node.name == "vector"
    assert node.svg_relationship_id == "rIdVector"


def test_a_choice_we_can_only_position_loses_to_the_fallback():
    """A ``cx:chartSpace`` frame parses to :class:`SourceUnsupported`; the Fallback is a
    picture of that chart, which is strictly more than an empty rectangle."""
    chartex = (
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="3" name="chartex"/>'
        "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/><a:graphic>"
        '<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'
        "</a:graphic></p:graphicFrame>"
    )
    node = parse_alternate_content(parse_xml(_alternate(chartex, PIC_FALLBACK)))
    assert isinstance(node, SourceImage) and node.name == "raster"


def test_a_choice_in_a_namespace_we_do_not_read_loses_to_the_fallback():
    vml = '<v:shape xmlns:v="urn:schemas-microsoft-com:vml" id="4"/>'
    node = parse_alternate_content(parse_xml(_alternate(vml, PIC_FALLBACK)))
    assert isinstance(node, SourceImage) and node.name == "raster"


def test_a_positionable_choice_is_kept_when_there_is_no_fallback_at_all():
    chartex = (
        '<p:graphicFrame><p:nvGraphicFramePr><p:cNvPr id="5" name="chartex"/>'
        "<p:cNvGraphicFramePr/><p:nvPr/></p:nvGraphicFramePr><p:xfrm/><a:graphic>"
        '<a:graphicData uri="http://schemas.microsoft.com/office/drawing/2014/chartex"/>'
        "</a:graphic></p:graphicFrame>"
    )
    xml = (
        f"<mc:AlternateContent {MC} {P} {A} {R}>"
        f'<mc:Choice Requires="cx1">{chartex}</mc:Choice>'
        "</mc:AlternateContent>"
    )
    node = parse_alternate_content(parse_xml(xml))
    assert isinstance(node, SourceUnsupported) and node.what == "chartex"


def test_every_choice_is_considered_not_only_the_first():
    xml = (
        f"<mc:AlternateContent {MC} {P} {A} {R}>"
        '<mc:Choice Requires="v"><v:shape xmlns:v="urn:schemas-microsoft-com:vml"/></mc:Choice>'
        f'<mc:Choice Requires="asvg">{PIC_CHOICE}</mc:Choice>'
        f"<mc:Fallback>{PIC_FALLBACK}</mc:Fallback>"
        "</mc:AlternateContent>"
    )
    node = parse_alternate_content(parse_xml(xml))
    assert isinstance(node, SourceImage) and node.name == "vector"


# --------------------------------------------------------------------------------------
# a:buBlip -- picture bullets
# --------------------------------------------------------------------------------------


def test_a_picture_bullet_parses_to_a_relationship_id():
    p_pr = parse_xml(
        f"<a:pPr {A} {R}><a:buBlip><a:blip r:embed=\"rIdBullet\"/></a:buBlip></a:pPr>"
    )
    bullet = parse_bullet(p_pr)
    assert isinstance(bullet, SourceBlipBullet) and bullet.relationship_id == "rIdBullet"


def test_bu_none_still_beats_a_picture_bullet():
    p_pr = parse_xml(
        f"<a:pPr {A} {R}><a:buNone/><a:buBlip><a:blip r:embed=\"rIdBullet\"/></a:buBlip></a:pPr>"
    )
    assert isinstance(parse_bullet(p_pr), m.NoBullet)


def _deck_with_picture_bullet(basic_theme: Path, *, image: bytes = PNG_1PX) -> bytes:
    shape = (
        f'<p:sp {P} {A} {R}><p:nvSpPr><p:cNvPr id="950" name="Picture bullets"/>'
        "<p:cNvSpPr txBox=\"1\"/><p:nvPr/></p:nvSpPr><p:spPr>"
        '<a:xfrm><a:off x="500000" y="500000"/><a:ext cx="4000000" cy="1500000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>'
        '<a:bodyPr wrap="square"/><a:lstStyle/>'
        '<a:p><a:pPr marL="457200" indent="-457200">'
        '<a:buBlip><a:blip r:embed="rIdBullet"/></a:buBlip></a:pPr>'
        '<a:r><a:rPr lang="en-US" sz="2000"/><a:t>Bulleted by a picture</a:t></a:r></a:p>'
        "</p:txBody></p:sp>"
    )
    return derive_deck(
        basic_theme,
        parts={"ppt/media/bullet.png": image},
        shapes_xml=shape,
        slide_relationships=[
            ("rIdBullet", "http://schemas.openxmlformats.org/officeDocument/2006/"
                          "relationships/image", "../media/bullet.png")
        ],
    )


def test_a_picture_bullet_reaches_the_model_carrying_its_image(basic_theme):
    resolved = convert_pptx_to_model(
        _deck_with_picture_bullet(basic_theme), ConvertOptions(slide_numbers=[1])
    )
    bullets = [
        paragraph.properties.bullet
        for element in resolved.slides[0].elements
        for paragraph in getattr(getattr(element, "text_body", None), "paragraphs", [])
    ]
    blip = next(b for b in bullets if isinstance(b, m.BlipBullet))
    assert blip.mime_type == "image/png"
    assert base64.b64decode(blip.image_data) == PNG_1PX


def test_a_picture_bullet_is_drawn_as_an_image_beside_the_text(basic_theme):
    """``<image>`` is not a text content element, so it cannot be the ``<tspan>`` every
    other bullet is -- it is emitted as a sibling of the ``<text>`` instead."""
    svg = convert_pptx_to_svg(
        _deck_with_picture_bullet(basic_theme), ConvertOptions(width=960, slide_numbers=[1])
    )[0]
    match = re.search(r'<image x="[\d.]+" y="[\d.]+" width="([\d.]+)" height="([\d.]+)" '
                      r'preserveAspectRatio="xMidYMid meet" xlink:href="data:image/png;base64',
                      svg)
    assert match is not None, "no picture bullet in the render"
    # Square, at the run's 20 pt: the same box a glyph bullet of that size would take.
    assert match.group(1) == match.group(2)
    assert float(match.group(1)) == pytest.approx(20 * 96 / 72 * 960 / 960, abs=0.5)


def test_an_unreadable_picture_bullet_draws_no_bullet_and_says_so(basic_theme):
    """No substituted character.

    A deck that asks for a picture and silently gets a black disc has been told something
    untrue about its own content; an absent bullet plus the warning is the honest answer.
    """
    shape = (
        f'<p:sp {P} {A} {R}><p:nvSpPr><p:cNvPr id="951" name="Broken bullet"/>'
        "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>"
        '<a:xfrm><a:off x="500000" y="500000"/><a:ext cx="4000000" cy="1000000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>'
        "<a:bodyPr/><a:lstStyle/>"
        '<a:p><a:pPr marL="457200" indent="-457200">'
        '<a:buBlip><a:blip r:embed="rIdMissing"/></a:buBlip></a:pPr>'
        "<a:r><a:rPr lang=\"en-US\"/><a:t>No bullet</a:t></a:r></a:p>"
        "</p:txBody></p:sp>"
    )
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(derive_deck(basic_theme, shapes_xml=shape), options)
    assert [w.code for w in options.warnings].count("unresolved-bullet-image") == 1
    bullets = [
        paragraph.properties.bullet
        for element in resolved.slides[0].elements
        for paragraph in getattr(getattr(element, "text_body", None), "paragraphs", [])
    ]
    assert not any(isinstance(b, m.BlipBullet) for b in bullets)


# --------------------------------------------------------------------------------------
# a:buSzPts -- the absolute spelling of a bullet's size
# --------------------------------------------------------------------------------------


def test_bu_sz_pts_is_read_in_points():
    properties = parse_paragraph_properties(parse_xml(f'<a:pPr {A}><a:buSzPts val="1800"/></a:pPr>'))
    assert properties is not None and properties.bullet_size_points == 18.0


def test_bu_sz_pts_sizes_the_bullet_independently_of_the_run():
    """This is why the gap was invisible: decks that set it usually set it *equal* to the
    run's size, so the wrong answer and the right one coincide.  ``real-basic-theme`` has
    184 of them, all 13 pt over 13 pt text -- which is why no snapshot moved when this
    was fixed, and why a test has to separate the two numbers to mean anything."""
    from pptx2svg.render.text import _bullet_size_pt

    percent = m.ParagraphProperties(bullet_size_pct=50000)
    absolute = m.ParagraphProperties(bullet_size_points=9.0)
    neither = m.ParagraphProperties()

    assert _bullet_size_pt(percent, 24.0) == 12.0
    assert _bullet_size_pt(absolute, 24.0) == 9.0
    assert _bullet_size_pt(neither, 24.0) == 24.0
    # Absolute wins over relative: it needs no context to apply.
    both = m.ParagraphProperties(bullet_size_pct=50000, bullet_size_points=9.0)
    assert _bullet_size_pt(both, 24.0) == 9.0


def test_a_bullet_sized_in_points_is_drawn_at_that_size(basic_theme):
    shape = (
        f'<p:sp {P} {A}><p:nvSpPr><p:cNvPr id="960" name="Sized bullet"/>'
        "<p:cNvSpPr/><p:nvPr/></p:nvSpPr><p:spPr>"
        '<a:xfrm><a:off x="500000" y="500000"/><a:ext cx="4000000" cy="1000000"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr><p:txBody>'
        "<a:bodyPr/><a:lstStyle/>"
        '<a:p><a:pPr marL="457200" indent="-457200">'
        '<a:buSzPts val="1000"/><a:buChar char="●"/></a:pPr>'
        '<a:r><a:rPr lang="en-US" sz="4000"/><a:t>Small bullet, big text</a:t></a:r></a:p>'
        "</p:txBody></p:sp>"
    )
    svg = convert_pptx_to_svg(
        derive_deck(basic_theme, shapes_xml=shape), ConvertOptions(width=960, slide_numbers=[1])
    )[0]
    # 10 pt in a 960 px render of a 10 in slide: 10 * 96/72 = 13.33 px.  Taking the run's
    # 40 pt instead -- what this did -- would give 53.33.
    bullet = re.search(r'<tspan [^>]*font-size="([\d.]+)"[^>]*>●</tspan>', svg)
    assert bullet is not None, "no character bullet in the render"
    assert float(bullet.group(1)) == pytest.approx(13.33, abs=0.05)
