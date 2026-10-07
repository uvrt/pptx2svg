"""The agent view (:mod:`pptx2svg.agent`): a compact SVG for a model to read."""

from __future__ import annotations

import re
import xml.etree.ElementTree as ET

import pytest

from deckbuilder import derive_deck
from pptx2svg import ConvertOptions, convert_pptx_to_agent_svg, convert_pptx_to_svg
from pptx2svg.opc import OpcPackage
from pptx2svg.parse.parts import read_presentation
from pptx2svg.resolve import resolve_presentation

from conftest import FIXTURE_DIR

SVG = "{http://www.w3.org/2000/svg}"
NUMBER_ATTRIBUTES = ("x", "y", "width", "height", "x1", "y1", "x2", "y2", "cx", "cy", "rx",
                     "ry", "dy", "font-size", "stroke-width")

FILLED_SHAPES = """
<p:sp><p:nvSpPr><p:cNvPr id="901" name="Accent box"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
 <p:spPr><a:xfrm><a:off x="1270000" y="635000"/><a:ext cx="2540000" cy="1270000"/></a:xfrm>
  <a:prstGeom prst="chevron"><a:avLst/></a:prstGeom>
  <a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>
  <a:ln w="25400"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:prstDash val="dash"/></a:ln>
 </p:spPr>
 <p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r><a:rPr lang="en-US" sz="1400" b="1"/>
  <a:t>Discover &amp; plan</a:t></a:r></a:p></p:txBody></p:sp>
<p:sp><p:nvSpPr><p:cNvPr id="902" name="Dot"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
 <p:spPr><a:xfrm rot="2700000"><a:off x="5080000" y="635000"/><a:ext cx="635000" cy="635000"/></a:xfrm>
  <a:prstGeom prst="ellipse"><a:avLst/></a:prstGeom>
  <a:solidFill><a:srgbClr val="123456"/></a:solidFill></p:spPr></p:sp>
"""


def _shapes_deck() -> bytes:
    return derive_deck(FIXTURE_DIR / "real-basic-theme.pptx", shapes_xml=FILLED_SHAPES)


def _numbers_are_rounded(svg: str) -> None:
    root = ET.fromstring(svg)
    for node in root.iter():
        for name in NUMBER_ATTRIBUTES:
            value = node.get(name)
            if value is None:
                continue
            assert re.fullmatch(r"-?\d+(\.\d)?", value), (node.tag, name, value)


def test_every_slide_shape_is_addressed_as_the_normal_render_addresses_it():
    deck = FIXTURE_DIR / "real-financial-report.pptx"
    agent = convert_pptx_to_agent_svg(deck)
    normal = convert_pptx_to_svg(deck, ConvertOptions(host_fonts=False))
    assert len(agent) == len(normal) == 4
    for compact, full in zip(agent, normal):
        ids = re.findall(r'data-pptx-id="([^"]*)"', compact)
        drawn = [i for i in re.findall(r'data-pptx-id="([^"]*)"', full) if "/" not in i]
        # The same shapes in the same document order; the normal render may draw a
        # chart's pieces too, which the agent view replaces with one placeholder.
        assert [i for i in drawn if i in ids] == ids
        assert ids


def test_the_view_is_in_slide_points_and_carries_no_fonts_images_or_base64():
    for deck in ("real-financial-report.pptx", "feature-sweep.pptx", "authoring-integration.pptx"):
        for svg in convert_pptx_to_agent_svg(FIXTURE_DIR / deck):
            assert svg.startswith('<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 ')
            for forbidden in ("<image", "base64", "@font-face", "<defs", "<filter", "<style",
                              "<path", "<clipPath"):
                assert forbidden not in svg, (deck, forbidden)
            _numbers_are_rounded(svg)
    assert 'viewBox="0 0 960 540"' in convert_pptx_to_agent_svg(FIXTURE_DIR / "sample.pptx")[0]


def test_pictures_charts_and_tables_are_placeholders_with_their_address():
    sweep = "".join(convert_pptx_to_agent_svg(FIXTURE_DIR / "feature-sweep.pptx"))
    pictures = re.findall(r"<rect[^>]*data-kind=\"picture\"[^>]*/>", sweep)
    assert len(pictures) >= 5
    assert all("data-pptx-id=" in rect for rect in pictures)
    report = convert_pptx_to_agent_svg(FIXTURE_DIR / "real-financial-report.pptx")[1]
    assert 'data-kind="chart"' in report and 'data-chart="' in report
    table = re.search(r'<g data-pptx-id="[^"]+" data-kind="table">.*?</g>', report).group(0)
    assert 'data-rows="6"' in table and 'data-cell="0,0"' in table


def test_theme_colours_keep_their_names_beside_the_hex():
    svg = convert_pptx_to_agent_svg(_shapes_deck(), ConvertOptions(slide_numbers=[1]))[0]
    chevron = re.search(r'<g data-pptx-id="[^"]*\.901">.*?</g>', svg).group(0)
    assert 'data-fill="accent1 lumMod=75%"' in chevron
    assert 'data-stroke="tx1"' in chevron and 'stroke-width="2"' in chevron
    assert 'data-dash="dash"' in chevron and 'data-preset="chevron"' in chevron
    assert ">Discover &amp; plan</tspan>" in chevron and 'font-weight="bold"' in chevron
    assert 'font-size="14"' in chevron and 'text-anchor="middle"' in chevron
    dot = re.search(r'<ellipse data-pptx-id="[^"]*\.902"[^>]*/>', svg).group(0)
    assert 'fill="#123456"' in dot and "data-fill" not in dot  # not a theme colour
    assert 'transform="rotate(45 ' in dot
    assert 'cx="425"' in dot and 'rx="25"' in dot  # 400 pt + 25 pt, in points


def test_inherited_decoration_is_marked_with_its_layer():
    svg = convert_pptx_to_agent_svg(FIXTURE_DIR / "real-basic-theme.pptx")[0]
    layers = re.findall(r'data-pptx-id="(lay|mst):[^"]*" data-layer="(layout|master)"', svg)
    assert layers and all({"lay": "layout", "mst": "master"}[a] == b for a, b in layers)


def test_smartart_is_one_placeholder_with_its_nodes_text():
    from test_diagram import deck_with_diagram

    svg = convert_pptx_to_agent_svg(deck_with_diagram(FIXTURE_DIR / "real-basic-theme.pptx"),
                                    ConvertOptions(slide_numbers=[1]))[0]
    group = re.search(r'<g data-pptx-id="[^"]*\.7001" data-kind="smartart">.*?</g>', svg)
    assert group is not None
    assert "/" not in "".join(re.findall(r'data-pptx-id="([^"]*)"', group.group(0)))
    assert "<text" in group.group(0)


def test_it_is_much_smaller_than_the_normal_render():
    deck = FIXTURE_DIR / "real-financial-report.pptx"
    agent = convert_pptx_to_agent_svg(deck)
    normal = convert_pptx_to_svg(deck, ConvertOptions(host_fonts=False))
    assert sum(map(len, agent)) < 0.5 * sum(map(len, normal))


def test_naming_colours_does_not_change_the_normal_render():
    deck = _shapes_deck()
    package = OpcPackage.open(deck)
    plain = resolve_presentation(package, read_presentation(package), slide_numbers=[1])
    names: dict = {}
    named = resolve_presentation(package, read_presentation(package), slide_numbers=[1],
                                 color_names=names)
    assert names
    assert repr(plain.slides) == repr(named.slides)


@pytest.mark.parametrize("deck", ["sample-cjk.pptx", "chart-gallery.pptx", "table test.pptx"])
def test_every_view_is_well_formed_xml(deck):
    for svg in convert_pptx_to_agent_svg(FIXTURE_DIR / deck):
        ET.fromstring(svg)
