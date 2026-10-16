"""Template views: a slide layout or master drawn without a slide (``ConvertOptions.templates``).

Every deck here is ``sample.pptx`` with standard-library zip edits: a decoration on the
master, a layout that hides the master's shapes, a layout background, an empty
placeholder.  The wireframe is this library's own drawing, so the tests check what it
promises -- what a slide on the layout starts as, and every placeholder outlined and
tagged -- not a PowerPoint rendering.
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_model, convert_pptx_to_svg

SAMPLE = Path(__file__).parent / "fixtures" / "sample.pptx"
MASTER = "ppt/slideMasters/slideMaster1.xml"
CONTENT = "ppt/slideLayouts/slideLayout2.xml"   # "Title and Content"
BLANK = "ppt/slideLayouts/slideLayout7.xml"     # "Blank": the master's footers only

#: A made-up decoration: a full-width band in a colour nothing else in the deck uses.
BAND = (
    '<p:sp><p:nvSpPr><p:cNvPr id="90" name="Band"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>'
    '<p:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="9144000" cy="228600"/></a:xfrm>'
    '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom>'
    '<a:solidFill><a:srgbClr val="12AB34"/></a:solidFill></p:spPr></p:sp>'
)


def _deck(edits: dict[str, "callable"] | None = None) -> bytes:
    edits = edits or {}
    source = zipfile.ZipFile(SAMPLE)
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as out:
        for name in source.namelist():
            data = source.read(name)
            if name in edits:
                data = edits[name](data.decode("utf-8")).encode("utf-8")
            out.writestr(name, data)
    return buffer.getvalue()


def _band_on_master(xml: str) -> str:
    return xml.replace("</p:spTree>", f"{BAND}</p:spTree>", 1)


def _view(data: bytes, *parts: str, slides=()) -> list[str]:
    return convert_pptx_to_svg(data, ConvertOptions(slide_numbers=list(slides),
                                                    templates=list(parts)))


def test_a_layout_view_outlines_and_tags_every_placeholder():
    svg = _view(_deck(), CONTENT)[0]
    assert 'data-template="layout"' in svg and f'data-template-part="{CONTENT}"' in svg
    assert 'data-template-name="Title and Content"' in svg
    # Title, content, date, footer, slide number: each a dashed grey box.
    assert svg.count('stroke-dasharray') == 5
    for tag in ("title", "content 1", "dt 10", "ftr 11", "sldNum 12"):
        assert f">{tag}</tspan>" in svg
    assert "Layout “Title and Content” · placeholders outlined" in svg
    # The prompts, in the style a slide's text there inherits.
    assert "Click to edit Master title style" in svg and "Second level" in svg


def test_a_template_view_is_no_slide_and_comes_after_the_slides():
    resolved = convert_pptx_to_model(_deck(), ConvertOptions(slide_numbers=[1],
                                                             templates=[CONTENT, MASTER]))
    assert [(s.slide_number, s.template, s.template_part) for s in resolved.slides] == [
        (1, None, None), (0, "layout", CONTENT), (0, "master", MASTER)]
    assert resolved.slides[1].template_name == "Title and Content"
    slide_svg = convert_pptx_to_svg(_deck(), ConvertOptions(slide_numbers=[1]))[0]
    assert "data-template" not in slide_svg


def test_the_masters_shapes_show_on_a_layout_unless_it_hides_them():
    shown = _view(_deck({MASTER: _band_on_master}), CONTENT)[0]
    assert "#12ab34" in shown.lower()
    hidden = _view(_deck({MASTER: _band_on_master,
                          CONTENT: lambda xml: xml.replace("<p:sldLayout ",
                                                           '<p:sldLayout showMasterSp="0" ', 1)}),
                   CONTENT)[0]
    assert "#12ab34" not in hidden.lower()


def test_a_master_view_draws_its_own_shapes_and_placeholders():
    svg = _view(_deck({MASTER: _band_on_master}), MASTER)[0]
    assert 'data-template="master"' in svg
    assert "#12ab34" in svg.lower()
    assert ">title</tspan>" in svg and ">body 1</tspan>" in svg
    # The sample's master has no cSld name: its theme's is shown.
    assert "Master “Office Theme”" in svg


def test_a_layout_background_is_drawn():
    def background(xml: str) -> str:
        return xml.replace("<p:cSld name=\"Title and Content\">",
                           '<p:cSld name="Title and Content"><p:bg><p:bgPr><a:solidFill>'
                           '<a:srgbClr val="FEDCBA"/></a:solidFill><a:effectLst/></p:bgPr>'
                           "</p:bg>", 1)

    svg = _view(_deck({CONTENT: background}), CONTENT)[0]
    assert re.search(r'<rect width="[\d.]+" height="[\d.]+" fill="#fedcba"', svg.lower())


def test_an_empty_placeholder_is_outlined_too():
    def empty_title(xml: str) -> str:
        return xml.replace("<a:t>Click to edit Master title style</a:t>", "<a:t></a:t>", 1)

    svg = _view(_deck({CONTENT: empty_title}), CONTENT)[0]
    assert svg.count('stroke-dasharray') == 5 and ">title</tspan>" in svg


def test_a_slide_number_field_shows_its_own_text():
    svg = _view(_deck(), BLANK)[0]
    assert "‹#›" in svg
    assert ">sldNum 12</tspan>" in svg


def test_a_part_that_is_no_layout_or_master_is_refused():
    with pytest.raises(ValueError, match="not a slide layout or master"):
        _view(_deck(), "ppt/slides/slide1.xml")


def test_a_layout_view_rasterises():
    pytest.importorskip("resvg_py")
    from pptx2svg import convert_pptx_to_png

    images = convert_pptx_to_png(_deck(), ConvertOptions(slide_numbers=[], templates=[CONTENT],
                                                         width=400))
    assert len(images) == 1 and images[0][:8] == b"\x89PNG\r\n\x1a\n"
