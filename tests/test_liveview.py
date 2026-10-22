"""A live, editable view of a slide (round 13, :mod:`pptx2svg.liveview`): pictures by
reference, edit marks and the id contract, the slide by shape, and no hyperlinks.

Every option is opt-in; with the defaults the snapshots in ``tests/vrt`` hold the output.
The PNG of a marked, referenced or link-free SVG is the plain one's, byte for byte."""

from __future__ import annotations

import base64
import re
import tempfile
import zipfile
from pathlib import Path
from xml.etree.ElementTree import fromstring

import pytest

from pptx2svg import ConvertOptions, embed_images, rasterise_svgs, render_svgs, slide_groups
from pptx2svg.liveview import externalise_images
from pptx2svg.png import available_backends

from deckbuilder import SLIDE_PATH, derive_deck

FIXTURES = Path(__file__).parent / "fixtures"
PICTURES = FIXTURES / "feature-sweep.pptx"      # PNG and SVG pictures, some on several slides
TABLE = FIXTURES / "table test.pptx"
LINKS = FIXTURES / "sample.pptx"                # text hyperlinks
SHAPES = FIXTURES / "real-financial-report.pptx"

needs_rasteriser = pytest.mark.skipif(not available_backends(), reason="no SVG rasteriser installed")


def _options(**extra) -> ConvertOptions:
    from datetime import datetime

    return ConvertOptions(host_fonts=False, now=datetime(2026, 1, 31, 9, 0), **extra)


def _pngs(svgs, resolved, **extra) -> list[bytes]:
    numbers = [(slide.slide_number, slide.template_part) for slide in resolved.slides]
    with tempfile.TemporaryDirectory() as folder:
        files = resolved.embedded_fonts.write(folder) if resolved.embedded_fonts else None
        return rasterise_svgs(svgs, _options(), numbers=numbers, font_files=files, host_fonts=False, **extra)


# -- R1: pictures by reference ---------------------------------------------------------


def test_pictures_are_written_once_by_their_hash_and_format():
    plain, _ = render_svgs(PICTURES, _options())
    images: dict[str, bytes] = {}
    referenced, _ = render_svgs(PICTURES, _options(images=images))

    assert not any("base64," in svg for svg in referenced)
    embedded = {uri for svg in plain for uri in re.findall(r'"data:[^"]+"', svg)}
    assert len(images) == len(embedded) < sum(svg.count('"data:') for svg in plain)   # shared ones once
    for name, data in images.items():
        assert re.fullmatch(r"img/[0-9a-f]{32}\.(png|svg)", name)
        extension = name.rsplit(".", 1)[1]
        assert data.startswith(b"\x89PNG") if extension == "png" else b"<svg" in data[:512]
    hrefs = {href for svg in referenced for href in re.findall(r'href="(img/[^"]+)"', svg)}
    assert hrefs == set(images)


def test_embedding_the_references_again_gives_the_plain_svg_back():
    plain, _ = render_svgs(PICTURES, _options())
    images: dict[str, bytes] = {}
    referenced, _ = render_svgs(PICTURES, _options(images=images))
    assert [embed_images(svg, images) for svg in referenced] == plain


def test_a_prefix_of_ones_own_and_a_store_shared_across_renders():
    images: dict[str, bytes] = {}
    first, _ = render_svgs(PICTURES, _options(slide_numbers=[1], images=images, image_prefix="media/"))
    count = len(images)
    render_svgs(PICTURES, _options(slide_numbers=[1], images=images, image_prefix="media/"))
    assert len(images) == count and all(name.startswith("media/") for name in images)
    assert 'href="media/' in first[0]


def test_a_picture_the_images_lack_is_an_error_not_a_blank():
    images: dict[str, bytes] = {}
    referenced, _ = render_svgs(PICTURES, _options(slide_numbers=[1], images=images))
    name = next(iter(images))
    with pytest.raises(KeyError, match=re.escape(name)):
        embed_images(referenced[0], {})


def test_the_format_is_what_the_bytes_are():
    png = base64.b64decode("iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==")
    store: dict[str, bytes] = {}
    svg = ('<svg xmlns="http://www.w3.org/2000/svg"><image href="data:image/jpeg;base64,'
           f'{base64.b64encode(png).decode()}"/></svg>')
    out = externalise_images(svg, store)
    assert list(store) == [re.search(r'href="([^"]+)"', out).group(1)] and out.count(".png") == 1


@needs_rasteriser
def test_the_png_of_the_referenced_svg_is_the_embedded_ones():
    plain, resolved = render_svgs(PICTURES, _options(slide_numbers=[1, 6]))
    images: dict[str, bytes] = {}
    referenced, _ = render_svgs(PICTURES, _options(slide_numbers=[1, 6], images=images))
    assert _pngs(referenced, resolved, images=images) == _pngs(plain, resolved)


# -- R3: edit marks and the id contract ------------------------------------------------


def test_marks_are_off_by_default():
    svg = render_svgs(TABLE, _options())[0][0]
    for mark in ("data-pptx-size", "data-pptx-source", "data-pptx-text", "/cell"):
        assert mark not in svg
    assert "data-pptx-path" in svg


def test_the_root_says_the_slide_size_in_points():
    svg = render_svgs(TABLE, _options(edit_marks=True))[0][0]
    assert re.match(r'<svg [^>]*data-pptx-size="960x540"', svg)


def test_every_id_says_where_the_shape_comes_from():
    # The slide draws a shape of its layout's and one of its master's.
    svg = render_svgs(FIXTURES / "authoring-integration.pptx", _options(edit_marks=True))[0][0]
    ids = re.findall(r'data-pptx-id="([^"]*)" data-pptx-source="([^"]*)"', svg)
    assert ids and len(ids) == svg.count("data-pptx-id=")
    assert {source for _, source in ids} == {"slide", "layout", "master"}
    for value, source in ids:
        assert source == ("layout" if value.startswith("lay:") else "master" if value.startswith("mst:") else "slide")


def test_table_cells_carry_their_address_and_text_is_marked():
    svg = render_svgs(TABLE, _options(edit_marks=True))[0][0]
    table = re.search(r'data-pptx-id="(257\.\d+)"[^>]*>(?=<g data-pptx-id="257\.\d+/cell0,0")', svg).group(1)
    cells = re.findall(rf'data-pptx-id="{re.escape(table)}/cell(\d+),(\d+)"', svg)
    assert cells == [(str(row), str(column)) for row in range(5) for column in range(5)]
    assert svg.count('data-pptx-text="cell"') == 7      # the cells with text
    assert '<g data-pptx-text="shape">' in svg          # the title's
    fromstring(svg)


def _slide_ids(deck: Path, number: int) -> list[str]:
    """``<sldId>.<cNvPr id>`` of every shape slide ``number`` draws (groups' members, graphic
    frames included), from the part itself."""
    from pptx2svg.opc import OpcPackage
    from pptx2svg.parse.parts import read_presentation

    slide = read_presentation(OpcPackage.open(str(deck))).slides[number - 1]
    root = fromstring(zipfile.ZipFile(deck).read(slide.part_path))
    tree = next(element for element in root.iter() if element.tag.endswith("}spTree"))
    found = []
    for element in tree.iter():
        # The tree's own properties are no shape's.
        if element.tag.endswith("}cNvPr") and element.get("hidden") not in ("1", "true") \
                and element is not tree[0][0]:
            found.append(f"{slide.slide_id}.{element.get('id')}")
    return found


@pytest.mark.parametrize("deck, number", [(SHAPES, 2), (TABLE, 1), (FIXTURES / "chart-gallery.pptx", 1),
                                          (PICTURES, 6)])
def test_every_drawn_shape_carries_its_id(deck, number):
    svg = render_svgs(deck, _options(slide_numbers=[number], edit_marks=True))[0][0]
    drawn = set(re.findall(r'data-pptx-id="([^"/]*)', svg))
    assert set(_slide_ids(deck, number)) <= drawn


def _edited(deck: Path, edit) -> bytes:
    original = zipfile.ZipFile(deck).read(SLIDE_PATH).decode("utf-8")
    return derive_deck(deck, parts={SLIDE_PATH: edit(original).encode("utf-8")})


def test_ids_survive_an_edit_of_text_and_a_move():
    def edit(xml: str) -> str:
        xml = re.sub(r"<a:t>[^<]*</a:t>", "<a:t>Rewritten</a:t>", xml, count=1)
        return re.sub(r'<a:off x="(\d+)"', lambda m: f'<a:off x="{int(m.group(1)) + 127000}"', xml, count=1)

    before = render_svgs(SHAPES, _options(slide_numbers=[1], edit_marks=True))[0][0]
    after = render_svgs(_edited(SHAPES, edit), _options(slide_numbers=[1], edit_marks=True))[0][0]
    ids = re.compile(r'data-pptx-id="([^"]*)"')
    assert before != after and ids.findall(before) == ids.findall(after)


def test_ids_survive_edits_made_with_pptx_agent():
    """The same through pptx-agent, whose ``Slide.render_svg`` puts its own ids -- the ones
    its tools take -- on the groups and cells.  Skipped where it is not installed (CI)."""
    agent = pytest.importorskip("pptx_agent")
    document = agent.Document.open(TABLE.read_bytes())
    slide = document.slides[0]
    ids = re.compile(r'data-pptx-id="([^"]*)"')
    before = slide.render_svg(edit_marks=True, host_fonts=False)
    title, table = list(slide.shapes)[:2]
    title.set_text("Rewritten")
    table.move_by(127000, 0)
    after = slide.render_svg(edit_marks=True, host_fonts=False)
    assert before != after and ids.findall(before) == ids.findall(after)
    cell = next(value for value in ids.findall(after) if value.endswith("/cell1,0"))
    assert document.resolve(cell).text == "Hello"


@needs_rasteriser
def test_marks_do_not_change_a_pixel():
    plain, resolved = render_svgs(TABLE, _options())
    marked, _ = render_svgs(TABLE, _options(edit_marks=True))
    assert marked != plain and _pngs(marked, resolved) == _pngs(plain, resolved)


# -- R4: the slide by shape ------------------------------------------------------------


def test_the_slide_by_shape_reassembles_to_the_slide():
    svg = render_svgs(SHAPES, _options(slide_numbers=[1], edit_marks=True))[0][0]
    groups = slide_groups(svg)
    assert groups["order"][0] == "background" and groups["root"].startswith("<svg ")
    assert len(groups["order"]) == len(set(groups["order"]))
    defs = re.search(r"<defs>(.*?)</defs>", svg).group(1) if "<defs>" in svg else ""
    body = "".join(groups["groups"][key] for key in groups["order"])
    assert svg == f"{groups['root']}" + (f"<defs>{defs}</defs>" if defs else "") + body + "</svg>"
    # Every reference a shape makes is among its definitions.
    for key in groups["order"]:
        for name in re.findall(r"url\(#([^)]+)\)", groups["groups"][key]):
            assert f'id="{name}"' in groups["defs"][key]


def _with_gradient_on_first_shape(xml: str) -> str:
    """The first shape's solid fill made a gradient: one more definition, before the others."""
    gradient = ('<a:gradFill><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs>'
                '<a:gs pos="100000"><a:srgbClr val="0000FF"/></a:gs></a:gsLst><a:lin ang="0"/></a:gradFill>')
    head, tree, shapes = xml.partition("<p:spTree>")
    return head + tree + re.sub(r"<a:solidFill>.*?</a:solidFill>", gradient, shapes, count=1)


@pytest.mark.parametrize("marks", [True, False])
def test_an_unchanged_shape_draws_the_same_bytes(marks):
    deck = FIXTURES / "real-product-page.pptx"      # shapes with effects after the first
    before = slide_groups(render_svgs(deck, _options(edit_marks=marks))[0][0])
    again = slide_groups(render_svgs(deck, _options(edit_marks=marks))[0][0])
    assert before == again
    edited = _edited(deck, _with_gradient_on_first_shape)
    after = slide_groups(render_svgs(edited, _options(edit_marks=marks))[0][0])
    changed = [key for key in before["order"]
               if (before["groups"][key], before["defs"][key]) != (after["groups"].get(key), after["defs"].get(key))]
    if marks:
        assert changed == ["256.2"]      # the shape whose fill changed, and no other
    else:
        assert len(changed) > 1          # numbered across the slide, the definitions after it move


# -- R5: no hyperlinks ------------------------------------------------------------------


def test_hyperlinks_can_be_left_out():
    linked, _ = render_svgs(LINKS, _options())
    bare, _ = render_svgs(LINKS, _options(hyperlinks=False))
    assert any("<a " in svg for svg in linked) and not any("<a " in svg for svg in bare)
    assert [re.sub(r"<a [^>]*>|</a>", "", svg) for svg in linked] == bare
