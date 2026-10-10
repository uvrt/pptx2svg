"""Relationship ids read from a layout, a master or a theme resolve in *that* part.

An ``r:embed`` is only meaningful next to the relationships of the part it was written
in.  A layout's ``<p:bg>`` picture, a theme's background image and a picture bullet from a
master's text styles all used to be looked up in the slide's relationships instead -- so
a title slide whose layout carries a photo rendered blank (the slide's rId2 is usually
its notes slide) or, when the ids happened to collide with a slide image, drew the wrong
picture without a word.

Every deck here is derived from ``sample.pptx`` with standard-library zip edits.  Each
case gives the *slide* a relationship with the same id, pointing at a differently
coloured picture, so a lookup in the wrong part is drawn rather than silently dropped.
"""

from __future__ import annotations

import base64
import io
import struct
import zipfile
import zlib
from pathlib import Path

from pptx2svg import convert_pptx_to_svg

SAMPLE = Path(__file__).parent / "fixtures" / "sample.pptx"
IMAGE_REL = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"

SLIDE = "ppt/slides/slide1.xml"
LAYOUT = "ppt/slideLayouts/slideLayout1.xml"
MASTER = "ppt/slideMasters/slideMaster1.xml"
THEME = "ppt/theme/theme1.xml"


def _png(rgb: tuple[int, int, int], size: int = 8) -> bytes:
    def chunk(kind: bytes, data: bytes) -> bytes:
        crc = zlib.crc32(kind + data) & 0xFFFFFFFF
        return struct.pack(">I", len(data)) + kind + data + struct.pack(">I", crc)

    rows = b"".join(b"\x00" + bytes(rgb) * size for _ in range(size))
    header = struct.pack(">IIBBBBB", size, size, 8, 2, 0, 0, 0)
    return (
        b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", header) + chunk(b"IDAT", zlib.compress(rows))
        + chunk(b"IEND", b"")
    )


RIGHT = _png((0, 0, 0))
WRONG = _png((255, 0, 0))


def _rels_path(part: str) -> str:
    folder, name = part.rsplit("/", 1)
    return f"{folder}/_rels/{name}.rels"


def _add_relationship(rels: str, rel_id: str, target: str) -> str:
    entry = f'<Relationship Id="{rel_id}" Type="{IMAGE_REL}" Target="{target}"/>'
    return rels.replace("</Relationships>", entry + "</Relationships>", 1)


def _deck(edits: dict[str, callable], images_for: dict[str, bytes]) -> bytes:
    """``sample.pptx`` with ``edits`` applied (part -> fn(xml) -> xml) and, for each part
    in ``images_for``, an ``rId90`` image relationship to the given PNG."""
    with zipfile.ZipFile(SAMPLE) as source:
        parts = {name: source.read(name) for name in source.namelist()}
    for index, (owner, png) in enumerate(images_for.items()):
        media = f"ppt/media/inherited{index}.png"
        parts[media] = png
        rels = _rels_path(owner)
        existing = parts.get(rels, b'<Relationships xmlns="http://schemas.openxmlformats.org/'
                             b'package/2006/relationships"></Relationships>')
        parts[rels] = _add_relationship(
            existing.decode(), "rId90", "../media/" + media.rsplit("/", 1)[1]
        ).encode()
    for name, edit in edits.items():
        parts[name] = edit(parts[name].decode()).encode()
    types = parts["[Content_Types].xml"].decode()
    if 'Extension="png"' not in types:
        types = types.replace(
            "<Default ", '<Default Extension="png" ContentType="image/png"/><Default ', 1
        )
        parts["[Content_Types].xml"] = types.encode()
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, payload in parts.items():
            archive.writestr(name, payload)
    return out.getvalue()


def _svg(deck: bytes) -> str:
    return convert_pptx_to_svg(deck)[0]


def _b64(png: bytes) -> str:
    return base64.b64encode(png).decode("ascii")


BG_IMAGE = (
    '<p:bg><p:bgPr><a:blipFill dpi="0" rotWithShape="1"><a:blip r:embed="rId90"/>'
    "<a:srcRect/><a:stretch><a:fillRect/></a:stretch></a:blipFill>"
    "<a:effectLst/></p:bgPr></p:bg>"
)


def _with_background(xml: str) -> str:
    return xml.replace("<p:cSld>", "<p:cSld>" + BG_IMAGE, 1).replace(
        '<p:cSld name="Title Slide">', '<p:cSld name="Title Slide">' + BG_IMAGE, 1
    )


def _without_background(xml: str) -> str:
    start = xml.find("<p:bg>")
    if start < 0:
        return xml
    return xml[:start] + xml[xml.index("</p:bg>", start) + len("</p:bg>"):]


def test_a_layout_background_image_is_read_from_the_layouts_relationships():
    deck = _deck(
        {LAYOUT: _with_background, MASTER: _without_background},
        {LAYOUT: RIGHT, SLIDE: WRONG},
    )
    svg = _svg(deck)
    assert _b64(RIGHT) in svg
    assert _b64(WRONG) not in svg


def test_a_master_background_image_is_read_from_the_masters_relationships():
    deck = _deck(
        {MASTER: lambda x: _with_background(_without_background(x))},
        {MASTER: RIGHT, SLIDE: WRONG, LAYOUT: WRONG},
    )
    svg = _svg(deck)
    assert _b64(RIGHT) in svg
    assert _b64(WRONG) not in svg


def test_a_theme_background_image_is_read_from_the_themes_relationships():
    """``p:bgRef idx="1003"`` picks the theme's third background fill -- here a picture
    related to the theme part."""
    theme_blip = (
        '<a:blipFill rotWithShape="1"><a:blip xmlns:r="http://schemas.openxmlformats.org/'
        'officeDocument/2006/relationships" r:embed="rId90"/><a:stretch><a:fillRect/>'
        "</a:stretch></a:blipFill></a:bgFillStyleLst>"
    )

    def theme(xml: str) -> str:
        head, tail = xml.split("</a:bgFillStyleLst>", 1)
        last = head.rindex("<a:gradFill")
        return head[:last] + theme_blip + tail

    def master(xml: str) -> str:
        return _without_background(xml).replace(
            "<p:cSld>",
            '<p:cSld><p:bg><p:bgRef idx="1003"><a:schemeClr val="bg1"/></p:bgRef></p:bg>',
            1,
        )

    deck = _deck({THEME: theme, MASTER: master}, {THEME: RIGHT, SLIDE: WRONG})
    svg = _svg(deck)
    assert _b64(RIGHT) in svg
    assert _b64(WRONG) not in svg


def test_a_picture_bullet_from_the_layouts_list_style_is_read_from_the_layout():
    def layout(xml: str) -> str:
        # The subtitle's first level: `buNone` becomes a picture bullet.
        return xml.replace(
            '<a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/>',
            '<a:lvl1pPr marL="342900" indent="-342900" algn="ctr"><a:buBlip>'
            '<a:blip r:embed="rId90"/></a:buBlip>',
            1,
        )

    deck = _deck({LAYOUT: layout}, {LAYOUT: RIGHT, SLIDE: WRONG})
    svg = _svg(deck)
    assert _b64(RIGHT) in svg
    assert _b64(WRONG) not in svg


def test_a_picture_bullet_from_the_masters_body_style_is_read_from_the_master():
    def layout(xml: str) -> str:
        # Let the subtitle fall through to the master's body style.
        return xml.replace(
            '<a:lvl1pPr marL="0" indent="0" algn="ctr"><a:buNone/>',
            '<a:lvl1pPr algn="ctr">',
            1,
        )

    def master(xml: str) -> str:
        head, tail = xml.split("<p:bodyStyle>", 1)
        tail = tail.replace(
            '<a:buFont typeface="Arial"/><a:buChar char="•"/>',
            '<a:buBlip><a:blip r:embed="rId90"/></a:buBlip>',
            1,
        )
        return head + "<p:bodyStyle>" + tail

    deck = _deck({LAYOUT: layout, MASTER: master}, {MASTER: RIGHT, SLIDE: WRONG})
    svg = _svg(deck)
    assert _b64(RIGHT) in svg
    assert _b64(WRONG) not in svg


def test_a_slides_own_background_image_still_reads_the_slides_relationships():
    deck = _deck({SLIDE: _with_background}, {SLIDE: RIGHT, LAYOUT: WRONG})
    svg = _svg(deck)
    assert _b64(RIGHT) in svg
    assert _b64(WRONG) not in svg
