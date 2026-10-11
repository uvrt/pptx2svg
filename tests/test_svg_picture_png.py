"""An SVG picture (``asvg:svgBlip``) in PNG output, whatever unit its root's size is in.

The SVG output embeds the picture's SVG as it is, and a browser reads any unit.  resvg-py
0.5 does not: a root of ``width="72pt"`` is "SVG has an invalid size", and the picture
vanished from the PNG without a word.  CairoSVG drew it, but drew an SVG with no
``viewBox`` unscaled at the frame's size.  The page now goes to the rasteriser with each
embedded SVG's size restated in pixels as PowerPoint reads it
(:mod:`ooxml_common.svgimage`); ``tools/make_svg_size_probe.py`` is the measurement.
"""

from __future__ import annotations

import base64
import io
import re
import struct
import zipfile
import zlib

import pytest

from deckbuilder import derive_deck
from pptx2svg import ConvertOptions, convert_pptx_to_png, convert_pptx_to_svg
from pptx2svg.png import available_backends

Image = pytest.importorskip("PIL.Image")

pytestmark = pytest.mark.skipif(
    not available_backends(),
    reason="no SVG rasterizer installed (pip install pptx2svg[png])",
)

SVG = 'xmlns="http://www.w3.org/2000/svg"'


def _green_png() -> bytes:
    """A one-pixel green PNG: the fallback raster, so drawing it instead shows."""
    def chunk(kind: bytes, body: bytes) -> bytes:
        return struct.pack(">I", len(body)) + kind + body + struct.pack(">I", zlib.crc32(kind + body))
    return (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0))
            + chunk(b"IDAT", zlib.compress(b"\x00\x00\xc8\x00")) + chunk(b"IEND", b""))


#: The frame: 2 x 1 inches at (1, 1) inches on a 10 x 5.625 inch slide.
EMU = 914400
WIDTH_PX = 1000  # 100 px to the inch


def _deck(basic_theme, svg: bytes) -> bytes:
    image = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    shape = (
        '<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" '
        'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" '
        'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<p:nvPicPr><p:cNvPr id="950" name="Graphic 1"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>'
        '<p:blipFill><a:blip r:embed="rIdUnitsPng"><a:extLst>'
        '<a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">'
        '<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" '
        'r:embed="rIdUnitsSvg"/></a:ext></a:extLst></a:blip>'
        "<a:stretch><a:fillRect/></a:stretch></p:blipFill><p:spPr>"
        f'<a:xfrm><a:off x="{EMU}" y="{EMU}"/><a:ext cx="{2 * EMU}" cy="{EMU}"/></a:xfrm>'
        '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>'
    )
    deck = derive_deck(
        basic_theme,
        parts={"ppt/media/units.png": _green_png(), "ppt/media/units.svg": svg},
        shapes_xml=shape,
        slide_relationships=[("rIdUnitsPng", image, "../media/units.png"),
                             ("rIdUnitsSvg", image, "../media/units.svg")],
        defaults={"svg": "image/svg+xml"},
    )
    # Only the picture on the slide: the fixture's placeholders would sit under it.
    source, out = zipfile.ZipFile(io.BytesIO(deck)), io.BytesIO()
    with zipfile.ZipFile(out, "w") as target:
        for part in source.namelist():
            data = source.read(part)
            if part == "ppt/slides/slide1.xml":
                data = re.sub(rb"<p:sp>.*?</p:sp>", b"", data, flags=re.S)
            target.writestr(part, data)
    return out.getvalue()


def _frame(basic_theme, svg: bytes, backend: str):
    """The picture's 200 x 100 px frame, from the slide rasterised at 100 px to the inch."""
    png = convert_pptx_to_png(_deck(basic_theme, svg), ConvertOptions(width=WIDTH_PX),
                              backend=backend)[0]
    return Image.open(io.BytesIO(png)).convert("RGB").crop((100, 100, 300, 200))


def _colour(pixel) -> str:
    r, g, b = pixel
    if b > 200 and r < 60 and g < 60:
        return "blue"
    if r > 200 and g < 60 and b < 60:
        return "red"
    if g > 150 and r < 60 and b < 60:
        return "green"
    return "other"


def _quarter(size: str) -> bytes:
    """A 20x20 viewBox, red, with a blue 10x10 square at its origin."""
    return (f'<svg {SVG} {size} viewBox="0 0 20 20"><rect width="20" height="20" fill="#f00"/>'
            '<rect width="10" height="10" fill="#00f"/></svg>').encode()


BACKENDS = available_backends()


@pytest.mark.parametrize("backend", BACKENDS)
@pytest.mark.parametrize("size", [
    'width="72pt" height="72pt"',
    'width="6pc" height="6pc"',
    'width="1in" height="1in"',
    'width="2.54cm" height="2.54cm"',
    'width="25.4mm" height="25.4mm"',
    'width="100%" height="100%"',
    'width="6em" height="6em"',
    "",  # no width or height: the viewBox alone
], ids=["pt", "pc", "in", "cm", "mm", "percent", "em", "viewbox-only"])
def test_an_svg_picture_is_drawn_whatever_unit_its_size_is_in(basic_theme, size, backend):
    """Each root is square, so PowerPoint draws the viewBox stretched to the 2:1 frame:
    the blue square is the frame's top-left quarter (measured, PowerPoint 16)."""
    frame = _frame(basic_theme, _quarter(size), backend)
    assert _colour(frame.getpixel((50, 25))) == "blue"
    assert _colour(frame.getpixel((150, 25))) == "red"
    assert _colour(frame.getpixel((50, 75))) == "red"
    assert _colour(frame.getpixel((150, 75))) == "red"


@pytest.mark.parametrize("backend", BACKENDS)
@pytest.mark.parametrize("size, side", [
    ('width="72pt" height="72pt"', 48),  # 96 px: 48 units is half of it
    ('width="6em" height="6em"', 36),  # 72 px: an em is 12 px to PowerPoint, not 16
    ('width="50%" height="50%"', 125),  # 250 px: half of PowerPoint's 500
    ("", 250),  # 500 px
])
def test_an_svg_without_a_viewbox_is_drawn_at_its_own_size_then_stretched(
    basic_theme, size, side, backend
):
    """No viewBox: the drawing is not scaled to the root, so the root's size is how much
    of the drawing reaches the frame.  Each blue square is half the root's side."""
    svg = (f'<svg {SVG} {size}><rect width="10000" height="10000" fill="#f00"/>'
           f'<rect width="{side}" height="{side}" fill="#00f"/></svg>').encode()
    frame = _frame(basic_theme, svg, backend)
    assert _colour(frame.getpixel((90, 45))) == "blue"
    assert _colour(frame.getpixel((110, 55))) == "red"


def test_a_root_out_of_proportion_with_its_viewbox_is_letterboxed(basic_theme):
    """2 x 1 inches over a square viewBox: fitted inside the root, centred, as PowerPoint
    draws it.  (CairoSVG stretches it instead: it reads no root's preserveAspectRatio.)"""
    if "resvg" not in BACKENDS:
        pytest.skip("resvg-py is not installed")
    frame = _frame(basic_theme, _quarter('width="2in" height="1in"'), "resvg")
    assert _colour(frame.getpixel((25, 50))) not in ("red", "blue")
    assert _colour(frame.getpixel((75, 25))) == "blue"
    assert _colour(frame.getpixel((125, 75))) == "red"


def test_the_svg_output_still_embeds_the_svg_as_it_is(basic_theme):
    """Only what the rasteriser is handed changes."""
    svg = _quarter('width="72pt" height="72pt"')
    page = convert_pptx_to_svg(_deck(basic_theme, svg))[0]
    assert base64.b64encode(svg).decode() in page
