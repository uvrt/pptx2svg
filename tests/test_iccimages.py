"""A picture's ICC profile honoured when rasterising, as PowerPoint honours it.

The pictures are the probe deck's own (``tools/make_exposed_probe.py``): flat patches under
two synthetic matrix/TRC profiles, and the colours below are the samples PowerPoint's PDF
export wrote for them -- converted to sRGB.
"""

from __future__ import annotations

import base64
import io
import sys
from pathlib import Path

import pytest

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import make_exposed_probe as probe  # noqa: E402

from pptx2svg import iccimages  # noqa: E402

ADOBE = probe.matrix_profile("Probe Adobe primaries 2.2", probe.ADOBE_PRIMARIES, 563 / 256)
P3 = probe.matrix_profile("Probe P3 primaries 1.8", probe.P3_PRIMARIES, 1.8)

#: Patch centres, and PowerPoint's samples for the PNG under each profile (every one but
#: the Adobe near-black patch, which ooxml-common's test_icc.py records).
CENTRES = [(16 + 32 * (i % 4), 16 + 32 * (i // 4)) for i in range(8)]
POWERPOINT = {
    "adobe": {1: (0, 255, 0), 3: (129, 129, 129), 4: (227, 100, 42), 5: (0, 57, 91), 6: (244, 231, 0)},
    "p3": {3: (146, 146, 146), 4: (224, 113, 49), 5: (24, 78, 112), 6: (245, 235, 0), 7: (13, 13, 13)},
}


def _pixels(data: bytes):
    Image = pytest.importorskip("PIL.Image")
    image = Image.open(io.BytesIO(data))
    assert "icc_profile" not in image.info
    return [image.convert("RGB").getpixel(centre) for centre in CENTRES]


@pytest.mark.parametrize("name, profile", [("adobe", ADOBE), ("p3", P3)])
def test_a_profiled_png_is_converted_to_srgb_as_powerpoint_converts_it(name, profile):
    pixels = _pixels(iccimages.srgb_picture(probe.png(profile)))
    for index, expected in POWERPOINT[name].items():
        assert all(abs(a - b) <= 1 for a, b in zip(pixels[index], expected)), (index, pixels[index])


@pytest.mark.parametrize("name, profile", [("adobe", ADOBE), ("p3", P3)])
def test_a_profiled_jpeg_is_converted_too(name, profile):
    pytest.importorskip("PIL")
    pixels = _pixels(iccimages.srgb_picture(probe.jpeg(profile)))
    for index, expected in POWERPOINT[name].items():
        # PowerPoint re-compressed its JPEG, and ours is decoded from Pillow's quality-100
        # one: two levels either way, on patches the conversion moves by up to 27.
        assert all(abs(a - b) <= 2 for a, b in zip(pixels[index], expected)), (index, pixels[index])


def test_an_untagged_or_srgb_picture_is_left_alone():
    from ooxml_common import icc

    assert iccimages.srgb_picture(probe.png(None)) is None
    srgb = probe.matrix_profile("sRGB-like", tuple(zip(*icc.SRGB_COLORANTS)), 2.2)
    # A 2.2 gamma is not sRGB's curve: it is converted.
    assert iccimages.srgb_picture(probe.png(srgb)) is not None
    assert iccimages.srgb_picture(b"GIF89a") is None


def test_a_palette_png_has_only_its_palette_converted():
    import struct
    import zlib

    def chunk(tag, body):
        return struct.pack(">I", len(body)) + tag + body + struct.pack(">I", zlib.crc32(tag + body))

    png = (b"\x89PNG\r\n\x1a\n" + chunk(b"IHDR", struct.pack(">IIBBBBB", 2, 1, 8, 3, 0, 0, 0))
           + chunk(b"iCCP", b"p\x00\x00" + zlib.compress(ADOBE))
           + chunk(b"PLTE", bytes([200, 100, 50, 128, 128, 128]))
           + chunk(b"IDAT", zlib.compress(b"\x00\x00\x01")) + chunk(b"IEND", b""))
    converted = iccimages.srgb_picture(png)
    assert bytes([227, 100, 42, 129, 129, 129]) in converted
    assert b"iCCP" not in converted
    assert zlib.compress(b"\x00\x00\x01") in converted


def test_the_rasteriser_is_handed_the_converted_picture():
    """The SVG keeps the picture as it came (a colour-managed reader converts it);
    ``svg_to_png`` hands resvg the converted one."""
    pytest.importorskip("resvg_py")
    Image = pytest.importorskip("PIL.Image")
    from pptx2svg.png import svg_to_png

    encoded = base64.b64encode(probe.png(ADOBE)).decode()
    svg = ('<svg xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" '
           'width="128" height="64"><image width="128" height="64" '
           f'xlink:href="data:image/png;base64,{encoded}"/></svg>')
    assert iccimages.srgb_images(svg) != svg
    assert iccimages.srgb_images(svg.replace(encoded, base64.b64encode(probe.png(None)).decode())) == \
        svg.replace(encoded, base64.b64encode(probe.png(None)).decode())
    image = Image.open(io.BytesIO(svg_to_png(svg, use_bundled_fonts=False)))
    assert image.convert("RGB").getpixel(CENTRES[4]) == (227, 100, 42)
