"""Pictures converted to sRGB through the ICC profile they carry, before rasterising.

**What PowerPoint does.**  A picture can carry a colour profile -- a PNG's ``iCCP``, a
JPEG's ``APP2`` -- and PowerPoint honours it.  ``tools/make_exposed_probe.py`` gave it
flat patches under two synthetic matrix/TRC profiles (Adobe RGB's primaries at gamma
563/256, Display P3's at 1.8): its PDF export wrote both the PNG and the JPEG converted
to sRGB, within a level of the exact conversion (ROADMAP.md 0.5, *What it exposed*).  A
picture it passes through untouched instead keeps its profile in the PDF, and the reader
converts it -- ``real-college-template``'s Adobe RGB photograph.

**What the SVG does.**  It embeds the picture's own bytes, profile and all, and SVG says
a picture's embedded profile is to be honoured (SVG 1.1, ``color-profile: auto``): a
browser converts it, as a PDF reader does.  That is exact, costs nothing, and is left as
it is.  The standard library cannot decode a JPEG, so converting the picture *in* the SVG
was never on the table for the pictures that matter most.

**What the rasteriser does not.**  resvg reads the samples as sRGB whatever the profile
says -- on that photograph a mean of 4 levels and a maximum of 39 off PowerPoint's, the
blue channel 3 low on average.  So :func:`srgb_images` is what :func:`pptx2svg.png.
svg_to_png` hands the rasteriser instead: each embedded picture whose profile is a
matrix/TRC one other than sRGB is decoded, converted exactly
(:class:`ooxml_common.icc.ToSRGB`: mean 0.03 of a level from little CMS, at most 1) and
re-embedded as an untagged PNG.

* **PNG**: decoded here, standard library only -- 8-bit truecolour with or without alpha,
  and palettes (only the palette is converted); 16-bit and interlaced ones are left alone.
* **JPEG**: needs a decoder, so it is converted when Pillow is importable (the
  ``fidelity`` tooling has it) and left as resvg would draw it when it is not.
* A profile that is sRGB already, that is not RGB, or whose conversion is a lookup table
  rather than a matrix is left alone: :func:`ooxml_common.icc.parse` refuses the last
  two rather than approximating them.
"""

from __future__ import annotations

import base64
import re
import struct
import zlib

from ooxml_common import icc
from ooxml_common.imagemeta import icc_profile

_DATA_URI = re.compile(r"data:image/(png|jpeg|jpg);base64,([A-Za-z0-9+/=]+)")

#: Converted pictures by their base64 text: a deck repeats its pictures across slides.
_CACHE: dict[str, str | None] = {}
_CACHE_LIMIT = 64


def srgb_images(svg: str) -> str:
    """``svg`` with every embedded picture that carries a non-sRGB matrix/TRC profile
    replaced by the same picture converted to sRGB (see the module docstring)."""
    if "base64," not in svg:
        return svg

    def replace(match: re.Match) -> str:
        encoded = match.group(2)
        if encoded not in _CACHE:
            if len(_CACHE) >= _CACHE_LIMIT:
                _CACHE.clear()
            converted = srgb_picture(base64.b64decode(encoded))
            _CACHE[encoded] = None if converted is None else base64.b64encode(converted).decode("ascii")
        found = _CACHE[encoded]
        if found is None:
            return match.group(0)
        return f"data:image/png;base64,{found}"

    return _DATA_URI.sub(replace, svg)


def srgb_picture(data: bytes) -> bytes | None:
    """``data`` converted to sRGB as a PNG, or ``None`` when there is nothing to convert
    or it cannot be converted here."""
    profile = icc.parse(icc_profile(data))
    if profile is None or icc.is_srgb(profile):
        return None
    transform = icc.ToSRGB(profile)
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return _png(data, transform)
    if data.startswith(b"\xff\xd8"):
        return _jpeg(data, transform)
    return None


def _jpeg(data: bytes, transform: icc.ToSRGB) -> bytes | None:
    try:
        import io

        from PIL import Image
    except Exception:  # pragma: no cover - exercised only without Pillow
        return None
    try:
        image = Image.open(io.BytesIO(data))
        if image.mode != "RGB":
            return None
        pixels = image.tobytes()
        width, height = image.size
    except Exception:
        return None
    return encode_png(width, height, transform.convert_pixels(pixels, 3), 3)


# --------------------------------------------------------------------------------------
# PNG, both ways
# --------------------------------------------------------------------------------------


def _chunks(data: bytes):
    offset = 8
    while offset + 8 <= len(data):
        (length,) = struct.unpack_from(">I", data, offset)
        yield data[offset + 4 : offset + 8], data[offset + 8 : offset + 8 + length]
        offset += 12 + length


def _png(data: bytes, transform: icc.ToSRGB) -> bytes | None:
    header = palette = None
    compressed = []
    for tag, body in _chunks(data):
        if tag == b"IHDR" and len(body) >= 13:
            header = struct.unpack_from(">IIBBBBB", body)
        elif tag == b"PLTE":
            palette = body
        elif tag == b"IDAT":
            compressed.append(body)
    if header is None:
        return None
    width, height, depth, colour, _, _, interlace = header
    if colour == 3:
        # A palette picture: only its colours need converting, and the rest is kept.
        if palette is None:
            return None
        converted = transform.convert_pixels(palette[: len(palette) // 3 * 3], 3)
        out = [b"\x89PNG\r\n\x1a\n"]
        for tag, body in _chunks(data):
            if tag in (b"iCCP", b"gAMA", b"cHRM", b"sRGB"):
                continue
            out.append(_chunk(tag, converted if tag == b"PLTE" else body))
        return b"".join(out)
    channels = {2: 3, 6: 4}.get(colour)
    if channels is None or depth != 8 or interlace:
        return None
    try:
        raw = zlib.decompress(b"".join(compressed))
    except zlib.error:
        return None
    pixels = _unfilter(raw, width, height, channels)
    if pixels is None:
        return None
    # A truecolour tRNS keys one exact colour, which the conversion moves: it is dropped
    # with the rest of the ancillary chunks, and that colour draws opaque.
    return encode_png(width, height, transform.convert_pixels(pixels, channels), channels)


def _unfilter(raw: bytes, width: int, height: int, channels: int) -> bytes | None:
    stride = width * channels
    if len(raw) < height * (stride + 1):
        return None
    out = bytearray(height * stride)
    previous = bytearray(stride)
    for row in range(height):
        start = row * (stride + 1)
        kind = raw[start]
        line = bytearray(raw[start + 1 : start + 1 + stride])
        if kind == 1:
            for i in range(channels, stride):
                line[i] = (line[i] + line[i - channels]) & 0xFF
        elif kind == 2:
            line = bytearray((a + b) & 0xFF for a, b in zip(line, previous))
        elif kind == 3:
            for i in range(stride):
                left = line[i - channels] if i >= channels else 0
                line[i] = (line[i] + ((left + previous[i]) >> 1)) & 0xFF
        elif kind == 4:
            for i in range(stride):
                a = line[i - channels] if i >= channels else 0
                b = previous[i]
                c = previous[i - channels] if i >= channels else 0
                p = a + b - c
                pa, pb, pc = abs(p - a), abs(p - b), abs(p - c)
                line[i] = (line[i] + (a if pa <= pb and pa <= pc else b if pb <= pc else c)) & 0xFF
        elif kind != 0:
            return None
        out[row * stride : (row + 1) * stride] = line
        previous = line
    return bytes(out)


def encode_png(width: int, height: int, pixels: bytes, channels: int) -> bytes:
    """Packed RGB (``channels`` 3) or RGBA (4) as an untagged 8-bit PNG, every row
    unfiltered, at a fixed zlib level so the bytes are the same everywhere."""
    stride = width * channels
    raw = b"".join(b"\x00" + pixels[row * stride : (row + 1) * stride] for row in range(height))
    header = struct.pack(">IIBBBBB", width, height, 8, 6 if channels == 4 else 2, 0, 0, 0)
    return b"".join((b"\x89PNG\r\n\x1a\n", _chunk(b"IHDR", header),
                     _chunk(b"IDAT", zlib.compress(raw, 6)), _chunk(b"IEND", b"")))


def _chunk(tag: bytes, body: bytes) -> bytes:
    return struct.pack(">I", len(body)) + tag + body + struct.pack(">I", zlib.crc32(tag + body) & 0xFFFFFFFF)
