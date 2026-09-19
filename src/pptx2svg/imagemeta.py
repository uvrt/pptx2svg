"""A picture's **natural size** -- the size it claims for itself, in points.

Only one thing in this library needs it, and it needs it badly: ``a:tile@sx`` scales the
picture's *own* size rather than the shape it fills, so a tiled fill cannot be drawn at all
without knowing what size the picture claims.  ``render/fill.py`` used to size a tile at
``sx`` of the shape's bounding box, and on ``tests/fixtures/feature-sweep.pptx`` slide 10
that drew a tile 8.3x too big.

A picture claims a size in two senses -- how many pixels it has, and how many of them it
says go in an inch -- and which one PowerPoint uses is the question the measurement below
answers.  It uses **both**::

    natural size in points = pixels * 72 / density

and when the file states **no** density, PowerPoint supplies **144 dpi**, not the 96 dpi
the rest of this library uses for the slide's own coordinate space.

Measured, not assumed, with ``tools/make_fill_probe.py`` (deck ``fill-tile``) and
``tools/read_fill_probe.py``, which read the cell straight out of the ``/XStep`` and
``/Matrix`` of the PDF tiling pattern PowerPoint exports.  Two independent sweeps:

* **Pixels, no density stated.**  8, 16, 32, 64 and 96 px square PNGs, and one 48x24,
  tiled at ``sx=sy=100%``, drew cells of 4, 8, 16, 32, 48 and 24x12 pt -- exactly half a
  point per pixel in every case, which is 144 dpi.
* **Density stated.**  The same 32 px PNG tagged ``pHYs`` 72, 96, 144 and 300 dpi drew
  32, 24, 16 and 7.68 pt -- ``32 * 72 / density`` to the last digit, and the 144 dpi row
  reproduces the untagged case exactly.

The law is the *picture's*, not the format's: the whole density sweep was repeated as
JPEGs carrying a JFIF ``APP0`` (deck ``fill-jpeg``, 32 px and 64 px at each density) and
every cell matched its PNG twin, with a JPEG written at no density -- JFIF ``units=0`` --
landing on the same 144 dpi default.

What is **not** measured here: a TIFF's or an Exif ``APP1``'s resolution tags, and EMF or
WMF, whose natural size lives in the metafile header rather than in a pixel count.  Those
return ``None`` and the caller falls back rather than guessing.
"""

from __future__ import annotations

import struct

#: PowerPoint's density for a picture that states none.  Not 96: measured at 144.
DEFAULT_DENSITY_DPI = 144.0

#: A ``pHYs`` or BMP header states pixels per *metre*; this converts to pixels per inch.
_INCHES_PER_METRE = 0.0254


def natural_size_pt(data: bytes) -> tuple[float, float] | None:
    """``(width, height)`` in points, or ``None`` for a format not read here."""
    found = _pixels_and_density(data)
    if found is None:
        return None
    width, height, density_x, density_y = found
    if not width or not height:
        return None
    return width * 72.0 / density_x, height * 72.0 / density_y


def _pixels_and_density(data: bytes) -> tuple[int, int, float, float] | None:
    if data.startswith(b"\x89PNG\r\n\x1a\n"):
        return _png(data)
    if data.startswith(b"\xff\xd8"):
        return _jpeg(data)
    if data.startswith((b"GIF87a", b"GIF89a")):
        return _gif(data)
    if data.startswith(b"BM"):
        return _bmp(data)
    return None


def _png(data: bytes) -> tuple[int, int, float, float] | None:
    width = height = 0
    density_x = density_y = DEFAULT_DENSITY_DPI
    offset = 8
    while offset + 8 <= len(data):
        (length,) = struct.unpack_from(">I", data, offset)
        tag = data[offset + 4 : offset + 8]
        body = data[offset + 8 : offset + 8 + length]
        if tag == b"IHDR" and len(body) >= 8:
            width, height = struct.unpack_from(">II", body)
        elif tag == b"pHYs" and len(body) >= 9:
            per_x, per_y, unit = struct.unpack_from(">IIB", body)
            # unit 1 is metres; unit 0 means "aspect ratio only", which states no size.
            if unit == 1 and per_x and per_y:
                density_x = per_x * _INCHES_PER_METRE
                density_y = per_y * _INCHES_PER_METRE
        elif tag == b"IDAT" or tag == b"IEND":
            break
        offset += 12 + length
    return (width, height, density_x, density_y) if width and height else None


#: Start-of-frame markers.  ``C4``/``C8``/``CC`` share the range but are tables, not frames.
_JPEG_FRAME = {0xC0, 0xC1, 0xC2, 0xC3, 0xC5, 0xC6, 0xC7, 0xC9, 0xCA, 0xCB, 0xCD, 0xCE, 0xCF}


def _jpeg(data: bytes) -> tuple[int, int, float, float] | None:
    density_x = density_y = DEFAULT_DENSITY_DPI
    offset = 2
    while offset + 4 <= len(data):
        if data[offset] != 0xFF:
            offset += 1
            continue
        marker = data[offset + 1]
        if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
            offset += 2
            continue
        (length,) = struct.unpack_from(">H", data, offset + 2)
        body = data[offset + 4 : offset + 2 + length]
        if marker == 0xE0 and body.startswith(b"JFIF\x00") and len(body) >= 12:
            units = body[7]
            per_x, per_y = struct.unpack_from(">HH", body, 8)
            # units 0 is "pixel aspect ratio only" -- the JPEG spelling of no size at all.
            if per_x and per_y:
                if units == 1:
                    density_x, density_y = float(per_x), float(per_y)
                elif units == 2:
                    density_x, density_y = per_x * 2.54, per_y * 2.54
        elif marker in _JPEG_FRAME and len(body) >= 5:
            height, width = struct.unpack_from(">HH", body, 1)
            return (width, height, density_x, density_y) if width and height else None
        elif marker == 0xDA:
            break
        offset += 2 + length
    return None


def _gif(data: bytes) -> tuple[int, int, float, float] | None:
    if len(data) < 10:
        return None
    width, height = struct.unpack_from("<HH", data, 6)
    if not width or not height:
        return None
    # GIF states no physical size at all, so the default density is the whole answer.
    return width, height, DEFAULT_DENSITY_DPI, DEFAULT_DENSITY_DPI


def _bmp(data: bytes) -> tuple[int, int, float, float] | None:
    if len(data) < 30:
        return None
    (header,) = struct.unpack_from("<I", data, 14)
    if header < 40 or len(data) < 54:
        # A BITMAPCOREHEADER carries 16-bit extents and no density.
        if header == 12 and len(data) >= 26:
            width, height = struct.unpack_from("<HH", data, 18)
            return (width, height, DEFAULT_DENSITY_DPI, DEFAULT_DENSITY_DPI)
        return None
    width, height = struct.unpack_from("<ii", data, 18)
    per_x, per_y = struct.unpack_from("<ii", data, 38)
    density_x = per_x * _INCHES_PER_METRE if per_x > 0 else DEFAULT_DENSITY_DPI
    density_y = per_y * _INCHES_PER_METRE if per_y > 0 else DEFAULT_DENSITY_DPI
    width, height = abs(width), abs(height)
    return (width, height, density_x, density_y) if width and height else None
