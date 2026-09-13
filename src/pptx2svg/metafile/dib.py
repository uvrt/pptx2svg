"""Device-independent bitmap -> PNG, standard library only.

A DIB is the bitmap format GDI passes around: a ``BITMAPINFOHEADER`` (or one of its two
later supersets), an optional colour table, and pixel rows.  Three things about it are
counter-intuitive and account for most of the code here:

* **Rows run bottom-up.**  A positive ``biHeight`` means the *first* row in the buffer is
  the *bottom* row of the image.  A negative one means top-down.  Getting this wrong
  flips the picture, which is easy to miss on a symmetric drawing.
* **Channel order is BGR, not RGB**, and each row is padded out to a 4-byte boundary
  regardless of the pixel width.
* **The "32-bit" alpha channel is usually not alpha.**  ``BI_RGB`` at 32 bpp declares the
  fourth byte *undefined*; Office fills it with zero.  Trusting it would render every
  such preview fully transparent, so it is only honoured when some pixel actually sets
  it -- see :func:`_alpha_is_meaningful`.

``BI_JPEG`` and ``BI_PNG`` are the easy cases: the "pixel data" is literally a JPEG or
PNG file, which is returned untouched.

Everything here is total: malformed input returns ``None`` rather than raising, because
the caller is handling attacker-supplied bytes and must degrade to a placeholder.
"""

from __future__ import annotations

import struct
import zlib
from dataclasses import dataclass

#: ``biCompression`` values (wingdi.h).
BI_RGB = 0
BI_RLE8 = 1
BI_RLE4 = 2
BI_BITFIELDS = 3
BI_JPEG = 4
BI_PNG = 5
BI_ALPHABITFIELDS = 6

#: Header sizes we understand.  12 is ``BITMAPCOREHEADER``, which has 16-bit dimensions
#: and no compression field; it predates Win32 and does not appear in Office metafiles.
HEADER_SIZE_INFO = 40
HEADER_SIZE_V4 = 108
HEADER_SIZE_V5 = 124

#: Refuse to allocate for an absurd bitmap.  A DIB declares its own dimensions, so a
#: 16-byte header can ask for a 100-gigapixel buffer; this is the guard against that.
#: 64 megapixels is far beyond any plausible slide-sized preview.
MAX_PIXELS = 64_000_000

#: zlib level for the PNG we emit.  Fixed (not ``-1``) so output stays byte-for-byte
#: reproducible across zlib builds, which the SVG golden tests depend on.
PNG_COMPRESS_LEVEL = 9


@dataclass(frozen=True)
class DecodedImage:
    """A ready-to-embed image payload and the MIME type that describes it."""

    data: bytes
    mime_type: str


def dib_to_image(header: bytes, pixels: bytes) -> DecodedImage | None:
    """Convert one DIB (header block + pixel block) to a PNG, JPEG or ``None``.

    ``header`` is the ``BITMAPINFO`` block -- the header struct plus any colour table or
    bitfield masks.  ``pixels`` is the separate pixel buffer, which is how EMF stores it
    (two offsets into the record rather than one contiguous blob).
    """
    info = _parse_header(header)
    if info is None:
        return None

    # BI_JPEG / BI_PNG mean the "pixels" are an entire encoded image file.  Hand it back
    # as-is: re-encoding would only lose quality, and the browser decodes it anyway.
    if info.compression == BI_JPEG:
        return DecodedImage(data=pixels, mime_type="image/jpeg") if pixels else None
    if info.compression == BI_PNG:
        return DecodedImage(data=pixels, mime_type="image/png") if pixels else None

    # Run-length encoded DIBs are legal but do not occur in Office previews, and the two
    # RLE decoders are pure liability without a real sample to test against.
    if info.compression in (BI_RLE8, BI_RLE4):
        return None
    if info.compression not in (BI_RGB, BI_BITFIELDS, BI_ALPHABITFIELDS):
        return None

    rgba = _decode_pixels(info, header, pixels)
    if rgba is None:
        return None

    return DecodedImage(
        data=encode_png(info.width, info.height, rgba, has_alpha=info.honour_alpha),
        mime_type="image/png",
    )


# --------------------------------------------------------------------------------------
# Header
# --------------------------------------------------------------------------------------


@dataclass
class _DibInfo:
    width: int
    height: int
    bit_count: int
    compression: int
    palette_offset: int
    palette_entries: int
    masks: tuple[int, int, int, int] | None
    top_down: bool
    honour_alpha: bool = False


def _parse_header(header: bytes) -> _DibInfo | None:
    if len(header) < HEADER_SIZE_INFO:
        return None
    (
        size,
        width,
        height,
        _planes,
        bit_count,
        compression,
        _size_image,
        _xppm,
        _yppm,
        clr_used,
        _clr_important,
    ) = struct.unpack_from("<IiiHHIIiiII", header, 0)

    if size not in (HEADER_SIZE_INFO, HEADER_SIZE_V4, HEADER_SIZE_V5) or size > len(header):
        return None
    if bit_count not in (1, 4, 8, 16, 24, 32):
        return None
    if width <= 0 or height == 0:
        return None
    # `height` is signed: negative means the rows are stored top-down.
    top_down = height < 0
    height = abs(height)
    if width * height > MAX_PIXELS:
        return None

    masks = None
    palette_offset = size
    if compression in (BI_BITFIELDS, BI_ALPHABITFIELDS):
        count = 4 if compression == BI_ALPHABITFIELDS else 3
        if size == HEADER_SIZE_INFO:
            # With the plain 40-byte header the masks live *after* it, in the space a
            # colour table would otherwise occupy -- and they push the palette along.
            if len(header) < size + count * 4:
                return None
            values = struct.unpack_from(f"<{count}I", header, size)
            palette_offset = size + count * 4
        else:
            # V4/V5 put the masks inside the header proper, always four of them.
            values = struct.unpack_from("<4I", header, 40)[:count]
        masks = (values[0], values[1], values[2], values[3] if count == 4 else 0)

    palette_entries = 0
    if bit_count <= 8:
        palette_entries = clr_used or (1 << bit_count)
        # A lying `biClrUsed` must not make us read past the block.
        available = max(0, (len(header) - palette_offset) // 4)
        palette_entries = min(palette_entries, available)

    return _DibInfo(
        width=width,
        height=height,
        bit_count=bit_count,
        compression=compression,
        palette_offset=palette_offset,
        palette_entries=palette_entries,
        masks=masks,
        top_down=top_down,
    )


# --------------------------------------------------------------------------------------
# Pixels
# --------------------------------------------------------------------------------------


def _decode_pixels(info: _DibInfo, header: bytes, pixels: bytes) -> bytearray | None:
    """Expand the pixel buffer to tightly-packed RGBA, top-down."""
    stride = ((info.width * info.bit_count + 31) // 32) * 4
    needed = stride * info.height
    if len(pixels) < needed:
        # Truncated buffers do occur; pad with zeroes rather than refusing, so a mostly
        # intact preview still renders.  Refusing here would send a whole picture back
        # to the grey placeholder over a handful of missing bytes.
        pixels = pixels + bytes(needed - len(pixels))

    palette = _palette(info, header)
    if info.bit_count <= 8 and not palette:
        return None

    out = bytearray(info.width * info.height * 4)
    for y in range(info.height):
        source_row = y if info.top_down else info.height - 1 - y
        row = pixels[source_row * stride : source_row * stride + stride]
        _decode_row(info, row, palette, out, y * info.width * 4)

    if info.bit_count == 32:
        info.honour_alpha = _alpha_is_meaningful(info, out)
    return out


def _palette(info: _DibInfo, header: bytes) -> list[tuple[int, int, int]]:
    """Colour table entries are stored BGRx, four bytes each."""
    palette: list[tuple[int, int, int]] = []
    for index in range(info.palette_entries):
        base = info.palette_offset + index * 4
        blue, green, red = header[base], header[base + 1], header[base + 2]
        palette.append((red, green, blue))
    return palette


def _decode_row(
    info: _DibInfo,
    row: bytes,
    palette: list[tuple[int, int, int]],
    out: bytearray,
    start: int,
) -> None:
    width, bpp = info.width, info.bit_count
    fallback = (0, 0, 0)

    if bpp in (1, 4, 8):
        per_byte = 8 // bpp
        mask = (1 << bpp) - 1
        for x in range(width):
            byte = row[x // per_byte] if x // per_byte < len(row) else 0
            # Indices are packed most-significant-bits first within each byte.
            shift = 8 - bpp * (x % per_byte + 1)
            index = (byte >> shift) & mask
            red, green, blue = palette[index] if index < len(palette) else fallback
            base = start + x * 4
            out[base : base + 4] = bytes((red, green, blue, 255))
        return

    if bpp == 16:
        # Default 16-bit layout is 5-5-5 with the top bit unused; BI_BITFIELDS overrides
        # it (5-6-5 being the other common one).
        masks = info.masks or (0x7C00, 0x03E0, 0x001F, 0)
        for x in range(width):
            base_in = x * 2
            if base_in + 2 > len(row):
                break
            value = row[base_in] | (row[base_in + 1] << 8)
            base = start + x * 4
            out[base : base + 4] = bytes(
                (
                    _scale(value, masks[0]),
                    _scale(value, masks[1]),
                    _scale(value, masks[2]),
                    255,
                )
            )
        return

    if bpp == 24:
        for x in range(width):
            base_in = x * 3
            if base_in + 3 > len(row):
                break
            base = start + x * 4
            out[base : base + 4] = bytes(
                (row[base_in + 2], row[base_in + 1], row[base_in], 255)
            )
        return

    # 32 bpp.  With bitfields, honour the declared masks; otherwise it is plain BGRx and
    # the fourth byte is carried through for _alpha_is_meaningful to judge.
    masks = info.masks
    for x in range(width):
        base_in = x * 4
        if base_in + 4 > len(row):
            break
        base = start + x * 4
        if masks is None:
            out[base : base + 4] = bytes(
                (row[base_in + 2], row[base_in + 1], row[base_in], row[base_in + 3])
            )
        else:
            value = int.from_bytes(row[base_in : base_in + 4], "little")
            alpha = _scale(value, masks[3]) if masks[3] else 255
            out[base : base + 4] = bytes(
                (_scale(value, masks[0]), _scale(value, masks[1]), _scale(value, masks[2]), alpha)
            )


def _scale(value: int, mask: int) -> int:
    """Extract a masked field and stretch it to the full 0-255 range."""
    if not mask:
        return 0
    shift = (mask & -mask).bit_length() - 1
    width = (mask >> shift).bit_length()
    field = (value & mask) >> shift
    maximum = (1 << width) - 1
    return (field * 255 + maximum // 2) // maximum if maximum else 0


def _alpha_is_meaningful(info: _DibInfo, rgba: bytes) -> bool:
    """Decide whether a 32-bit DIB's fourth byte is really alpha.

    ``BI_RGB`` leaves it undefined and Office writes zeros, so an all-zero channel means
    "opaque", not "invisible".  An explicit ``BI_ALPHABITFIELDS`` mask settles it; short
    of that, any non-zero byte is taken as evidence the channel was actually populated.
    """
    if info.compression == BI_ALPHABITFIELDS and info.masks and info.masks[3]:
        return True
    return any(rgba[index] for index in range(3, len(rgba), 4))


# --------------------------------------------------------------------------------------
# PNG
# --------------------------------------------------------------------------------------


def encode_png(width: int, height: int, rgba: bytes, *, has_alpha: bool) -> bytes:
    """Minimal PNG writer: one IHDR, one IDAT, one IEND, filter type 0 throughout.

    Filtering exists to help compression, not correctness, and a preview bitmap is not
    worth a filter-selection heuristic; "None" for every row keeps this short and the
    output deterministic.
    """
    color_type = 6 if has_alpha else 2  # 6 = truecolour with alpha, 2 = truecolour

    raw = bytearray()
    for y in range(height):
        raw.append(0)  # per-scanline filter type: None
        row_start = y * width * 4
        if has_alpha:
            raw += rgba[row_start : row_start + width * 4]
        else:
            for x in range(width):
                base = row_start + x * 4
                raw += rgba[base : base + 3]

    header = struct.pack(">IIBBBBB", width, height, 8, color_type, 0, 0, 0)
    return b"".join(
        (
            b"\x89PNG\r\n\x1a\n",
            _chunk(b"IHDR", header),
            _chunk(b"IDAT", zlib.compress(bytes(raw), PNG_COMPRESS_LEVEL)),
            _chunk(b"IEND", b""),
        )
    )


def _chunk(tag: bytes, payload: bytes) -> bytes:
    return (
        struct.pack(">I", len(payload))
        + tag
        + payload
        + struct.pack(">I", zlib.crc32(tag + payload) & 0xFFFFFFFF)
    )
