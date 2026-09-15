"""Embedded OpenType (EOT) containers -- the shape of a ``ppt/fonts/*.fntdata`` part.

PowerPoint does not put a font file in a deck.  It puts an EOT: Microsoft's web-font
wrapper, a little-endian header carrying the family name, the weight, the embedding
permissions and a set of flags, followed by the font data, which may be XOR-obfuscated,
MicroType Express compressed, or both.

**What the local payloads actually are, measured rather than assumed.**  Every one of the
21 ``.fntdata`` parts across the two third-party template decks on hand reads::

    Version       0x00020002   (EOT 2.2)
    Flags         0x00000004   (TTEMBED_TTCOMPRESSED; no SUBSET, no XORENCRYPTDATA)
    fsType        0x0000       (installable embedding)
    MagicNumber   0x504C
    EOTSize == the part length, to the byte, on all 21

The compression flag was previously only an inference -- "nothing has actually
decompressed one of these files".  It has now been settled by decoding: all 21 payloads
decompress through :mod:`pptx2svg.fonts.mtx` to valid TrueType files, byte-identical to
what libEOT's ``eot2ttf`` produces from the same input.  So MTX is not a hypothesis about
these decks; it is what they contain.

One surprise worth recording, since it shapes the design: the payloads are **not**
subsetted (``TTEMBED_SUBSET`` is clear, and Arimo decodes with all 3237 glyphs and a
3010-entry cmap).  PowerPoint embedded the whole face.  That is why measuring from the
extracted file is worth doing at all -- a subset would only cover the characters already
on the slides.

The uncompressed and XOR-only variants are handled here too, in three lines, even though
nothing local exercises them: Apache POI's developer list records that both MTX and
non-MTX occur in the wild, and the synthetic fixture in ``tests/fixtures`` covers the
uncompressed path precisely because no real deck on hand does.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass

from .mtx import MtxDecodeError, decode_mtx

__all__ = [
    "EMBEDDING_BITMAP_ONLY",
    "EMBEDDING_EDITABLE",
    "EMBEDDING_NO_SUBSETTING",
    "EMBEDDING_PREVIEW_PRINT",
    "EMBEDDING_RESTRICTED",
    "EotError",
    "EotHeader",
    "decode_eot",
    "embedding_refusal",
    "read_eot_header",
]

#: ``TTEMBED_*`` flags, from the EOT specification.  Only the two we act on are named;
#: the rest (SUBSET 0x1, FAILIFVARIATIONSIMULATED 0x10, EMBEDEUDC 0x20,
#: VALIDATIONTESTS 0x40, WEBOBJECT 0x80) describe how the file was produced and do not
#: change how it is read.
TTEMBED_TTCOMPRESSED = 0x00000004
TTEMBED_XORENCRYPTDATA = 0x10000000

#: The byte the XOR variant obfuscates with.  A constant in the format, not a key.
_XOR_KEY = 0x50

#: ``OS/2.fsType`` bits (OpenType spec, "OS/2 -- fsType").  Bit 0 is reserved and must be
#: zero; bits 1-3 are the mutually exclusive permission level, although real files do set
#: combinations, which is why the check below masks rather than compares.
EMBEDDING_RESTRICTED = 0x0002
EMBEDDING_PREVIEW_PRINT = 0x0004
EMBEDDING_EDITABLE = 0x0008
EMBEDDING_NO_SUBSETTING = 0x0100
EMBEDDING_BITMAP_ONLY = 0x0200

#: The permission field, bits 0-3.
_PERMISSION_MASK = 0x000F


class EotError(Exception):
    """An EOT payload could not be read."""


@dataclass(frozen=True)
class EotHeader:
    """The fixed part of an EOT header.

    Everything here lives at a constant offset, so it can be read without walking the
    four variable-length name strings that follow.
    """

    eot_size: int
    font_data_size: int
    version: int
    flags: int
    fs_type: int
    weight: int
    italic: bool
    family_name: str
    style_name: str

    @property
    def font_data_offset(self) -> int:
        """Where the font data starts.

        Derived as ``EOTSize - FontDataSize`` rather than by summing the header fields,
        which is what libEOT does (``EOTgetMetadataLength``) and is robust against the
        version-dependent tail (root strings, signature, EUDC block).  On all 21 local
        payloads it agrees with a full header walk exactly.
        """
        return self.eot_size - self.font_data_size

    @property
    def compressed(self) -> bool:
        return bool(self.flags & TTEMBED_TTCOMPRESSED)

    @property
    def encrypted(self) -> bool:
        return bool(self.flags & TTEMBED_XORENCRYPTDATA)


def read_eot_header(data: bytes) -> EotHeader:
    """Parse the fixed header of a ``.fntdata`` payload."""
    if len(data) < 84:
        raise EotError("payload is too short to be an EOT font")
    eot_size, font_data_size, version, flags = struct.unpack_from("<IIII", data, 0)
    weight = struct.unpack_from("<I", data, 28)[0]
    fs_type, magic = struct.unpack_from("<HH", data, 32)
    italic = bool(data[27])
    if magic != 0x504C:
        raise EotError(f"not an EOT font: magic number is 0x{magic:04X}, expected 0x504C")
    if font_data_size > eot_size or eot_size - font_data_size < 84:
        raise EotError("EOT header sizes are inconsistent")
    if eot_size > len(data):
        raise EotError(
            f"EOT claims {eot_size} bytes but the part carries {len(data)}"
        )
    # Padding1 at 80, FamilyNameSize at 82, FamilyName at 84; the rest of the names
    # follow in the same Padding/Size/UTF-16LE shape.
    family, offset = _read_utf16(data, 80)
    style, _ = _read_utf16(data, offset)
    return EotHeader(
        eot_size=eot_size,
        font_data_size=font_data_size,
        version=version,
        flags=flags,
        fs_type=fs_type,
        weight=weight,
        italic=italic,
        family_name=family,
        style_name=style,
    )


def _read_utf16(data: bytes, offset: int) -> tuple[str, int]:
    """A ``Padding`` + ``Size`` + UTF-16LE string, returning the value and the next offset.

    Best-effort: the names are used for reporting only, so a truncated header yields an
    empty string rather than an error, and the caller still gets the fixed fields.
    """
    if offset + 4 > len(data):
        return "", offset
    size = struct.unpack_from("<H", data, offset + 2)[0]
    start = offset + 4
    if size % 2 or start + size > len(data):
        return "", start
    return data[start : start + size].decode("utf-16-le", "replace"), start + size


def embedding_refusal(fs_type: int) -> str | None:
    """Why this font may not be used, or ``None`` when it may.

    The foundry's restrictions travel with the file, and a renderer that ignores them is
    redistributing someone's font under the cover of "it was already in the deck".  This
    project publishes advance widths of fonts it will not ship precisely because it takes
    that line seriously, so the gate is not optional and there is no bypass parameter.

    The rule is LibreOffice's, followed deliberately rather than invented::

        // EmbeddedFontsHelper::sufficientTTFRights, vcl/source/gdi/embeddedfontshelper.cxx
        case FontRights::ViewingAllowed:
            return copyright == 0 || ( copyright & 0x0e ) != 0x02;

    That is: **refuse only a font whose permission field says "restricted licence" and
    says nothing else.**  Rendering a deck to a fixed image is a viewing use -- the output
    is not editable and carries no font file -- so "Preview & Print" (0x0004) and
    "Editable" (0x0008) both permit it, and 0x0000 ("installable") is unrestricted.

    Two further bits are read:

    * ``0x0200`` *bitmap embedding only* -- refused.  We draw outlines; that is exactly
      what the bit forbids.  LibreOffice does not check this one, and refusing where it
      permits is the direction to err in.
    * ``0x0100`` *no subsetting* -- permitted, and noted here so nobody adds a check for
      it later.  Nothing in this library subsets a font: the extracted face is written
      out whole.

    Returning a sentence rather than a bool because the caller puts it in a warning, and
    "this font was refused" without the reason is the kind of message that gets ignored.
    """
    if fs_type & EMBEDDING_BITMAP_ONLY:
        return (
            "its OS/2 fsType permits bitmap embedding only (0x0200), and this renderer "
            "draws outlines"
        )
    permissions = fs_type & _PERMISSION_MASK
    if permissions and (permissions & 0x0E) == EMBEDDING_RESTRICTED:
        return (
            "its OS/2 fsType is restricted-licence embedding (0x0002), which permits no "
            "use without the foundry's permission"
        )
    return None


def decode_eot(data: bytes) -> bytes:
    """Turn a ``.fntdata`` payload into a TrueType file.

    Raises :class:`EotError` for anything this cannot read, including a payload whose
    embedding permissions refuse it -- the caller degrades to substitution and says why.
    """
    header = read_eot_header(data)

    # Gate before decoding, not after: refusing a restricted font should not cost a
    # second of decompression first.  The authoritative check is on the decoded OS/2
    # table (see fonts.sfnt.read_sfnt), because this header field is written by whatever
    # produced the EOT and the font's own table is the foundry's statement.  Both must
    # permit.
    refusal = embedding_refusal(header.fs_type)
    if refusal is not None:
        raise EotError(f"embedding is not permitted: {refusal}")

    payload = data[header.font_data_offset : header.font_data_offset + header.font_data_size]
    if len(payload) != header.font_data_size:
        raise EotError("EOT font data is truncated")
    if header.encrypted:
        payload = bytes(byte ^ _XOR_KEY for byte in payload)
    if not header.compressed:
        return payload
    try:
        return decode_mtx(payload)
    except MtxDecodeError as error:
        raise EotError(f"MicroType Express payload could not be decoded: {error}") from error
