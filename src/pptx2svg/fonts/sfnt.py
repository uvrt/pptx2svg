"""Just enough TrueType to measure a font and to give it the name a deck calls it by.

This exists because of the invariant the rest of the font subsystem is built on: **we
measure with what we draw with.**  An extracted embedded face has to reach text
measurement, and measurement runs inside ``convert_pptx_to_svg``, which is
standard-library-only.  ``fontTools`` would do all of this and more, but it lives behind
``pptx2svg[measure]``, and making embedded fonts depend on an extra would mean a bare
install draws the deck's own face at guessed widths -- precisely the right-glyphs,
wrong-line-breaks failure recorded against ``--font-dir`` in the roadmap.

So: a table directory reader, ``head``/``hhea``/``maxp``/``OS/2``/``hmtx``/``cmap``/``name``,
and a writer.  Around 300 lines, against the 4700 lines of C it took to *decompress* the
same file.  Nothing here interprets an outline.

The other half is :func:`relabel`, which is not a convenience.  See its docstring.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass

__all__ = ["SfntError", "SfntFace", "read_sfnt", "relabel", "write_sfnt"]


class SfntError(Exception):
    """A font file could not be read."""


@dataclass(frozen=True)
class SfntFace:
    """The parts of a font this library needs, and the raw tables to write back out."""

    tables: dict[bytes, bytes]
    units_per_em: int
    #: ``hhea``, not ``OS/2.sTypo*``: the first-baseline rule in
    #: :mod:`pptx2svg.text.measure` was calibrated against PowerPoint using hhea's
    #: descent, and the two disagree (Arial -434 vs -431, Carlito -550 vs -512).
    ascender: int
    descender: int
    fs_type: int
    weight_class: int
    family_name: str
    subfamily_name: str
    #: Unicode code point -> advance width in font units.
    advances: dict[int, int]


# --------------------------------------------------------------------------------------
# Reading
# --------------------------------------------------------------------------------------


def _directory(data: bytes) -> dict[bytes, bytes]:
    if len(data) < 12:
        raise SfntError("font file is too short to hold an offset table")
    tag = data[:4]
    if tag == b"ttcf":
        raise SfntError("TrueType collections are not supported here")
    if tag not in (b"\x00\x01\x00\x00", b"OTTO", b"true", b"typ1"):
        raise SfntError(f"not a font file: leading bytes are {tag!r}")
    count = struct.unpack_from(">H", data, 4)[0]
    if 12 + 16 * count > len(data):
        raise SfntError("font table directory runs past the end of the file")
    tables: dict[bytes, bytes] = {}
    for i in range(count):
        entry = 12 + 16 * i
        name = data[entry : entry + 4]
        offset, length = struct.unpack_from(">II", data, entry + 8)
        if offset + length > len(data):
            raise SfntError(f"table {name!r} runs past the end of the file")
        tables[name] = data[offset : offset + length]
    return tables


def read_sfnt(data: bytes) -> SfntFace:
    """Parse a TrueType/OpenType file into the pieces measurement and relabelling need."""
    tables = _directory(data)
    for required in (b"head", b"hhea", b"maxp", b"hmtx"):
        if required not in tables:
            raise SfntError(f"font has no {required.decode()} table")

    head = tables[b"head"]
    if len(head) < 54:
        raise SfntError("head table is too short")
    units_per_em = struct.unpack_from(">H", head, 18)[0]
    if not units_per_em:
        raise SfntError("head table declares unitsPerEm of 0")

    hhea = tables[b"hhea"]
    if len(hhea) < 36:
        raise SfntError("hhea table is too short")
    ascender, descender = struct.unpack_from(">hh", hhea, 4)
    num_h_metrics = struct.unpack_from(">H", hhea, 34)[0]

    maxp = tables[b"maxp"]
    if len(maxp) < 6:
        raise SfntError("maxp table is too short")
    num_glyphs = struct.unpack_from(">H", maxp, 4)[0]

    os2 = tables.get(b"OS/2", b"")
    # A font with no OS/2 table states no restriction; the spec's default is installable
    # embedding, which is also how LibreOffice treats one it cannot read.
    fs_type = struct.unpack_from(">H", os2, 8)[0] if len(os2) >= 10 else 0
    weight_class = struct.unpack_from(">H", os2, 4)[0] if len(os2) >= 6 else 400

    advances = _advances(tables[b"hmtx"], num_h_metrics, num_glyphs, tables.get(b"cmap"))
    family, subfamily = _names(tables.get(b"name"))

    return SfntFace(
        tables=tables,
        units_per_em=units_per_em,
        ascender=ascender,
        descender=descender,
        fs_type=fs_type,
        weight_class=weight_class,
        family_name=family,
        subfamily_name=subfamily,
        advances=advances,
    )


def _advances(
    hmtx: bytes, num_h_metrics: int, num_glyphs: int, cmap: bytes | None
) -> dict[int, int]:
    if not num_h_metrics or cmap is None:
        return {}
    per_glyph = [
        struct.unpack_from(">H", hmtx, 4 * i)[0]
        for i in range(min(num_h_metrics, len(hmtx) // 4))
    ]
    if not per_glyph:
        return {}
    # Glyphs past numberOfHMetrics all share the last advance: that is the point of the
    # field, and it is how a monospaced face's hmtx stays four bytes long.
    last = per_glyph[-1]
    result: dict[int, int] = {}
    for code_point, glyph_id in _cmap(cmap).items():
        if glyph_id >= num_glyphs:
            continue
        result[code_point] = per_glyph[glyph_id] if glyph_id < len(per_glyph) else last
    return result


def _cmap(data: bytes) -> dict[int, int]:
    """Code point -> glyph id, from the best subtable present."""
    if len(data) < 4:
        return {}
    count = struct.unpack_from(">H", data, 2)[0]
    best: tuple[int, int] | None = None
    for i in range(count):
        entry = 4 + 8 * i
        if entry + 8 > len(data):
            break
        platform, encoding, offset = struct.unpack_from(">HHI", data, entry)
        # Preference order: full Unicode, then BMP Unicode, then Windows symbol, then
        # anything else.  A symbol font -- Wingdings and its relatives -- has only (3, 0).
        if (platform, encoding) in ((3, 10), (0, 4), (0, 6)):
            rank = 3
        elif (platform, encoding) in ((3, 1), (0, 3), (0, 2), (0, 1), (0, 0)):
            rank = 2
        elif (platform, encoding) == (3, 0):
            rank = 1
        else:
            rank = 0
        if best is None or rank > best[0]:
            best = (rank, offset)
    if best is None:
        return {}
    rank, offset = best
    mapping = _cmap_subtable(data, offset)
    if rank == 1:
        # Windows symbol subtables live in the private-use block at 0xF000, but decks
        # address those characters as plain 0x20-0xFF, so publish both spellings.
        for code_point in list(mapping):
            if 0xF000 <= code_point <= 0xF0FF:
                mapping.setdefault(code_point & 0xFF, mapping[code_point])
    return mapping


#: Upper bound on how many code points one subtable may map.
#:
#: Unicode has 0x110000 of them, so a well-formed font cannot exceed this.  The bound is
#: here because a deck is untrusted input and both variable-length subtable formats invite
#: a cheap denial of service: a format 12 header can declare four billion groups, and a
#: format 4 segment can span 65,536 code points, so a kilobyte of malicious cmap can ask
#: for hours of work.  Truncating is the right failure -- a font this malformed has no
#: meaningful advance widths to lose.
_MAX_MAPPED = 0x110000


def _cmap_subtable(data: bytes, offset: int) -> dict[int, int]:
    if offset + 4 > len(data):
        return {}
    fmt = struct.unpack_from(">H", data, offset)[0]
    if fmt == 0:
        if offset + 262 > len(data):
            return {}
        return {i: data[offset + 6 + i] for i in range(256) if data[offset + 6 + i]}
    if fmt == 4:
        return _cmap_format4(data, offset)
    if fmt == 6:
        if offset + 10 > len(data):
            return {}
        first, count = struct.unpack_from(">HH", data, offset + 6)
        out = {}
        for i in range(count):
            at = offset + 10 + 2 * i
            if at + 2 > len(data):
                break
            glyph = struct.unpack_from(">H", data, at)[0]
            if glyph:
                out[first + i] = glyph
        return out
    if fmt == 12:
        return _cmap_format12(data, offset)
    return {}


def _cmap_format4(data: bytes, offset: int) -> dict[int, int]:
    if offset + 14 > len(data):
        return {}
    seg_x2 = struct.unpack_from(">H", data, offset + 6)[0]
    segments = seg_x2 // 2
    ends = offset + 14
    starts = ends + seg_x2 + 2
    deltas = starts + seg_x2
    ranges = deltas + seg_x2
    if ranges + seg_x2 > len(data):
        return {}
    out: dict[int, int] = {}
    for i in range(segments):
        end = struct.unpack_from(">H", data, ends + 2 * i)[0]
        start = struct.unpack_from(">H", data, starts + 2 * i)[0]
        delta = struct.unpack_from(">h", data, deltas + 2 * i)[0]
        range_offset = struct.unpack_from(">H", data, ranges + 2 * i)[0]
        if start > end or (start == 0xFFFF and end == 0xFFFF):
            continue
        for code_point in range(start, end + 1):
            if range_offset == 0:
                glyph = (code_point + delta) & 0xFFFF
            else:
                at = ranges + 2 * i + range_offset + 2 * (code_point - start)
                if at + 2 > len(data):
                    continue
                glyph = struct.unpack_from(">H", data, at)[0]
                if glyph:
                    glyph = (glyph + delta) & 0xFFFF
            if glyph:
                out[code_point] = glyph
            if len(out) >= _MAX_MAPPED:
                return out
    return out


def _cmap_format12(data: bytes, offset: int) -> dict[int, int]:
    if offset + 16 > len(data):
        return {}
    groups = struct.unpack_from(">I", data, offset + 12)[0]
    out: dict[int, int] = {}
    for i in range(groups):
        at = offset + 16 + 12 * i
        if at + 12 > len(data):
            break
        start, end, glyph = struct.unpack_from(">III", data, at)
        if end < start or end > 0x10FFFF:
            continue
        end = min(end, start + _MAX_MAPPED - len(out) - 1)
        for code_point in range(start, end + 1):
            out[code_point] = glyph + (code_point - start)
        if len(out) >= _MAX_MAPPED:
            break
    return out


#: ``name`` records this library reads or replaces.  1/2 are the RIBBI family and style,
#: 4 the full name, 6 the PostScript name, 16/17 the typographic family and style, which
#: override 1/2 for a family with more than four cuts.
_NAME_FAMILY = 1
_NAME_SUBFAMILY = 2
_NAME_UNIQUE_ID = 3
_NAME_FULL = 4
_NAME_POSTSCRIPT = 6
_NAME_TYPOGRAPHIC_FAMILY = 16
_NAME_TYPOGRAPHIC_SUBFAMILY = 17


def _name_records(data: bytes | None):
    if not data or len(data) < 6:
        return
    count, string_offset = struct.unpack_from(">HH", data, 2)
    for i in range(count):
        at = 6 + 12 * i
        if at + 12 > len(data):
            return
        platform, encoding, language, name_id, length, offset = struct.unpack_from(
            ">HHHHHH", data, at
        )
        start = string_offset + offset
        yield platform, encoding, language, name_id, data[start : start + length]


def _decode_name(platform: int, encoding: int, raw: bytes) -> str:
    if platform in (0, 3) or (platform == 2 and encoding == 1):
        return raw.decode("utf-16-be", "replace")
    return raw.decode("mac-roman" if platform == 1 else "latin-1", "replace")


def _names(data: bytes | None) -> tuple[str, str]:
    family = subfamily = ""
    typographic_family = typographic_subfamily = ""
    for platform, encoding, _language, name_id, raw in _name_records(data):
        value = _decode_name(platform, encoding, raw)
        if not value:
            continue
        if name_id == _NAME_FAMILY and not family:
            family = value
        elif name_id == _NAME_SUBFAMILY and not subfamily:
            subfamily = value
        elif name_id == _NAME_TYPOGRAPHIC_FAMILY and not typographic_family:
            typographic_family = value
        elif name_id == _NAME_TYPOGRAPHIC_SUBFAMILY and not typographic_subfamily:
            typographic_subfamily = value
    return (typographic_family or family), (typographic_subfamily or subfamily)


# --------------------------------------------------------------------------------------
# Writing
# --------------------------------------------------------------------------------------


def write_sfnt(tables: list[tuple[bytes, bytes]]) -> bytes:
    """Assemble tables into a TrueType file with correct checksums.

    The directory is sorted by tag, as the specification requires, and ``head``'s
    ``checkSumAdjustment`` is recomputed last, which is the only order in which it can
    come out right.
    """
    ordered = sorted(tables, key=lambda item: item[0])
    count = len(ordered)
    largest_power, entry_selector, remaining = 1, 0, count
    while remaining > 1:
        largest_power *= 2
        entry_selector += 1
        remaining //= 2
    search_range = largest_power * 16

    out = bytearray()
    out += struct.pack(
        ">IHHHH", 0x00010000, count, search_range, entry_selector,
        count * 16 - search_range,
    )
    directory_at = len(out)
    out += bytes(16 * count)

    total = 0
    head_at = None
    entries = []
    for tag, buf in ordered:
        offset = len(out)
        if tag == b"head":
            head_at = offset
        padded = buf + bytes(-len(buf) % 4)
        checksum = 0
        for i in range(0, len(padded), 4):
            checksum = (checksum + struct.unpack_from(">I", padded, i)[0]) & 0xFFFFFFFF
        entries.append((tag, checksum, offset, len(buf)))
        out += padded
        total = (total + checksum) & 0xFFFFFFFF

    at = directory_at
    for tag, checksum, offset, length in entries:
        struct.pack_into(">4sIII", out, at, tag, checksum, offset, length)
        at += 16

    prefix = bytes(out[: directory_at + 16 * count])
    for i in range(0, len(prefix), 4):
        total = (total + struct.unpack_from(">I", prefix, i)[0]) & 0xFFFFFFFF

    if head_at is not None:
        # 0xB1B0AFBA is the constant the TrueType specification defines for this field.
        struct.pack_into(">I", out, head_at + 8, (0xB1B0AFBA - total) & 0xFFFFFFFF)
    return bytes(out)


_STYLE_NAMES = {
    (False, False): "Regular",
    (True, False): "Bold",
    (False, True): "Italic",
    (True, True): "Bold Italic",
}


def relabel(data: bytes, family: str, *, bold: bool, italic: bool) -> bytes:
    """Rewrite a face's identity to the family and slot the deck assigned it.

    This is not cosmetic.  ``<p:embeddedFont><p:font typeface="X"/><p:bold r:id="Y"/>``
    is an assertion: *Y is the bold cut of X, for this deck*.  The file behind Y need not
    agree, and in the local corpus it frequently does not --
    ``ppt/fonts/MerriweatherSansLight-bold.fntdata`` decodes to a font whose ``name``
    table reads "Merriweather Sans / Regular" at ``usWeightClass`` 400, because PowerPoint
    had no Light Bold cut to embed and substituted one.  PowerPoint still draws bold
    "Merriweather Sans Light" runs with that file.

    Left alone, such a file is unaddressable: the SVG asks for
    ``font-family: 'Merriweather Sans Light'; font-weight: bold`` and the rasteriser's
    font database has it filed under "Merriweather Sans" at weight 400.  The face would be
    handed over and then silently passed over -- the same silent-substitution failure this
    subsystem exists to prevent, and one that would have looked fine in a screenshot
    because a *similar* face would have drawn in its place.

    So the identity fields are rewritten to what the deck says:

    * ``name`` 1/2/4/6 become the deck's typeface and the slot's style name, and 16/17
      (typographic family/subfamily) are dropped, because their whole purpose is to
      override 1/2 and here 1/2 are now the correct answer.  Every other record --
      copyright (0), trademark (7), licence (13) and licence URL (14) -- is preserved
      byte for byte, which matters for a file this library writes to disk.
    * ``OS/2.usWeightClass`` becomes 700 or 400.  Discarding the original is deliberate:
      a Light face filling a family's regular slot has to match ``font-weight: normal``,
      or the same non-match recurs one level down.
    * ``OS/2.fsSelection`` and ``head.macStyle`` get the slot's bold and italic bits.

    The measured widths come from this same file, so measure-equals-draw holds whatever
    the label says.
    """
    tables = dict(_directory(data))

    tables[b"name"] = _rebuild_name(
        tables.get(b"name"), family, _STYLE_NAMES[(bold, italic)]
    )

    os2 = tables.get(b"OS/2")
    if os2 is not None and len(os2) >= 64:
        os2 = bytearray(os2)
        struct.pack_into(">H", os2, 4, 700 if bold else 400)
        selection = struct.unpack_from(">H", os2, 62)[0]
        selection &= ~0x0061  # clear ITALIC, BOLD, REGULAR and OBLIQUE
        selection |= (0x20 if bold else 0) | (0x01 if italic else 0)
        if not bold and not italic:
            selection |= 0x40
        struct.pack_into(">H", os2, 62, selection)
        tables[b"OS/2"] = bytes(os2)

    head = tables.get(b"head")
    if head is not None and len(head) >= 46:
        head = bytearray(head)
        mac_style = struct.unpack_from(">H", head, 44)[0] & ~0x0003
        struct.pack_into(
            ">H", head, 44, mac_style | (0x01 if bold else 0) | (0x02 if italic else 0)
        )
        tables[b"head"] = bytes(head)

    return write_sfnt(list(tables.items()))


def _postscript_name(family: str, style: str) -> str:
    """``Family-Style`` with the characters the PostScript name grammar forbids removed."""
    forbidden = set(" [](){}<>/%")
    clean = "".join(
        c for c in f"{family}-{style}" if 33 <= ord(c) <= 126 and c not in forbidden
    )
    return clean[:63] or "EmbeddedFont-Regular"


def _rebuild_name(data: bytes | None, family: str, style: str) -> bytes:
    full = family if style == "Regular" else f"{family} {style}"
    replacements = {
        _NAME_FAMILY: family,
        _NAME_SUBFAMILY: style,
        _NAME_UNIQUE_ID: f"pptx2svg embedded: {full}",
        _NAME_FULL: full,
        _NAME_POSTSCRIPT: _postscript_name(family, style),
    }
    dropped = set(replacements) | {_NAME_TYPOGRAPHIC_FAMILY, _NAME_TYPOGRAPHIC_SUBFAMILY}

    records: list[tuple[int, int, int, int, bytes]] = [
        record for record in _name_records(data) if record[3] not in dropped
    ]
    # Windows Unicode (3, 1, 0x409) and Macintosh Roman (1, 0, 0): fontconfig, DirectWrite
    # and fontdb between them read one or the other, and writing both costs ten records.
    for name_id, value in replacements.items():
        records.append((3, 1, 0x409, name_id, value.encode("utf-16-be")))
        records.append((1, 0, 0, name_id, value.encode("mac-roman", "replace")))
    records.sort(key=lambda record: record[:4])

    strings = bytearray()
    header = bytearray(struct.pack(">HHH", 0, len(records), 6 + 12 * len(records)))
    for platform, encoding, language, name_id, raw in records:
        header += struct.pack(
            ">HHHHHH", platform, encoding, language, name_id, len(raw), len(strings)
        )
        strings += raw
    return bytes(header + strings)
