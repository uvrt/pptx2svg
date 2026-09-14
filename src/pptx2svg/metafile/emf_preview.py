"""Pull the embedded preview out of an Enhanced Metafile.

An EMF is a list of self-describing records: ``(type: u32, size: u32, payload)``, where
``size`` counts the eight header bytes and is always a multiple of four.  That is the
whole reason this module can be short -- every record can be *skipped* by its own size
without understanding it, so finding the two interesting ones costs a linear scan.

The two interesting ones:

* **``EMR_COMMENT`` (70)** with the ``GDIC`` comment identifier.  Office uses these to
  smuggle a complete PDF rendition of the artwork through the metafile.  The public
  comment type is ``BEGINGROUP`` (2) or ``MULTIFORMATS`` (0x40000004), and the PDF may be
  split across several consecutive comment records, so the payloads are accumulated and
  then searched as one buffer rather than record by record.  The structured
  ``EmrFormat`` descriptors inside ``MULTIFORMATS`` are deliberately *not* trusted:
  implementations disagree over whether ``offData`` is relative to the record or to the
  comment data, and a ``%PDF`` ... ``%%EOF`` scan is immune to that disagreement.
* **``EMR_STRETCHDIBITS`` (81)**, which carries a device-independent bitmap.  Office
  writes one covering the whole canvas when the artwork came from a raster source.

Neither needs a single drawing record to be interpreted.

**This parses attacker-controlled binary.**  A ``.pptx`` is a zip a stranger sent you and
its media parts are unvalidated bytes.  So: hard caps on input size, per-record size and
record count; every read bounds-checked; and no path out of here raises -- malformed
input yields ``None`` and the caller draws its placeholder.
"""

from __future__ import annotations

import struct
from dataclasses import dataclass

from .dib import dib_to_image

#: ``EMR_HEADER.dSignature``, the four bytes " EMF" at offset 40 of the file.
EMF_SIGNATURE = 0x464D4520

#: Record types we act on.  Everything else is skipped by its declared size.
EMR_HEADER = 1
EMR_EOF = 14
EMR_COMMENT = 70
EMR_STRETCHDIBITS = 81

#: ``EMR_COMMENT`` identifier "GDIC", marking a comment with a public (documented)
#: meaning rather than an application-private blob.
GDIC_IDENTIFIER = 0x43494447

#: Public comment types under which Office has been observed to embed a PDF.
COMMENT_BEGINGROUP = 0x00000002
COMMENT_MULTIFORMATS = 0x40000004
PDF_COMMENT_TYPES = frozenset({COMMENT_BEGINGROUP, COMMENT_MULTIFORMATS})

# --- Hard limits ----------------------------------------------------------------------
# Carried over verbatim from pptx-glimpse, whose numbers were chosen to bound the work a
# hostile metafile can force.  They are generous for real Office output: the largest EMF
# in a normal deck is a few hundred kilobytes with a few thousand records.

#: Largest metafile we will look at at all.
MAX_INPUT_BYTES = 8 * 1024 * 1024
#: Largest single record.  A record longer than this is treated as corruption.
MAX_RECORD_BYTES = 4 * 1024 * 1024
#: Stop after this many records rather than walking a crafted multi-million-record file.
MAX_RECORDS = 50_000
#: pptx-glimpse also caps geometry at 200,000 points.  There is no equivalent here
#: because no geometry is interpreted; the bound that matters in its place is the decoded
#: pixel count, enforced by :data:`pptx2svg.metafile.dib.MAX_PIXELS`.
MAX_GEOMETRY_POINTS = 200_000

#: Smallest slice that could plausibly be a PDF ("%PDF-1.x" plus "%%EOF" plus a body).
MIN_PDF_BYTES = 32


@dataclass(frozen=True)
class MetafilePreview:
    """A preview payload lifted out of a metafile, with the MIME type describing it."""

    data: bytes
    mime_type: str


def extract_metafile_preview(data: bytes) -> MetafilePreview | None:
    """Return the embedded preview of an EMF, or ``None`` if there is not one.

    Also accepts WMF bytes, for which it always returns ``None``: WMF has no preview
    convention, but a ``.wmf`` part occasionally contains EMF bytes and the signature
    check below sorts that out without the caller having to guess.
    """
    if not isinstance(data, (bytes, bytearray)) or len(data) > MAX_INPUT_BYTES:
        return None
    data = bytes(data)
    if not _has_emf_signature(data):
        return None

    try:
        comment_data, bitmaps = _walk_records(data)
    except (struct.error, IndexError, ValueError):
        # Defence in depth: the walk bounds-checks every read, but a metafile is hostile
        # input and a crash here would take down the whole conversion.
        return None

    pdf = _find_pdf(comment_data)
    if pdf is not None:
        return MetafilePreview(data=pdf, mime_type="application/pdf")

    # Several bitmaps mean the metafile draws a composition rather than carrying one
    # preview.  The largest is the best single-image approximation of it -- usually it is
    # the full-canvas backdrop and the rest are small decorations.
    for header, pixels in sorted(bitmaps, key=_bitmap_area, reverse=True):
        image = dib_to_image(header, pixels)
        if image is not None:
            return MetafilePreview(data=image.data, mime_type=image.mime_type)

    return None


def _has_emf_signature(data: bytes) -> bool:
    """``EMR_HEADER`` is record type 1 and carries " EMF" as a magic number at offset 40.

    Checking the type as well as the signature matters: the signature alone is four bytes
    that could appear anywhere, and a file that does not *start* with the header record
    is not a walkable EMF regardless of what it contains.
    """
    if len(data) < 44:
        return False
    record_type, record_size = struct.unpack_from("<II", data, 0)
    signature = struct.unpack_from("<I", data, 40)[0]
    return record_type == EMR_HEADER and record_size >= 44 and signature == EMF_SIGNATURE


def _walk_records(data: bytes) -> tuple[bytes, list[tuple[bytes, bytes]]]:
    """Linear pass over the record list, collecting comment payloads and DIBs."""
    comment_chunks: list[bytes] = []
    bitmaps: list[tuple[bytes, bytes]] = []

    offset = 0
    for _ in range(MAX_RECORDS):
        if offset + 8 > len(data):
            break
        record_type, record_size = struct.unpack_from("<II", data, offset)
        # A record must be at least its own header, four-byte aligned, and fit in the
        # file.  Any of these failing means the list is corrupt from here on; stop with
        # whatever was already collected rather than hunting for a resync point.
        if record_size < 8 or record_size % 4 or record_size > MAX_RECORD_BYTES:
            break
        if offset + record_size > len(data):
            break

        record = data[offset : offset + record_size]
        if record_type == EMR_COMMENT:
            chunk = _comment_payload(record)
            if chunk:
                comment_chunks.append(chunk)
        elif record_type == EMR_STRETCHDIBITS:
            bitmap = _stretch_dib_bits(record)
            if bitmap is not None:
                bitmaps.append(bitmap)
        elif record_type == EMR_EOF:
            break

        offset += record_size

    return b"".join(comment_chunks), bitmaps


def _comment_payload(record: bytes) -> bytes | None:
    """Bytes of an ``EMR_COMMENT`` after its public-type field, or ``None`` to ignore it.

    Layout::

        0   iType              u32   = 70
        4   nSize              u32
        8   DataSize           u32   bytes of comment data that follow
        12  CommentIdentifier  u32   "GDIC" for a public comment
        16  PublicCommentType  u32
        20  ...                      type-specific, then the payload

    Returning everything from offset 20 keeps a PDF split across several ``BEGINGROUP``
    records byte-adjacent in the accumulated buffer.  For ``MULTIFORMATS`` it also drags
    in the ``EmrFormat`` descriptor table, but that sits *before* the ``%PDF`` marker and
    the scan steps over it.
    """
    if len(record) < 20:
        return None
    data_size, identifier, public_type = struct.unpack_from("<III", record, 8)
    if identifier != GDIC_IDENTIFIER or public_type not in PDF_COMMENT_TYPES:
        return None
    # `DataSize` counts from `CommentIdentifier`; clamp to what the record really holds
    # so a lying length cannot read into the next record.
    end = min(len(record), 12 + max(0, data_size))
    return record[20:end] if end > 20 else None


def _stretch_dib_bits(record: bytes) -> tuple[bytes, bytes] | None:
    """Split an ``EMR_STRETCHDIBITS`` into its ``BITMAPINFO`` and pixel blocks.

    Layout (MS-EMF 2.3.1.7)::

        0   iType     u32 = 81      8   Bounds    RECTL (16 bytes)
        24  xDest i32    28  yDest i32    32  xSrc i32    36  ySrc i32
        40  cxSrc i32    44  cySrc i32
        48  offBmiSrc  u32   52  cbBmiSrc   u32
        56  offBitsSrc u32   60  cbBitsSrc  u32
        64  UsageSrc u32     68  BitBltRasterOperation u32
        72  cxDest i32       76  cyDest i32

    Both offsets are from the start of *this record*, not the file.
    """
    if len(record) < 80:
        return None
    off_bmi, cb_bmi, off_bits, cb_bits = struct.unpack_from("<IIII", record, 48)
    if cb_bmi == 0 or cb_bits == 0:
        return None
    if off_bmi + cb_bmi > len(record) or off_bits + cb_bits > len(record):
        return None
    if off_bmi < 8 or off_bits < 8:
        return None
    return record[off_bmi : off_bmi + cb_bmi], record[off_bits : off_bits + cb_bits]


def _bitmap_area(bitmap: tuple[bytes, bytes]) -> int:
    """Pixel area declared by a DIB header, for picking the largest of several."""
    header = bitmap[0]
    if len(header) < 16:
        return 0
    width, height = struct.unpack_from("<ii", header, 4)
    return abs(width) * abs(height)


def _find_pdf(data: bytes) -> bytes | None:
    """Carve ``%PDF`` ... ``%%EOF`` out of the accumulated comment payload.

    The *last* ``%%EOF`` is used, not the first: an incrementally-updated PDF carries one
    per revision and only the final one closes the file.

    When there is no ``%%EOF`` at all the rest of the buffer is taken instead, rather
    than giving up.  A PDF whose trailer was truncated is usually still renderable --
    pdfium rebuilds the cross-reference table from the object offsets -- and any trailing
    bytes from later records are junk a PDF reader skips.  The alternative is a grey
    placeholder in place of artwork that would have rendered, which is worse.
    """
    if not data:
        return None
    start = data.find(b"%PDF")
    if start < 0:
        return None
    end = data.rfind(b"%%EOF")
    pdf = data[start : end + len(b"%%EOF")] if end >= start else data[start:]
    return pdf if len(pdf) >= MIN_PDF_BYTES else None
