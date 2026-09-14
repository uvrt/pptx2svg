"""EMF preview extraction, DIB decoding, and the resolver path that uses them."""

from __future__ import annotations

import io
import struct
import zlib

import pytest

from deckbuilder import derive_deck
from pptx2svg import ConvertOptions, convert_pptx_to_model, convert_pptx_to_svg
from pptx2svg.metafile import extract_metafile_preview
from pptx2svg.metafile.dib import encode_png
from pptx2svg.metafile.emf_preview import MAX_INPUT_BYTES
from pptx2svg.metafile.pdf import pdf_rasterizer_available, rasterise_pdf

# --------------------------------------------------------------------------------------
# Synthetic EMF construction
#
# There is no EMF in the fixture corpus and no pure-Python writer to borrow, so the test
# inputs are built here from the record layouts in MS-EMF.  Building them by hand is also
# what makes the adversarial cases possible: a truncated or lying record is exactly what
# a real corrupt file looks like, and nothing else can produce one on demand.
# --------------------------------------------------------------------------------------

EMF_HEADER_SIZE = 88


def emf_header(record_count: int, total_bytes: int) -> bytes:
    """``EMR_HEADER``: the 88-byte minimum form, signature " EMF" at offset 40."""
    return struct.pack(
        "<II4i4iIIIIHHIIIiiii",
        1,
        EMF_HEADER_SIZE,
        0, 0, 100, 100,           # rclBounds
        0, 0, 100, 100,           # rclFrame
        0x464D4520,               # dSignature
        0x10000,                  # nVersion
        total_bytes,
        record_count,
        0, 0,                     # nHandles, sReserved
        0, 0, 0,                  # nDescription, offDescription, nPalEntries
        1000, 1000, 300, 300,     # szlDevice, szlMillimeters
    )


def emf_eof() -> bytes:
    return struct.pack("<IIIII", 14, 20, 0, 0, 20)


def comment_record(public_type: int, payload: bytes, *, identifier: int = 0x43494447) -> bytes:
    """``EMR_COMMENT``.  ``DataSize`` is the *unpadded* length, as real writers emit it;
    ``nSize`` is then rounded up to the four-byte record alignment."""
    body = struct.pack("<II", identifier, public_type) + payload
    data_size = len(body)
    body += b"\0" * ((-len(body)) % 4)
    return struct.pack("<III", 70, 12 + len(body), data_size) + body


def stretch_dib_record(width: int, height: int, bpp: int, pixels: bytes, palette: bytes = b"") -> bytes:
    """``EMR_STRETCHDIBITS`` carrying a BITMAPINFOHEADER DIB."""
    bmi = struct.pack("<IiiHHIIiiII", 40, width, height, 1, bpp, 0, len(pixels), 0, 0, 0, 0)
    bmi += palette
    off_bmi = 80
    off_bits = off_bmi + len(bmi)
    size = off_bits + len(pixels)
    size += (-size) % 4

    record = bytearray(size)
    struct.pack_into("<II", record, 0, 81, size)
    struct.pack_into("<4i", record, 8, 0, 0, width, height)          # Bounds
    struct.pack_into("<6i", record, 24, 0, 0, 0, 0, width, height)   # xDest..cySrc
    struct.pack_into("<IIII", record, 48, off_bmi, len(bmi), off_bits, len(pixels))
    struct.pack_into("<II", record, 64, 0, 0x00CC0020)               # UsageSrc, SRCCOPY
    struct.pack_into("<ii", record, 72, width, height)               # cxDest, cyDest
    record[off_bmi : off_bmi + len(bmi)] = bmi
    record[off_bits : off_bits + len(pixels)] = pixels
    return bytes(record)


def build_emf(records: list[bytes]) -> bytes:
    body = b"".join(records) + emf_eof()
    return emf_header(len(records) + 2, EMF_HEADER_SIZE + len(body)) + body


#: A minimal but genuinely loadable PDF: catalog, one page tree, one page, one filled
#: rectangle.  72x36 pt so the aspect ratio of a render is unambiguous (exactly 2:1).
MINI_PDF = b"""%PDF-1.4
1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj
2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj
3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 72 36]/Contents 4 0 R/Resources<<>>>>endobj
4 0 obj<</Length 24>>stream
1 0 0 rg 0 0 72 36 re f
endstream
endobj
trailer<</Size 5/Root 1 0 R>>
%%EOF"""

COMMENT_BEGINGROUP = 0x00000002
COMMENT_MULTIFORMATS = 0x40000004


def png_size(data: bytes) -> tuple[int, int]:
    assert data[:8] == b"\x89PNG\r\n\x1a\n"
    return struct.unpack_from(">II", data, 16)


def png_pixels(data: bytes) -> tuple[int, int, int, list[int]]:
    """(width, height, colour type, unfiltered sample bytes) of a PNG this code wrote."""
    width, height = png_size(data)
    color_type = data[25]
    offset, idat = 8, b""
    while offset < len(data):
        length, tag = struct.unpack_from(">I4s", data, offset)
        if tag == b"IDAT":
            idat += data[offset + 8 : offset + 8 + length]
        offset += 12 + length
    raw = zlib.decompress(idat)
    channels = 4 if color_type == 6 else 3
    samples: list[int] = []
    stride = width * channels + 1
    for y in range(height):
        assert raw[y * stride] == 0, "only filter type 0 is emitted"
        samples.extend(raw[y * stride + 1 : (y + 1) * stride])
    return width, height, color_type, samples


# --------------------------------------------------------------------------------------
# Embedded PDF
# --------------------------------------------------------------------------------------


def test_multiformats_comment_yields_the_embedded_pdf():
    # The 20 leading bytes stand in for the OutputRect + CountFormats + EmrFormat table
    # that MULTIFORMATS puts before the payload; the scan must step over them.
    emf = build_emf([comment_record(COMMENT_MULTIFORMATS, b"\0" * 20 + MINI_PDF)])
    preview = extract_metafile_preview(emf)
    assert preview is not None
    assert preview.mime_type == "application/pdf"
    assert preview.data == MINI_PDF


def test_pdf_split_across_begingroup_comments_is_rejoined():
    """Office splits a large PDF over consecutive comment records; they must be glued."""
    half = len(MINI_PDF) // 2
    emf = build_emf(
        [
            comment_record(COMMENT_BEGINGROUP, MINI_PDF[:half]),
            comment_record(COMMENT_BEGINGROUP, MINI_PDF[half:]),
        ]
    )
    preview = extract_metafile_preview(emf)
    assert preview is not None and preview.data == MINI_PDF


def test_last_eof_wins_for_an_incrementally_updated_pdf():
    updated = MINI_PDF + b"\n5 0 obj<<>>endobj\ntrailer<</Size 6>>\n%%EOF"
    emf = build_emf([comment_record(COMMENT_MULTIFORMATS, updated)])
    preview = extract_metafile_preview(emf)
    assert preview is not None and preview.data.endswith(b"%%EOF")
    assert preview.data == updated


def test_a_pdf_with_no_end_marker_is_still_returned():
    """A truncated trailer usually still renders -- pdfium rebuilds the cross-reference
    table from the object offsets -- so giving up would put a grey placeholder in place
    of artwork that would have drawn."""
    truncated = MINI_PDF[: MINI_PDF.index(b"%%EOF")]
    emf = build_emf([comment_record(COMMENT_MULTIFORMATS, truncated)])
    preview = extract_metafile_preview(emf)
    assert preview is not None
    assert preview.mime_type == "application/pdf"
    assert preview.data.startswith(b"%PDF")


def test_a_comment_holding_only_a_pdf_marker_is_not_mistaken_for_a_pdf():
    """Without a length floor, four stray bytes anywhere in a comment would be carved
    out and handed to the rasteriser as a document."""
    emf = build_emf([comment_record(COMMENT_MULTIFORMATS, b"%PDF")])
    assert extract_metafile_preview(emf) is None


def test_comment_without_the_gdic_identifier_is_ignored():
    """An application-private comment is not ours to interpret, whatever it contains."""
    emf = build_emf([comment_record(COMMENT_MULTIFORMATS, MINI_PDF, identifier=0x11223344)])
    assert extract_metafile_preview(emf) is None


def test_comment_with_an_unlisted_public_type_is_ignored():
    emf = build_emf([comment_record(0x00000003, MINI_PDF)])  # ENDGROUP
    assert extract_metafile_preview(emf) is None


# --------------------------------------------------------------------------------------
# Embedded DIB
# --------------------------------------------------------------------------------------


def test_24bpp_dib_is_decoded_bottom_up_and_bgr_swapped():
    """The two things a DIB decoder gets wrong: row order and channel order."""
    bottom = bytes([0, 0, 255, 0, 255, 0]) + b"\0\0"     # BGR red, green + row padding
    top = bytes([255, 0, 0, 255, 255, 255]) + b"\0\0"    # BGR blue, white
    emf = build_emf([stretch_dib_record(2, 2, 24, bottom + top)])

    preview = extract_metafile_preview(emf)
    assert preview is not None and preview.mime_type == "image/png"
    width, height, color_type, samples = png_pixels(preview.data)
    assert (width, height, color_type) == (2, 2, 2)
    # First PNG row is the *top* of the image, which is the *last* row in the DIB.
    assert samples == [0, 0, 255, 255, 255, 255, 255, 0, 0, 0, 255, 0]


def test_negative_height_dib_is_stored_top_down():
    rows = bytes([0, 0, 255, 0, 0, 0]) + b"\0\0" + bytes([255, 255, 255, 0, 255, 0]) + b"\0\0"
    record = stretch_dib_record(2, 2, 24, rows)
    # Flip biHeight's sign in place; everything else about the record is unchanged.
    record = bytearray(record)
    off_bmi = struct.unpack_from("<I", record, 48)[0]
    struct.pack_into("<i", record, off_bmi + 8, -2)
    preview = extract_metafile_preview(build_emf([bytes(record)]))
    assert preview is not None
    _, _, _, samples = png_pixels(preview.data)
    assert samples[:3] == [255, 0, 0]  # first stored row is now the top row


def test_8bpp_dib_resolves_through_its_palette():
    palette = bytes([255, 0, 0, 0, 0, 255, 0, 0])  # BGRx entries: blue, green
    emf = build_emf([stretch_dib_record(2, 1, 8, bytes([0, 1, 0, 0]), palette=palette)])
    preview = extract_metafile_preview(emf)
    assert preview is not None
    _, _, _, samples = png_pixels(preview.data)
    assert samples == [0, 0, 255, 0, 255, 0]


def test_1bpp_dib_unpacks_bits_most_significant_first():
    palette = bytes([0, 0, 0, 0, 255, 255, 255, 0])  # index 0 black, index 1 white
    # 0b10000000 -> first pixel white, remaining seven black.
    emf = build_emf([stretch_dib_record(8, 1, 1, bytes([0x80, 0, 0, 0]), palette=palette)])
    preview = extract_metafile_preview(emf)
    assert preview is not None
    _, _, _, samples = png_pixels(preview.data)
    assert samples[:3] == [255, 255, 255]
    assert samples[3:6] == [0, 0, 0]


def test_32bpp_zero_alpha_is_treated_as_opaque():
    """BI_RGB leaves the fourth byte undefined; Office writes zero and means opaque."""
    pixels = bytes([0, 0, 255, 0, 0, 255, 0, 0])  # two BGRx pixels, alpha byte zero
    emf = build_emf([stretch_dib_record(2, 1, 32, pixels)])
    preview = extract_metafile_preview(emf)
    assert preview is not None
    _, _, color_type, samples = png_pixels(preview.data)
    assert color_type == 2, "an all-zero alpha channel must not produce an RGBA image"
    assert samples == [255, 0, 0, 0, 255, 0]


def test_32bpp_populated_alpha_is_honoured():
    pixels = bytes([0, 0, 255, 128, 0, 255, 0, 255])
    emf = build_emf([stretch_dib_record(2, 1, 32, pixels)])
    preview = extract_metafile_preview(emf)
    assert preview is not None
    _, _, color_type, samples = png_pixels(preview.data)
    assert color_type == 6
    assert samples == [255, 0, 0, 128, 0, 255, 0, 255]


def test_the_largest_of_several_bitmaps_is_chosen():
    small = stretch_dib_record(1, 1, 24, bytes([0, 0, 255, 0]))
    large = stretch_dib_record(4, 1, 24, bytes([255, 0, 0] * 4))
    preview = extract_metafile_preview(build_emf([small, large]))
    assert preview is not None
    assert png_size(preview.data) == (4, 1)


def test_an_embedded_pdf_outranks_an_embedded_bitmap():
    """A PDF is vector; falling back to the bitmap when both exist would lose detail."""
    emf = build_emf(
        [
            stretch_dib_record(4, 1, 24, bytes([255, 0, 0] * 4)),
            comment_record(COMMENT_MULTIFORMATS, MINI_PDF),
        ]
    )
    preview = extract_metafile_preview(emf)
    assert preview is not None and preview.mime_type == "application/pdf"


def test_bi_png_pixel_data_is_passed_through_untouched():
    inner = encode_png(1, 1, bytes([1, 2, 3, 255]), has_alpha=False)
    record = bytearray(stretch_dib_record(1, 1, 24, inner))
    off_bmi = struct.unpack_from("<I", record, 48)[0]
    struct.pack_into("<I", record, off_bmi + 16, 5)  # biCompression = BI_PNG
    preview = extract_metafile_preview(build_emf([bytes(record)]))
    assert preview is not None
    assert preview.mime_type == "image/png" and preview.data == inner


# --------------------------------------------------------------------------------------
# Hostile and malformed input
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "payload",
    [
        pytest.param(b"", id="empty"),
        pytest.param(b"\x00" * 200, id="zeroes"),
        pytest.param(b"\xd7\xcd\xc6\x9a" + b"\x00" * 200, id="placeable-wmf"),
        pytest.param(b"%PDF-1.4 not an emf %%EOF", id="bare-pdf"),
        pytest.param("not bytes", id="wrong-type"),
    ],
)
def test_non_emf_input_returns_none(payload):
    assert extract_metafile_preview(payload) is None


def test_truncation_at_any_offset_never_raises():
    """Every prefix of a valid EMF is a plausible corrupt file; none may throw."""
    emf = build_emf([comment_record(COMMENT_MULTIFORMATS, MINI_PDF)])
    for length in range(len(emf)):
        extract_metafile_preview(emf[:length])
    # Cutting mid-comment specifically must lose the preview, not return a partial PDF.
    assert extract_metafile_preview(emf[: EMF_HEADER_SIZE + 30]) is None


def test_a_zero_length_record_cannot_loop_forever():
    """The size field is what advances the walk; zero would pin it in place."""
    emf = emf_header(2, 96) + struct.pack("<II", 70, 0) + b"\0" * 8
    assert extract_metafile_preview(emf) is None


def test_a_record_claiming_more_bytes_than_exist_is_rejected():
    emf = emf_header(2, 96) + struct.pack("<II", 70, 0xFFFFFFF0) + b"\0" * 8
    assert extract_metafile_preview(emf) is None


def test_input_above_the_size_cap_is_refused_without_parsing():
    assert extract_metafile_preview(b"\0" * (MAX_INPUT_BYTES + 1)) is None


def test_a_dib_declaring_an_enormous_size_is_refused():
    """A 60-byte record must not be able to ask for a gigapixel allocation."""
    record = bytearray(stretch_dib_record(2, 1, 24, bytes([0, 0, 0, 0, 0, 0, 0, 0])))
    off_bmi = struct.unpack_from("<I", record, 48)[0]
    struct.pack_into("<ii", record, off_bmi + 4, 100000, 100000)
    assert extract_metafile_preview(build_emf([bytes(record)])) is None


def test_a_dib_with_a_truncated_pixel_buffer_still_decodes():
    """Padding a short buffer beats discarding a mostly-intact picture."""
    record = bytearray(stretch_dib_record(4, 4, 24, bytes(48)))
    off_bmi = struct.unpack_from("<I", record, 48)[0]
    struct.pack_into("<ii", record, off_bmi + 4, 4, 8)  # claim twice the rows we carry
    preview = extract_metafile_preview(build_emf([bytes(record)]))
    assert preview is not None and png_size(preview.data) == (4, 8)


def test_rle_compressed_dibs_are_declined():
    record = bytearray(stretch_dib_record(2, 1, 8, bytes([0, 1, 0, 0]), palette=bytes(8)))
    off_bmi = struct.unpack_from("<I", record, 48)[0]
    struct.pack_into("<I", record, off_bmi + 16, 1)  # BI_RLE8
    assert extract_metafile_preview(build_emf([bytes(record)])) is None


# --------------------------------------------------------------------------------------
# PNG encoder
# --------------------------------------------------------------------------------------


def test_encode_png_is_byte_stable():
    """Output determinism is load-bearing: the SVG golden tests embed these bytes."""
    rgba = bytes(range(0, 64)) * 4
    first = encode_png(8, 8, rgba, has_alpha=True)
    assert first == encode_png(8, 8, rgba, has_alpha=True)


def test_encode_png_round_trips_through_a_third_party_decoder():
    pytest.importorskip("PIL")
    from PIL import Image

    rgba = bytes([255, 0, 0, 255, 0, 255, 0, 128])
    image = Image.open(io.BytesIO(encode_png(2, 1, rgba, has_alpha=True)))
    assert image.mode == "RGBA" and image.size == (2, 1)
    assert list(image.getdata()) == [(255, 0, 0, 255), (0, 255, 0, 128)]


# --------------------------------------------------------------------------------------
# PDF rasterisation (optional extra)
# --------------------------------------------------------------------------------------


@pytest.mark.skipif(not pdf_rasterizer_available(), reason="pypdfium2 not installed")
def test_rasterise_pdf_sizes_the_render_from_the_on_slide_extent():
    png = rasterise_pdf(MINI_PDF, width_emu=914400)  # exactly one inch
    assert png is not None
    # One inch at PREVIEW_DPI (192), and the page's own 2:1 aspect ratio.
    assert png_size(png) == (192, 96)


@pytest.mark.skipif(not pdf_rasterizer_available(), reason="pypdfium2 not installed")
def test_rasterise_pdf_returns_none_for_a_broken_document():
    assert rasterise_pdf(b"%PDF-1.4 truncated", width_emu=914400) is None


# --------------------------------------------------------------------------------------
# Resolver integration
# --------------------------------------------------------------------------------------

EMF_PICTURE_XML = """<p:pic xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
 xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"
 xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
 <p:nvPicPr><p:cNvPr id="9001" name="Vector artwork"/><p:cNvPicPr/><p:nvPr/></p:nvPicPr>
 <p:blipFill><a:blip r:embed="rIdEmf"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>
 <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="1828800" cy="914400"/></a:xfrm>
 <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr></p:pic>"""

IMAGE_REL_TYPE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"


def deck_with_emf(basic_theme, emf_bytes: bytes) -> bytes:
    return derive_deck(
        basic_theme,
        parts={"ppt/media/vector1.emf": emf_bytes},
        shapes_xml=EMF_PICTURE_XML,
        slide_relationships=[("rIdEmf", IMAGE_REL_TYPE, "../media/vector1.emf")],
        # PowerPoint's own spelling for the EMF default content type.
        defaults={"emf": "image/x-emf"},
    )


def _emf_images(slide):
    """The pictures resolved from our spliced-in shape, found by its `cNvPr@id`.

    Slide shape ids are namespaced by the slide id (``"256.9001"``), which is what keeps
    `data-pptx-id` unique across a deck, so the match is on the suffix.
    """
    import pptx2svg.model as m

    return [
        e
        for e in slide.elements
        if isinstance(e, m.ImageElement) and (e.element_id or "").endswith(".9001")
    ]


def test_an_emf_with_a_dib_preview_resolves_to_an_image(basic_theme):
    emf = build_emf([stretch_dib_record(4, 2, 24, bytes([255, 0, 0] * 4 + [0, 255, 0] * 4))])
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck_with_emf(basic_theme, emf), options)

    images = _emf_images(resolved.slides[0])
    assert len(images) == 1
    assert images[0].mime_type == "image/png"
    assert not [w for w in options.warnings if w.code == "metafile-image"]


def test_an_emf_without_a_preview_keeps_the_placeholder_and_warns(basic_theme):
    emf = build_emf([])  # a valid but empty metafile: header, EOF, nothing between
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck_with_emf(basic_theme, emf), options)

    assert not _emf_images(resolved.slides[0])
    assert [w for w in options.warnings if w.code == "metafile-image"]


@pytest.mark.skipif(not pdf_rasterizer_available(), reason="pypdfium2 not installed")
def test_an_emf_with_a_pdf_preview_resolves_to_a_rasterised_image(basic_theme):
    emf = build_emf([comment_record(COMMENT_MULTIFORMATS, MINI_PDF)])
    options = ConvertOptions(slide_numbers=[1])
    resolved = convert_pptx_to_model(deck_with_emf(basic_theme, emf), options)

    images = _emf_images(resolved.slides[0])
    assert len(images) == 1 and images[0].mime_type == "image/png"


def test_a_metafile_converter_takes_precedence_over_the_embedded_preview(basic_theme):
    emf = build_emf([stretch_dib_record(4, 2, 24, bytes([255, 0, 0] * 8))])
    seen: list[tuple[int, str]] = []

    def converter(payload: bytes, mime_type: str):
        seen.append((len(payload), mime_type))
        return b'<svg xmlns="http://www.w3.org/2000/svg"/>', "image/svg+xml"

    options = ConvertOptions(slide_numbers=[1], metafile_converter=converter)
    resolved = convert_pptx_to_model(deck_with_emf(basic_theme, emf), options)

    assert seen == [(len(emf), "image/x-emf")]
    images = _emf_images(resolved.slides[0])
    assert len(images) == 1 and images[0].mime_type == "image/svg+xml"


def test_a_converter_returning_none_falls_back_to_the_embedded_preview(basic_theme):
    emf = build_emf([stretch_dib_record(4, 2, 24, bytes([255, 0, 0] * 8))])
    options = ConvertOptions(slide_numbers=[1], metafile_converter=lambda data, mime: None)
    resolved = convert_pptx_to_model(deck_with_emf(basic_theme, emf), options)

    images = _emf_images(resolved.slides[0])
    assert len(images) == 1 and images[0].mime_type == "image/png"


def test_a_raising_converter_warns_and_does_not_abort_the_deck(basic_theme):
    emf = build_emf([stretch_dib_record(4, 2, 24, bytes([255, 0, 0] * 8))])

    def converter(payload: bytes, mime_type: str):
        raise RuntimeError("inkscape exited 1")

    options = ConvertOptions(slide_numbers=[1], metafile_converter=converter)
    resolved = convert_pptx_to_model(deck_with_emf(basic_theme, emf), options)

    assert [w for w in options.warnings if w.code == "metafile-converter-failed"]
    images = _emf_images(resolved.slides[0])
    assert len(images) == 1 and images[0].mime_type == "image/png"


def test_the_rendered_svg_embeds_the_extracted_preview(basic_theme):
    emf = build_emf([stretch_dib_record(4, 2, 24, bytes([255, 0, 0] * 8))])
    svg = convert_pptx_to_svg(
        deck_with_emf(basic_theme, emf), ConvertOptions(slide_numbers=[1])
    )[0]
    assert 'data-pptx-id="256.9001"' in svg
    assert "data:image/png;base64," in svg


def test_svg_output_for_a_metafile_is_reproducible(basic_theme):
    emf = build_emf([stretch_dib_record(4, 2, 24, bytes([255, 0, 0] * 8))])
    deck = deck_with_emf(basic_theme, emf)
    first = convert_pptx_to_svg(deck, ConvertOptions(slide_numbers=[1]))[0]
    second = convert_pptx_to_svg(deck, ConvertOptions(slide_numbers=[1]))[0]
    assert first == second
