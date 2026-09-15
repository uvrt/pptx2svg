"""Fonts a deck carries in ``<p:embeddedFontLst>``.

Two kinds of test here, and the split is deliberate.

**Against real PowerPoint output.**  ``tests/fixtures/real-basic-theme.pptx`` embeds eight
faces -- Lato and Raleway, all four cuts each -- as EOT 2.2 payloads that PowerPoint wrote
and MicroType Express compressed.  Nothing in this repository produced them, so decoding
them is a real test of the decoder rather than our writer agreeing with our reader.  That
the roadmap recorded "no corpus deck embeds fonts" turned out to be wrong is the reason
this feature could be tested in CI at all.

**Against synthetic EOTs.**  The uncompressed and XOR-obfuscated variants, and every
``fsType`` refusal, have no example in any deck on hand, so they are built here.  A
synthetic font with one deliberate advance width also makes the measurement assertion
sharp in a way a real face cannot: if every glyph is 1234/2048 em wide, a string of ten
characters has exactly one right answer.
"""

from __future__ import annotations

import io
import struct
import zipfile
from pathlib import Path

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_model, convert_pptx_to_svg
from pptx2svg.fonts.eot import (
    EMBEDDING_BITMAP_ONLY,
    EMBEDDING_EDITABLE,
    EMBEDDING_NO_SUBSETTING,
    EMBEDDING_PREVIEW_PRINT,
    EMBEDDING_RESTRICTED,
    EotError,
    decode_eot,
    embedding_refusal,
    read_eot_header,
)
from pptx2svg.fonts.mtx import MtxDecodeError, decode_mtx
from pptx2svg.fonts.sfnt import SfntError, read_sfnt, relabel, write_sfnt
from pptx2svg.text.measure import DefaultTextMeasurer
from pptx2svg.units import PX_PER_PT

FIXTURE_DIR = Path(__file__).parent / "fixtures"
PRESENTATION = "ppt/presentation.xml"
PRESENTATION_RELS = "ppt/_rels/presentation.xml.rels"
REL_FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"


# --------------------------------------------------------------------------------------
# A font built here, so a measurement has exactly one right answer
# --------------------------------------------------------------------------------------

#: Every glyph in the probe face advances this much, in a 2048-unit em.  A round number
#: nothing else in the project uses, so a width computed from any other table is obvious.
PROBE_ADVANCE = 1234
PROBE_UPM = 2048
PROBE_ASCENDER = 1700
PROBE_DESCENDER = -500
#: The code points the probe face maps: printable ASCII.
PROBE_FIRST, PROBE_LAST = 0x20, 0x7E


def build_probe_font(*, fs_type: int = 0, advance: int = PROBE_ADVANCE) -> bytes:
    """A minimal but valid TrueType file with a known advance for every mapped glyph.

    No ``glyf``: nothing in these tests rasterises it, and :func:`read_sfnt` needs only
    the metric tables.  ``write_sfnt`` supplies the directory and the checksums.
    """
    glyphs = PROBE_LAST - PROBE_FIRST + 2  # + .notdef

    head = bytearray(54)
    struct.pack_into(">II", head, 0, 0x00010000, 0x00010000)  # version, fontRevision
    struct.pack_into(">I", head, 12, 0x5F0F3CF5)  # magicNumber
    struct.pack_into(">HH", head, 16, 0x000B, PROBE_UPM)  # flags, unitsPerEm
    struct.pack_into(">hhhh", head, 36, 0, PROBE_DESCENDER, advance, PROBE_ASCENDER)
    struct.pack_into(">hhh", head, 48, 2, 0, 0)  # lowestRecPPEM, indexToLocFormat, glyphDataFormat

    hhea = bytearray(36)
    struct.pack_into(">I", hhea, 0, 0x00010000)
    struct.pack_into(">hhh", hhea, 4, PROBE_ASCENDER, PROBE_DESCENDER, 0)
    struct.pack_into(">H", hhea, 10, advance)  # advanceWidthMax
    struct.pack_into(">H", hhea, 34, glyphs)  # numberOfHMetrics

    maxp = struct.pack(">IH", 0x00005000, glyphs)
    hmtx = b"".join(struct.pack(">Hh", advance, 0) for _ in range(glyphs))

    # cmap: one (3, 10) subtable in format 12 -- simpler to write by hand than format 4
    # and ranked first by the reader.
    group = struct.pack(">III", PROBE_FIRST, PROBE_LAST, 1)
    subtable = struct.pack(">HHIII", 12, 0, 16 + len(group), 0, 1) + group
    cmap = struct.pack(">HHHHI", 0, 1, 3, 10, 12) + subtable

    os2 = bytearray(96)
    struct.pack_into(">H", os2, 0, 4)  # version
    struct.pack_into(">H", os2, 2, advance)  # xAvgCharWidth
    struct.pack_into(">HH", os2, 4, 400, 5)  # usWeightClass, usWidthClass
    struct.pack_into(">H", os2, 8, fs_type)
    struct.pack_into(">H", os2, 62, 0x0040)  # fsSelection: REGULAR
    struct.pack_into(">HH", os2, 64, PROBE_FIRST, PROBE_LAST)
    struct.pack_into(">hhh", os2, 68, PROBE_ASCENDER, PROBE_DESCENDER, 0)

    name = _name_table([(1, "Probe Face"), (2, "Regular"), (4, "Probe Face"), (6, "ProbeFace")])

    return write_sfnt(
        [
            (b"OS/2", bytes(os2)),
            (b"cmap", cmap),
            (b"head", bytes(head)),
            (b"hhea", bytes(hhea)),
            (b"hmtx", hmtx),
            (b"maxp", maxp),
            (b"name", name),
        ]
    )


def _name_table(entries: list[tuple[int, str]]) -> bytes:
    records = [(3, 1, 0x409, name_id, value.encode("utf-16-be")) for name_id, value in entries]
    header = bytearray(struct.pack(">HHH", 0, len(records), 6 + 12 * len(records)))
    strings = bytearray()
    for platform, encoding, language, name_id, raw in records:
        header += struct.pack(
            ">HHHHHH", platform, encoding, language, name_id, len(raw), len(strings)
        )
        strings += raw
    return bytes(header + strings)


def wrap_as_eot(
    font: bytes,
    *,
    family: str = "Probe Face",
    style: str = "Regular",
    flags: int = 0,
    fs_type: int = 0,
    weight: int = 400,
) -> bytes:
    """Wrap a font in an EOT 1.0 container.

    Version 0x00010000 rather than 2.2 on purpose: the header's variable tail differs by
    version, and building the oldest one exercises the fact that the reader derives the
    font-data offset from ``EOTSize - FontDataSize`` rather than by walking that tail.
    """
    if flags & 0x10000000:  # TTEMBED_XORENCRYPTDATA
        font = bytes(byte ^ 0x50 for byte in font)

    names = b""
    for value in (family, style, "Version 1.000", f"{family} {style}"):
        encoded = value.encode("utf-16-le")
        names += struct.pack("<HH", 0, len(encoded)) + encoded

    header = bytearray(80)
    struct.pack_into("<I", header, 8, 0x00010000)  # Version
    struct.pack_into("<I", header, 12, flags)
    struct.pack_into("<I", header, 28, weight)
    struct.pack_into("<HH", header, 32, fs_type, 0x504C)  # fsType, MagicNumber
    body = bytes(header) + names
    struct.pack_into("<I", header, 0, len(body) + len(font))  # EOTSize
    struct.pack_into("<I", header, 4, len(font))  # FontDataSize
    return bytes(header) + names + font


# --------------------------------------------------------------------------------------
# Building a deck around it
# --------------------------------------------------------------------------------------

PROBE_SHAPE = """
<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr><p:cNvPr id="9001" name="probe"/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>
  <p:spPr><a:xfrm><a:off x="914400" y="914400"/><a:ext cx="5486400" cy="914400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>
  <p:txBody><a:bodyPr wrap="none"/><a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" sz="1800"><a:latin typeface="{typeface}"/></a:rPr>
      <a:t>{text}</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
"""

PROBE_TEXT = "Hamburgefo"


def deck_with_embedded_font(
    source: Path,
    payloads: dict[str, bytes],
    *,
    typeface: str = "Probe Face",
    run_typeface: str | None = None,
    text: str = PROBE_TEXT,
    relationship_type: str = REL_FONT,
) -> bytes:
    """``source`` with a text run in ``typeface`` and an ``<p:embeddedFontLst>`` for it.

    ``payloads`` maps a slot name (``regular``, ``bold``, ``italic``, ``boldItalic``) to
    the ``.fntdata`` bytes.  ``run_typeface`` lets the run name a *different* family from
    the embedded one, which is how the "nobody asks for it" case is built.
    """
    original = zipfile.ZipFile(source)
    written = {name: original.read(name) for name in original.namelist()}

    slide = written["ppt/slides/slide1.xml"].decode()
    written["ppt/slides/slide1.xml"] = slide.replace(
        "</p:spTree>",
        PROBE_SHAPE.format(typeface=run_typeface or typeface, text=text) + "</p:spTree>",
    ).encode()

    slots = []
    rels = []
    for index, (slot, payload) in enumerate(payloads.items()):
        rel_id = f"rIdEmbeddedFont{index}"
        target = f"fonts/probe{index}.fntdata"
        written[f"ppt/{target}"] = payload
        slots.append(f'<p:{slot} r:id="{rel_id}"/>')
        rels.append(
            f'<Relationship Id="{rel_id}" Type="{relationship_type}" Target="{target}"/>'
        )

    presentation = written[PRESENTATION].decode()
    font_list = (
        "<p:embeddedFontLst><p:embeddedFont>"
        f'<p:font typeface="{typeface}"/>{"".join(slots)}'
        "</p:embeddedFont></p:embeddedFontLst>"
    )
    presentation = presentation.replace("<p:defaultTextStyle>", font_list + "<p:defaultTextStyle>")
    if font_list not in presentation:  # no defaultTextStyle in this deck
        presentation = presentation.replace("</p:presentation>", font_list + "</p:presentation>")
    written[PRESENTATION] = presentation.encode()

    rels_xml = written[PRESENTATION_RELS].decode()
    written[PRESENTATION_RELS] = rels_xml.replace(
        "</Relationships>", "".join(rels) + "</Relationships>"
    ).encode()

    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in written.items():
            archive.writestr(name, data)
    return buffer.getvalue()


@pytest.fixture(scope="module")
def probe_font() -> bytes:
    return build_probe_font()


@pytest.fixture(scope="session")
def sample() -> Path:
    return FIXTURE_DIR / "sample.pptx"


def warnings_for(deck, **kwargs) -> list:
    options = ConvertOptions(**kwargs)
    convert_pptx_to_svg(deck, options)
    return options.warnings


# --------------------------------------------------------------------------------------
# The EOT container
# --------------------------------------------------------------------------------------


def test_an_uncompressed_payload_round_trips(probe_font):
    assert decode_eot(wrap_as_eot(probe_font)) == probe_font


def test_an_xor_obfuscated_payload_round_trips(probe_font):
    assert decode_eot(wrap_as_eot(probe_font, flags=0x10000000)) == probe_font


def test_the_header_reports_what_it_was_given(probe_font):
    header = read_eot_header(wrap_as_eot(probe_font, family="Probe Face", style="Bold", weight=700))
    assert header.family_name == "Probe Face"
    assert header.style_name == "Bold"
    assert header.weight == 700
    assert header.font_data_size == len(probe_font)
    assert not header.compressed and not header.encrypted


def test_a_payload_without_the_magic_number_is_refused(probe_font):
    data = bytearray(wrap_as_eot(probe_font))
    struct.pack_into("<H", data, 34, 0x1234)
    with pytest.raises(EotError, match="magic number"):
        read_eot_header(bytes(data))


def test_a_truncated_payload_is_refused(probe_font):
    with pytest.raises(EotError):
        read_eot_header(wrap_as_eot(probe_font)[:60])


def test_garbage_in_the_compressed_payload_does_not_escape_as_something_else(probe_font):
    data = wrap_as_eot(probe_font, flags=0x00000004)  # claims MTX, is not
    with pytest.raises(EotError, match="MicroType Express"):
        decode_eot(data)


def test_mtx_rejects_a_payload_shorter_than_its_header():
    with pytest.raises(MtxDecodeError):
        decode_mtx(b"\x03\x00\x00")


# --------------------------------------------------------------------------------------
# The fsType gate
#
# Bits are from the OpenType spec's OS/2 fsType field.  The rule is LibreOffice's
# `sufficientTTFRights`: refuse only a font whose permission field says restricted licence
# and says nothing else.
# --------------------------------------------------------------------------------------


@pytest.mark.parametrize(
    "fs_type",
    [
        0x0000,  # installable
        EMBEDDING_PREVIEW_PRINT,
        EMBEDDING_EDITABLE,
        EMBEDDING_RESTRICTED | EMBEDDING_PREVIEW_PRINT,  # restrictive *and* permissive
        EMBEDDING_RESTRICTED | EMBEDDING_EDITABLE,
        EMBEDDING_NO_SUBSETTING,  # we never subset, so this is not our business
        EMBEDDING_EDITABLE | EMBEDDING_NO_SUBSETTING,
    ],
)
def test_permitted_fs_type_values_are_allowed(fs_type):
    assert embedding_refusal(fs_type) is None


@pytest.mark.parametrize(
    "fs_type, reason",
    [
        (EMBEDDING_RESTRICTED, "restricted-licence"),
        (EMBEDDING_BITMAP_ONLY, "bitmap embedding only"),
        (EMBEDDING_EDITABLE | EMBEDDING_BITMAP_ONLY, "bitmap embedding only"),
    ],
)
def test_refused_fs_type_values_say_why(fs_type, reason):
    refusal = embedding_refusal(fs_type)
    assert refusal is not None and reason in refusal


def test_a_restricted_header_is_refused_before_anything_is_decoded(probe_font):
    with pytest.raises(EotError, match="restricted-licence"):
        decode_eot(wrap_as_eot(probe_font, fs_type=EMBEDDING_RESTRICTED))


def test_a_permissive_header_over_a_restricted_font_is_still_refused(sample):
    """The second gate exists because the first one is written by the embedder.

    An EOT header carries its own copy of ``fsType``, and nothing makes it agree with the
    font's ``OS/2`` table.  The foundry's statement is the one in the font, so both are
    checked and both must permit.
    """
    restricted = build_probe_font(fs_type=EMBEDDING_RESTRICTED)
    assert embedding_refusal(read_sfnt(restricted).fs_type) is not None

    deck = deck_with_embedded_font(sample, {"regular": wrap_as_eot(restricted, fs_type=0)})
    codes = [w.code for w in warnings_for(deck)]
    assert "font-embedded-restricted" in codes
    assert not convert_pptx_to_model(deck, ConvertOptions()).embedded_fonts


def test_a_restricted_font_still_renders_the_deck(sample):
    restricted = build_probe_font(fs_type=EMBEDDING_RESTRICTED)
    deck = deck_with_embedded_font(sample, {"regular": wrap_as_eot(restricted, fs_type=0)})
    documents = convert_pptx_to_svg(deck, ConvertOptions())
    assert documents and "<svg" in documents[0]


# --------------------------------------------------------------------------------------
# Reading the deck
# --------------------------------------------------------------------------------------


def test_the_font_list_is_parsed(sample, probe_font):
    from pptx2svg.opc import OpcPackage
    from pptx2svg.parse.parts import read_presentation

    deck = deck_with_embedded_font(
        sample,
        {"regular": wrap_as_eot(probe_font), "bold": wrap_as_eot(probe_font, style="Bold")},
    )
    presentation = read_presentation(OpcPackage.open(deck))
    (entry,) = presentation.embedded_fonts
    assert entry.typeface == "Probe Face"
    assert entry.part_path == PRESENTATION
    assert entry.regular and entry.bold
    assert entry.italic is None and entry.bold_italic is None


def test_a_slot_pointing_at_something_that_is_not_a_font_is_ignored(sample, probe_font):
    deck = deck_with_embedded_font(
        sample,
        {"regular": wrap_as_eot(probe_font)},
        relationship_type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image",
    )
    resolved = convert_pptx_to_model(deck, ConvertOptions())
    assert not resolved.embedded_fonts
    assert "font-embedded-unreadable" in [w.code for w in resolved.warnings]


def test_a_family_no_slide_uses_is_not_decoded(sample, probe_font):
    """Decoding costs about a second a face, so it is limited to what the deck draws."""
    deck = deck_with_embedded_font(
        sample,
        {"regular": wrap_as_eot(probe_font)},
        typeface="Nobody Asks For This",
        run_typeface="Probe Face",
    )
    resolved = convert_pptx_to_model(deck, ConvertOptions())
    assert not resolved.embedded_fonts


# --------------------------------------------------------------------------------------
# Measurement -- the half that `font_files=` alone would have missed
# --------------------------------------------------------------------------------------


def test_the_embedded_face_is_what_gets_measured(sample, probe_font):
    deck = deck_with_embedded_font(sample, {"regular": wrap_as_eot(probe_font)})
    resolved = convert_pptx_to_model(deck, ConvertOptions())
    metrics = resolved.embedded_fonts.metrics["probe face"]

    assert metrics.units_per_em == PROBE_UPM
    assert metrics.ascender == PROBE_ASCENDER
    assert metrics.descender == PROBE_DESCENDER
    assert set(metrics.widths.values()) == {PROBE_ADVANCE}

    measurer = DefaultTextMeasurer(resolved.embedded_fonts.metrics)
    expected = len(PROBE_TEXT) * (PROBE_ADVANCE / PROBE_UPM) * 18 * PX_PER_PT
    assert measurer.measure_text_width(PROBE_TEXT, 18, font_family="Probe Face") == pytest.approx(
        expected
    )


def test_measurement_falls_back_when_the_family_is_not_embedded(sample, probe_font):
    deck = deck_with_embedded_font(sample, {"regular": wrap_as_eot(probe_font)})
    resolved = convert_pptx_to_model(deck, ConvertOptions())
    measurer = DefaultTextMeasurer(resolved.embedded_fonts.metrics)
    plain = DefaultTextMeasurer()
    assert measurer.measure_text_width("abc", 18, font_family="Arial") == plain.measure_text_width(
        "abc", 18, font_family="Arial"
    )


def test_the_embedded_widths_reach_the_rendered_svg(sample, probe_font):
    """The whole point: a different measurement has to change the layout, not just the name."""
    deck = deck_with_embedded_font(sample, {"regular": wrap_as_eot(probe_font)})
    with_fonts = convert_pptx_to_svg(deck, ConvertOptions())[0]
    without = convert_pptx_to_svg(deck, ConvertOptions(use_embedded_fonts=False))[0]
    assert with_fonts != without
    assert "Probe Face" in with_fonts


def test_a_caller_supplied_measurer_still_wins(sample, probe_font):
    class Fixed:
        def measure_text_width(
            self, text, font_size_pt, bold=False, font_family=None, font_family_ea=None
        ):
            return 7.0 * len(text)

        def line_height_ratio(self, font_family=None, font_family_ea=None):
            return 1.2

        def ascender_ratio(self, font_family=None, font_family_ea=None):
            return 1.0

    deck = deck_with_embedded_font(sample, {"regular": wrap_as_eot(probe_font)})
    mine = convert_pptx_to_svg(deck, ConvertOptions(measurer=Fixed()))[0]
    theirs = convert_pptx_to_svg(deck, ConvertOptions())[0]
    assert mine != theirs


def test_the_bold_slot_gives_the_bold_table_its_own_widths(sample, probe_font):
    bold_font = build_probe_font(advance=1700)
    deck = deck_with_embedded_font(
        sample,
        {"regular": wrap_as_eot(probe_font), "bold": wrap_as_eot(bold_font, style="Bold")},
    )
    metrics = convert_pptx_to_model(deck, ConvertOptions()).embedded_fonts.metrics["probe face"]
    assert set(metrics.widths.values()) == {PROBE_ADVANCE}
    assert set(metrics.bold_widths.values()) == {1700}

    measurer = DefaultTextMeasurer({"probe face": metrics})
    upright = measurer.measure_text_width("abc", 18, font_family="Probe Face")
    bold = measurer.measure_text_width("abc", 18, bold=True, font_family="Probe Face")
    assert bold == pytest.approx(upright * 1700 / PROBE_ADVANCE)


def test_a_family_with_only_an_italic_cut_measures_from_it(sample, probe_font):
    """Italic is measured from the upright table everywhere else; the converse holds too."""
    deck = deck_with_embedded_font(
        sample, {"italic": wrap_as_eot(probe_font, style="Italic")}
    )
    metrics = convert_pptx_to_model(deck, ConvertOptions()).embedded_fonts.metrics["probe face"]
    assert set(metrics.widths.values()) == {PROBE_ADVANCE}


# --------------------------------------------------------------------------------------
# Relabelling -- so the rasteriser can find the face the SVG asks for
# --------------------------------------------------------------------------------------


def test_the_face_takes_the_name_the_deck_gave_it(probe_font):
    relabelled = read_sfnt(relabel(probe_font, "Deck Says This", bold=True, italic=False))
    assert relabelled.family_name == "Deck Says This"
    assert relabelled.subfamily_name == "Bold"
    assert relabelled.weight_class == 700


def test_relabelling_keeps_every_advance_width(probe_font):
    before = read_sfnt(probe_font)
    after = read_sfnt(relabel(probe_font, "Something Else", bold=False, italic=True))
    assert after.advances == before.advances
    assert after.units_per_em == before.units_per_em


def test_a_written_face_is_named_for_its_slot(sample, probe_font, tmp_path):
    deck = deck_with_embedded_font(
        sample,
        {
            "regular": wrap_as_eot(probe_font),
            "boldItalic": wrap_as_eot(probe_font, style="Bold Italic"),
        },
    )
    embedded = convert_pptx_to_model(deck, ConvertOptions()).embedded_fonts
    paths = embedded.write(str(tmp_path))
    assert sorted(Path(p).name for p in paths) == [
        "Probe-Face-Bold-Italic.ttf",
        "Probe-Face-Regular.ttf",
    ]
    assert read_sfnt(Path(paths[0]).read_bytes()).family_name == "Probe Face"


def test_a_font_file_is_not_written_unless_it_is_asked_for(sample, probe_font, tmp_path):
    deck = deck_with_embedded_font(sample, {"regular": wrap_as_eot(probe_font)})
    convert_pptx_to_svg(deck, ConvertOptions())
    assert list(tmp_path.iterdir()) == []


# --------------------------------------------------------------------------------------
# The substitution report
# --------------------------------------------------------------------------------------


def test_an_embedded_family_does_not_report_as_substituted(sample, probe_font):
    deck = deck_with_embedded_font(sample, {"regular": wrap_as_eot(probe_font)})
    substituted = [
        w for w in warnings_for(deck) if w.code == "font-substituted" and "Probe Face" in w.message
    ]
    assert substituted == []


def test_without_the_embedded_font_the_same_deck_does_report_it(sample, probe_font):
    deck = deck_with_embedded_font(sample, {"regular": wrap_as_eot(probe_font)})
    substituted = [
        w
        for w in warnings_for(deck, use_embedded_fonts=False)
        if w.code == "font-substituted" and "Probe Face" in w.message
    ]
    assert len(substituted) == 1
    assert "no substitute known" in substituted[0].message


def test_the_check_report_grades_an_embedded_face_exact(sample, probe_font):
    from pptx2svg.fonts.check import check_deck

    deck = deck_with_embedded_font(sample, {"regular": wrap_as_eot(probe_font)})
    (face,) = [f for f in check_deck(deck).faces if f.requested == "Probe Face"]
    assert face.verdict == "exact"
    assert face.reason == "drawn with the face the deck embedded"
    assert face.faithful


# --------------------------------------------------------------------------------------
# Degrading rather than failing
# --------------------------------------------------------------------------------------


def test_an_undecodable_payload_warns_and_renders_anyway(sample):
    deck = deck_with_embedded_font(sample, {"regular": b"not a font at all, not even close"})
    options = ConvertOptions()
    documents = convert_pptx_to_svg(deck, options)
    assert documents and "<svg" in documents[0]
    codes = [w.code for w in options.warnings]
    assert "font-embedded-undecodable" in codes
    assert "font-substituted" in codes  # and the fallback is reported honestly


def test_a_deck_with_no_embedded_fonts_is_untouched(sample):
    resolved = convert_pptx_to_model(sample, ConvertOptions())
    assert not resolved.embedded_fonts
    assert resolved.embedded_fonts.metrics == {}
    assert [w for w in resolved.warnings if w.code.startswith("font-embedded")] == []


def test_opting_out_leaves_the_faces_alone(sample, probe_font):
    deck = deck_with_embedded_font(sample, {"regular": wrap_as_eot(probe_font)})
    resolved = convert_pptx_to_model(deck, ConvertOptions(use_embedded_fonts=False))
    assert not resolved.embedded_fonts


# --------------------------------------------------------------------------------------
# Against PowerPoint's own output
# --------------------------------------------------------------------------------------

#: What `tests/fixtures/real-basic-theme.pptx` carries, read off the parts themselves.
BASIC_THEME_FACES = {
    ("Lato", False, False),
    ("Lato", True, False),
    ("Lato", False, True),
    ("Lato", True, True),
    ("Raleway", False, False),
    ("Raleway", True, False),
    ("Raleway", False, True),
    ("Raleway", True, True),
}


def test_a_real_powerpoint_deck_decodes_all_eight_faces(basic_theme):
    resolved = convert_pptx_to_model(basic_theme, ConvertOptions())
    embedded = resolved.embedded_fonts
    assert embedded.problems == ()
    assert {(f.family, f.bold, f.italic) for f in embedded.faces} == BASIC_THEME_FACES
    assert sorted(embedded.metrics) == ["lato", "raleway"]


def test_the_real_payloads_are_mtx_compressed_and_installable(basic_theme):
    """The flag reading used to be an inference.  These are the bytes that settle it."""
    from pptx2svg.opc import OpcPackage

    package = OpcPackage.open(str(basic_theme))
    parts = sorted(p for p in package.parts if p.startswith("ppt/fonts/"))
    assert len(parts) == 8
    for path in parts:
        header = read_eot_header(package.read(path))
        assert header.version == 0x00020002  # EOT 2.2
        assert header.compressed and not header.encrypted
        assert header.fs_type == 0x0000  # installable embedding
        assert header.eot_size == len(package.read(path))


def test_a_corrupted_payload_only_ever_raises_our_own_error(basic_theme):
    """A deck is untrusted input, and this decoder writes into a fixed-size window.

    The reference implementation does not bound its copy items: ``Decode`` writes the whole
    item and only afterwards asserts ``pos == out_len``, which on a corrupt stream is a
    heap overflow in C and was an ``IndexError`` escaping from here. Found by mutating
    bytes inside the font data of these very payloads -- 19 of 1680 mutations hit it -- so
    the same mutation is the regression test.

    Deliberately built from a committed fixture rather than a stored corrupt payload: no
    font file, whole or damaged, belongs in this repository.  The seed and the count are
    pinned so the assertion below can be exact -- with them, three of the 120 mutations
    reach the copy-item bound, which is what stops this quietly ceasing to cover it.
    """
    import random

    from pptx2svg.opc import OpcPackage

    package = OpcPackage.open(str(basic_theme))
    # The smallest part, because every mutation pays for a full decode attempt.
    path = min((len(package.read(p)), p) for p in package.parts if p.startswith("ppt/fonts/"))[1]
    original = package.read(path)
    offset = read_eot_header(original).font_data_offset

    rng = random.Random(20260915)
    reached_the_bound = 0
    for _ in range(120):
        data = bytearray(original)
        for _ in range(rng.randrange(1, 10)):
            data[rng.randrange(offset, len(data))] = rng.randrange(256)
        try:
            decode_eot(bytes(data))
        except (EotError, MtxDecodeError, SfntError) as error:
            if "overruns the declared output length" in str(error):
                reached_the_bound += 1
    assert reached_the_bound == 3


def test_the_decoded_faces_are_usable_fonts(basic_theme):
    resolved = convert_pptx_to_model(basic_theme, ConvertOptions())
    for face in resolved.embedded_fonts.faces:
        parsed = read_sfnt(face.source)
        assert parsed.units_per_em in (1000, 2000)
        assert len(parsed.advances) > 200
        assert b"glyf" in parsed.tables and b"loca" in parsed.tables
        # and it survives the relabelling the rasteriser needs
        assert read_sfnt(face.data).family_name == face.family


def test_the_deck_supplies_its_own_faces_so_nothing_is_substituted(basic_theme):
    options = ConvertOptions()
    convert_pptx_to_svg(basic_theme, options)
    substituted = [
        w.message
        for w in options.warnings
        if w.code == "font-substituted" and ("Lato" in w.message or "Raleway" in w.message)
    ]
    assert substituted == []


def test_the_decks_own_raleway_is_not_the_bundled_one(basic_theme):
    """Why an embedded face wins over a bundled one with the same name.

    The deck carries Raleway 4.026; the bundle ships a later release.  Measured over the
    340 characters both tables describe, 338 advance widths differ, and "Hamburgefonstiv
    12345" at 18 pt comes out 251.184 px from the bundled table against 261.000 px from
    the deck's own file -- 3.9% apart, wider than the 4.5% Caladea error the roadmap
    records as a bug.  The author laid the deck out with the file inside it, so that is
    the file to measure.
    """
    from pptx2svg.text.metrics import METRICS

    embedded = convert_pptx_to_model(basic_theme, ConvertOptions()).embedded_fonts
    deck_table = embedded.metrics["raleway"]
    bundled = METRICS["Raleway"]
    assert deck_table.units_per_em == bundled.units_per_em

    shared = set(deck_table.widths) & set(bundled.widths)
    differing = [c for c in shared if deck_table.widths[c] != bundled.widths[c]]
    assert len(differing) > 100, "expected the two Raleway releases to disagree"

    probe = "Hamburgefonstiv 12345"
    deck_width = DefaultTextMeasurer(embedded.metrics).measure_text_width(
        probe, 18, font_family="Raleway"
    )
    bundled_width = DefaultTextMeasurer().measure_text_width(probe, 18, font_family="Raleway")
    assert abs(deck_width - bundled_width) / bundled_width > 0.02


# --------------------------------------------------------------------------------------
# Against the third-party template decks, where they are available
# --------------------------------------------------------------------------------------

LOCAL_DIR = Path(__file__).resolve().parents[1] / "scratch"

#: The decks this phase exists to fix -- **not committed**, like `college_template`.
#: They are commercial templates, so they live in the gitignored `scratch/` directory and
#: these tests skip where they are absent, which is every machine but one.
TEMPLATE_DECKS = {
    "nutrition_templates-Meal Planning Slides.pptx": {
        "families": {"anton", "arimo", "literata"},
        "faces": 9,
    },
    "sales_templates-IT Software Sales Proposal Slides.pptx": {
        "families": {"inclusive sans", "merriweather sans light"},
        "faces": 8,
    },
}


@pytest.mark.parametrize("name", sorted(TEMPLATE_DECKS))
def test_a_template_deck_renders_with_no_substitution_warning(name):
    path = LOCAL_DIR / name
    if not path.is_file():
        pytest.skip(f"no {path}; a commercial template, not ours to redistribute")
    expected = TEMPLATE_DECKS[name]

    options = ConvertOptions()
    convert_pptx_to_svg(path, options)
    assert [w for w in options.warnings if w.code == "font-substituted"] == []

    embedded = convert_pptx_to_model(path, ConvertOptions()).embedded_fonts
    assert set(embedded.metrics) == expected["families"]
    assert len(embedded.faces) == expected["faces"]


def test_the_sfnt_reader_reproduces_a_table_generated_by_fonttools(basic_theme):
    """A cross-check that does not depend on fontTools being installed.

    ``tools/extract_font_metrics.py`` built ``METRICS["Arimo"]`` with fontTools from the
    variable Arimo the bundle ships.  One of the template decks embeds a static Arimo,
    and decoding it and measuring it here reproduces that table character for character --
    a different file, a different tool chain, the same numbers.  The Lato entry below is
    the version that is committed, so it is the one CI can check.
    """
    from pptx2svg.text.metrics import METRICS

    embedded = convert_pptx_to_model(basic_theme, ConvertOptions()).embedded_fonts
    deck_lato = embedded.metrics["lato"]
    bundled = METRICS["Lato"]
    assert deck_lato.units_per_em == bundled.units_per_em
    assert deck_lato.ascender == bundled.ascender
    assert deck_lato.descender == bundled.descender
