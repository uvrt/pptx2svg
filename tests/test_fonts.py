"""The font bundle, and the invariant that makes it worth having.

The bug these guard against is not a crash.  It is that layout was computed from one
face's advance widths and drawn with another's, and nothing anywhere said so: resvg
substitutes silently, so the output looked exactly like correct output.  Every test here
is really the same assertion from a different angle -- *we draw with what we measured
with, and where we cannot, we say so out loud*.
"""

from __future__ import annotations

import re
import subprocess
import sys
from pathlib import Path

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_svg
from pptx2svg.fonts import (
    BUNDLED_FAMILIES,
    GENERIC_FAMILY_DEFAULTS,
    available_families,
    bundle_dir,
    bundle_mode,
    font_dirs,
)
from pptx2svg.fonts.check import check_deck, check_families
from pptx2svg.text.fontmap import (
    SUBSTITUTIONS,
    font_family_value,
    generic_family,
    substitution_for,
    typographic_family,
)
from pptx2svg.text.measure import DefaultTextMeasurer
from pptx2svg.text.metrics import METRICS

ROOT = Path(__file__).resolve().parents[1]


def _bundle() -> Path:
    directory = bundle_dir()
    if directory is None:
        pytest.skip("pptx2svg-fonts is not importable; run `pip install 'pptx2svg[fonts]'`")
    return directory


# --------------------------------------------------------------------------------------
# The invariant
# --------------------------------------------------------------------------------------

def test_a_metric_compatible_substitute_is_also_what_we_measure_with():
    """The whole point, in one assertion.

    A substitution claiming metric compatibility must name the *same* family for drawing
    and for measuring.  If those two fields can differ while the flag is set, the flag
    means nothing and the silent-substitution bug is back.
    """
    for substitution in SUBSTITUTIONS.values():
        if substitution.metric_compatible:
            assert substitution.metrics == substitution.substitute, substitution.office


def test_every_substitution_points_at_a_family_we_can_draw():
    for substitution in SUBSTITUTIONS.values():
        assert substitution.substitute in BUNDLED_FAMILIES, substitution.office


def test_every_substitution_points_at_a_metrics_table_that_exists():
    for substitution in SUBSTITUTIONS.values():
        assert substitution.metrics in METRICS, substitution.office


# --------------------------------------------------------------------------------------
# The other direction: a family we ship, named by a deck
# --------------------------------------------------------------------------------------
#
# The table was written from the Office side -- "a deck asked for Calibri, what do we
# draw?" -- and for a long time that was the only direction it worked in.  A family was a
# legal *input* only if some Office face happened to be spelled that way, so the five
# faces we ship, measure and draw with were names we did not recognise.  Every test below
# is the same assertion as the ones above, arriving from the other side: what we can draw,
# we can be asked for.

#: Names a deck may spell that must reach the bundle, beyond the bundled families
#: themselves.  Liberation 2.x is built from the Chrome OS core fonts, which is why
#: ``tools/install-fonts-debian.sh`` installs ``fonts-liberation2`` deliberately.
LIBERATION_ALIASES = {
    "Liberation Sans": "Arimo",
    "Liberation Serif": "Tinos",
    "Liberation Mono": "Cousine",
}

#: The generic fallback width for this string at this size: 21 characters of
#: NORMAL_RATIO/NARROW_RATIO guess at 18 pt, which is what *every* unrecognised face
#: measures.  Hard-coded rather than computed so that a change to the heuristic does not
#: quietly make the assertions below vacuous.
PROBE = "Hamburgefonstiv 12345"
PROBE_SIZE_PT = 18.0
PROBE_GENERIC_WIDTH = 280.800


def test_the_generic_fallback_width_is_what_we_think_it_is():
    """The control for the three tests after it -- an assertion about nothing real."""
    measurer = DefaultTextMeasurer()
    width = measurer.measure_text_width(PROBE, PROBE_SIZE_PT, False, "Nonexistent Face")
    assert width == pytest.approx(PROBE_GENERIC_WIDTH, abs=0.001)


def test_a_bundled_family_named_by_a_deck_is_measured_not_guessed():
    """The bug: Carlito measured identically to a face that does not exist.

    Decks really do name these.  Carlito and Caladea exist because LibreOffice ships them
    as its Calibri and Cambria substitutes, so anything round-tripped through LibreOffice
    names them directly, and Arimo/Tinos/Cousine are the Chrome OS core fonts that Debian
    packages as ``fonts-croscore``.  All five, plus the three Liberation aliases, came
    back at exactly 280.800 px -- the generic guess -- while their own tables were sitting
    in ``METRICS`` all along.
    """
    measurer = DefaultTextMeasurer()
    for family in sorted(BUNDLED_FAMILIES) + sorted(LIBERATION_ALIASES):
        width = measurer.measure_text_width(PROBE, PROBE_SIZE_PT, False, family)
        assert width != pytest.approx(PROBE_GENERIC_WIDTH, abs=0.001), family


def test_a_substitute_measures_the_same_by_its_own_name_as_by_the_office_name():
    """Two routes to one table.  If they disagree, one of them is not using the table."""
    measurer = DefaultTextMeasurer()
    pairs = (
        ("Calibri", "Carlito"), ("Arial", "Arimo"), ("Times New Roman", "Tinos"),
        ("Courier New", "Cousine"),
        # ...and the aliases of the substitutes, which are the same designs again.
        ("Arial", "Liberation Sans"), ("Times New Roman", "Liberation Serif"),
        ("Courier New", "Liberation Mono"),
    )
    for office, own in pairs:
        for bold in (False, True):
            by_office = measurer.measure_text_width(PROBE, PROBE_SIZE_PT, bold, office)
            by_own = measurer.measure_text_width(PROBE, PROBE_SIZE_PT, bold, own)
            assert by_office == pytest.approx(by_own, abs=0.001), (office, own, bold)

    # Cambria is deliberately *not* in that list: Caladea runs 4.5% narrow, so the deck
    # asking for Cambria is measured from Cambria's own table and the deck asking for
    # Caladea is measured from Caladea's.  Both are right; they are different faces.
    cambria = measurer.measure_text_width(PROBE, PROBE_SIZE_PT, False, "Cambria")
    caladea = measurer.measure_text_width(PROBE, PROBE_SIZE_PT, False, "Caladea")
    assert cambria != pytest.approx(caladea, abs=0.001)


def test_a_bundled_family_named_by_a_deck_reports_as_faithful():
    report = check_families(sorted(BUNDLED_FAMILIES))
    if report.mode != "bundled":
        pytest.skip("pptx2svg-fonts is not importable")
    assert report.faithful
    assert {face.verdict for face in report.faces} == {"exact"}
    # The old verdict said the opposite of the truth: "no substitute known" for a face
    # whose file is in the wheel and whose widths are in METRICS.
    for face in report.faces:
        assert "no substitute known" not in face.reason, face.requested


def test_the_liberation_aliases_report_as_compatible_and_name_the_design():
    report = check_families(sorted(LIBERATION_ALIASES))
    if report.mode != "bundled":
        pytest.skip("pptx2svg-fonts is not importable")
    assert report.faithful
    for face in report.faces:
        assert face.verdict == "compatible"
        assert face.substitute == LIBERATION_ALIASES[face.requested]
        assert "same design" in face.reason


def test_a_bundled_family_is_named_first_in_its_own_font_stack():
    """A correct width behind a wrong ``font-family`` still renders wrong.

    resvg resolves per text chunk, so the two halves have to be checked separately: the
    measurement can be right while the stack sends the chunk somewhere else entirely.
    """
    for family in sorted(BUNDLED_FAMILIES):
        names = font_family_value([family]).split(", ")
        assert names[0].strip("'") == family, family
        assert names[-1] in ("sans-serif", "serif", "monospace"), family
    for alias, design in sorted(LIBERATION_ALIASES.items()):
        names = font_family_value([alias]).split(", ")
        assert names[0].strip("'") == alias, alias
        assert design in names, alias


def test_the_generic_behind_a_bundled_family_matches_its_design():
    """Nothing in "Tinos" or "Cousine" says serif or monospace, so guessing gets it wrong.

    It is the last resort in the stack, which is exactly where it matters: in ``system``
    mode a deck set in Tinos used to end its stack in ``sans-serif`` and degrade to
    resvg's sans default.
    """
    assert generic_family("Tinos") == "serif"
    assert generic_family("Caladea") == "serif"
    assert generic_family("Cousine") == "monospace"
    assert generic_family("Carlito") == "sans-serif"
    assert generic_family("Arimo") == "sans-serif"
    # Case-folded like every other lookup in the module.
    assert generic_family("cousine") == "monospace"
    # Hand-written, unlike the identity rows, so it needs a guard against drift: a ninth
    # bundled family would otherwise fall back to reading its name.
    from pptx2svg.text.fontmap import _BUNDLED_GENERICS

    assert set(_BUNDLED_GENERICS) == {family.lower() for family in BUNDLED_FAMILIES}


def test_a_family_name_resolves_however_the_deck_spells_it():
    """OOXML puts no constraint on the capitalisation of a ``typeface`` attribute.

    A name we fail to match does not fail loudly -- it becomes the 0.6 em guess -- so the
    lookup is deliberately forgiving.  Inner whitespace is deliberately *not* normalised;
    see ``_key``.
    """
    measurer = DefaultTextMeasurer()
    expected = measurer.measure_text_width(PROBE, PROBE_SIZE_PT, False, "Carlito")
    for spelling in ("carlito", "CARLITO", "Carlito ", " carlito"):
        width = measurer.measure_text_width(PROBE, PROBE_SIZE_PT, False, spelling)
        assert width == pytest.approx(expected, abs=0.001), spelling
        # ...and the stack still names the canonical spelling, so the chunk resolves.
        assert "Carlito" in font_family_value([spelling]).split(", ")


def test_every_family_we_can_draw_is_a_family_we_can_be_asked_for():
    """The invariant, stated once.  A ninth bundled family cannot ship without a row."""
    for family in BUNDLED_FAMILIES:
        substitution = substitution_for(family)
        assert substitution is not None, family
        assert substitution.substitute == family, family
        assert substitution.exact, family
        assert substitution.metrics in METRICS, family


#: A 24 pt line of Carlito in a 250 px box with no insets.  Carlito's own advance widths
#: put "Hamburgefonstiv" at 225.078 px, so it fits; the 0.6 em per-character guess puts it
#: at 278.400 px, so it does not.  The deck is derived rather than committed -- no font
#: file, and nothing opaque, enters the corpus.
CARLITO_PROBE_XML = """
<p:sp xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"
      xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main">
  <p:nvSpPr>
    <p:cNvPr id="9001" name="Carlito probe"/>
    <p:cNvSpPr txBox="1"/>
    <p:nvPr/>
  </p:nvSpPr>
  <p:spPr>
    <a:xfrm><a:off x="0" y="0"/><a:ext cx="2381250" cy="914400"/></a:xfrm>
    <a:prstGeom prst="rect"><a:avLst/></a:prstGeom>
  </p:spPr>
  <p:txBody>
    <a:bodyPr wrap="square" lIns="0" tIns="0" rIns="0" bIns="0"><a:noAutofit/></a:bodyPr>
    <a:lstStyle/>
    <a:p><a:r><a:rPr lang="en-US" sz="2400"><a:latin typeface="Carlito"/></a:rPr>
      <a:t>Hamburgefonstiv</a:t></a:r></a:p>
  </p:txBody>
</p:sp>
"""


def test_a_deck_naming_a_bundled_face_lays_out_from_its_own_widths(authoring):
    """End to end, because the unit widths being right is only half of it.

    This is the failure as a reader of the output would meet it, which is to say not at
    all: the text still drew, in the right family, at the right size -- and broke
    "Hamburgefonstiv" across two lines mid-word, because the layout had been computed
    from a per-character guess 23.7% wider than Carlito.  The first baseline moved too
    (y=32 from the no-metrics default, against y=29.806 from Carlito's own descender).
    Nothing in the SVG says the widths were invented.
    """
    from deckbuilder import derive_deck

    deck = derive_deck(authoring, shapes_xml=CARLITO_PROBE_XML)
    (svg,) = convert_pptx_to_svg(deck, ConvertOptions(warn_on_font_substitution=False))

    elements = re.findall(r"<text[^>]*>.*?</text>", svg, re.S)
    (text,) = [element for element in elements if "Hamburge" in element]
    assert text.count("<tspan") == 1, text
    assert ">Hamburgefonstiv</tspan>" in text
    # ...and the chunk is sent to the family the widths came from.
    assert 'font-family="Carlito, sans-serif"' in text


#: Faces measured from one font and drawn with another, listed by name so that adding a
#: ninth is a deliberate act with a test to update rather than a quiet slide.
#:
#: Every one of them is proprietary with no metric-compatible clone, so there is nothing
#: we could ship that draws them correctly -- and measuring them with whatever we *can*
#: draw is worse than measuring them properly, because PowerPoint laid the deck out with
#: the real advances.  The four MS Japanese faces joined the list when ``sample.pptx``
#: showed what the alternative costs: MS PGothic is proportional (katakana 0.648-1.0 em)
#: and Noto Sans JP is not (uniformly 1.0), so its lines were measured up to a third too
#: wide and wrapped early.
DIVERGENT = {
    "Aptos", "Aptos Display", "Aptos Narrow", "Cambria",
    "MS Gothic", "MS ゴシック", "MS PGothic", "MS Pゴシック",
    "MS Mincho", "MS 明朝", "MS PMincho", "MS P明朝",
}

#: Tables with no font behind them.  See the note on :data:`DIVERGENT`.
MEASURED_ONLY = {
    "Aptos", "Aptos Display", "Cambria",
    "ＭＳ ゴシック", "ＭＳ Ｐゴシック", "ＭＳ 明朝", "ＭＳ Ｐ明朝",
}


def test_measuring_and_drawing_diverge_only_where_it_is_declared():
    divergent = {
        substitution.office
        for substitution in SUBSTITUTIONS.values()
        if substitution.metrics != substitution.substitute
    }
    assert divergent == DIVERGENT


def test_metrics_tables_exist_only_for_families_we_ship_or_deliberately_measure():
    assert set(METRICS) == set(BUNDLED_FAMILIES) | MEASURED_ONLY


def test_the_proportional_japanese_faces_are_measured_as_proportional():
    """The bug this catches: every kana measured at one em because the face is CJK.

    ``ＭＳ Ｐゴシック`` is proportional and ``ＭＳ ゴシック`` is not, from the same file --
    they are two faces inside ``msgothic.ttc``.  A generator that took face 0 for both
    would produce two identical tables and look perfectly healthy.
    """
    proportional = METRICS["ＭＳ Ｐゴシック"]
    monospaced = METRICS["ＭＳ ゴシック"]

    # A kanji is a full em in both; the katakana that gave the bug away are not.
    assert proportional.cjk_width == proportional.units_per_em
    assert proportional.widths["ト"] < 0.7 * proportional.units_per_em
    assert proportional.widths["、"] < 0.7 * proportional.units_per_em

    # The monospaced cut keeps every full-width kana at one em, so it stores no row for
    # them at all and falls through to cjk_width.
    assert "ト" not in monospaced.widths
    assert monospaced.cjk_width == monospaced.units_per_em


# --------------------------------------------------------------------------------------
# The bundle on disk
# --------------------------------------------------------------------------------------

def test_every_declared_family_resolves_to_a_readable_file():
    """A wheel that declares eight families and ships none is the failure mode."""
    directory = _bundle()
    files = list(directory.glob("*.ttf"))
    assert files, f"no font files in {directory}"
    stems = " ".join(path.stem for path in files).replace(" ", "")
    for family in BUNDLED_FAMILIES:
        # File names drop the spaces ("NotoSansJP"), so compare without them.
        assert family.replace(" ", "") in stems, family
    for path in files:
        assert path.stat().st_size > 1024, path


def test_every_font_ships_its_licence():
    """SIL OFL 1.1 requires the licence to travel with the font."""
    directory = _bundle()
    licenses = directory.parent / "licenses"
    assert licenses.is_dir()
    for family in BUNDLED_FAMILIES:
        path = licenses / f"{family.replace(' ', '')}-OFL.txt"
        assert path.is_file(), path
        assert "SIL OPEN FONT LICENSE Version 1.1" in path.read_text(encoding="utf-8")


def test_font_dirs_and_available_families_agree_with_the_mode():
    if bundle_dir() is None:
        assert bundle_mode() == "system"
        assert font_dirs() == []
        assert available_families() == frozenset()
    else:
        assert bundle_mode() == "bundled"
        assert font_dirs() == [str(bundle_dir())]
        assert available_families() == BUNDLED_FAMILIES


def test_generic_families_resolve_into_the_bundle():
    """With skip_system_fonts, an unresolvable family draws *nothing*.

    Every font-family this library emits ends in a CSS generic keyword, and resvg maps
    those to Arial / Times New Roman / Courier New by default -- none of which we ship.
    Pointing them at the bundle is what stops an unknown face becoming a blank paragraph.
    """
    for family in GENERIC_FAMILY_DEFAULTS.values():
        assert family in BUNDLED_FAMILIES


# --------------------------------------------------------------------------------------
# Generated metrics stay in step with the files
# --------------------------------------------------------------------------------------

def test_metrics_table_still_matches_the_bundled_fonts():
    """Re-derives the table and fails on drift.  This is the guard, not the docstring."""
    pytest.importorskip("fontTools", reason="fontTools is not installed")
    _bundle()
    result = subprocess.run(
        [sys.executable, str(ROOT / "tools" / "extract_font_metrics.py"), "--check"],
        capture_output=True,
        text=True,
    )
    assert result.returncode == 0, result.stdout + result.stderr


def test_bold_is_measured_from_the_bold_cut_not_a_multiplier():
    """A flat 1.05 was wrong in both directions; monospace bold is the clearest case."""
    measurer = DefaultTextMeasurer()
    text = "Handgloves"
    for family, expected in (("Courier New", 1.0), ("Calibri", 1.023)):
        regular = measurer.measure_text_width(text, 18, False, family)
        bold = measurer.measure_text_width(text, 18, True, family)
        assert bold / regular == pytest.approx(expected, abs=0.01), family


# --------------------------------------------------------------------------------------
# The diagnostic
# --------------------------------------------------------------------------------------

def test_office_faces_report_as_compatible():
    report = check_families(["Calibri", "Arial", "Times New Roman", "Courier New"])
    if report.mode != "bundled":
        pytest.skip("pptx2svg-fonts is not importable")
    assert report.faithful
    assert {face.verdict for face in report.faces} == {"compatible"}


def test_aptos_reports_as_approximate_and_says_why():
    report = check_families(["Aptos"])
    if report.mode != "bundled":
        pytest.skip("pptx2svg-fonts is not importable")
    (face,) = report.faces
    assert face.verdict == "approximate"
    assert face.substitute == "Carlito"
    assert face.metrics == "Aptos"
    assert "no metric-compatible clone" in face.reason
    assert not report.faithful


def test_an_unknown_face_is_reported_rather_than_guessed_at():
    report = check_families(["Wingdings 3"])
    (face,) = report.faces
    assert face.verdict == "missing"
    assert face.substitute is None
    assert not report.faithful


def test_theme_pointers_are_not_reported_as_fonts():
    """``+mn-cs`` survives resolution when a font scheme has no ``cs`` entry."""
    report = check_families(["Calibri", "+mn-cs", "+mj-lt"])
    assert [face.requested for face in report.faces] == ["Calibri"]


def test_without_the_bundle_nothing_is_faithful_or_reproducible(monkeypatch):
    monkeypatch.setattr("pptx2svg.fonts.bundle_dir", lambda: None)
    report = check_families(["Calibri"])
    assert report.mode == "system"
    assert not report.reproducible
    assert not report.faithful
    (face,) = report.faces
    assert "pptx2svg[fonts]" in face.reason


def test_faithful_and_reproducible_are_different_questions():
    """A deck can be drawn faithfully and still not reproducibly."""
    report = check_families(["Calibri"], system_fonts=True)
    if report.mode != "bundled":
        pytest.skip("pptx2svg-fonts is not importable")
    assert report.faithful          # Carlito has Calibri's widths either way
    assert not report.reproducible  # ... but the host's fonts are in play too


def test_checking_a_deck_finds_the_faces_it_really_uses(authoring):
    report = check_deck(authoring)
    assert {face.requested for face in report.faces} >= {"Aptos", "Aptos Display"}
    assert not report.faithful


# --------------------------------------------------------------------------------------
# Warnings
# --------------------------------------------------------------------------------------

def test_a_substituted_face_warns_through_the_normal_channel(authoring):
    options = ConvertOptions()
    convert_pptx_to_svg(authoring, options)
    codes = {warning.code for warning in options.warnings}
    assert codes & {"font-substituted", "font-bundle-missing"}
    assert any("Aptos" in warning.message for warning in options.warnings)


def test_a_missing_bundle_warns_once_and_names_the_fix(monkeypatch, authoring):
    """One warning, not one per face: without the bundle they all have the same cause."""
    monkeypatch.setattr("pptx2svg.fonts.bundle_dir", lambda: None)
    options = ConvertOptions()
    convert_pptx_to_svg(authoring, options)
    font_warnings = [w for w in options.warnings if w.code.startswith("font")]
    assert len(font_warnings) == 1
    assert font_warnings[0].code == "font-bundle-missing"
    assert "pptx2svg[fonts]" in font_warnings[0].message


def test_warnings_can_be_turned_off(authoring):
    options = ConvertOptions(warn_on_font_substitution=False)
    convert_pptx_to_svg(authoring, options)
    assert not [w for w in options.warnings if w.code.startswith("font")]


# --------------------------------------------------------------------------------------
# What the SVG asks for
# --------------------------------------------------------------------------------------

def test_the_font_stack_asks_for_the_original_face_before_the_substitute(authoring):
    """A host that really has Aptos should use it; the substitute is the fallback."""
    (svg,) = convert_pptx_to_svg(
        authoring, ConvertOptions(warn_on_font_substitution=False)
    )
    assert 'font-family="Aptos, Carlito, sans-serif"' in svg


def test_a_superfamily_face_names_its_typographic_family_too():
    """resvg indexes by OpenType name ID 16, which a deck never spells.

    ``Aptos Display.ttf`` carries ID 1 "Aptos Display" but ID 16 "Aptos", and fontdb
    files it under the latter only -- measured, not assumed: with just that file loaded,
    resvg draws nothing for ``font-family="Aptos Display"``.  Without "Aptos" in the
    stack the request fell past every named face to the generic and drew ~10% wide.
    """
    value = font_family_value(["Aptos Display"])
    names = value.split(", ")
    assert names[0] == "'Aptos Display'"
    assert names[1] == "Aptos"


def test_typographic_family_strips_one_style_word_only():
    assert typographic_family("Calibri Light") == "Calibri"
    assert typographic_family("Open Sans SemiBold") == "Open Sans"
    assert typographic_family("\u6e38\u30b4\u30b7\u30c3\u30af Light") == "\u6e38\u30b4\u30b7\u30c3\u30af"
    # Not a style word, and a real family in its own right: Aptos Narrow's own ID 16 is
    # "Aptos Narrow", so inventing a fallback to "Aptos" would be wrong.
    assert typographic_family("Aptos Narrow") is None
    assert typographic_family("Calibri") is None
    # A different glyph set rather than a weight -- see _TYPOGRAPHIC_STYLE_WORDS.
    assert typographic_family("Hoefler Text Ornaments") is None


def test_a_plain_text_box_falls_back_to_arial_not_the_theme(authoring):
    """PowerPoint does not consult the theme for a text box with nothing specified.

    Measured from PowerPoint's own export of this fixture: its text boxes are drawn in
    Arial (/BaseFont ArialMT, "MASTER CONTRACT" inked 178.28 pt against Arial's 180.00 pt
    advance at 18 pt) while its table cells in the same export are Aptos-Bold.  Getting
    this wrong is not cosmetic -- Arial is wider, so PowerPoint wraps "LAYOUT CONTRACT"
    onto two lines where the theme face fits it on one.
    """
    (svg,) = convert_pptx_to_svg(
        authoring, ConvertOptions(warn_on_font_substitution=False)
    )
    assert 'font-family="Arial, Arimo, sans-serif">MASTER CONTRACT<' in svg
    # ...and the wider face makes the layout's box wrap, as PowerPoint's does.
    assert ">LAYOUT</tspan>" in svg and ">CONTRACT</tspan>" in svg
    # The table is a table, not a text box: it keeps the theme face.
    assert "Aptos" in svg


def test_every_font_stack_ends_in_a_generic_family(pptx_path):
    generics = ("sans-serif", "serif", "monospace")
    documents = convert_pptx_to_svg(
        pptx_path, ConvertOptions(warn_on_font_substitution=False)
    )
    for document in documents:
        for value in re.findall(r'font-family="([^"]*)"', document):
            assert value.rsplit(", ", 1)[-1] in generics, value


def test_no_theme_pointer_reaches_a_font_family_stack(pptx_path):
    """``+`` cannot start a CSS identifier, and resvg drops the whole declaration.

    A theme that writes ``<a:cs typeface=""/>`` -- which every deck in the corpus does --
    used to resolve ``+mn-cs`` to itself, and the literal pointer was emitted in the
    stack.  The cost was not one dead entry: resvg rejected the entire ``font-family``
    and fell back to its default face, so *every* named face in the stack was lost.
    """
    for document in convert_pptx_to_svg(
        pptx_path, ConvertOptions(warn_on_font_substitution=False)
    ):
        for value in re.findall(r'font-family="([^"]*)"', document):
            assert "+" not in value, value


def test_synthetic_bold_widens_a_cjk_face_with_no_bold_cut():
    """PowerPoint emboldens such a face itself, and the advance grows with it.

    Not by a ratio -- by a constant.  Read off the pen origins in PowerPoint's PDF export
    of ``sample.pptx``, which sets the same Japanese text bold at three sizes: 24 pt goes
    24.000 -> 24.1248, 28 pt goes 28.000 -> 28.1260, 32 pt goes 32.000 -> 32.1248.  A
    1/256 em model would have predicted +0.094 at 24 pt.
    """
    from pptx2svg.text.measure import DefaultTextMeasurer
    from pptx2svg.units import PX_PER_PT

    measurer = DefaultTextMeasurer()
    text = "テンプレート"
    for size in (24, 28, 32):
        upright = measurer.measure_text_width(text, size, False, None, "ＭＳ Ｐゴシック")
        bold = measurer.measure_text_width(text, size, True, None, "ＭＳ Ｐゴシック")
        per_glyph = (bold - upright) / PX_PER_PT / len(text)
        assert per_glyph == pytest.approx(0.125, abs=0.002), size

    # A face that does ship a bold cut is measured from it, not nudged.
    upright = measurer.measure_text_width(text, 32, False, None, "Noto Sans JP")
    bold = measurer.measure_text_width(text, 32, True, None, "Noto Sans JP")
    assert bold == pytest.approx(upright, abs=0.01)


# -- The East Asian cascade --------------------------------------------------------------
#
# Read out of PowerPoint's own PDF export of `real-financial-report.pptx`, which is the
# one place in the corpus where every competing explanation is ruled out: its charts name
# `<a:latin typeface="Arial"/>` and no `<a:ea>` at all, its theme writes
# `<a:ea typeface=""/>` in both font collections, and its `<a:font script="Jpan"/>` says
# 游ゴシック.  The export embeds YuGothic-Regular for the Japanese and ArialMT for the
# Latin, inside the same labels.


def test_the_east_asian_cascade_takes_the_script_face_over_the_latin_one():
    """An empty `<a:ea>` falls through to the theme's script list, not to `<a:latin>`."""
    from pptx2svg.text.fontmap import east_asian_family

    assert east_asian_family("", "游ゴシック", "Arial") == "游ゴシック"
    assert east_asian_family(None, "游ゴシック", "Arial") == "游ゴシック"
    # A run that names one of its own keeps it.
    assert east_asian_family("ＭＳ Ｐ明朝", "游ゴシック", "Arial") == "ＭＳ Ｐ明朝"
    # An unexpanded theme pointer is not a face name; `+` cannot start a CSS identifier
    # either, so letting one through costs the whole `font-family` declaration.
    assert east_asian_family("+mn-ea", None, None) is None
    # Nothing East Asian anywhere: the first real name, so the stack is never empty.
    assert east_asian_family(None, None, "Arial") == "Arial"
    assert east_asian_family(None, None, None) is None


def test_a_latin_face_named_as_the_east_asian_one_loses_to_the_script_face():
    """`real-basic-theme.pptx` writes `<a:ea typeface="Raleway"/>` on nine of its runs.

    PowerPoint drew their Japanese in MS Gothic and MS Mincho -- not in Raleway, which has
    no kana at all.  Taking the name at face value measured every one of those glyphs at
    Raleway's `cjk_width`, and that number is 1.0 em only because
    `tools/extract_font_metrics.py` writes `units_per_em` when the probe kanji is missing
    from the face.  It is the absence of a measurement, dressed as one.
    """
    from pptx2svg.text.fontmap import covers_east_asian, east_asian_family

    assert not covers_east_asian("Raleway")
    assert not covers_east_asian("Arial")
    assert covers_east_asian("游ゴシック")
    assert covers_east_asian("ＭＳ ゴシック")
    assert east_asian_family("Raleway", "ＭＳ Ｐゴシック", "Raleway") == "ＭＳ Ｐゴシック"


def test_covers_east_asian_cannot_be_read_off_the_metric_tables():
    """The reason the flag lives on `Substitution` rather than being derived.

    `cjk_width` is 1.0 em in every table, Latin ones included, and a per-character row is
    written only where the advance *disagrees* with it -- so the two monospaced MS faces
    carry no East Asian rows while Cambria carries four, its bracket forms.  Counting rows
    would call Cambria Japanese and ＭＳ ゴシック not.
    """
    from pptx2svg.text.fontmap import covers_east_asian
    from pptx2svg.text.measure import is_cjk

    for table in METRICS.values():
        assert table.cjk_width == table.units_per_em

    rows = {
        name: sum(1 for char in table.widths if is_cjk(ord(char)))
        for name, table in METRICS.items()
    }
    assert rows["Cambria"] > 0 and not covers_east_asian("Cambria")
    assert rows["ＭＳ ゴシック"] == 0 and covers_east_asian("ＭＳ ゴシック")


def test_a_mixed_run_is_measured_face_by_face():
    """`DX投資額` is two faces in one label, and PowerPoint drew it that way.

    Its export puts `DX` in ArialMT and `投資額` in YuGothic-Regular: D advances 8.665 pt
    and X 8.004 pt at 12 pt, then three ideographs at a full em each.  52.669 pt in total,
    which is what the measurer has to return for a *single* string carrying both faces.
    """
    from pptx2svg.units import PX_PER_PT

    measurer = DefaultTextMeasurer()
    width = measurer.measure_text_width("DX投資額", 12, False, "Arial", "游ゴシック")
    assert width / PX_PER_PT == pytest.approx(52.669, abs=0.01)
    # CO2削減 on the same axis: 8.666 + 9.334 + 6.674 of Arial, then two ideographs.
    width = measurer.measure_text_width("CO2削減", 12, False, "Arial", "游ゴシック")
    assert width / PX_PER_PT == pytest.approx(48.674, abs=0.01)
