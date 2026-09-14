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
from pptx2svg.text.fontmap import SUBSTITUTIONS
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
