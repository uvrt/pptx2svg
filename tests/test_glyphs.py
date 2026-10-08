"""Text the rasteriser will not draw is said out loud (pptx2svg.glyphs).

Production: rendered without ``pptx2svg-fonts``, a deck's Japanese text vanished from its
PNGs and nothing said so.  The absence is simulated here -- the bundle's ``bundle_dir``
patched away, the host's fonts skipped -- and the faces the rasteriser does get are chosen
from the bundle's files, so the test means the same on every machine.
"""

from __future__ import annotations

import io
import warnings
from pathlib import Path

import pytest

import pptx2svg
from pptx2svg import ConvertOptions, MissingGlyphsWarning, convert_pptx_to_png, svg_to_png
from pptx2svg.fonts import GENERIC_FAMILY_DEFAULTS, bundle_dir
from pptx2svg.glyphs import MissingGlyphs, missing_glyphs, script_of, text_by_stack

FIXTURES = Path(__file__).parent / "fixtures"
BUNDLE = bundle_dir()

needs_bundle = pytest.mark.skipif(BUNDLE is None, reason="needs pptx2svg-fonts' files")
needs_resvg = pytest.mark.skipif("resvg" not in pptx2svg.available_backends(),
                                 reason="needs resvg-py")


def _svg(*runs: tuple[str, str]) -> str:
    spans = "".join(f'<tspan font-family="{family}">{text}</tspan>' for family, text in runs)
    return ('<svg xmlns="http://www.w3.org/2000/svg" width="400" height="60">'
            '<rect width="400" height="60" fill="#fff"/>'
            f'<text x="5" y="40" font-size="30">{spans}</text></svg>')


def _ink(png: bytes) -> int:
    from PIL import Image

    image = Image.open(io.BytesIO(png)).convert("L")
    return sum(1 for value in image.get_flattened_data() if value < 128) \
        if hasattr(image, "get_flattened_data") else \
        sum(1 for value in image.getdata() if value < 128)


@needs_bundle
@needs_resvg
def test_cjk_without_the_font_bundle_is_reported_once_not_dropped_silently(monkeypatch):
    """sample-cjk's Japanese, with Carlito for its Latin and no CJK face anywhere: one
    warning for the face and script, from the whole deck, in both channels."""
    carlito = [str(BUNDLE / name) for name in ("Carlito-Regular.ttf", "Carlito-Bold.ttf")]
    monkeypatch.setattr("pptx2svg.fonts.bundle_dir", lambda: None)
    options = ConvertOptions(host_fonts=False)
    with pytest.warns(MissingGlyphsWarning) as caught:
        images = convert_pptx_to_png(FIXTURES / "sample-cjk.pptx", options,
                                     skip_system_fonts=True, font_files=carlito)
    assert len(images) == 6
    found = [w for w in options.warnings if w.code == "glyphs-missing"]
    assert len(found) == 1 and found[0].slide_number == 1
    assert (found[0].detail.face, found[0].detail.script) == ("Noto Sans JP", "CJK")
    assert found[0].message.startswith("Noto Sans JP: no font this render loads has its CJK")
    assert "サンプル" in found[0].message and "pptx2svg-fonts" in found[0].message
    assert [str(w.message) for w in caught.list] == [found[0].message]


@needs_bundle
@needs_resvg
def test_with_the_bundle_the_same_deck_draws_and_says_nothing():
    options = ConvertOptions(host_fonts=False)
    with warnings.catch_warnings():
        warnings.simplefilter("error", MissingGlyphsWarning)
        convert_pptx_to_png(FIXTURES / "sample-cjk.pptx", options)
    assert not [w for w in options.warnings if w.code == "glyphs-missing"]


@needs_bundle
def test_hangul_the_bundle_cannot_draw_is_reported():
    """Noto Sans JP draws Japanese but no Hangul (FONTS.md), so Korean in a deck rendered
    with the bundle alone is lost -- and says so, once."""
    svg = _svg(("'Noto Sans JP', sans-serif", "日本語"), ("'Noto Sans JP', sans-serif", "한국어"),
               ("'Noto Sans JP', sans-serif", "회의"))
    found = missing_glyphs(svg, font_dirs=[str(BUNDLE)],
                           generic_families=GENERIC_FAMILY_DEFAULTS)
    assert found == [MissingGlyphs("Noto Sans JP", "Hangul", "한국어회의", False,
                                   "Noto Sans JP, sans-serif")]


@needs_bundle
@needs_resvg
def test_svg_to_png_warns_once_per_face_and_script(recwarn):
    arimo = str(BUNDLE / "Arimo[wght].ttf")
    svg = _svg(("Arimo", "Ab"), ("Arimo", "サンプル"), ("Arimo", "拡張"), ("Arimo", "한"))
    svg_to_png(svg, backend="resvg", font_files=[arimo], use_bundled_fonts=False,
               skip_system_fonts=True, host_fonts=False)
    messages = [str(w.message) for w in recwarn.list if w.category is MissingGlyphsWarning]
    assert len(messages) == 2
    assert messages[0].startswith("Arimo: no font this render loads has its CJK glyphs "
                                  "(サンプル拡張)")
    assert messages[1].startswith("Arimo: no font this render loads has its Hangul glyphs")
    svg_to_png(svg, backend="resvg", font_files=[arimo], use_bundled_fonts=False,
               skip_system_fonts=True, host_fonts=False, check_glyphs=False)
    assert len([w for w in recwarn.list if w.category is MissingGlyphsWarning]) == 2


@needs_bundle
@needs_resvg
def test_the_two_ways_resvg_leaves_text_out_are_the_two_the_check_reports():
    """Measured: a stack no loaded face answers to draws nothing at all, Latin included;
    a character the loaded faces lack is not drawn as itself (Arimo's empty box), and is
    drawn once a face that has it is loaded, whatever the stack names."""
    arimo, noto = str(BUNDLE / "Arimo[wght].ttf"), str(BUNDLE / "NotoSansJP[wght].ttf")

    def ink(svg, files, generic=None):
        from pptx2svg.png import _render_with_resvg

        return _ink(_render_with_resvg(svg, width=None, height=None, scale=None,
                                       background=None, font_dirs=[], font_files=files,
                                       skip_system_fonts=True, generic_families=generic))

    unloaded = _svg(("Calibri, sans-serif", "Abc"))
    assert ink(unloaded, [arimo]) == 0
    assert missing_glyphs(unloaded, font_files=[arimo])[0].unloaded
    assert ink(unloaded, [arimo], {"sans_serif_family": "Arimo"}) > 0
    assert missing_glyphs(unloaded, font_files=[arimo],
                          generic_families={"sans_serif_family": "Arimo"}) == []
    kana = _svg(("Arimo", "サ"))
    assert ink(kana, [arimo, noto]) > ink(kana, [arimo])
    assert missing_glyphs(kana, font_files=[arimo])[0].script == "CJK"
    assert missing_glyphs(kana, font_files=[arimo, noto]) == []


def test_text_by_stack_reads_every_run_once_each():
    svg = ('<svg xmlns="http://www.w3.org/2000/svg"><g font-family="A, serif">'
           '<text>ab <tspan font-family="B">cd</tspan> ba​</text></g>'
           '<text font-family="B">dc</text></svg>')
    assert text_by_stack(svg) == {"A, serif": "ab", "B": "cd"}


@pytest.mark.parametrize("char,script", [("漢", "CJK"), ("ひ", "CJK"), ("カ", "CJK"),
                                         ("ー", "CJK"), ("。", "CJK"), ("Ａ", "CJK"),
                                         ("한", "Hangul"), ("ㄱ", "Hangul"), ("a", "Latin"),
                                         ("7", "Latin"), ("Ж", "Cyrillic"), ("ب", "Arabic"),
                                         ("ก", "Thai"), ("★", "Symbols"), ("😀", "Symbols")])
def test_script_of(char, script):
    assert script_of(char) == script
