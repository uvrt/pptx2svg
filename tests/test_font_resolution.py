"""Which face PowerPoint resolves a run to, held to what it drew for
``tools/make_font_resolution_probe.py``.

The probe decks are built here exactly as the tool builds them, resolved by pptx2svg, and
each box's East Asian face held to the face PowerPoint's PDF drew its kana in
(``tests/fixtures/font-resolution-probe.json``, written by
``tools/read_font_resolution_probe.py --json``).  Two ways:

* **the reproducible render** (``host_fonts=False``), on every platform: every probe whose
  faces this library knows without reading the machine's fonts;
* **with the host's faces**, on a Mac with Office: every probe, each face found as
  PowerPoint finds it and compared by its PostScript name.
"""

from __future__ import annotations

import json
import sys
from pathlib import Path

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_model, model as m
from pptx2svg.text.measure import is_cjk

TOOLS = Path(__file__).resolve().parents[1] / "tools"
sys.path.insert(0, str(TOOLS))

import make_font_resolution_probe as probe  # noqa: E402

OBSERVED = json.loads((Path(__file__).parent / "fixtures" / "font-resolution-probe.json").read_text("utf-8"))

#: The PostScript name PowerPoint's PDF draws each of these names in.
POSTSCRIPT = {
    "MS Gothic": "MS-Gothic", "ＭＳ ゴシック": "MS-Gothic", "MS Mincho": "MS-Mincho", "ＭＳ 明朝": "MS-Mincho",
    "ＭＳ Ｐゴシック": "MS-PGothic", "MS PGothic": "MS-PGothic", "ＭＳ Ｐ明朝": "MS-PMincho",
    "MS PMincho": "MS-PMincho", "游ゴシック": "YuGothic-Regular", "Yu Gothic": "YuGothic-Regular",
    "游ゴシック Light": "YuGothic-Light", "游明朝": "YuMincho-Regular", "Yu Mincho": "YuMincho-Regular",
    "メイリオ": "Meiryo", "Meiryo": "Meiryo", "Noto Sans JP": "NotoSansJP-Thin_Regular",
    "Hiragino Sans": "HiraginoSans-W3", "Hiragino Mincho ProN": "HiraMinProN-W3",
}

#: Probes the reproducible render cannot get right, because the answer is a fact about
#: the machine: a Latin face this library has no PANOSE for (macOS's Helvetica Neue,
#: Optima, Menlo, Didot; Office's Rockwell, Comic Sans MS, Brush Script MT), or a
#: Japanese face it does not know draws Japanese.
NEEDS_THE_HOST = {
    "latin:Helvetica Neue", "latin:Optima", "latin:Menlo", "latin:Didot", "latin:Rockwell",
    "latin:Comic Sans MS", "latin:Brush Script MT", "ea:MS UI Gothic", "ea:Arial Unicode MS",
    "ea:Hiragino Kaku Gothic ProN", "ea:ヒラギノ明朝 ProN",
}

#: Measured and not reproduced.  PowerPoint does not find Hiragino Kaku Gothic ProN by its
#: typographic family, nor by its Japanese name ID 1, although the face's name records
#: have the same shape as Hiragino Sans's, which it finds both ways; we find it.  And
#: Hiragino Sans has ten weights, W0 to W9, none of them flagged regular in its tables;
#: PowerPoint draws W3, this lookup the first, W0.
UNEXPLAINED = {
    ("jpan-script", "ea:Hiragino Kaku Gothic ProN"), ("names", "ea:ヒラギノ角ゴ ProN W3"),
    ("jpan-script", "ea:Hiragino Sans"), ("names", "ea:Hiragino Sans"),
}

JAPANESE = [deck for deck in probe.JAPANESE_DECKS if deck not in ("names", "bold-sizes")]


def _runs(element):
    if isinstance(element, m.GroupElement):
        for child in element.children:
            yield from _runs(child)
    elif isinstance(element, m.TableElement):
        for row in element.table.rows:
            for cell in row.cells:
                for paragraph in cell.text_body.paragraphs:
                    yield from paragraph.runs
    elif getattr(element, "text_body", None) is not None:
        for paragraph in element.text_body.paragraphs:
            yield from paragraph.runs


def _resolved(deck: str, directory: Path, host: bool) -> dict[str, str | None]:
    """``probe key -> the East Asian face`` of its kana."""
    target = directory / f"frp-{deck}.pptx"
    probe.build(deck, target)
    resolved = convert_pptx_to_model(str(target), ConvertOptions(host_fonts=host))
    elements = [element for slide in resolved.slides for element in slide.elements]
    specs = probe.deck_probes(deck)
    assert len(elements) == len(specs)
    out = {}
    for element, spec in zip(elements, specs):
        faces = {run.properties.font_family_ea for run in _runs(element)
                 if any(is_cjk(ord(char)) for char in run.text)}
        if spec["key"].startswith("sym:"):
            continue
        assert len(faces) == 1, (spec["key"], faces)
        out[spec["key"]] = faces.pop()
    return out


def _drawn(deck: str, key: str) -> str:
    return OBSERVED[deck][key]["drawn"]["east-asian"][0]


@pytest.fixture(scope="module")
def reproducible(tmp_path_factory):
    directory = tmp_path_factory.mktemp("frp")
    return {deck: _resolved(deck, directory, host=False) for deck in JAPANESE}


@pytest.mark.parametrize("deck", JAPANESE)
def test_the_reproducible_render_draws_japanese_where_powerpoint_did(reproducible, deck):
    wrong = {
        key: (face, _drawn(deck, key))
        for key, face in reproducible[deck].items()
        if not key.startswith("sym:") and key not in NEEDS_THE_HOST
        and POSTSCRIPT.get(face, face) != _drawn(deck, key)
    }
    assert wrong == {}


def test_the_japanese_script_entry_is_the_japanese_runs_own(reproducible):
    """``lang="ja-JP"`` takes the theme's ``Jpan`` entry; ``en-US`` never does."""
    assert reproducible["jpan-yu"]["lang-ja:Calibri"] == "游ゴシック"
    assert reproducible["jpan-yu"]["latin:Calibri"] == "MS Gothic"
    assert reproducible["ea-yu"]["latin:Calibri"] == "游ゴシック"
    assert reproducible["ea-yu"]["lang-ja:Calibri"] == "ＭＳ Ｐゴシック"


def test_a_latin_face_decides_between_gothic_and_mincho(reproducible):
    faces = reproducible["jpan-script"]
    assert faces["latin:Calibri"] == faces["latin:Aptos"] == faces["latin:Lato"] == "MS Gothic"
    assert faces["latin:Times New Roman"] == faces["latin:Courier New"] == faces["latin:Raleway"] == "MS Mincho"
    assert faces["cross:Calibri/Courier New"] == "MS Mincho"
    assert faces["courier+ea:NoSuchJapaneseFace"] == "MS Gothic"
    assert faces["table:Courier New"] == "MS Mincho"


def _postscript(name: str) -> str | None:
    from ooxml_common.fonts.office import Face, font_offsets, read_file

    from pptx2svg.fonts import office

    faces = [face for face in office.find(name) if not face.bold and not face.italic] or list(office.find(name))
    if not faces:
        return None
    data = read_file(faces[0].path)
    names = dict(Face(data, font_offsets(data)[faces[0].number]).names)
    return names.get(6)


def _office_with_the_probed_faces() -> bool:
    from pptx2svg.fonts import office

    return office.available() and all(office.find(name) for name in ("MS Gothic", "MS Mincho", "MS PGothic"))


@pytest.mark.skipif(not _office_with_the_probed_faces(), reason="needs Office for Mac's own faces")
@pytest.mark.parametrize("deck", [*JAPANESE, "names"])
def test_with_the_host_faces_every_probe_draws_where_powerpoint_did(tmp_path, deck):
    wrong = {}
    for key, face in _resolved(deck, tmp_path, host=True).items():
        if key.startswith("sym:") or (deck, key) in UNEXPLAINED:
            continue
        drawn = _drawn(deck, key)
        ours = _postscript(face)
        if ours is None and drawn not in ("MS-Gothic", "MS-Mincho"):
            continue  # the face PowerPoint drew is not installed on this machine
        if ours != drawn.split("_")[0]:
            wrong[key] = (face, ours, drawn)
    assert wrong == {}


#: ``bold-sizes``: PowerPoint's advance per kana for MS Gothic and MS Mincho (no bold cut
#: in either), read off the export's pen origins: bold advances its point size up to
#: 15.5 pt, and 0.1242--0.1260 pt more from 16 pt.
BOLD = [(key, entry["advance_pt"]) for key, entry in OBSERVED["bold-sizes"].items()]


@pytest.mark.parametrize("key, advance", BOLD, ids=[key for key, _ in BOLD])
def test_synthetic_bold_widens_an_advance_from_16_pt_only(key, advance):
    face, size, style = key.split(":")
    size = int(size) / 100
    from pptx2svg.text.measure import DefaultTextMeasurer
    from pptx2svg.units import PX_PER_PT

    width = DefaultTextMeasurer().measure_text_width("かなカナかなカナかな", size, style == "bold", "Calibri", face)
    assert width / PX_PER_PT / 10 == pytest.approx(advance, abs=0.002)


# -- Embedded against installed ---------------------------------------------------------


def test_an_installed_family_is_not_taken_from_the_deck(basic_theme, monkeypatch):
    """PowerPoint drew ``real-basic-theme.pptx``'s Lato with the installed 2.015, not the
    1.104 the deck embeds -- also when the embedded copy claimed a later version."""
    for name in ("embed-older", "embed-newer"):
        errors = OBSERVED[name]["Lato:regular"]["max_error_pt"]
        installed = min(value for label, value in errors.items() if not label.startswith("embedded"))
        embedded = min(value for label, value in errors.items() if label.startswith("embedded"))
        assert installed < embedded / 4
    assert OBSERVED["embed-absent"]["Latoprobe:regular"]["max_error_pt"]  # drawn from the deck

    from ooxml_common.fonts import bundle_dir

    from pptx2svg.fonts import office

    if bundle_dir() is None:
        pytest.skip("the pptx2svg-fonts bundle stands in for an installed Lato")
    lato = tuple(office._faces_in(Path(bundle_dir()) / "Lato-Regular.ttf", "cloud"))
    monkeypatch.setattr(office, "find", lambda family: lato if family == "Lato" else ())
    host = convert_pptx_to_model(basic_theme, ConvertOptions(host_fonts=True)).embedded_fonts
    assert {face.family for face in host.faces} == {"Raleway"}
    reproducible = convert_pptx_to_model(basic_theme, ConvertOptions(host_fonts=False)).embedded_fonts
    assert {face.family for face in reproducible.faces} == {"Lato", "Raleway"}


# -- Emoji ---------------------------------------------------------------------------------


def _emoji():
    from pptx2svg.fonts import office

    return office.emoji_face()


@pytest.mark.skipif(_emoji() is None, reason="Apple Color Emoji is not installed")
def test_an_emoji_the_face_lacks_is_named_in_apple_color_emoji():
    from pptx2svg.fonts import office

    for key in ("Calibri", "Aptos", "Noto Sans JP", "Lato"):
        drawn = OBSERVED["jpan-script"][f"sym:{key}:⚡📱🔒"]["drawn"]
        assert drawn == {"emoji": ["AppleColorEmoji"]}
    svg = ('<svg><text><tspan font-family="\'Noto Sans JP\', sans-serif">A⚡B↔</tspan></text></svg>')
    out, files = office.with_emoji(svg)
    assert files == [str(office.APPLE_COLOR_EMOJI)]
    # ⚡ is not Noto Sans JP's, ↔ is: only the first changes face.
    assert ">A<tspan font-family=\"'Apple Color Emoji'\">⚡</tspan>B↔</tspan>" in out
    assert office.with_emoji(svg.replace("⚡", "")) == (svg.replace("⚡", ""), [])


def test_no_emoji_face_without_the_host(monkeypatch):
    from pptx2svg.fonts import office

    monkeypatch.setenv("PPTX2SVG_OFFICE_FONTS", "0")
    svg = '<svg><text><tspan font-family="Lato">⚡</tspan></text></svg>'
    assert office.with_emoji(svg) == (svg, [])


def test_the_bundled_panose_table_is_the_bundles():
    from ooxml_common.fonts import bundle_dir

    from pptx2svg.fonts import office
    from pptx2svg.resolve.east_asian import _BUNDLED_PANOSE

    if bundle_dir() is None:
        pytest.skip("needs the pptx2svg-fonts bundle")
    for path in sorted(Path(bundle_dir()).iterdir()):
        for face in office._faces_in(path, "bundle"):
            if face.bold or face.italic:
                continue
            for key in face.families | face.typographic:
                if key in _BUNDLED_PANOSE:
                    assert face.panose[:4] == _BUNDLED_PANOSE[key], key
