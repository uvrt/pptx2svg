"""A symbol face the conversion cannot draw is drawn as the Unicode it stands for
(``pptx2svg/symbols.py``), and a face it can draw is drawn as itself.

The deck is ``tools/make_symbol_probe.py``'s: the bullets and marks of the Office bullet
libraries in Wingdings, Wingdings 2/3, Webdings and Symbol, as runs and as bullets, spelled
as the code and as its private-use alias.  What each must become is what PowerPoint 16
drew for that deck (see ``ooxml-common``'s ``tests/test_symbol_fonts.py``).
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_svg

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import make_symbol_probe as probe  # noqa: E402

#: Wingdings' bullet-library codes and the characters drawn for them.
EXPECTED = {
    "q": "❑", "l": "●", "n": "■", "§": "▪", "Ø": "➢",
    "ü": "✓", "þ": "☑", "v": "❖",
}


@pytest.fixture(scope="module")
def deck() -> bytes:
    return probe.build_deck(probe.slides())


def _render(deck: bytes, **options) -> tuple[list[str], list]:
    convert = ConvertOptions(**options)
    return convert_pptx_to_svg(deck, convert), convert.warnings


def test_an_absent_symbol_face_is_drawn_as_the_unicode_it_stands_for(deck):
    svgs, warnings = _render(deck, host_fonts=False)
    assert len(svgs) == 4
    for svg in svgs:  # runs and bullets, as codes and as private-use aliases
        for drawn in EXPECTED.values():
            assert drawn in svg
        assert "•" in svg  # Symbol's 0xB7
        assert 'font-family="Wingdings' not in svg
    mapped = [w for w in warnings if w.code == "symbol-font-mapped"]
    assert {w.message.split(" is not available")[0] for w in mapped} == {
        "Wingdings", "Wingdings 2", "Wingdings 3", "Webdings", "Symbol",
    }
    assert "'q'->'❑'" in next(w.message for w in mapped if w.message.startswith("Wingdings "))


def test_a_symbol_face_this_host_has_is_drawn_as_itself(deck, monkeypatch):
    from pptx2svg.fonts import office

    real_find = office.find
    monkeypatch.setattr(
        office, "find", lambda family, *a, **k: ("face",) if family in {
            "Wingdings", "Wingdings 2", "Wingdings 3", "Webdings", "Symbol"
        } else real_find(family, *a, **k),
    )
    svgs, warnings = _render(deck, host_fonts=True)
    assert not [w for w in warnings if w.code == "symbol-font-mapped"]
    assert "❑" not in svgs[0]
    assert 'font-family="Wingdings' in svgs[0]


def test_a_symbol_face_in_the_applications_font_folder_is_drawn_as_itself(deck, monkeypatch):
    import ooxml_common.fonts.office as common_office

    monkeypatch.setattr(common_office, "user_families", lambda dirs=None: frozenset({"wingdings"}))
    svgs, warnings = _render(deck, host_fonts=False, font_dirs=["/nonexistent"])
    families = {w.message.split(" is not available")[0] for w in warnings
                if w.code == "symbol-font-mapped"}
    assert "Wingdings" not in families and "Symbol" in families
