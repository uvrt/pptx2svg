"""Command-line interface."""

from __future__ import annotations

import pytest

from pptx2svg.cli import main, parse_slide_selection
from pptx2svg.png import available_backends


@pytest.mark.parametrize(
    "value,expected",
    [
        (None, None),
        ("", None),
        ("1", [1]),
        ("1,3", [1, 3]),
        ("2-5", [2, 3, 4, 5]),
        ("1,3-4, 7", [1, 3, 4, 7]),
    ],
)
def test_slide_selection_parsing(value, expected):
    assert parse_slide_selection(value) == expected


def test_writes_one_svg_per_slide(tmp_path, product_page):
    assert main([str(product_page), "-o", str(tmp_path)]) == 0
    assert (tmp_path / "real-product-page-1.svg").is_file()


def test_slide_selection_limits_output(tmp_path, pptx_path):
    assert main([str(pptx_path), "-o", str(tmp_path), "-s", "1"]) == 0
    assert len(list(tmp_path.glob("*.svg"))) == 1


def test_missing_input_is_an_error(tmp_path, capsys):
    assert main([str(tmp_path / "nope.pptx"), "-o", str(tmp_path)]) == 1
    assert "no such file" in capsys.readouterr().err


@pytest.mark.skipif(not available_backends(), reason="no rasterizer installed")
def test_both_formats_writes_svg_and_png(tmp_path, product_page):
    assert main([str(product_page), "-o", str(tmp_path), "-f", "both"]) == 0
    assert (tmp_path / "real-product-page-1.svg").is_file()
    assert (tmp_path / "real-product-page-1.png").is_file()


# --------------------------------------------------------------------------------------
# `pptx2svg fonts`
# --------------------------------------------------------------------------------------


def test_fonts_reports_on_the_bundle_with_no_deck(capsys):
    assert main(["fonts"]) == 0
    out = capsys.readouterr().out
    assert "mode:" in out
    assert "Calibri" in out


def test_fonts_check_fails_a_deck_that_cannot_be_drawn_faithfully(capsys, authoring):
    """Aptos has no metric-compatible clone, so this deck can never pass --check."""
    assert main(["fonts", "--check", str(authoring)]) == 1
    out = capsys.readouterr().out
    assert "Aptos" in out
    assert "approximate" in out


def test_fonts_without_check_reports_but_does_not_fail(capsys, authoring):
    """The report is useful on its own; only --check turns it into a gate."""
    assert main(["fonts", str(authoring)]) == 0
    assert "Aptos" in capsys.readouterr().out


def test_fonts_rejects_a_missing_deck(capsys, tmp_path):
    assert main(["fonts", str(tmp_path / "nope.pptx")]) == 1
    assert "no such file" in capsys.readouterr().err


def test_a_file_named_fonts_is_still_reachable(tmp_path, product_page, capsys):
    """`fonts` diverts only as a bare first word, so ./fonts still renders."""
    deck = tmp_path / "fonts"
    deck.write_bytes(product_page.read_bytes())
    assert main([str(deck), "-o", str(tmp_path)]) == 0
    assert (tmp_path / "fonts-1.svg").is_file()
