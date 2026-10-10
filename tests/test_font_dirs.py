"""The application's own font folders: ``font_dirs``, or ``OOXML_FONT_DIRS``.

Production: an application keeps its licensed faces in a folder of its own, not one the
system searches.  Named with ``font_dirs`` (or the environment variable), a face there is
measured from and drawn with; not named, it is missing.  The face here is an open one --
the bundle's Cousine, relabelled in a temporary folder as a family nothing else answers
to, so no system copy can stand in for it -- and the host's faces are left out
(``host_fonts=False``), so the test means the same on every machine.
"""

from __future__ import annotations

import os
import zipfile
import io
from pathlib import Path

import pytest

import pptx2svg
from pptx2svg import ConvertOptions, convert_pptx_to_png, convert_pptx_to_svg, svg_to_png
from pptx2svg.fonts import bundle_dir

from ooxml_common.fonts.office import FONT_DIRS_ENV
from ooxml_common.fonts.sfnt import relabel

FIXTURES = Path(__file__).parent / "fixtures"
BUNDLE = bundle_dir()
FAMILY = "Fontdirs Probe Mono"

needs_bundle = pytest.mark.skipif(BUNDLE is None, reason="needs pptx2svg-fonts' files")
needs_resvg = pytest.mark.skipif("resvg" not in pptx2svg.available_backends(),
                                 reason="needs resvg-py")


@pytest.fixture(autouse=True)
def _no_environment(monkeypatch):
    monkeypatch.delenv(FONT_DIRS_ENV, raising=False)


@pytest.fixture
def folder(tmp_path) -> Path:
    """The application's folder: one face, in a subfolder as a licensed set often is."""
    target = tmp_path / "app-fonts" / "probe"
    target.mkdir(parents=True)
    data = relabel((BUNDLE / "Cousine-Regular.ttf").read_bytes(), FAMILY, bold=False, italic=False)
    (target / "FontdirsProbeMono-Regular.ttf").write_bytes(data)
    return tmp_path / "app-fonts"


def _deck() -> bytes:
    """authoring-integration.pptx with its theme's Latin faces set to the probe's family."""
    source = zipfile.ZipFile(FIXTURES / "authoring-integration.pptx")
    out = io.BytesIO()
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "ppt/theme/theme1.xml":
                text = data.decode("utf-8")
                for name in ("Aptos Display", "Aptos"):
                    text = text.replace(f'<a:latin typeface="{name}"', f'<a:latin typeface="{FAMILY}"')
                data = text.encode("utf-8")
            target.writestr(item, data)
    return out.getvalue()


def _named(options: ConvertOptions) -> list[str]:
    return [w.message for w in options.warnings if FAMILY in w.message]


@needs_bundle
def test_the_folder_is_measured_from_only_when_named(folder, monkeypatch):
    deck = _deck()
    plain = ConvertOptions(slide_numbers=[1], host_fonts=False)
    unconfigured = convert_pptx_to_svg(deck, plain)
    assert _named(plain), "an unnamed folder's face is reported as not drawn faithfully"

    given = ConvertOptions(slide_numbers=[1], host_fonts=False, font_dirs=[str(folder)])
    configured = convert_pptx_to_svg(deck, given)
    assert not _named(given)
    assert configured != unconfigured  # measured as the monospaced face, not guessed

    # The environment variable: the same, when nothing is given explicitly.
    monkeypatch.setenv(FONT_DIRS_ENV, os.pathsep.join([str(folder / "absent"), str(folder)]))
    from_env = ConvertOptions(slide_numbers=[1], host_fonts=False)
    assert convert_pptx_to_svg(deck, from_env) == configured and not _named(from_env)

    # An explicit empty list wins over the variable: none.
    none = ConvertOptions(slide_numbers=[1], host_fonts=False, font_dirs=[])
    assert convert_pptx_to_svg(deck, none) == unconfigured and _named(none)


@needs_bundle
@needs_resvg
def test_the_folder_is_drawn_with_and_its_argument_wins(folder, monkeypatch):
    deck = _deck()
    unconfigured = convert_pptx_to_png(deck, ConvertOptions(slide_numbers=[1], host_fonts=False))
    configured = convert_pptx_to_png(deck, ConvertOptions(slide_numbers=[1], host_fonts=False),
                                     font_dirs=[str(folder)])
    assert configured != unconfigured
    assert convert_pptx_to_png(
        deck, ConvertOptions(slide_numbers=[1], host_fonts=False, font_dirs=[str(folder)])) == configured

    monkeypatch.setenv(FONT_DIRS_ENV, str(folder))
    assert convert_pptx_to_png(deck, ConvertOptions(slide_numbers=[1], host_fonts=False)) == configured
    # convert_pptx_to_png's own argument takes precedence over the options and the variable.
    assert convert_pptx_to_png(deck, ConvertOptions(slide_numbers=[1], host_fonts=False, font_dirs=[str(folder)]),
                               font_dirs=[]) == unconfigured


@needs_bundle
@needs_resvg
def test_svg_to_png_reads_the_variable(folder, monkeypatch):
    svg = ('<svg xmlns="http://www.w3.org/2000/svg" width="300" height="60">'
           '<rect width="300" height="60" fill="#fff"/>'
           f'<text x="5" y="40" font-size="30" font-family="{FAMILY}, sans-serif">iiiiii</text></svg>')
    unconfigured = svg_to_png(svg)
    assert svg_to_png(svg, font_dirs=[str(folder)]) != unconfigured
    monkeypatch.setenv(FONT_DIRS_ENV, str(folder))
    assert svg_to_png(svg) == svg_to_png(svg, font_dirs=[str(folder)])
    assert svg_to_png(svg, font_dirs=[]) == unconfigured


def test_the_old_variable_name_is_not_read(monkeypatch, tmp_path):
    """``PPTX_FONT_DIR`` was never read by pptx2svg; ``OOXML_FONT_DIRS`` is the one."""
    monkeypatch.setenv("PPTX_FONT_DIR", str(tmp_path))
    assert pptx2svg.user_font_dirs(ConvertOptions()) == []
    monkeypatch.setenv(FONT_DIRS_ENV, str(tmp_path))
    assert pptx2svg.user_font_dirs(ConvertOptions()) == [str(tmp_path)]
    x, y = str(tmp_path / "x"), str(tmp_path / "y")
    assert pptx2svg.user_font_dirs(ConvertOptions(font_dirs=[x])) == [x]
    assert pptx2svg.user_font_dirs(ConvertOptions(font_dirs=[x]), [y]) == [y]
