"""Slides rasterised in threads where the GIL is off (free threading, step 2).

:func:`pptx2svg.rasterise_svgs` and :func:`pptx2svg.convert_pptx_to_png` draw several
slides at once in threads of this process when the interpreter has the GIL off
(``python3.15t``); with it on -- every regular build -- one at a time, as before, since resvg
holds the GIL while it draws.  The images, and the warnings and their order, are the same
either way.  Here the GIL's state is patched where a test needs threads on a regular build:
they take turns there, but draw the same."""

from __future__ import annotations

import gc
import hashlib
import threading
import warnings
from pathlib import Path

import pytest

import pptx2svg
from pptx2svg import ConvertOptions, MissingGlyphsWarning, convert_pptx_to_png
from pptx2svg.fonts import bundle_dir

FIXTURES = Path(__file__).parent / "fixtures"
BUNDLE = bundle_dir()

needs_resvg = pytest.mark.skipif("resvg" not in pptx2svg.available_backends(),
                                 reason="needs resvg-py")
needs_bundle = pytest.mark.skipif(BUNDLE is None, reason="needs pptx2svg-fonts' files")


def test_with_the_gil_on_slides_are_drawn_one_at_a_time(monkeypatch):
    monkeypatch.setattr(pptx2svg, "_gil_enabled", lambda: True)
    assert pptx2svg._raster_threads(None, 8) == 1
    assert pptx2svg._raster_threads(4, 8) == 1


@needs_resvg
def test_with_the_gil_off_as_many_threads_as_cpus_at_most_one_a_slide(monkeypatch):
    monkeypatch.setattr(pptx2svg, "_gil_enabled", lambda: False)
    monkeypatch.setattr("os.process_cpu_count", lambda: 4, raising=False)
    monkeypatch.setattr("os.cpu_count", lambda: 4)
    assert pptx2svg._raster_threads(None, 8) == 4
    assert pptx2svg._raster_threads(None, 3) == 3
    assert pptx2svg._raster_threads(2, 8) == 2
    assert pptx2svg._raster_threads(16, 8) == 8
    for one in [(None, 1), (None, 0), (1, 8), (0, 8)]:
        assert pptx2svg._raster_threads(*one) == 1, one


def test_cairosvg_draws_one_at_a_time(monkeypatch):
    monkeypatch.setattr(pptx2svg, "_gil_enabled", lambda: False)
    assert pptx2svg._raster_threads(None, 8, "cairosvg") == 1
    assert pptx2svg._raster_threads(None, 8, "no such backend") == 1


def _drawn(monkeypatch, threads):
    """sample-cjk's six slides, with Carlito and no CJK face: one warning, on slide 1."""
    carlito = [str(BUNDLE / name) for name in ("Carlito-Regular.ttf", "Carlito-Bold.ttf")]
    monkeypatch.setattr("pptx2svg.fonts.bundle_dir", lambda: None)
    options = ConvertOptions(host_fonts=False)
    gc.collect()
    with warnings.catch_warnings(record=True) as caught:
        warnings.simplefilter("always", MissingGlyphsWarning)
        images = convert_pptx_to_png(FIXTURES / "sample-cjk.pptx", options,
                                     skip_system_fonts=True, font_files=carlito,
                                     threads=threads)
    return ([hashlib.sha256(image).hexdigest() for image in images],
            [(w.code, w.slide_number, w.part_path, w.message) for w in options.warnings],
            [str(w.message) for w in caught if issubclass(w.category, MissingGlyphsWarning)])


@needs_bundle
@needs_resvg
def test_threads_draw_the_same_images_and_say_the_same_in_the_same_order(monkeypatch):
    names = set()
    draw = pptx2svg.png._svg_to_png

    def recording(*args, **kwargs):
        names.add(threading.current_thread().name)
        return draw(*args, **kwargs)

    monkeypatch.setattr("pptx2svg.png._svg_to_png", recording)
    monkeypatch.setattr(pptx2svg, "_gil_enabled", lambda: True)
    alone = _drawn(monkeypatch, None)
    assert names == {threading.current_thread().name}
    names.clear()
    monkeypatch.setattr(pptx2svg, "_gil_enabled", lambda: False)
    side_by_side = _drawn(monkeypatch, 3)
    assert len(names) == 3 and all(name.startswith("pptx2svg-raster") for name in names)
    assert side_by_side == alone
    images, said, warned = alone
    missing = [entry for entry in said if entry[0] == "glyphs-missing"]
    assert len(images) == 6 and [entry[1] for entry in missing] == [1] and len(warned) == 1


@needs_resvg
def test_rasterise_svgs_in_threads_keeps_the_order(monkeypatch):
    svgs = [('<svg xmlns="http://www.w3.org/2000/svg" width="40" height="20">'
             f'<rect width="40" height="20" fill="#{shade:02x}3366"/></svg>')
            for shade in range(0, 250, 25)]
    one = pptx2svg.rasterise_svgs(svgs, threads=1)
    monkeypatch.setattr(pptx2svg, "_gil_enabled", lambda: False)
    assert pptx2svg.rasterise_svgs(svgs, threads=4) == one
    assert len(set(one)) == len(svgs)
