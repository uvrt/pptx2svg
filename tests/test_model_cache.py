"""Resolved slides kept across conversions (``pptx2svg.cache.ModelCache``), parts handed
over in memory (``open_parts``) and only the asked-for slides read: the same output as a
conversion that keeps nothing, with less work.

The decks are the fixtures' own, edited in memory part by part, as an editor holding a deck
would hand them over between its edits.
"""

from __future__ import annotations

import io
import threading
import zipfile
from datetime import datetime
from pathlib import Path

import pytest

from pptx2svg import (
    ConvertOptions,
    ModelCache,
    OpcPackage,
    convert_pptx_to_model,
    convert_pptx_to_svg,
    open_parts,
)
from pptx2svg.parse.parts import read_presentation

FIXTURES = Path(__file__).resolve().parent / "fixtures"
NOW = datetime(2026, 11, 7, 9, 30)

#: A text box holding a date field, appended to a slide's shape tree.
DATE_BOX = (
    '<p:sp><p:nvSpPr><p:cNvPr id="9001" name="Stamp"/><p:cNvSpPr txBox="1"/><p:nvPr/>'
    '</p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="457200"/><a:ext cx="3657600" '
    'cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="{0F5F4C2B-3D1E-4A55-9C34-1A2B3C4D5E6F}" '
    'type="datetime11"><a:rPr lang="en-US"/><a:t>00:00:00</a:t></a:fld></a:p></p:txBody></p:sp>'
)

#: A plain text box, appended to a slide's or layout's shape tree.
TEXT_BOX = (
    '<p:sp><p:nvSpPr><p:cNvPr id="9002" name="Note"/><p:cNvSpPr txBox="1"/><p:nvPr/>'
    '</p:nvSpPr><p:spPr><a:xfrm><a:off x="457200" y="1371600"/><a:ext cx="3657600" '
    'cy="457200"/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>'
    '<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:r><a:rPr lang="en-US"/><a:t>An added note</a:t>'
    '</a:r></a:p></p:txBody></p:sp>'
)


def parts_of(name: str) -> dict[str, bytes]:
    with zipfile.ZipFile(FIXTURES / name) as archive:
        return {info.filename: archive.read(info) for info in archive.infolist()
                if not info.is_dir()}


def zipped(parts: dict[str, bytes]) -> bytes:
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w", zipfile.ZIP_DEFLATED) as archive:
        for name, data in parts.items():
            archive.writestr(name, data)
    return buffer.getvalue()


def with_shape(parts: dict[str, bytes], part: str, shape: str) -> dict[str, bytes]:
    changed = dict(parts)
    changed[part] = parts[part].replace(b"</p:spTree>", shape.encode() + b"</p:spTree>")
    return changed


def drawn(source, **options) -> tuple[list[str], list[str]]:
    """The SVGs and the warnings, as text."""
    settings = ConvertOptions(now=NOW, **options)
    svgs = convert_pptx_to_svg(source, settings)
    return svgs, [str(warning) for warning in settings.warnings]


@pytest.mark.parametrize("fixture", ["sample.pptx", "feature-sweep.pptx", "chart-gallery.pptx"])
def test_a_kept_slide_draws_as_a_slide_resolved_anew(fixture):
    data = (FIXTURES / fixture).read_bytes()
    cache = ModelCache()
    plain = drawn(data)
    first = drawn(data, cache=cache)
    misses = cache.misses
    second = drawn(data, cache=cache)
    assert plain == first == second
    assert cache.misses == misses and cache.hits == len(plain[0])
    # One slide at a time, from what the whole deck kept.
    for number in range(1, len(plain[0]) + 1):
        assert drawn(data, cache=cache, slide_numbers=[number])[0] == [plain[0][number - 1]]
    assert cache.misses == misses


def test_an_edited_slide_is_the_only_one_resolved_again():
    parts = parts_of("sample.pptx")
    cache = ModelCache()
    drawn(open_parts(parts), cache=cache)
    resolved = cache.misses
    edited = with_shape(parts, "ppt/slides/slide3.xml", TEXT_BOX)
    after = drawn(open_parts(edited), cache=cache)
    assert cache.misses == resolved + 1
    assert after == drawn(zipped(edited))
    assert "An added note" in after[0][2]


def test_a_changed_layout_resolves_the_slides_that_reach_it():
    parts = parts_of("sample.pptx")
    cache = ModelCache()
    count = len(drawn(open_parts(parts), cache=cache)[0])
    resolved = cache.misses
    edited = with_shape(parts, "ppt/slideLayouts/slideLayout2.xml", TEXT_BOX)
    after = drawn(open_parts(edited), cache=cache)
    # The master lists every layout, so each slide reaches each layout: all resolved again.
    assert cache.misses == resolved + count
    assert after == drawn(zipped(edited))


def test_a_slide_showing_the_clock_is_resolved_every_time():
    parts = with_shape(parts_of("sample.pptx"), "ppt/slides/slide1.xml", DATE_BOX)
    cache = ModelCache()
    morning = convert_pptx_to_svg(open_parts(parts), ConvertOptions(now=NOW, cache=cache))
    later = convert_pptx_to_svg(open_parts(parts),
                                ConvertOptions(now=datetime(2026, 11, 7, 17, 45, 5), cache=cache))
    assert ">9:30:00<" in morning[0] and ">17:45:05<" in later[0]
    assert morning[1:] == later[1:]
    assert len(cache) == len(morning) - 1
    model = convert_pptx_to_model(open_parts(parts), ConvertOptions(now=NOW))
    assert model.dated == frozenset({1})


def test_only_the_slides_asked_for_are_read():
    parts = parts_of("sample.pptx")
    parts["ppt/slides/slide2.xml"] = b"<p:sld this is not XML"
    package = open_parts(parts)
    assert [slide.slide_number for slide in read_presentation(package, [1, 3]).slides] == [1, 3]
    assert convert_pptx_to_svg(zipped(parts), ConvertOptions(slide_numbers=[1]))
    with pytest.raises(Exception):
        convert_pptx_to_svg(zipped(parts), ConvertOptions(slide_numbers=[2]))


def test_parts_in_memory_draw_as_their_zip_does():
    parts = parts_of("feature-sweep.pptx")
    assert drawn(open_parts(parts)) == drawn(zipped(parts)) == drawn(parts)
    assert isinstance(open_parts(parts), OpcPackage)


def test_decks_made_from_one_template_share_its_parses():
    parts = parts_of("sample.pptx")
    other = with_shape(parts, "ppt/slides/slide1.xml", TEXT_BOX)
    cache = ModelCache()
    drawn(open_parts(parts), cache=cache)
    parsed = dict(cache._parts)
    drawn(open_parts(other), cache=cache)
    assert cache._parts.keys() == parsed.keys()
    assert all(cache._parts[key] is value for key, value in parsed.items())


def test_what_is_kept_is_bounded():
    data = (FIXTURES / "sample.pptx").read_bytes()
    cache = ModelCache(slides=2, parts=1)
    plain = drawn(data)
    assert drawn(data, cache=cache) == plain
    assert len(cache) == 2 and len(cache._parts) == 1
    assert drawn(data, cache=cache) == plain
    cache.clear()
    assert len(cache) == 0


def test_threads_share_one_cache():
    data = (FIXTURES / "feature-sweep.pptx").read_bytes()
    plain = drawn(data)[0]
    cache = ModelCache()
    found: dict[int, list[str]] = {}
    errors: list[BaseException] = []

    def draw(number: int) -> None:
        try:
            for _ in range(3):
                found[number] = drawn(data, cache=cache, slide_numbers=[number])[0]
        except BaseException as error:  # noqa: BLE001
            errors.append(error)

    threads = [threading.Thread(target=draw, args=(number,))
               for number in range(1, len(plain) + 1)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    assert not errors
    assert [found[number][0] for number in range(1, len(plain) + 1)] == plain
