"""Autofit is what the file stores, through the whole read, resolve and render path.

PowerPoint fits nothing when it opens a file (ROADMAP.md, 5.8, measured on
``tools/make_autofit_probe.py``): ``a:normAutofit`` applies the ``fontScale`` and
``lnSpcReduction`` it stores, each scaled size rounded to a whole point, and overflowing
text with nothing stored is drawn at full size; ``a:spAutoFit`` keeps the stored extent.
"""

from __future__ import annotations

import re

from deckbuilder import derive_deck

from pptx2svg import ConvertOptions, convert_pptx_to_svg

PX = 4 / 3


def _box(name: str, fit: str, lines: int = 12) -> str:
    paragraphs = "".join(
        f"<a:p><a:r><a:rPr lang='en-US' sz='1800'/><a:t>Line {n:02d}</a:t></a:r></a:p>"
        for n in range(lines)
    )
    return (
        f"<p:sp><p:nvSpPr><p:cNvPr id='9000' name='{name}'/><p:cNvSpPr txBox='1'/>"
        "<p:nvPr/></p:nvSpPr><p:spPr><a:xfrm><a:off x='0' y='0'/>"
        "<a:ext cx='5080000' cy='1905000'/></a:xfrm><a:prstGeom prst='rect'><a:avLst/>"
        "</a:prstGeom><a:solidFill><a:srgbClr val='D9D9D9'/></a:solidFill></p:spPr>"
        f"<p:txBody><a:bodyPr wrap='square'>{fit}</a:bodyPr><a:lstStyle/>{paragraphs}"
        "</p:txBody></p:sp>"
    )


def _render(financial, *boxes: str) -> str:
    deck = derive_deck(financial, shapes_xml="".join(boxes))
    return convert_pptx_to_svg(deck, ConvertOptions(slide_numbers=[1]))[0]


def _shape(svg: str, name: str) -> str:
    return re.search(rf'<g [^>]*aria-label="{name}".*?</g>', svg, re.S).group(0)


def _sizes(shape: str) -> set[float]:
    return {round(float(size) / PX, 3) for size in re.findall(r'font-size="([^"]+)"', shape)}


def test_overflowing_text_with_nothing_stored_is_drawn_at_full_size(financial):
    svg = _render(financial, _box("probe-norm", "<a:normAutofit/>"))
    assert _sizes(_shape(svg, "probe-norm")) == {18.0}


def test_a_stored_font_scale_is_applied_rounded_and_not_shrunk_further(financial):
    svg = _render(financial, _box("probe-scale", "<a:normAutofit fontScale='62500'/>", 30))
    assert _sizes(_shape(svg, "probe-scale")) == {11.0}


def test_a_shape_that_fits_its_text_keeps_its_stored_extent(financial):
    svg = _render(financial, _box("probe-grow", "<a:spAutoFit/>"))
    height = re.search(r'<rect fill="#d9d9d9"[^>]* height="([^"]+)"', _shape(svg, "probe-grow"))
    assert float(height.group(1)) == 200.0
