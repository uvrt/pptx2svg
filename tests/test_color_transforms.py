"""Colour transforms, held to what PowerPoint drew for ``tools/make_color_probe.py``.

The probe deck is built here exactly as the tool builds it, read and resolved by pptx2svg,
and a selection of its 3,108 swatches is held to the level PowerPoint 16 wrote for each in
its PDF export -- a vector fill, so the number is the level drawn, not a raster reading of
it (``tools/read_color_probe.py``).  ROADMAP.md, "5.6 Colour transforms, measured".
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

from pptx2svg import convert_pptx_to_model, model as m

TOOLS = Path(__file__).resolve().parents[1] / "tools"
sys.path.insert(0, str(TOOLS))

import make_color_probe as probe  # noqa: E402

#: (swatch, the level PowerPoint drew).  The swatch is its colour expression, as
#: ``make_color_probe.key`` spells it.
DRAWN = [
    # In document order, the saturation unbounded above.
    ("accent2 tint=95000 satMod=170000", "#ff7818"),
    ("accent1 satMod=200000", "#0460ff"),
    ("accent1 satMod=200000 satMod=200000", "#003eff"),
    ("accent1 satMod=200000 satMod=50000", "#4371c0"),
    ("accent1 lumOff=40000 lumMod=60000", "#517cc8"),
    # The Office themes' own gradient stops.
    ("accent1 tint=93000 satMod=150000 shade=98000", "#4878e0"),
    ("accent1 lumMod=110000 satMod=105000 tint=67000", "#a8b7df"),
    ("accent1 lumMod=99000 satMod=120000 shade=78000", "#2e61ba"),
    # Kept in scRGB percentages: 127.5 levels is drawn 7F.
    ("dk1 lumOff=50000", "#7f7f7f"),
    # White after the clamp has no hue left to keep.
    ("FFE699 lumOff=40000 lumMod=60000", "#999999"),
    # Unbounded below, and a grey given a saturation.
    ("accent6 satOff=-50000", "#7c7084"),
    ("808080 satOff=25000", "#a06000"),
    # Rec. 709 grey, gamma and invGamma.
    ("accent2 gray", "#8f8f8f"),
    ("accent2 gray tint=50000", "#d1d1d1"),
    ("accent1 gamma", "#8db2e3"),
    ("accent1 invGamma", "#0f2b8d"),
    ("accent1 gamma tint=50000", "#d0ddf2"),
    ("accent1 invGamma shade=50000", "#081d66"),
    # The other colour elements.
    ("scrgb:50000,20000,0", "#bc7c00"),
    ("scrgb:50000,20000,0 satMod=150000", "#ea8a00"),
    ("hsl:14400000,50000,25000 satMod=150000", "#101070"),
    ("prst:orange lumMod=50000", "#7f5200"),
    ("sys:windowText:000000 tint=50000", "#bcbcbc"),
]


@pytest.fixture(scope="module")
def swatches(tmp_path_factory) -> dict[str, m.ShapeElement]:
    deck = tmp_path_factory.mktemp("color") / "color-probe.pptx"
    probe.write_deck(deck, probe.SOURCE, probe.slides())
    resolved = convert_pptx_to_model(str(deck))
    by_name = {
        element.alt_text: element
        for slide in resolved.slides
        for element in slide.elements
        if isinstance(element, m.ShapeElement)
    }
    return {
        probe.key(base, transforms): by_name[f"Swatch {100 + index}"]
        for index, (base, transforms) in enumerate(probe.PROBES)
    }


@pytest.mark.parametrize("swatch, drawn", DRAWN)
def test_a_swatch_is_filled_as_powerpoint_drew_it(swatches, swatch, drawn):
    fill = swatches[swatch].fill
    assert isinstance(fill, m.SolidFill)
    assert fill.color.hex == drawn


def test_alpha_mod_and_alpha_off_change_the_opacity(swatches):
    # Drawn at 64/255 and 191/255.
    assert swatches["accent1 alpha=50000 alphaMod=50000"].fill.color.alpha == 0.25
    assert swatches["accent1 alpha=50000 alphaOff=25000"].fill.color.alpha == 0.75
