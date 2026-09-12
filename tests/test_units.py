"""Unit conversions and the OOXML colour algebra."""

from __future__ import annotations

import pytest

from pptx2svg.resolve.color import (
    DEFAULT_COLOR_MAP,
    ColorContext,
    build_effective_color_map,
    resolve_color,
)
from pptx2svg.parse.source import ColorTransform, SchemeColor, SourceTheme, SrgbColor, SystemColor
from pptx2svg.units import emu_to_pt, emu_to_px, px_to_emu, rotation_to_degrees


def test_emu_to_px_uses_96_dpi():
    # A 16:9 slide is 9,144,000 x 5,143,500 EMU = 960 x 540 px.
    assert emu_to_px(9144000) == 960
    assert emu_to_px(5143500) == 540


def test_emu_to_pt():
    assert emu_to_pt(12700) == 1


def test_px_to_emu_roundtrips():
    assert px_to_emu(emu_to_px(914400)) == pytest.approx(914400)


def test_rotation_unit_is_sixtythousandths_of_a_degree():
    assert rotation_to_degrees(2700000) == 45


def _context(**scheme) -> ColorContext:
    theme = SourceTheme(part_path="theme1.xml", color_scheme=scheme)
    return ColorContext(theme, dict(DEFAULT_COLOR_MAP))


def test_srgb_colour_passes_through_normalised():
    resolved = resolve_color(_context(), SrgbColor(hex="FF8800"))
    assert resolved.hex == "#ff8800"
    assert resolved.alpha == 1.0


def test_system_colour_uses_last_known_value():
    resolved = resolve_color(_context(), SystemColor(value="windowText", last_color="000000"))
    assert resolved.hex == "#000000"


def test_scheme_colour_goes_through_the_colour_map():
    # tx1 maps to dk1 by default, so a `tx1` reference resolves to the dk1 entry.
    context = _context(dk1=SrgbColor(hex="123456"))
    assert resolve_color(context, SchemeColor(scheme="tx1")).hex == "#123456"


def test_unknown_scheme_colour_falls_back_to_the_office_default():
    assert resolve_color(_context(), SchemeColor(scheme="accent1")).hex == "#4472c4"


def test_colour_map_override_layers_over_the_master():
    from pptx2svg.parse.source import SourceColorMap

    mapping = build_effective_color_map(
        SourceColorMap(mapping={"tx1": "dk2"}),
        SourceColorMap(mapping={"bg1": "dk1"}),
    )
    assert mapping["tx1"] == "dk2"  # from the master
    assert mapping["bg1"] == "dk1"  # from the layout override
    assert mapping["accent1"] == "accent1"  # schema default survives


def test_alpha_transform():
    color = SrgbColor(hex="FF0000", transforms=[ColorTransform(kind="alpha", value=50000)])
    assert resolve_color(_context(), color).alpha == 0.5


def test_tint_blends_toward_white():
    # tint 50% keeps half the colour and makes the rest up with white, blended in
    # linear light -- so half-way between black and white is #bcbcbc, not #808080.
    # Every value here was read off PowerPoint's own render of a swatch sheet.
    color = SrgbColor(hex="000000", transforms=[ColorTransform(kind="tint", value=50000)])
    assert resolve_color(_context(), color).hex == "#bcbcbc"

    blue = SchemeColor(scheme="accent1", transforms=[ColorTransform(kind="tint", value=40000)])
    assert resolve_color(_context(), blue).hex == "#cfd5ea"


def test_shade_scales_toward_black():
    color = SrgbColor(hex="FFFFFF", transforms=[ColorTransform(kind="shade", value=50000)])
    assert resolve_color(_context(), color).hex == "#bcbcbc"

    blue = SchemeColor(scheme="accent1", transforms=[ColorTransform(kind="shade", value=50000)])
    assert resolve_color(_context(), blue).hex == "#2f528f"


def test_lum_mod_and_lum_off_apply_together():
    # A common theme pattern: lumMod 60% + lumOff 40% lightens a colour.
    color = SrgbColor(
        hex="000000",
        transforms=[
            ColorTransform(kind="lumMod", value=60000),
            ColorTransform(kind="lumOff", value=40000),
        ],
    )
    assert resolve_color(_context(), color).hex == "#666666"


def test_colour_map_cycle_terminates():
    # tx1 -> dk1 -> tx1 would recurse forever without cycle detection.
    context = ColorContext(
        SourceTheme(part_path="t", color_scheme={"dk1": SchemeColor(scheme="tx1")}),
        dict(DEFAULT_COLOR_MAP),
    )
    assert resolve_color(context, SchemeColor(scheme="tx1")).hex == "#000000"
