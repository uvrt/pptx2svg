"""A shape's theme style references (``p:style``), held to what PowerPoint drew.

Every expectation here was read off PowerPoint's own PDF export of the decks
``tools/make_style_probe.py`` writes -- the same probes, on the same two themes, spliced
into the fixtures that lend those decks their themes.  ``tools/read_style_probe.py`` reads
the export; the probe key each test names is the shape it read.

``real-financial-report.pptx`` carries the 2013 Office theme (accent1 ``#4472C4``, accent2
``#ED7D31``, Calibri Light over Calibri) and ``sample.pptx`` the 2007 one (accent1
``#4F81BD``, accent2 ``#C0504D``, a shadow on every effect style).
"""

from __future__ import annotations

import re

import pytest

from pptx2svg import convert_pptx_to_model, model as m

from tests.conftest import FIXTURE_DIR
from tests.deckbuilder import derive_deck

THEME_2013 = FIXTURE_DIR / "real-financial-report.pptx"
THEME_2007 = FIXTURE_DIR / "sample.pptx"


def clr(value: str, **mods: int) -> str:
    children = "".join(f"<a:{name} val='{val}'/>" for name, val in mods.items())
    tag = "srgbClr" if re.fullmatch(r"[0-9A-F]{6}", value) else "schemeClr"
    return f"<a:{tag} val='{value}'>{children}</a:{tag}>"


def styled(
    *,
    ln=(0, clr("accent1")),
    fill=(1, clr("accent1")),
    effect=(0, clr("accent1")),
    font=("minor", clr("lt1")),
    sp_pr: str = "",
    run: str = "",
    list_style: str = "<a:lstStyle/>",
) -> str:
    """A rectangle saying "Styled", formatted only by its ``p:style`` unless told."""
    return (
        "<p:sp><p:nvSpPr><p:cNvPr id='{id}' name='{name}'/><p:cNvSpPr/><p:nvPr/></p:nvSpPr>"
        "<p:spPr><a:xfrm><a:off x='0' y='0'/><a:ext cx='1905000' cy='812800'/></a:xfrm>"
        f"<a:prstGeom prst='rect'><a:avLst/></a:prstGeom>{sp_pr}</p:spPr>"
        f"<p:style><a:lnRef idx='{ln[0]}'>{ln[1]}</a:lnRef>"
        f"<a:fillRef idx='{fill[0]}'>{fill[1]}</a:fillRef>"
        f"<a:effectRef idx='{effect[0]}'>{effect[1]}</a:effectRef>"
        f"<a:fontRef idx='{font[0]}'>{font[1]}</a:fontRef></p:style>"
        f"<p:txBody><a:bodyPr anchor='ctr'/>{list_style}<a:p><a:pPr algn='ctr'/>"
        f"<a:r><a:rPr lang='en-US' sz='2000'>{run}</a:rPr><a:t>Styled</a:t></a:r></a:p>"
        "</p:txBody></p:sp>"
    )


def resolve_shapes(source, **shapes: str) -> dict[str, m.ShapeElement]:
    """Splice each shape into slide 1 of ``source`` under its key, and resolve them."""
    xml = "".join(
        shape.replace("{id}", str(900 + n)).replace("{name}", key)
        for n, (key, shape) in enumerate(shapes.items())
    )
    resolved = convert_pptx_to_model(derive_deck(source, shapes_xml=xml))
    found = {
        element.alt_text: element
        for element in resolved.slides[0].elements
        if isinstance(element, m.ShapeElement)
    }
    return {key: found[key] for key in shapes}


def run_of(shape: m.ShapeElement) -> m.RunProperties:
    return shape.text_body.paragraphs[0].runs[0].properties


def near(hex_value: str, measured: str, tolerance: int = 4) -> bool:
    a, b = (bytes.fromhex(v.lstrip("#")) for v in (hex_value, measured))
    return all(abs(x - y) <= tolerance for x, y in zip(a, b))


def test_a_default_shape_is_filled_and_inked_from_its_style():
    """``default-shape``: what PowerPoint inserts.  Filled accent1, inked lt1, 1 pt
    outline of accent1 shaded to 15 %.  It used to be drawn unfilled and in black."""
    shape = resolve_shapes(
        THEME_2013, default=styled(ln=(2, clr("accent1", shade=15000)))
    )["default"]
    assert isinstance(shape.fill, m.SolidFill)
    assert shape.fill.color.hex == "#4472c4"
    assert run_of(shape).color.hex == "#ffffff"
    assert shape.outline.width == 12700
    assert near(shape.outline.fill.color.hex, "#172b51")
    assert shape.effects is None


def test_fill_reference_indices_name_both_fill_lists():
    """``fill-*``: 0 and 1000 are no fill; 1-3 the fill styles, 1001-1003 the
    background fill styles, each over the reference's own colour."""
    shapes = resolve_shapes(
        THEME_2013,
        none=styled(fill=(0, clr("accent1"))),
        none_bg=styled(fill=(1000, clr("accent2"))),
        solid=styled(fill=(1, clr("accent1"))),
        bg_solid=styled(fill=(1001, clr("accent2"))),
        gradient=styled(fill=(3, clr("accent1"))),
    )
    assert shapes["none"].fill is None
    assert shapes["none_bg"].fill is None
    assert shapes["solid"].fill.color.hex == "#4472c4"
    assert shapes["bg_solid"].fill.color.hex == "#ed7d31"
    gradient = shapes["gradient"].fill
    assert isinstance(gradient, m.GradientFill)
    # Measured at 50 % down the shape, where the theme's middle stop sits.
    assert near(gradient.stops[1].color.hex, "#3e70ca")


def test_a_gradient_style_derives_each_stop_from_the_shaded_reference():
    """``mod-2-shade50``: the reference's ``shade`` applies first and then each stop's
    own ``lumMod``/``satMod``/``tint`` -- three colours, not the one the reference names
    (ROADMAP 5.3's "Gradient stop overrides")."""
    fill = resolve_shapes(
        THEME_2013, shaded=styled(fill=(2, clr("accent2", shade=50000)))
    )["shaded"].fill
    assert isinstance(fill, m.GradientFill)
    assert len({stop.color.hex for stop in fill.stops}) == 3
    assert near(fill.stops[1].color.hex, "#cfa090")


def test_modifiers_and_alpha_on_the_reference_carry_through():
    """``mod-1-*``: ``shade``, ``lumMod``/``lumOff`` and ``alpha`` on the reference."""
    shapes = resolve_shapes(
        THEME_2013,
        shade=styled(fill=(1, clr("accent1", shade=50000))),
        lum=styled(fill=(1, clr("accent1", lumMod=60000, lumOff=40000))),
        alpha=styled(fill=(1, clr("accent1", alpha=50000))),
    )
    assert near(shapes["shade"].fill.color.hex, "#2e518e")
    assert near(shapes["lum"].fill.color.hex, "#8eaadc")
    assert shapes["alpha"].fill.color.hex == "#4472c4"
    assert shapes["alpha"].fill.color.alpha == 0.5


def test_an_explicit_fill_or_line_wins_over_the_reference():
    """``own-*``: the shape's own ``a:spPr`` over its style.  A local line that names
    only a colour keeps the referenced width (2 pt under the 2007 theme's ``idx="2"``),
    and a local ``a:noFill`` line turns the outline off."""
    shapes = resolve_shapes(
        THEME_2007,
        solid=styled(sp_pr="<a:solidFill><a:srgbClr val='00B050'/></a:solidFill>"),
        nofill=styled(sp_pr="<a:noFill/>"),
        line=styled(
            ln=(2, clr("accent1")),
            sp_pr="<a:ln><a:solidFill><a:srgbClr val='FF0000'/></a:solidFill></a:ln>",
        ),
        no_line=styled(ln=(3, clr("accent2")), sp_pr="<a:ln><a:noFill/></a:ln>"),
        width=styled(ln=(1, clr("accent2")), fill=(0, clr("accent1")),
                     sp_pr="<a:ln w='76200'/>"),
    )
    assert shapes["solid"].fill.color.hex == "#00b050"
    assert isinstance(shapes["nofill"].fill, m.NoFill)
    assert shapes["line"].outline.fill.color.hex == "#ff0000"
    assert shapes["line"].outline.width == 25400
    assert shapes["no_line"].outline is None
    # The 2007 theme's first line style shades its placeholder: #BE4B48, not #C0504D.
    assert shapes["width"].outline.width == 76200
    assert near(shapes["width"].outline.fill.color.hex, "#be4b48")


@pytest.mark.parametrize(
    "source, shadowed",
    [(THEME_2007, {1, 2, 3}), (THEME_2013, {3})],
    ids=["2007", "2013"],
)
def test_effect_references_count_from_one(source, shadowed):
    """``effect-*``: ``idx="0"`` is no effect, and ``idx="1"`` the first effect style.
    Reading the list from 0 put a shadow under every default shape on a 2007 theme."""
    shapes = resolve_shapes(
        source, **{f"e{idx}": styled(effect=(idx, clr("accent1"))) for idx in range(4)}
    )
    drawn = {
        idx
        for idx in range(4)
        if shapes[f"e{idx}"].effects is not None
        and shapes[f"e{idx}"].effects.outer_shadow is not None
    }
    assert drawn == shadowed


def test_the_font_reference_sits_over_the_masters_other_style():
    """``font-*``: the reference's colour and collection beat ``p:otherStyle`` (``tx1``,
    ``+mn-lt``); the shape's ``lstStyle`` and its runs beat the reference."""
    shapes = resolve_shapes(
        THEME_2013,
        major=styled(fill=(0, clr("accent1")), font=("major", clr("accent2"))),
        no_colour=styled(fill=(0, clr("accent1")), font=("minor", "")),
        no_face=styled(fill=(0, clr("accent1")), font=("none", clr("accent2"))),
        modified=styled(fill=(0, clr("accent1")), font=("minor", clr("accent1", lumMod=50000))),
        run=styled(font=("minor", clr("accent2")),
                   run="<a:solidFill><a:srgbClr val='00B050'/></a:solidFill>"),
        list_style=styled(
            font=("minor", clr("accent2")),
            list_style="<a:lstStyle><a:lvl1pPr><a:defRPr><a:solidFill>"
            "<a:srgbClr val='7030A0'/></a:solidFill></a:defRPr></a:lvl1pPr></a:lstStyle>",
        ),
    )
    assert run_of(shapes["major"]).color.hex == "#ed7d31"
    assert run_of(shapes["major"]).font_family == "Calibri Light"
    assert run_of(shapes["no_colour"]).color.hex == "#000000"
    assert run_of(shapes["no_colour"]).font_family == "Calibri"
    assert run_of(shapes["no_face"]).color.hex == "#ed7d31"
    assert run_of(shapes["no_face"]).font_family == "Calibri"
    assert near(run_of(shapes["modified"]).color.hex, "#203864")
    assert run_of(shapes["run"]).color.hex == "#00b050"
    assert run_of(shapes["list_style"]).color.hex == "#7030a0"
