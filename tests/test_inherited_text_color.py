"""A text box's inherited colour, size and face, in the order PowerPoint was measured to walk.

Measured on PowerPoint 16.106 for Mac, 23 October 2026 (round 15): made-up decks giving
each candidate layer its own colour and size, read from the slide show on screen, the PDF
export and AppleScript's ``font color`` of the text range, which agreed on every probe
(``ooxml_common.drawingml.text_chain``).  The decks here repeat the cases on
``authoring-integration.pptx``, whose own ``defaultTextStyle`` and ``otherStyle`` are empty.
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

from pptx2svg import convert_pptx_to_model

FIXTURE = Path(__file__).resolve().parent / "fixtures" / "authoring-integration.pptx"


def _fill(color):
    return f'<a:solidFill><a:srgbClr val="{color}"/></a:solidFill>' if color else ""


def _style(tag: str, levels: dict) -> str:
    body = ""
    for key in ["def"] + list(range(1, 10)):
        if key in levels:
            color, size = levels[key]
            name = "defPPr" if key == "def" else f"lvl{key}pPr"
            sz = f' sz="{size}"' if size else ""
            body += f"<a:{name}><a:defRPr{sz}>{_fill(color)}</a:defRPr></a:{name}>"
    return f"<p:{tag}>{body}</p:{tag}>"


def _shape(sid: int, name: str, *, level: int | None = None, lst: str = "",
           font_ref: str | None = None, textbox: bool = True) -> str:
    ppr = f'<a:pPr lvl="{level}"/>' if level is not None else ""
    style = ""
    if font_ref is not None:
        color = f'<a:schemeClr val="{font_ref}"/>' if font_ref else ""
        style = f'<p:style><a:lnRef idx="0"/><a:fillRef idx="0"/><a:effectRef idx="0"/>' \
                f'<a:fontRef idx="minor">{color}</a:fontRef></p:style>'
    box = ' txBox="1"' if textbox else ""
    return (f'<p:sp><p:nvSpPr><p:cNvPr id="{sid}" name="{name}"/>'
            f'<p:cNvSpPr{box}/><p:nvPr/></p:nvSpPr>'
            f'<p:spPr><a:xfrm><a:off x="0" y="{sid * 500000}"/><a:ext cx="2000000" cy="400000"/>'
            f'</a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>{style}'
            f'<p:txBody><a:bodyPr/><a:lstStyle>{lst}</a:lstStyle><a:p>{ppr}<a:r>'
            f'<a:rPr lang="en-US"/><a:t>Hello</a:t></a:r></a:p></p:txBody></p:sp>')


def _runs(default_style: str, other_style: str, shapes: list[str]) -> dict:
    out = io.BytesIO()
    with zipfile.ZipFile(FIXTURE) as source, zipfile.ZipFile(out, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "ppt/presentation.xml":
                data = re.sub(rb"<p:defaultTextStyle>.*?</p:defaultTextStyle>",
                              default_style.encode(), data, flags=re.S)
            elif item.filename == "ppt/slideMasters/slideMaster1.xml":
                assert b"<p:otherStyle/>" in data
                data = data.replace(b"<p:otherStyle/>", other_style.encode())
            elif item.filename == "ppt/slides/slide1.xml":
                data = data.replace(b"</p:spTree>", "".join(shapes).encode() + b"</p:spTree>")
            target.writestr(item, data)
    slide = convert_pptx_to_model(out.getvalue()).slides[0]
    found = {}
    for element in slide.elements:
        if element.alt_text and element.alt_text.startswith("P-"):
            run = element.text_body.paragraphs[0].runs[0].properties
            found[element.alt_text] = (run.color.hex.upper() if run.color else None, run.font_size)
    return found


def test_the_repro_a_white_default_text_style_draws_a_text_box_white():
    # defPPr black, every level white; the master's otherStyle tx1: PowerPoint draws white.
    default = _style("defaultTextStyle", {"def": ("000000", None),
                                          **{n: ("FFFFFF", 1800) for n in range(1, 10)}})
    other = '<p:otherStyle><a:lvl1pPr><a:defRPr sz="1800"><a:solidFill><a:schemeClr val="tx1"/>' \
            '</a:solidFill></a:defRPr></a:lvl1pPr></p:otherStyle>'
    found = _runs(default, other, [_shape(900, "P-box"), _shape(901, "P-lvl1", level=1)])
    assert found == {"P-box": ("#FFFFFF", 18.0), "P-lvl1": ("#FFFFFF", 18.0)}


def test_other_style_is_not_on_a_shapes_way_and_levels_map_one_up():
    default = _style("defaultTextStyle", {"def": ("000000", 1200), 1: ("FF0000", 2000),
                                          2: ("00B050", 2400), 3: ("0000FF", 3200)})
    other = _style("otherStyle", {"def": ("808000", None), 1: ("FFC000", 2800),
                                  2: ("7030A0", 1600), 3: ("00B0F0", 4000)})
    found = _runs(default, other, [
        _shape(900, "P-0"), _shape(901, "P-1", level=1), _shape(902, "P-2", level=2),
        # level 3: defaultTextStyle has no lvl4pPr, and its defPPr does not stand in
        _shape(903, "P-3", level=3),
        _shape(904, "P-rect", textbox=False),
    ])
    assert found == {"P-0": ("#FF0000", 20.0), "P-1": ("#00B050", 24.0),
                     "P-2": ("#0000FF", 32.0), "P-3": (None, None), "P-rect": ("#FF0000", 20.0)}


def test_font_ref_sits_over_the_default_style_and_under_the_list_style():
    default = _style("defaultTextStyle", {1: ("FF0000", 2000)})
    other = _style("otherStyle", {1: ("FFC000", 2800)})
    found = _runs(default, other, [
        _shape(900, "P-ref", font_ref="lt1", textbox=False),
        _shape(901, "P-ref-none", font_ref="", textbox=False),
        _shape(902, "P-ref-list", font_ref="lt1", textbox=False,
               lst='<a:lvl1pPr><a:defRPr><a:solidFill><a:srgbClr val="FF00FF"/></a:solidFill>'
                   '</a:defRPr></a:lvl1pPr>'),
        # a list style's defPPr is not a level either; a size alone leaves the colour below
        _shape(903, "P-list-def", lst='<a:defPPr><a:defRPr sz="3600"><a:solidFill>'
                                      '<a:srgbClr val="996633"/></a:solidFill></a:defRPr></a:defPPr>'),
        _shape(904, "P-list-size", lst='<a:lvl1pPr><a:defRPr sz="3600"/></a:lvl1pPr>'),
    ])
    assert found["P-ref"][0] == "#FFFFFF" and found["P-ref"][1] == 20.0
    assert found["P-ref-none"] == ("#FF0000", 20.0)
    assert found["P-ref-list"][0] == "#FF00FF"
    assert found["P-list-def"] == ("#FF0000", 20.0)
    assert found["P-list-size"] == ("#FF0000", 36.0)
