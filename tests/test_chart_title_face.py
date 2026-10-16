"""A rich chart title that states no font is drawn in Arial, as PowerPoint draws it.

Measured, PowerPoint 16 for Mac's PDF export, 16 October 2026: a ``c:title/c:tx/c:rich``
whose runs, paragraphs, list style and ``c:txPr`` state no face and no size came out
**18 pt ArialMT, regular, black**, at a chart-wide size of none, 8 and 24 pt and under an
unsized title ``c:txPr`` -- in a deck whose text boxes get the theme's minor face (Aptos).
A title that names a face (``+mn-lt``, ``Georgia``) keeps it; an automatic title stays the
minor face, bold (``ooxml_common.chart.layout.title_states_no_font``).
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_svg

FIXTURES = Path(__file__).resolve().parent / "fixtures"
#: chart-gallery.pptx's slide 2: one series ("Units") of horizontal bars, its title deleted,
#: the chart's text 10 pt.
SLIDE = 2
DELETED = "<c:autoTitleDeleted val='1'/>"
#: The deck's default text style, made to name the theme's minor face (Aptos) as the
#: probe deck's did -- so a text box there is Aptos, and the title is not.
DEFAULT_STYLE = ('<a:defRPr lang="en-US"/>', '<a:defRPr lang="en-US"><a:latin typeface="+mn-lt"/></a:defRPr>')


def _rich(rpr: str = "<a:rPr lang='en-US'/>", txpr: str = "") -> str:
    return ("<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>" + rpr
            + "<a:t>Made-up Title</a:t></a:r></a:p></c:rich></c:tx><c:overlay val='0'/>" + txpr
            + "</c:title><c:autoTitleDeleted val='0'/>")


def _deck(title: str) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(FIXTURES / "chart-gallery.pptx") as source, zipfile.ZipFile(out, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "ppt/charts/chart2.xml":
                text = data.decode()
                assert DELETED in text
                data = text.replace(DELETED, title).encode()
            elif item.filename == "ppt/presentation.xml":
                text = data.decode()
                assert DEFAULT_STYLE[0] in text
                data = text.replace(*DEFAULT_STYLE).encode()
            target.writestr(item, data)
    return out.getvalue()


def _title(title: str) -> str:
    svg = convert_pptx_to_svg(_deck(title), ConvertOptions(slide_numbers=[SLIDE]))[0]
    spans = [attrs for attrs, text in re.findall(r"<tspan([^>]*)>([^<]*)</tspan>", svg)
             if text == "Made-up Title"]
    assert len(spans) == 1
    return spans[0]


def test_the_decks_text_boxes_are_aptos():
    # What the cascade gives an unstyled title -- and what pptx2svg drew before.
    assert 'font-family="Aptos,' in _title(_rich("<a:rPr lang='en-US' sz='1800'/>"))


@pytest.mark.parametrize("txpr", ["", "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr>"
                                      "<a:endParaRPr lang='en-US'/></a:p></c:txPr>"])
def test_a_title_that_states_no_font_is_arial_18_regular(txpr):
    attrs = _title(_rich(txpr=txpr))
    assert 'font-family="Arial,' in attrs
    assert 'font-size="24"' in attrs          # 18 pt
    assert "font-weight" not in attrs


@pytest.mark.parametrize("rpr, family", [
    ("<a:rPr lang='en-US'><a:latin typeface='+mn-lt'/></a:rPr>", "Aptos"),   # the fix check suggests
    ("<a:rPr lang='en-US' sz='1862'><a:latin typeface='+mn-lt'/></a:rPr>", "Aptos"),
    ("<a:rPr lang='en-US'><a:latin typeface='Georgia'/></a:rPr>", "Georgia"),
])
def test_a_title_that_names_a_face_keeps_it(rpr, family):
    assert f'font-family="{family}' in _title(_rich(rpr))
