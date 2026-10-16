"""A rich chart title is drawn in PowerPoint's own fonts, not the text cascade's.

Measured, PowerPoint 16 for Mac's PDF export, 16 October 2026, in a deck whose text boxes
take the theme's minor face (Aptos) -- oracle probes P1-P7, S1-S5 and B1-B4:

* no paragraph ``a:defRPr`` and no face on the run: **ArialMT, regular**, at the run's size
  or 18 pt, whatever the chart-wide ``c:txPr`` says (``+mn-lt`` there: still Arial);
* a paragraph ``a:defRPr``: the theme's **minor face** unless one is stated, **bold** unless
  ``b`` is stated (``sz`` alone: Aptos bold; ``b="0"``: Aptos regular);
* a run naming a face, no paragraph defaults: that face, regular.

(``ooxml_common.chart.layout.title_states_no_font``, ``title_states_defaults``.)
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
#: probe decks' did -- so a text box there is Aptos.
DEFAULT_STYLE = ('<a:defRPr lang="en-US"/>', '<a:defRPr lang="en-US"><a:latin typeface="+mn-lt"/></a:defRPr>')
FACES = "<a:latin typeface='+mn-lt'/><a:ea typeface='+mn-ea'/><a:cs typeface='+mn-cs'/>"
WIDE_FACE = ("<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz='1000'/>",
             "<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz='1000'>"
             "<a:latin typeface='+mn-lt'/></a:defRPr>")


def _rich(rpr: str = "<a:rPr lang='en-US'/>", ppr: str = "", txpr: str = "") -> str:
    return ("<c:title><c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p>" + ppr + "<a:r>" + rpr
            + "<a:t>Made-up Title</a:t></a:r></a:p></c:rich></c:tx><c:overlay val='0'/>" + txpr
            + "</c:title><c:autoTitleDeleted val='0'/>")


def _deck(title: str, *, wide_face: bool = False) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(FIXTURES / "chart-gallery.pptx") as source, zipfile.ZipFile(out, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "ppt/charts/chart2.xml":
                text = data.decode()
                assert DELETED in text
                text = text.replace(DELETED, title)
                if wide_face:
                    assert text.count(WIDE_FACE[0]) == 1
                    text = text.replace(*WIDE_FACE)
                data = text.encode()
            elif item.filename == "ppt/presentation.xml":
                text = data.decode()
                assert DEFAULT_STYLE[0] in text
                data = text.replace(*DEFAULT_STYLE).encode()
            target.writestr(item, data)
    return out.getvalue()


def _spans(title: str, **options) -> list[tuple[str, str]]:
    svg = convert_pptx_to_svg(_deck(title, **options), ConvertOptions(slide_numbers=[SLIDE]))[0]
    return [(attrs, text) for attrs, text in re.findall(r"<tspan([^>]*)>([^<]*)</tspan>", svg)]


def _drawn(title: str, **options) -> tuple[str, str, bool]:
    """``(face, size in px, bold)`` of the title."""
    spans = [attrs for attrs, text in _spans(title, **options) if text == "Made-up Title"]
    assert len(spans) == 1
    attrs = spans[0]
    face = re.search(r'font-family="([^",]*)', attrs).group(1)
    size = re.search(r'font-size="([^"]*)"', attrs).group(1)
    return face, size, 'font-weight="bold"' in attrs


def test_the_decks_text_boxes_are_aptos():
    # The slide's caption: what the text cascade gives text with no face of its own.
    captions = [attrs for attrs, text in _spans(_rich()) if "barChart" in text]
    assert captions and all('font-family="Aptos,' in attrs for attrs in captions)


UNSIZED_TXPR = ("<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr>"
                "<a:endParaRPr lang='en-US'/></a:p></c:txPr>")


@pytest.mark.parametrize("title, options, drawn", [
    (_rich(), {}, ("Arial", "24", False)),                                   # S1, P4
    (_rich(txpr=UNSIZED_TXPR), {}, ("Arial", "24", False)),                  # P3
    (_rich("<a:rPr lang='en-US' sz='2400'/>"), {}, ("Arial", "32", False)),  # S4: 24 pt
    (_rich(), {"wide_face": True}, ("Arial", "24", False)),                  # S5
    # B1: a paragraph a:defRPr with a size only: the minor face, bold.
    (_rich("<a:rPr lang='en-US' sz='1862'/>", "<a:pPr><a:defRPr sz='1862'/></a:pPr>"), {},
     ("Aptos", "24.827", True)),
    # S2: +mn-lt and 18.62 on both, no b: bold.
    (_rich(f"<a:rPr lang='en-US' sz='1862'>{FACES}</a:rPr>",
           f"<a:pPr><a:defRPr sz='1862'>{FACES}</a:defRPr></a:pPr>"), {}, ("Aptos", "24.827", True)),
    # B2: b="0" and +mn-lt: regular.
    (_rich(f"<a:rPr lang='en-US' sz='1862'>{FACES}</a:rPr>",
           f"<a:pPr><a:defRPr sz='1862' b='0'>{FACES}</a:defRPr></a:pPr>"), {},
     ("Aptos", "24.827", False)),
    # B4: b="0" alone: the minor face, regular.
    (_rich("<a:rPr lang='en-US' sz='1862'/>", "<a:pPr><a:defRPr b='0'/></a:pPr>"), {},
     ("Aptos", "24.827", False)),
    # B3: no paragraph defaults, the run names +mn-lt: regular.
    (_rich(f"<a:rPr lang='en-US' sz='1862'>{FACES}</a:rPr>"), {}, ("Aptos", "24.827", False)),
    # S3: the run names Georgia.
    (_rich("<a:rPr lang='en-US'><a:latin typeface='Georgia'/></a:rPr>"), {}, ("Georgia", "24", False)),
])
def test_a_rich_title_is_drawn_as_powerpoint_drew_it(title, options, drawn):
    assert _drawn(title, **options) == drawn
