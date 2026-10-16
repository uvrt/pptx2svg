"""Date categories and PowerPoint's automatic chart title, drawn through ooxml-common 0.17.

Measured on PowerPoint 16 for Mac, 16 October 2026: serials under ``[$-413]d\\ mmmm\\ yyyy``
read ``5 januari 2026`` (not ``46027``); a ``c:title`` with no text over one series is the
series' name, bold, in the theme's minor face, at 1.2 times the chart's text size rounded
to the point (10 pt text gives 12 pt), 18 pt where the chart states none.
"""

from __future__ import annotations

import io
import re
import zipfile
from pathlib import Path

from ooxml_common.locales import Locale

from pptx2svg import ConvertOptions, convert_pptx_to_svg

FIXTURES = Path(__file__).resolve().parent / "fixtures"
#: chart-gallery.pptx's slide 2: one series ("Units") of horizontal bars, its title
#: deleted, the chart's text 10 pt.
SLIDE = 2
SERIALS = (46027, 46063, 46096, 46113, 46143)
DUTCH = Locale(decimal_separator=",", thousands_separator=".", date_order="dmy",
               language="nl-BE", list_separator=";")


def _deck(*, code: str | None = None, auto_title: bool = False, wide: bool = True) -> bytes:
    out = io.BytesIO()
    with zipfile.ZipFile(FIXTURES / "chart-gallery.pptx") as source, zipfile.ZipFile(out, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "ppt/charts/chart2.xml":
                text = data.decode()
                if code is not None:
                    points = "".join(f"<c:pt idx='{i}'><c:v>{v}</c:v></c:pt>" for i, v in enumerate(SERIALS))
                    text = re.sub(r"<c:cat>.*?</c:cat>",
                                  f"<c:cat><c:numRef><c:numCache><c:formatCode>{code}</c:formatCode>"
                                  f"<c:ptCount val='5'/>{points}</c:numCache></c:numRef></c:cat>", text,
                                  flags=re.S)
                if auto_title:
                    text = text.replace("<c:autoTitleDeleted val='1'/>",
                                        "<c:title><c:overlay val='0'/></c:title><c:autoTitleDeleted val='0'/>")
                if not wide:
                    text = re.sub(r"<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr><a:defRPr sz='1000'/>.*?</c:txPr>"
                                  r"</c:chartSpace>", "</c:chartSpace>", text, flags=re.S)
                data = text.encode()
            target.writestr(item, data)
    return out.getvalue()


def _tspans(svg: str) -> list[tuple[str, str]]:
    """``(text, attributes)`` of every tspan."""
    return [(text, attrs) for attrs, text in re.findall(r"<tspan([^>]*)>([^<]*)</tspan>", svg)]


def _svg(deck: bytes, **options) -> str:
    return convert_pptx_to_svg(deck, ConvertOptions(slide_numbers=[SLIDE], **options))[0]


def test_date_categories_are_drawn_as_dates():
    texts = [text for text, _ in _tspans(_svg(_deck(code=r"[$-413]d\ mmmm\ yyyy")))]
    assert {"5 januari 2026", "10 februari 2026", "15 maart 2026", "1 april 2026", "1 mei 2026"} <= set(texts)
    assert not any(str(serial) in texts for serial in SERIALS)


def test_a_slash_code_and_a_code_without_a_prefix():
    texts = {text for text, _ in _tspans(_svg(_deck(code="[$-409]dd/mm/yyyy")))}
    assert {"05/01/2026", "10/02/2026", "15/03/2026"} <= texts
    # No prefix: the names are the caller's locale's (en-US unless given).
    assert {"5 Jan", "15 Mar"} <= {text for text, _ in _tspans(_svg(_deck(code="d mmm")))}
    assert {"5 jan", "15 mrt"} <= {text for text, _ in _tspans(_svg(_deck(code="d mmm"), locale=DUTCH))}


def test_the_automatic_title_is_drawn_bold_at_1_2_times_the_charts_text():
    for wide, pixels in ((True, "16"), (False, "24")):   # 12 pt; 18 pt without a chart size
        spans = [attrs for text, attrs in _tspans(_svg(_deck(auto_title=True, wide=wide))) if text == "Units"]
        assert len(spans) == 1, wide
        assert f'font-size="{pixels}"' in spans[0]
        assert 'font-weight="bold"' in spans[0]
        assert 'font-family="Aptos,' in spans[0]   # the theme's minor face


def test_a_deleted_title_stays_off():
    assert not [text for text, _ in _tspans(_svg(_deck())) if text == "Units"]
