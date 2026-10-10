"""Text fields evaluated per slide, as PowerPoint draws them (``resolve/fields.py``).

The decks are ``tools/make_field_probe.py``'s, built here from ``sample.pptx``; the
expected strings are what PowerPoint 16 for Mac drew for the same decks, exported
2026-10-10 at 21:26 (the module docstring of ``resolve/fields.py`` has the details).  The
clock is pinned to that moment, so every reading is reproduced exactly.
"""

from __future__ import annotations

import re
import sys
from datetime import datetime
from pathlib import Path

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_model, convert_pptx_to_svg
from pptx2svg.cli import main as cli_main
from pptx2svg.resolve.fields import evaluate_field, format_date

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import make_field_probe as probe  # noqa: E402

#: PowerPoint's readings.  en-US was exported at 21:26:37, the other pages at :40.
READINGS = {
    "en-US": (37, {
        "datetime": "10/10/26", "datetime1": "10/10/26",
        "datetime2": "Saturday, October 10, 2026", "datetime3": "10 October 2026",
        "datetime4": "October 10, 2026", "datetime5": "10-Oct-26", "datetime6": "October 26",
        "datetime7": "Oct-26", "datetime8": "10/10/26 9:26 PM",
        "datetime9": "10/10/26 9:26:37 PM", "datetime10": "21:26", "datetime11": "21:26:37",
        "datetime12": "9:26 PM", "datetime13": "9:26:37 PM", "datetimeFigureOut": "10/10/26",
    }),
    "en-GB": (40, {
        "datetime": "10/10/2026", "datetime1": "10/10/2026",
        "datetime2": "Saturday, 10 October 2026", "datetime3": "10 October, 2026",
        "datetime4": "10 October 2026", "datetime5": "10-Oct-26", "datetime6": "October 26",
        "datetime7": "Oct-26", "datetime8": "10/10/2026 21:26",
        "datetime9": "10/10/2026 21:26:40", "datetime10": "21:26", "datetime11": "21:26:40",
        "datetime12": "9:26 pm", "datetime13": "9:26:40 pm", "datetimeFigureOut": "10/10/2026",
    }),
    "nl-NL": (40, {
        "datetime": "10-10-2026", "datetime1": "10-10-2026",
        "datetime2": "zaterdag 10 oktober 2026", "datetime3": "10/10/26",
        "datetime4": "10 oktober 2026", "datetime5": "10-okt-26", "datetime6": "oktober ’26",
        "datetime7": "okt-26", "datetime8": "10-10-2026 21:26",
        "datetime9": "10-10-2026 21:26:40", "datetime10": "21:26", "datetime11": "21:26:40",
        "datetime12": "9:26 PM", "datetime13": "9:26:40 PM", "datetimeFigureOut": "10-10-2026",
    }),
    "de-DE": (40, {
        "datetime": "10.10.26", "datetime1": "10.10.26",
        "datetime2": "Samstag, 10. Oktober 2026", "datetime3": "10/10/2026",
        "datetime4": "10. Oktober 2026", "datetime5": "26-10-10", "datetime6": "Oktober 26",
        "datetime7": "Okt-26", "datetime8": "10.10.26 21:26", "datetime9": "10.10.26 21:26:40",
        "datetime10": "21:26", "datetime11": "21:26:40", "datetime12": "9:26 PM",
        "datetime13": "9:26:40 PM", "datetimeFigureOut": "10.10.26",
    }),
    "fr-FR": (40, {
        "datetime": "10/10/2026", "datetime1": "10/10/2026",
        "datetime2": "samedi 10 octobre 2026", "datetime3": "10.10.26",
        "datetime4": "10 octobre 2026", "datetime5": "10-oct.-26", "datetime6": "octobre 26",
        "datetime7": "oct.-26", "datetime8": "10/10/2026 21:26",
        "datetime9": "10/10/2026 21:26:40", "datetime10": "21:26", "datetime11": "21:26:40",
        "datetime12": "9:26 PM", "datetime13": "9:26:40 PM", "datetimeFigureOut": "10/10/2026",
    }),
}

CASES = [
    (lang, kind, second, text)
    for lang, (second, table) in READINGS.items()
    for kind, text in table.items()
]


@pytest.mark.parametrize("lang,kind,second,expected", CASES)
def test_a_date_field_is_formatted_as_powerpoint_drew_it(lang, kind, second, expected):
    now = datetime(2026, 10, 10, 21, 26, second)
    text, missing = evaluate_field(kind, "1/1/2020", slide_number=1, now=now, lang=lang)
    assert (text, missing) == (expected, None)


def test_day_month_and_padding_follow_the_locale_pattern_off_the_diagonal():
    """The one reading had day == month == 10; off it, the patterns' own order and
    padding decide (CLDR's short date: en-US ``M/d/yy``, de-DE ``dd.MM.yy``)."""
    now = datetime(2026, 3, 5, 9, 7, 4)
    assert format_date(now, "M/d/yy h:mm a", "en-US") == "3/5/26 9:07 AM"
    assert format_date(now, "dd.MM.yy HH:mm", "de-DE") == "05.03.26 09:07"
    assert format_date(now, "d-MMM-yy", "de-DE") == "5-Mär-26"


def _texts(svg: str) -> str:
    return "".join(re.findall(r">([^<]*)<", svg))


def _probe_svgs(first: int = 1, now=datetime(2026, 10, 10, 21, 26, 40)):
    deck = probe.build(first)
    return convert_pptx_to_svg(deck, ConvertOptions(now=now))


def test_a_slide_number_in_a_layout_and_a_master_text_box_is_the_slides_own():
    svgs = _probe_svgs()
    assert len(svgs) == 8
    for number, svg in enumerate(svgs, start=1):
        text = _texts(svg)
        assert f"layout.slidenum=[{number}]" in text
        assert f"master.slidenum=[{number}]" in text
        assert "‹#›" not in text
        assert "layout.datetime1=[10/10/26]" in text
        assert "1/1/2020" not in text


def test_first_slide_num_offsets_every_slide_number():
    svgs = _probe_svgs(first=5)
    assert "layout.slidenum=[5]" in _texts(svgs[0])
    last = _texts(svgs[-1])
    assert "layout.slidenum=[12]" in last and "ph.sldNum=[12]" in last


def test_a_slide_placeholders_stale_cache_is_re_evaluated():
    last = _texts(_probe_svgs()[-1])
    assert "ph.sldNum=[8]" in last
    assert "ph.dt=[10/10/26]" in last


def test_a_field_type_powerpoint_does_not_evaluate_keeps_its_cached_text():
    assert "en-US.unknown=[CACHED]" in _texts(_probe_svgs()[0])


def test_the_probe_slides_read_as_powerpoint_drew_them():
    svgs = _probe_svgs()
    for page, lang in enumerate(("en-US", "en-GB", "nl-NL", "de-DE", "fr-FR")):
        second, table = READINGS[lang]
        text = _texts(
            convert_pptx_to_svg(
                probe.build(1),
                ConvertOptions(slide_numbers=[page + 1], now=datetime(2026, 10, 10, 21, 26, second)),
            )[0]
        )
        for kind, expected in table.items():
            assert f"{lang}.{kind}=[{expected}]" in text, (lang, kind)
    assert len(svgs) == 8


def test_a_language_without_a_table_falls_back_to_en_us_and_says_so():
    model = convert_pptx_to_model(
        probe.build(1), ConvertOptions(slide_numbers=[6], now=datetime(2026, 3, 5, 9, 7, 4))
    )
    runs = [
        run.text
        for element in model.slides[0].elements
        if getattr(element, "text_body", None)
        for paragraph in element.text_body.paragraphs
        for run in paragraph.runs
    ]
    assert "Thursday, March 5, 2026" in runs  # ja-JP datetime2, in en-US's format
    codes = [(w.code, w.message) for w in model.warnings if w.code == "field-date-format"]
    assert len(codes) == 1 and "ja-JP" in codes[0][1]


def test_the_clock_defaults_to_the_time_of_conversion():
    before = datetime.now()
    text = _texts(convert_pptx_to_svg(probe.build(1), ConvertOptions(slide_numbers=[1]))[0])
    after = datetime.now()
    shown = re.search(r"en-US\.datetime4=\[([^\]]*)\]", text).group(1)
    assert shown in {format_date(moment, "MMMM d, yyyy", "en-US") for moment in (before, after)}


def test_the_cli_pins_the_clock(tmp_path):
    deck = tmp_path / "fields.pptx"
    deck.write_bytes(probe.build(1))
    assert cli_main([str(deck), "-o", str(tmp_path), "--slides", "1", "--now",
                     "2026-03-05T09:07", "-q"]) == 0
    text = _texts((tmp_path / "fields-1.svg").read_text(encoding="utf-8"))
    assert "en-US.datetime1=[3/5/26]" in text
