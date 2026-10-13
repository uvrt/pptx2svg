"""The caller's locale (``ConvertOptions.locale``): a chart's number separators and the
language of a date field with none of its own.  Without one the conversion is en-US's,
never the machine's, and draws exactly what it drew before the option existed."""

from __future__ import annotations

import io
import pickle
import re
import sys
import zipfile
from dataclasses import replace
from datetime import datetime
from pathlib import Path

from ooxml_common.locales import EN_US, Locale

from pptx2svg import ConvertOptions, convert_pptx_to_svg

sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "tools"))
import make_field_probe as probe  # noqa: E402

FIXTURES = Path(__file__).resolve().parent / "fixtures"

#: A caller-built locale with a decimal comma (not a measured one: the separators are
#: the point here, not what Office does in any particular region).
COMMA = Locale(
    decimal_separator=",", thousands_separator=".", date_order="dmy", language="nl-NL",
    list_separator=";",
)

#: The clock, off the day == month diagonal.
NOW = datetime(2026, 3, 5, 9, 7, 4)

#: The probe deck's slide whose fields state no ``lang``.
NO_LANG_SLIDE = probe.LANGUAGES.index(None) + 1


def _texts(svg: str) -> list[str]:
    return ["".join(re.findall(r">([^<]*)<", text)) for text in re.findall(r"<text[^>]*>.*?</text>", svg, re.S)]


def _bar_chart_deck(format_code: str) -> bytes:
    """``chart-gallery.pptx`` with slide 2's data-labelled bars in ``format_code``."""
    out = io.BytesIO()
    with zipfile.ZipFile(FIXTURES / "chart-gallery.pptx") as source, zipfile.ZipFile(out, "w") as target:
        for item in source.infolist():
            data = source.read(item.filename)
            if item.filename == "ppt/charts/chart2.xml":
                data = data.replace(
                    b"<c:formatCode>General</c:formatCode>",
                    f"<c:formatCode>{format_code}</c:formatCode>".encode(),
                )
            target.writestr(item, data)
    return out.getvalue()


def test_the_default_locale_is_en_us_and_the_options_pickle_with_it():
    options = ConvertOptions()
    assert options.locale is EN_US
    restored = pickle.loads(pickle.dumps(ConvertOptions(locale=COMMA)))
    assert restored.locale == COMMA


def test_a_chart_prints_its_numbers_in_the_callers_separators():
    deck = _bar_chart_deck("#,##0.0")
    default = _texts(convert_pptx_to_svg(deck, ConvertOptions(slide_numbers=[2]))[0])
    comma = _texts(convert_pptx_to_svg(deck, ConvertOptions(slide_numbers=[2], locale=COMMA))[0])
    assert {"1,240.0", "860.0", "1,980.0"} <= set(default)
    assert {"1.240,0", "860,0", "1.980,0"} <= set(comma)
    assert "1,240.0" not in comma


def test_an_explicit_en_us_locale_draws_what_no_locale_draws():
    deck = _bar_chart_deck("#,##0.0")
    assert convert_pptx_to_svg(deck, ConvertOptions(slide_numbers=[2])) == convert_pptx_to_svg(
        deck, ConvertOptions(slide_numbers=[2], locale=replace(EN_US))
    )


def _no_lang_texts(locale: Locale) -> tuple[str, list]:
    options = ConvertOptions(slide_numbers=[NO_LANG_SLIDE], now=NOW, locale=locale)
    text = "".join(_texts(convert_pptx_to_svg(probe.build(1), options)[0]))
    return text, [w.message for w in options.warnings if w.code == "field-date-format"]


def test_a_date_field_without_a_lang_is_drawn_in_the_locales_language():
    text, warnings = _no_lang_texts(COMMA)
    assert "none.datetime1=[05-03-2026]" in text
    assert "none.datetime2=[donderdag 5 maart 2026]" in text
    assert len(warnings) == 1 and "nl-NL's" in warnings[0]


def test_without_a_locale_a_date_field_without_a_lang_is_en_us():
    text, warnings = _no_lang_texts(EN_US)
    assert "none.datetime1=[3/5/26]" in text
    assert "none.datetime2=[Thursday, March 5, 2026]" in text
    assert len(warnings) == 1 and "en-US's" in warnings[0]


def test_a_locale_language_without_a_table_falls_back_to_en_us():
    text, _ = _no_lang_texts(replace(COMMA, language="pt-BR"))
    assert "none.datetime1=[3/5/26]" in text


def test_a_fields_own_lang_beats_the_locale():
    options = ConvertOptions(slide_numbers=[1], now=NOW, locale=COMMA)  # the en-US slide
    text = "".join(_texts(convert_pptx_to_svg(probe.build(1), options)[0]))
    assert "en-US.datetime1=[3/5/26]" in text
    assert "en-US.datetime2=[Thursday, March 5, 2026]" in text

