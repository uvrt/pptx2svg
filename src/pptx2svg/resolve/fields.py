"""Text fields (``a:fld``) evaluated for the slide they are drawn on.

A field's ``a:t`` is a cache: ``‹#›`` for a slide number on a layout, and for a date the
day the file was last saved.  PowerPoint draws neither.  It evaluates the field for every
slide it is drawn on, wherever the field sits -- a slide's own ``sldNum`` and ``dt``
placeholders, and equally a plain text box on the layout or the master, the usual
corporate footer ("‹#› | Copyright ...").

Measured (``tools/make_field_probe.py``; PowerPoint 16 for Mac, exported 2026-10-10
21:26 CEST on a machine whose own locale is nl-BE):

* ``slidenum`` is the slide's position in the deck plus ``p:presentation@firstSlideNum``
  minus one -- 5, 6, ... 12 for eight slides with ``firstSlideNum="5"`` -- in a layout or
  master text box and in a slide placeholder alike, whatever the cached text says
  (``‹#›``, ``99``).
* ``datetime``, ``datetime1`` to ``datetime13`` and ``datetimeFigureOut`` are the clock
  at the time of drawing, formatted by the field's ``lang``: :data:`DATE_FORMATS`, read off
  the export field by field.  ``datetime`` and ``datetimeFigureOut`` are ``datetime1``.
* A type PowerPoint does not know keeps its cached text, and so does every type here
  that is not a slide number or a date.

The clock is :attr:`pptx2svg.ConvertOptions.now`, the time of conversion unless a caller
pins it -- as the tests and the snapshot goldens do, since output that holds a date is
otherwise different every day.

What one reading on 10 October cannot settle: the order of day and month and whether
either is zero-padded, since both are 10.  There the patterns follow the locale's own
short date pattern (Unicode CLDR as macOS ships it), which every measured numeric
reading agrees with (en-US ``M/d/yy`` gives ``10/10/26``, de-DE ``dd.MM.yy`` gives
``10.10.26``).  A language with no measured table draws in en-US's formats and says so
(``field-date-format``): PowerPoint itself falls back to the machine's locale, which a
renderer has no business reproducing.  Japanese is not in the table for that reason --
PowerPoint for Mac drew three of its thirteen formats from the machine's Dutch.
"""

from __future__ import annotations

from datetime import datetime

#: ``datetime1`` .. ``datetime13`` per language, in the pattern letters of
#: :func:`format_date`.  Measured: see the module docstring.
DATE_FORMATS: dict[str, tuple[str, ...]] = {
    "en-US": (
        "M/d/yy", "EEEE, MMMM d, yyyy", "d MMMM yyyy", "MMMM d, yyyy", "d-MMM-yy",
        "MMMM yy", "MMM-yy", "M/d/yy h:mm a", "M/d/yy h:mm:ss a", "H:mm", "H:mm:ss",
        "h:mm a", "h:mm:ss a",
    ),
    "en-GB": (
        "dd/MM/yyyy", "EEEE, d MMMM yyyy", "d MMMM, yyyy", "d MMMM yyyy", "d-MMM-yy",
        "MMMM yy", "MMM-yy", "dd/MM/yyyy HH:mm", "dd/MM/yyyy HH:mm:ss", "HH:mm",
        "HH:mm:ss", "h:mm a", "h:mm:ss a",
    ),
    "nl-NL": (
        "dd-MM-yyyy", "EEEE d MMMM yyyy", "dd/MM/yy", "d MMMM yyyy", "d-MMM-yy",
        "MMMM ’yy", "MMM-yy", "dd-MM-yyyy HH:mm", "dd-MM-yyyy HH:mm:ss", "HH:mm",
        "HH:mm:ss", "h:mm a", "h:mm:ss a",
    ),
    "de-DE": (
        "dd.MM.yy", "EEEE, d. MMMM yyyy", "dd/MM/yyyy", "d. MMMM yyyy", "yy-MM-dd",
        "MMMM yy", "MMM-yy", "dd.MM.yy HH:mm", "dd.MM.yy HH:mm:ss", "HH:mm", "HH:mm:ss",
        "h:mm a", "h:mm:ss a",
    ),
    "fr-FR": (
        "dd/MM/yyyy", "EEEE d MMMM yyyy", "dd.MM.yy", "d MMMM yyyy", "d-MMM-yy",
        "MMMM yy", "MMM-yy", "dd/MM/yyyy HH:mm", "dd/MM/yyyy HH:mm:ss", "HH:mm",
        "HH:mm:ss", "h:mm a", "h:mm:ss a",
    ),
}

#: Month names, abbreviated month names, weekday names (Monday first) and the AM/PM
#: markers, per language of :data:`DATE_FORMATS`.  The abbreviations are the ones
#: PowerPoint drew (German ``Okt``, not CLDR's ``Okt.``; French ``oct.``).
NAMES: dict[str, tuple[tuple[str, ...], tuple[str, ...], tuple[str, ...], tuple[str, str]]] = {
    "en": (
        ("January", "February", "March", "April", "May", "June", "July", "August",
         "September", "October", "November", "December"),
        ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"),
        ("Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"),
        ("AM", "PM"),
    ),
    "nl": (
        ("januari", "februari", "maart", "april", "mei", "juni", "juli", "augustus",
         "september", "oktober", "november", "december"),
        ("jan", "feb", "mrt", "apr", "mei", "jun", "jul", "aug", "sep", "okt", "nov", "dec"),
        ("maandag", "dinsdag", "woensdag", "donderdag", "vrijdag", "zaterdag", "zondag"),
        ("AM", "PM"),
    ),
    "de": (
        ("Januar", "Februar", "März", "April", "Mai", "Juni", "Juli", "August",
         "September", "Oktober", "November", "Dezember"),
        ("Jan", "Feb", "Mär", "Apr", "Mai", "Jun", "Jul", "Aug", "Sep", "Okt", "Nov", "Dez"),
        ("Montag", "Dienstag", "Mittwoch", "Donnerstag", "Freitag", "Samstag", "Sonntag"),
        ("AM", "PM"),
    ),
    "fr": (
        ("janvier", "février", "mars", "avril", "mai", "juin", "juillet", "août",
         "septembre", "octobre", "novembre", "décembre"),
        ("janv.", "févr.", "mars", "avr.", "mai", "juin", "juil.", "août", "sept.", "oct.",
         "nov.", "déc."),
        ("lundi", "mardi", "mercredi", "jeudi", "vendredi", "samedi", "dimanche"),
        ("AM", "PM"),
    ),
}
#: en-GB writes its markers in lower case ("9:26 pm"); measured.
_MARKERS = {"en-GB": ("am", "pm")}

#: The language a field with no (known) ``lang`` is formatted in.
DEFAULT_LANGUAGE = "en-US"

SLIDE_NUMBER = "slidenum"
_DATE_ALIASES = {"datetime": 1, "datetimefigureout": 1}


def is_date_field(field_type: str) -> bool:
    return date_format_index(field_type) is not None


def date_format_index(field_type: str) -> int | None:
    """1..13 for a date field type, ``None`` for anything else."""
    kind = field_type.lower()
    if kind in _DATE_ALIASES:
        return _DATE_ALIASES[kind]
    if kind.startswith("datetime") and kind[8:].isdigit():
        number = int(kind[8:])
        return number if 1 <= number <= 13 else None
    return None


def table_language(lang: str | None) -> str | None:
    """The :data:`DATE_FORMATS` key for ``lang``, matched without regard to case;
    ``None`` when there is no table for it."""
    if not lang:
        return None
    wanted = lang.replace("_", "-").lower()
    for key in DATE_FORMATS:
        if key.lower() == wanted:
            return key
    return None


def evaluate_field(
    field_type: str | None,
    cached: str,
    *,
    slide_number: int,
    now: datetime,
    lang: str | None,
) -> tuple[str, str | None]:
    """The text PowerPoint draws for a field, and the language it fell back from when
    it had no table for the field's own (``None`` when it did, or for a non-date)."""
    if not field_type:
        return cached, None
    if field_type.lower() == SLIDE_NUMBER:
        return str(slide_number), None
    index = date_format_index(field_type)
    if index is None:
        return cached, None
    language = table_language(lang)
    missing = None
    if language is None:
        missing, language = (lang or ""), DEFAULT_LANGUAGE
    return format_date(now, DATE_FORMATS[language][index - 1], language), missing


def format_date(moment: datetime, pattern: str, language: str) -> str:
    """``moment`` in ``pattern``: ``d dd M MM MMM MMMM yy yyyy EEEE H HH h mm ss a``, as
    in Unicode's date patterns; any other character is literal."""
    months, short_months, weekdays, markers = NAMES[language.split("-")[0]]
    markers = _MARKERS.get(language, markers)
    out: list[str] = []
    index = 0
    while index < len(pattern):
        letter = pattern[index]
        end = index
        while end < len(pattern) and pattern[end] == letter:
            end += 1
        count = end - index
        index = end
        if letter == "d":
            out.append(f"{moment.day:0{count}d}")
        elif letter == "M":
            if count >= 4:
                out.append(months[moment.month - 1])
            elif count == 3:
                out.append(short_months[moment.month - 1])
            else:
                out.append(f"{moment.month:0{count}d}")
        elif letter == "y":
            out.append(f"{moment.year % 100:02d}" if count == 2 else f"{moment.year:04d}")
        elif letter == "E":
            out.append(weekdays[moment.weekday()])
        elif letter == "H":
            out.append(f"{moment.hour:0{count}d}")
        elif letter == "h":
            out.append(f"{(moment.hour % 12) or 12:0{count}d}")
        elif letter == "m":
            out.append(f"{moment.minute:02d}")
        elif letter == "s":
            out.append(f"{moment.second:02d}")
        elif letter == "a":
            out.append(markers[moment.hour >= 12])
        else:
            out.append(letter * count)
    return "".join(out)
