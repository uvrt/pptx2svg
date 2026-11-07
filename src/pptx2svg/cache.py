"""Resolved slides kept by what they were resolved from: :class:`ModelCache`.

A program that converts the same deck again and again -- after every edit, to measure or
draw one slide -- resolves mostly the same slides each time.  A slide's resolution reads
its own part and what that part's relationships reach (its layout, the master, the
master's other layouts, the theme, pictures, charts, diagrams), the presentation's own
settings (slide size, default text style, table styles, embedded font list, first slide
number, the application version) and the slide's place in the deck -- nothing of another
slide.  So a resolved slide is kept under a key made of exactly those: the bytes of every
part in that reach (and of their relationships), their content types, the presentation's
settings as read, the slide's number, id and part, and the conversion's locale and font
choice.  Equal key, equal slide: a conversion handed ``ConvertOptions(cache=...)``
resolves only the slides whose key it has not seen, and gives back for the others the
very objects it resolved before -- which therefore must be treated as read-only.

Not kept:

* a slide on which a date field (``a:fld type="datetime…"``) was drawn: it shows the
  conversion's clock (``ConvertOptions.now``), so it is resolved every time;
* every slide of a conversion with a ``metafile_converter`` (a callable; its answers are
  not known to depend on the bytes alone) or with ``templates`` (layout and master
  views): such a conversion runs as without a cache.

Layouts, masters and themes are parsed once per their bytes, too (:meth:`ModelCache.parsed`),
so decks made from one template share those parses.  Both stores are bounded, least
recently used first out, and safe to share between threads.
"""

from __future__ import annotations

import hashlib
import threading
from collections import OrderedDict
from datetime import datetime
from typing import Callable

from . import model as m
from .opc import REL_SLIDE, OpcPackage, rels_path_for
from .parse.parts import read_presentation, slide_entries

__all__ = ["ModelCache"]

#: What :class:`ModelCache` keeps by default: resolved slides, and parsed layouts, masters
#: and themes.
SLIDES = 512
PARTS = 256


def _digest(data: bytes | None) -> bytes:
    return b"-" if data is None else hashlib.blake2b(data, digest_size=16).digest()


class ModelCache:
    """Resolved slides, and parsed layouts, masters and themes, by what they are made of
    (see the module docstring).  Hand one to every conversion that should share it::

        cache = ModelCache()
        options = ConvertOptions(slide_numbers=[3], cache=cache)
        model = convert_pptx_to_model(deck_bytes, options)

    ``slides`` and ``parts`` bound what is kept.  A slide given back from it is the object
    an earlier conversion returned: read it, do not change it."""

    def __init__(self, slides: int = SLIDES, parts: int = PARTS) -> None:
        self._limit_slides = max(0, int(slides))
        self._limit_parts = max(0, int(parts))
        self._slides: OrderedDict[bytes, tuple] = OrderedDict()
        self._parts: OrderedDict[tuple, object] = OrderedDict()
        self._lock = threading.Lock()
        #: Slides given back from the cache, and slides resolved, since it was made.
        self.hits = 0
        self.misses = 0

    def __len__(self) -> int:
        with self._lock:
            return len(self._slides)

    def clear(self) -> None:
        with self._lock:
            self._slides.clear()
            self._parts.clear()

    # -- parsed parts ----------------------------------------------------------------------

    def parsed(self, package: OpcPackage, path: str, reader: Callable) -> object:
        """``reader(package, path)``, once per the part's bytes and its relationships' (a
        :data:`~pptx2svg.parse.parts.PartReader`)."""
        key = (f"{reader.__module__}.{reader.__qualname__}", path,
               _digest(package.read(path)), _digest(package.read(rels_path_for(path))))
        with self._lock:
            found = self._parts.get(key)
            if found is not None:
                self._parts.move_to_end(key)
                return found
        value = reader(package, path)
        if value is not None and self._limit_parts:
            with self._lock:
                self._parts[key] = value
                while len(self._parts) > self._limit_parts:
                    self._parts.popitem(last=False)
        return value

    # -- resolved slides -------------------------------------------------------------------

    def resolve(self, package: OpcPackage, options, *, east_asian, host_fonts: bool,
                now: datetime | None):
        """``(presentation, resolved)`` as :func:`pptx2svg.convert_pptx_to_model` reads and
        resolves them, the slides kept here taken from here."""
        from .resolve.view import ResolvedPresentation, _font_scheme, resolve_presentation

        header = read_presentation(package, (), parsed=self.parsed)
        entries = slide_entries(package)
        count = len(entries)
        if options.slide_numbers is None:
            numbers = list(range(1, count + 1))
        else:
            numbers = sorted({n for n in options.slide_numbers if 1 <= n <= count})
        # What every slide's resolution reads besides its own reach.
        settings = repr((header.part_path, header.slide_width, header.slide_height,
                         header.default_text_style, header.table_styles,
                         header.embedded_fonts, header.first_slide_number,
                         header.app_version, options.locale, bool(host_fonts)))
        digests: dict[str, tuple] = {}
        keys = {number: self._key(package, settings, number, *entries[number - 1], digests)
                for number in numbers}
        found: dict[int, tuple] = {}
        with self._lock:
            for number, key in keys.items():
                entry = self._slides.get(key) if key is not None else None
                if entry is not None:
                    self._slides.move_to_end(key)
                    found[number] = entry
            self.hits += len(found)
            self.misses += len(numbers) - len(found)
        missing = [number for number in numbers if number not in found]
        if missing:
            now = now if now is not None else datetime.now()
            presentation = read_presentation(package, missing, parsed=self.parsed)
            for source in presentation.slides:
                one = resolve_presentation(package, presentation,
                                           slide_numbers=[source.slide_number],
                                           east_asian=east_asian, now=now,
                                           locale=options.locale)
                layout = presentation.layouts.get(source.layout_part_path or "")
                master = presentation.masters.get(layout.master_part_path or "") if layout else None
                theme = presentation.themes.get(master.theme_part_path or "") if master else None
                entry = (one.slides[0] if one.slides else None, tuple(one.warnings),
                         _font_scheme(theme) if theme is not None else None)
                found[source.slide_number] = entry
                key = keys.get(source.slide_number)
                if key is not None and self._limit_slides and not one.dated:
                    with self._lock:
                        self._slides[key] = entry
                        while len(self._slides) > self._limit_slides:
                            self._slides.popitem(last=False)
        slides: list[m.Slide] = []
        warnings: list = []
        font_scheme = m.FontScheme()
        for number in numbers:
            entry = found.get(number)
            if entry is None or entry[0] is None:      # no such part: not drawn, as before
                continue
            slides.append(entry[0])
            warnings.extend(entry[1])
            if entry[2] is not None:
                font_scheme = entry[2]
        size = m.SlideSize(width=header.slide_width, height=header.slide_height)
        resolved = ResolvedPresentation(slide_size=size, slides=slides, warnings=warnings,
                                        font_scheme=font_scheme)
        return header, resolved

    def _key(self, package: OpcPackage, settings: str, number: int, slide_id: int | None,
             path: str, digests: dict) -> bytes:
        """The slide's key: what its resolution reads."""
        reach: dict[str, tuple] = {}
        stack = [path]
        while stack:
            part = stack.pop()
            if part in reach:
                continue
            known = digests.get(part)
            if known is None:
                known = (_digest(package.read(part)),
                         _digest(package.read(rels_path_for(part))), package.content_type(part))
                digests[part] = known
            reach[part] = known
            if known[0] == b"-":
                continue
            for relationship in package.relationships(part).values():
                # Another slide is reached only as a hyperlink's target: its own bytes do
                # not draw this one.
                if relationship.is_external or relationship.target_part is None \
                        or relationship.type == REL_SLIDE:
                    continue
                stack.append(relationship.target_part)
        made = repr((settings, number, slide_id, path, sorted(reach.items())))
        return hashlib.blake2b(made.encode(), digest_size=20).digest()
