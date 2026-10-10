"""Render PowerPoint (.pptx) slides to SVG, and SVG to PNG.

Parsing, layout and SVG generation use only the standard library; PNG output delegates to
an existing rasteriser (resvg-py by default).

Typical use::

    from pptx2svg import convert_pptx_to_svg, convert_pptx_to_png

    svgs = convert_pptx_to_svg("deck.pptx")           # one SVG string per slide
    pngs = convert_pptx_to_png("deck.pptx", width=1920)

To render a single slide, pass ``slide_numbers=[1]`` (1-based, presentation order).

For a model to *read* a slide rather than look at it, :func:`convert_pptx_to_agent_svg`
writes a compact view: slide points, shape ids, theme-colour names, text without fonts,
and placeholders for pictures, charts and tables (:mod:`pptx2svg.agent`).

The pipeline is ``.pptx -> OPC package -> source model -> render model -> SVG -> PNG``:

* :mod:`pptx2svg.opc` unpacks the ZIP and resolves relationships;
* :mod:`pptx2svg.parse` reads the OOXML into an unresolved source model;
* :mod:`pptx2svg.resolve` applies theme colours and the placeholder/text inheritance
  cascades, producing a fully-resolved render model;
* :mod:`pptx2svg.render` lays out text and writes the SVG;
* :mod:`pptx2svg.png` rasterises.

Each stage is usable on its own -- call :func:`convert_pptx_to_model` when you want the
resolved slide structure rather than markup.
"""

from __future__ import annotations

import dataclasses
import os
from datetime import datetime
import tempfile
from dataclasses import dataclass, field
from typing import Callable, Sequence

from . import model
from .model import Slide, SlideSize
from .opc import OpcPackage
from .parse.parts import read_presentation
from . import fonts
from .fonts.check import FontReport, check_families, resolved_families
from .fonts.embedded import EmbeddedFonts, extract_embedded_fonts
from .glyphs import MissingGlyphsWarning
from .png import RasterizerNotAvailable, available_backends, svg_to_png
from .render.context import RenderContext
from .render.svg import render_slide_to_svg
from .resolve import ResolvedPresentation, Warning, resolve_presentation
from .resolve.east_asian import EastAsianFaces
from .tables import table_row_heights
from .text.fontmap import DEFAULT_FONT_MAPPING, create_font_mapping, family_key
from .text.measure import DefaultTextMeasurer, FontToolsTextMeasurer, TextMeasurer
from ooxml_common.drawingml.rules import POWERPOINT as POWERPOINT_RULES

__version__ = "0.1.0"

__all__ = [
    "ConvertOptions",
    "DEFAULT_FONT_MAPPING",
    "DefaultTextMeasurer",
    "EmbeddedFonts",
    "FontReport",
    "FontToolsTextMeasurer",
    "MissingGlyphsWarning",
    "OpcPackage",
    "RasterizerNotAvailable",
    "RenderContext",
    "ResolvedPresentation",
    "Slide",
    "SlideSize",
    "TextMeasurer",
    "Warning",
    "available_backends",
    "check_families",
    "convert_pptx_to_agent_svg",
    "convert_pptx_to_model",
    "convert_pptx_to_png",
    "convert_pptx_to_svg",
    "create_font_mapping",
    "model",
    "render_slide_to_svg",
    "svg_to_png",
    "table_row_heights",
    "user_font_dirs",
]


@dataclass
class ConvertOptions:
    """Knobs shared by the SVG and PNG entry points."""

    #: 1-based slide numbers in presentation order; ``None`` renders every slide.
    slide_numbers: Sequence[int] | None = None
    #: Output width in pixels.  The slide's own size is used when both are ``None``.
    width: int | None = None
    height: int | None = None
    #: PPTX font name -> replacement font name, merged over :data:`DEFAULT_FONT_MAPPING`.
    font_mapping: dict[str, str] | None = None
    #: Override text measurement (see :class:`~pptx2svg.text.measure.FontToolsTextMeasurer`).
    measurer: TextMeasurer | None = None
    #: Convert an EMF/WMF metafile to something embeddable, as
    #: ``(payload, mime_type) -> (payload, mime_type) | None``.  Opt-in: without it the
    #: pure-Python path extracts the preview Office embedded in the metafile, which
    #: covers most real files.  Supply one to shell out to Inkscape or ``libemf2svg``
    #: for the rest -- returning ``image/svg+xml`` embeds the vectors directly.  Return
    #: ``None`` to decline a particular file and fall back to the built-in path.
    metafile_converter: Callable[[bytes, str], "tuple[bytes, str] | None"] | None = None
    #: Collects warnings for unsupported content; also returned by
    #: :func:`convert_pptx_to_model`.
    warnings: list[Warning] = field(default_factory=list)
    #: Raise a ``font-substituted`` warning for every face the render cannot draw at the
    #: widths it measured.  On by default, because a substitution nobody is told about is
    #: how this library shipped wrong layout for months.
    warn_on_font_substitution: bool = True
    #: Use the faces a deck carries in ``<p:embeddedFontLst>``, for both measurement and
    #: drawing.  On by default: the deck's own font is what PowerPoint drew with, so it
    #: beats any substitute.  Turn it off to reproduce pre-0.2 output, or when a deck's
    #: embedded fonts are known to be damaged -- decoding costs about a second per face.
    use_embedded_fonts: bool = True
    #: Measure (and draw) with the faces PowerPoint would use on this machine -- its own
    #: bundle, macOS's fonts, Office's cloud-font cache, read in place
    #: (:mod:`pptx2svg.fonts.office`) -- for every family the static tables do not
    #: measure as itself.  ``None`` (the default) does so whenever Office's fonts are on
    #: this machine (:func:`pptx2svg.fonts.office.available`), with or without the font
    #: bundle, because that most closely resembles PowerPoint; the output then depends on
    #: the machine.  Elsewhere it does so when the PNG is drawn from this machine's fonts
    #: anyway: without the font bundle, or with ``skip_system_fonts=False``.  ``False``
    #: (or ``PPTX2SVG_OFFICE_FONTS=0``) is the reproducible render: no host face is read.
    host_fonts: bool | None = None
    #: Check, before rasterising, that some font the rasteriser loads has a glyph for every
    #: character of the slide's text, and say which face and script it lacks: a
    #: ``glyphs-missing`` warning here and a :class:`~pptx2svg.glyphs.MissingGlyphsWarning`
    #: (:mod:`pptx2svg.glyphs`).  resvg draws nothing for such a character, silently --
    #: every CJK character of a deck, rendered without ``pptx2svg-fonts`` on a host with no
    #: CJK face.  PNG output only: an SVG viewer draws with its own fonts.
    check_glyphs: bool = True
    #: The application's own font folders, searched before every other: a face in them
    #: is drawn from its file and, where the static tables do not measure its family as
    #: itself, measured from it too (:func:`pptx2svg.fonts.office.user_layout_metrics`),
    #: whatever ``host_fonts`` says.  ``None`` (the default) reads ``OOXML_FONT_DIRS``
    #: (``os.pathsep``-separated); an empty list means none, the variable not read.
    #: Added to the system's folders, never in their place.  ``convert_pptx_to_png``'s
    #: own ``font_dirs`` argument, when given, takes precedence.
    font_dirs: Sequence[str] | None = None
    #: The clock a date field shows (``a:fld type="datetime1"`` .. ``"datetime13"``, as in
    #: a footer's date), formatted in the field's language
    #: (:mod:`pptx2svg.resolve.fields`).  ``None`` (the default) is the local time of the
    #: conversion, as PowerPoint shows the time it draws -- so output holding such a field
    #: differs from day to day.  Pin it (``datetime(2026, 1, 31, 9, 0)``) for output that
    #: must be reproducible: tests, snapshot goldens, a cache keyed by the deck.  A naive
    #: value is used as it stands; an aware one in its own time zone.
    now: datetime | None = None


def user_font_dirs(options: "ConvertOptions | None" = None, font_dirs=None) -> list[str]:
    """The application's font folders a conversion uses, as strings: ``font_dirs`` when
    given, else ``options.font_dirs``, else ``OOXML_FONT_DIRS``
    (:func:`ooxml_common.fonts.office.user_font_dirs`)."""
    from ooxml_common.fonts.office import user_font_dirs as _user_font_dirs

    if font_dirs is None and options is not None:
        font_dirs = options.font_dirs
    return [str(path) for path in _user_font_dirs(font_dirs)]


def _host_fonts(options: ConvertOptions) -> bool:
    """Whether this conversion measures (and draws) with the host's faces: as asked, or,
    left to the default, wherever Office's fonts are installed, and otherwise when the
    host's fonts draw the PNG anyway (no font bundle)."""
    if options.host_fonts is not None:
        return options.host_fonts
    from .fonts import office

    return office.available() or fonts.bundle_mode() == "system"


def _open_package(source) -> OpcPackage:
    if isinstance(source, OpcPackage):
        return source
    if isinstance(source, (str, os.PathLike)):
        return OpcPackage.open(os.fspath(source))
    return OpcPackage.open(source)


def convert_pptx_to_model(
    source, options: ConvertOptions | None = None
) -> ResolvedPresentation:
    """Parse and resolve a deck without rendering it.

    Returns the render model: slide size, one :class:`~pptx2svg.model.Slide` per slide
    with fully-resolved colours, fonts and geometry, and any warnings raised on the way.
    """
    options = options or ConvertOptions()
    package = _open_package(source)
    presentation = read_presentation(package)
    host_fonts = _host_fonts(options)
    resolved = resolve_presentation(
        package,
        presentation,
        slide_numbers=options.slide_numbers,
        metafile_converter=options.metafile_converter,
        # Which face draws a run's Japanese depends on what is installed: as PowerPoint
        # would find it here when the host's faces are in use, else from what we know.
        east_asian=EastAsianFaces(host=host_fonts),
        now=options.now,
    )
    if options.use_embedded_fonts and presentation.embedded_fonts:
        # After resolution, not during it: `resolved_families` is what tells us which of
        # the embedded families a slide actually asks for, and decoding one costs about a
        # second.  A template deck routinely embeds a family that no slide uses.
        #
        # **An installed face beats the embedded one**, measured: PowerPoint draws a
        # deck's embedded Lato 1.104 with the Lato 2.015 Office's cloud cache holds --
        # glyph origins 0.12 pt from 2.015's advances over a 20-character line, 0.55 pt
        # from 1.104's -- and still does when the embedded copy claims version 9.000; a
        # face it embeds under a name nothing installs is drawn from the deck
        # (``tools/make_font_resolution_probe.py``).  So with the host's faces in use, a
        # family PowerPoint would find installed here is not taken from the deck at all.
        wanted = resolved_families(resolved)
        if host_fonts:
            from .fonts import office

            embedded = {family_key(entry.typeface or "") for entry in presentation.embedded_fonts}
            installed = {family for family in wanted if family_key(family) in embedded and office.find(family)}
            resolved.installed_families = frozenset(family_key(family) for family in installed)
            wanted = [family for family in wanted if family not in installed]
        resolved.embedded_fonts = extract_embedded_fonts(
            package,
            presentation.embedded_fonts,
            wanted_families=wanted,
        )
        resolved.warnings.extend(
            Warning(code=code, message=message, part_path=presentation.part_path)
            for code, message in resolved.embedded_fonts.problems
        )
    options.warnings.extend(resolved.warnings)
    return resolved


def convert_pptx_to_svg(source, options: ConvertOptions | None = None) -> list[str]:
    """Render a deck to one SVG document per slide."""
    return _render(source, options or ConvertOptions())[0]


def _render(source, options: ConvertOptions) -> "tuple[list[str], ResolvedPresentation]":
    """The SVG pass, keeping the resolved model.

    ``convert_pptx_to_png`` needs it: the deck's embedded faces have to be written to
    disk for the rasteriser, and they live on the resolved presentation.  Re-parsing to
    get them would decode every font a second time.
    """
    resolved = convert_pptx_to_model(source, options)

    font_mapping = create_font_mapping(options.font_mapping)
    # No script-list fallback for text: a run's East Asian face is decided at resolution,
    # and the theme's `Jpan` entry reaches a run only as `+mn-ea` for Japanese text --
    # measured (pptx2svg.resolve.east_asian).  Offering it again behind every East Asian
    # chunk put 游ゴシック into stacks PowerPoint never draws from.
    jpan_fallback = None

    user_dirs = user_font_dirs(options)
    if options.warn_on_font_substitution:
        options.warnings.extend(_font_warnings(resolved, user_dirs))

    # The embedded faces have to reach *measurement*, not only the rasteriser: by the
    # time `svg_to_png` sees a font file every line has already been wrapped, autofitted
    # and centred.  Handing them to the measurer here is what keeps measure-equals-draw
    # true for a face the deck brought with it.  A caller-supplied measurer wins, because
    # it was asked for explicitly.
    #
    # The faces PowerPoint itself would use on this machine come next, under the deck's
    # own: a family the static tables only guess at, or measure from another face, is
    # measured from the installed file the PNG will be drawn with.
    #
    # Kerned as PowerPoint kerns (``DrawingRules.kerning``): a static face's legacy
    # ``kern`` table, a variable face's GPOS pairs -- measured, and not the OpenType
    # feature the tables used to charge everywhere (ooxml_common.text.kerning).
    #
    # The application's own folders (``font_dirs``, ``OOXML_FONT_DIRS``) come before the
    # host's: the rasteriser is handed them ahead of every installed face.
    measurer = options.measurer
    if measurer is None:
        extra = dict(resolved.embedded_fonts.metrics)
        if user_dirs:
            from .fonts import office

            for key, table in office.user_layout_metrics(resolved_families(resolved), user_dirs).items():
                extra.setdefault(key, table)
        if _host_fonts(options):
            from .fonts import office

            for key, table in office.layout_metrics(resolved_families(resolved)).items():
                extra.setdefault(key, table)
        measurer = DefaultTextMeasurer(extra, kerning=POWERPOINT_RULES.kerning)

    documents: list[str] = []
    for slide in resolved.slides:
        # A fresh context per slide keeps ids stable and defs scoped to their document.
        context = RenderContext(
            measurer=measurer,
            font_mapping=font_mapping,
            jpan_fallback_font=jpan_fallback,
        )
        documents.append(
            render_slide_to_svg(
                slide,
                resolved.slide_size,
                context,
                width=options.width,
                height=options.height,
            )
        )
    return documents, resolved


def _font_warnings(resolved: ResolvedPresentation, user_dirs: Sequence[str] = ()) -> list[Warning]:
    """Say out loud what the rasteriser would otherwise do silently.

    Deliberately warnings rather than errors: a deck that names Aptos still renders, and
    refusing to render it would help nobody.  What was missing was any signal at all --
    resvg substitutes without a word, so wrong output looked exactly like right output.
    ``pptx2svg fonts --check`` is the same information with an exit status.

    Two shapes, because two different things go wrong.  Without the font bundle *nothing*
    is reproducible and every face would report the same cause, so that is one warning,
    not one per face.  With the bundle, the remaining gaps are per-face and worth naming
    individually.
    """
    embedded = resolved.embedded_fonts.families | resolved.installed_families
    supplied = frozenset()
    if user_dirs:
        from ooxml_common.fonts.office import user_families

        supplied = user_families(user_dirs)
    report = check_families(resolved_families(resolved), embedded=embedded, supplied=supplied)

    if report.mode != "bundled":
        unsupplied = [face.requested for face in report.faces if not face.faithful]
        if not unsupplied:
            # Every face this deck draws with came out of the deck itself, so the bundle
            # has nothing left to supply and saying it is missing would be noise.
            return []
        names = ", ".join(unsupplied)
        return [
            Warning(
                code="font-bundle-missing",
                message=(
                    "no font bundle installed: PNG output will use this machine's fonts "
                    "and will not be reproducible elsewhere. Run "
                    f"`{fonts.INSTALL_HINT}`, or pass font_dirs= explicitly. "
                    f"Faces this deck needs: {names}"
                ),
            )
        ]

    return [
        Warning(code="font-substituted", message=f"{face.requested}: {face.reason}")
        for face in report.faces
        if not face.faithful
    ]


def convert_pptx_to_png(
    source,
    options: ConvertOptions | None = None,
    *,
    backend: str = "auto",
    font_dirs: Sequence[str] | None = None,
    font_files: Sequence[str] | None = None,
    skip_system_fonts: bool | None = None,
    use_bundled_fonts: bool = True,
) -> list[bytes]:
    """Render a deck to one PNG image per slide.

    Needs a rasteriser: ``pip install pptx2svg[png]`` for resvg-py (prebuilt wheels), or
    ``pip install pptx2svg[cairo]``.  ``width``/``height`` on ``options`` set the output size.

    By default this draws with the fonts bundled in :mod:`pptx2svg.fonts` and ignores the
    host's, so the same deck rasterises to the same bytes on a laptop and on a Debian
    box.  ``use_bundled_fonts=False`` (or ``skip_system_fonts=False``) opts back into
    whatever the machine happens to have installed, which is faster to set up and
    impossible to reproduce.  Either way, faces the render could not draw faithfully are
    appended to ``options.warnings``.

    ``font_dirs`` -- the application's own folders, handed to the rasteriser ahead of
    the bundle and the system's fonts, and measured from as ``options.font_dirs`` is --
    takes precedence over ``options.font_dirs``; without either, ``OOXML_FONT_DIRS`` is
    read.
    """
    options = options or ConvertOptions()
    if font_dirs is not None:
        # Measured as drawn: the folders the rasteriser is given are the ones measured.
        options = dataclasses.replace(options, font_dirs=list(font_dirs))
    font_dirs = user_font_dirs(options)
    if options.host_fonts is None and skip_system_fonts is not None:
        # The rasteriser is told explicitly whether to read this machine's fonts, so
        # measure as it will draw.
        # A shallow copy: `warnings` is still the caller's own list.
        options = dataclasses.replace(options, host_fonts=not skip_system_fonts)
    documents, resolved = _render(source, options)
    host_fonts = _host_fonts(options)
    numbers = [slide.slide_number for slide in resolved.slides]

    if not resolved.embedded_fonts:
        return _rasterise(
            documents, backend, font_dirs, font_files, skip_system_fonts, use_bundled_fonts,
            host_fonts, options, numbers,
        )

    # The rasteriser's font database indexes files, so the extracted faces have to touch
    # disk.  That is not a workaround -- it is what LibreOffice does for the same reason
    # (`EmbeddedFontsHelper::addEmbeddedFont` writes a temp file, then `AddTempDevFont`).
    # The directory lives exactly as long as the rasterisation that needs it.
    with tempfile.TemporaryDirectory(prefix="pptx2svg-fonts-") as directory:
        extracted = resolved.embedded_fonts.write(directory)
        return _rasterise(
            documents,
            backend,
            font_dirs,
            [*extracted, *(font_files or ())],
            skip_system_fonts,
            use_bundled_fonts,
            host_fonts,
            options,
            numbers,
        )


def _rasterise(
    documents, backend, font_dirs, font_files, skip_system_fonts, use_bundled_fonts,
    host_fonts, options: ConvertOptions, numbers: Sequence[int | None],
) -> list[bytes]:
    """Each slide's PNG; text no loaded font can draw reported once per face and script
    for the whole deck, on the first slide that has it (:mod:`pptx2svg.glyphs`)."""
    import warnings

    from .glyphs import MissingGlyphsWarning
    from .png import _svg_to_png

    images: list[bytes] = []
    reported: set[tuple[str, str]] = set()
    for document, number in zip(documents, numbers):
        png, missing = _svg_to_png(
            document,
            width=None,
            height=None,
            scale=None,
            background=None,
            backend=backend,  # type: ignore[arg-type]
            font_dirs=font_dirs,
            font_files=font_files,
            skip_system_fonts=skip_system_fonts,
            use_bundled_fonts=use_bundled_fonts,
            host_fonts=host_fonts,
            check_glyphs=options.check_glyphs,
        )
        images.append(png)
        for item in missing:
            if (item.face, item.script) in reported:
                continue
            reported.add((item.face, item.script))
            options.warnings.append(
                Warning(code="glyphs-missing", message=item.message(), slide_number=number,
                        detail=item)
            )
            warnings.warn(item.message(), MissingGlyphsWarning, stacklevel=3)
    return images


from .agent import convert_pptx_to_agent_svg  # noqa: E402  (needs the names above)
