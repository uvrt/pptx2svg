"""Render PowerPoint (.pptx) slides to SVG, and SVG to PNG.

Parsing, layout and SVG generation use only the standard library; PNG output delegates to
an existing rasteriser (resvg-py by default).

Typical use::

    from pptx2svg import convert_pptx_to_svg, convert_pptx_to_png

    svgs = convert_pptx_to_svg("deck.pptx")           # one SVG string per slide
    pngs = convert_pptx_to_png("deck.pptx", width=1920)

To render a single slide, pass ``slide_numbers=[1]`` (1-based, presentation order).

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

import os
from dataclasses import dataclass, field
from typing import Callable, Sequence

from . import model
from .model import Slide, SlideSize
from .opc import OpcPackage
from .parse.parts import read_presentation
from . import fonts
from .fonts.check import FontReport, check_families, resolved_families
from .png import RasterizerNotAvailable, available_backends, svg_to_png
from .render.context import RenderContext
from .render.svg import render_slide_to_svg
from .resolve import ResolvedPresentation, Warning, resolve_presentation
from .text.fontmap import DEFAULT_FONT_MAPPING, create_font_mapping
from .text.measure import DefaultTextMeasurer, FontToolsTextMeasurer, TextMeasurer

__version__ = "0.1.0"

__all__ = [
    "ConvertOptions",
    "DEFAULT_FONT_MAPPING",
    "DefaultTextMeasurer",
    "FontReport",
    "FontToolsTextMeasurer",
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
    "convert_pptx_to_model",
    "convert_pptx_to_png",
    "convert_pptx_to_svg",
    "create_font_mapping",
    "model",
    "render_slide_to_svg",
    "svg_to_png",
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
    resolved = resolve_presentation(
        package,
        presentation,
        slide_numbers=options.slide_numbers,
        metafile_converter=options.metafile_converter,
    )
    options.warnings.extend(resolved.warnings)
    return resolved


def convert_pptx_to_svg(source, options: ConvertOptions | None = None) -> list[str]:
    """Render a deck to one SVG document per slide."""
    options = options or ConvertOptions()
    resolved = convert_pptx_to_model(source, options)

    font_mapping = create_font_mapping(options.font_mapping)
    jpan_fallback = resolved.font_scheme.major_font_jpan or resolved.font_scheme.minor_font_jpan

    if options.warn_on_font_substitution:
        options.warnings.extend(_font_warnings(resolved))

    documents: list[str] = []
    for slide in resolved.slides:
        # A fresh context per slide keeps ids stable and defs scoped to their document.
        context = RenderContext(
            measurer=options.measurer or DefaultTextMeasurer(),
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
    return documents


def _font_warnings(resolved: ResolvedPresentation) -> list[Warning]:
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
    report = check_families(resolved_families(resolved))

    if report.mode != "bundled":
        names = ", ".join(face.requested for face in report.faces) or "none"
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
    """
    options = options or ConvertOptions()
    documents = convert_pptx_to_svg(source, options)
    return [
        svg_to_png(
            document,
            backend=backend,  # type: ignore[arg-type]
            font_dirs=font_dirs,
            font_files=font_files,
            skip_system_fonts=skip_system_fonts,
            use_bundled_fonts=use_bundled_fonts,
        )
        for document in documents
    ]
