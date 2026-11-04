"""SVG to PNG rasterisation.

Rasterising is delegated to an existing renderer rather than reimplemented.  Two backends
are supported and tried in this order:

* **resvg-py** (default) -- Python bindings for `resvg <https://github.com/linebender/resvg>`_,
  the same Rust renderer pptx-glimpse rasterises with.  It ships prebuilt wheels, so
  ``pip install pptx2svg[png]`` needs no compiler and no system libraries, and it handles
  the filter/pattern/marker features this library emits.
* **cairosvg** -- fallback for environments that already have Cairo.  Note that Cairo's
  SVG filter support is weaker, so shadows and glows may differ.

Install one with ``pip install pptx2svg[png]`` (resvg) or ``pip install pptx2svg[cairo]``.

**Fonts default to the bundle when there is one.**  A rasteriser reading the host's font
list makes the same deck rasterise differently on every machine, and worse, does it
quietly: resvg substitutes a missing face without a word, so text appears at widths
nothing computed.  So when ``pptx2svg-fonts`` is installed, :func:`svg_to_png` renders
from it with ``skip_system_fonts`` on, and points the generic families at it as well --
with system fonts off, a family resvg cannot resolve draws *nothing*, and the generic
keyword at the end of every stack this library emits is what stops a blank slide.

When it is not installed there is nothing to be deterministic with, so rendering falls
back to the host's fonts.  That is a real downgrade and it is announced rather than
inferred: :func:`pptx2svg.convert_pptx_to_svg` puts a ``font-bundle-missing`` warning in
``ConvertOptions.warnings``, and ``pptx2svg fonts`` says so at the top of its report.

**Pictures are drawn in the colours their profile says.**  Neither rasteriser honours a
picture's embedded ICC profile, and PowerPoint does, so a picture carrying one other than
sRGB is converted before it is handed over (:mod:`pptx2svg.iccimages`); the SVG itself
keeps the picture as it came, profile and all, for a colour-managed reader to convert.
"""

from __future__ import annotations

import warnings
from typing import Sequence

from ooxml_common import raster
from ooxml_common.metafile import inline_metafile_images
from ooxml_common.raster import (  # noqa: F401  (this module's public names, kept)
    Backend,
    RasterizerNotAvailable,
    available_backends,
)
from ooxml_common.svgimage import with_pixel_sized_images

from .fonts import GENERIC_FAMILY_DEFAULTS, font_dirs as bundled_font_dirs
from .glyphs import MissingGlyphs, MissingGlyphsWarning, missing_glyphs
from .iccimages import srgb_images

# Probing, the backends' order, the call to each and the passes are ooxml-common's
# (:mod:`ooxml_common.raster`), shared with docx2svg; the names this module had stay.
_BACKEND_MODULES = raster.BACKENDS
_importable = raster.importable
_preferred_backend = raster.preferred_backend
_as_bytes = raster.as_bytes


def svg_to_png(
    svg: str,
    *,
    width: int | None = None,
    height: int | None = None,
    scale: float | None = None,
    background: str | None = None,
    backend: Backend = "auto",
    font_dirs: Sequence[str] | None = None,
    font_files: Sequence[str] | None = None,
    skip_system_fonts: bool | None = None,
    use_bundled_fonts: bool = True,
    host_fonts: bool | None = None,
    check_glyphs: bool = True,
) -> bytes:
    """Rasterise an SVG document to PNG bytes.

    ``width``/``height`` set the output size in pixels; give one and resvg keeps the
    aspect ratio.  ``scale`` is an alternative multiplier on the SVG's own size -- pass
    ``scale=2`` for a 2x render.  ``background`` paints a CSS colour behind the image
    (slides already paint their own background, so this is rarely needed).

    Fonts, in the order resvg searches them: ``font_dirs``/``font_files`` first, then the
    bundled families unless ``use_bundled_fonts=False``, then the host's own fonts unless
    ``skip_system_fonts`` is set.  ``font_dirs`` left out reads ``OOXML_FONT_DIRS``
    (``os.pathsep``-separated); an empty list means none.  ``skip_system_fonts`` defaults to *whether we have a
    bundle to be reproducible with*: on with ``pptx2svg-fonts`` installed, off without it,
    because skipping system fonts with no bundle would render every slide blank.  Set it
    to ``False`` explicitly to let bundled and installed fonts both take part, which is
    useful when a deck names a face the bundle does not carry.

    ``host_fonts`` draws each family the SVG names with the face PowerPoint would use on
    this machine -- its own bundle's, macOS's or Office's cloud-font cache's, read where
    it is installed (:mod:`pptx2svg.fonts.office`) -- ahead of everything but
    ``font_files`` and ``font_dirs``.  It defaults to on wherever Office's fonts are
    installed (:func:`pptx2svg.fonts.office.available`), bundle or not, and elsewhere to
    whether this machine's fonts take part at all (``not skip_system_fonts``); where the
    folders are absent it changes nothing.  ``host_fonts=False`` (or
    ``PPTX2SVG_OFFICE_FONTS=0``) keeps a render reproducible.
    :func:`pptx2svg.convert_pptx_to_svg` measures with the same faces under the same
    condition (``ConvertOptions.host_fonts``).

    Only the resvg backend takes any of this; cairosvg reads the host's fontconfig and
    cannot be pointed at a directory, so it cannot render reproducibly.

    Text no loaded font has glyphs for is not drawn at all, so ``check_glyphs`` (on by
    default) raises a :class:`~pptx2svg.glyphs.MissingGlyphsWarning` for it, once per face
    and script (:mod:`pptx2svg.glyphs`).
    """
    from ooxml_common.fonts.office import user_font_dirs

    png, missing = _svg_to_png(
        svg, width=width, height=height, scale=scale, background=background,
        backend=backend, font_dirs=[str(path) for path in user_font_dirs(font_dirs)],
        font_files=font_files,
        skip_system_fonts=skip_system_fonts, use_bundled_fonts=use_bundled_fonts,
        host_fonts=host_fonts, check_glyphs=check_glyphs,
    )
    for item in missing:
        warnings.warn(item.message(), MissingGlyphsWarning, stacklevel=2)
    return png


def _svg_to_png(
    svg: str,
    *,
    width: int | None,
    height: int | None,
    scale: float | None,
    background: str | None,
    backend: Backend,
    font_dirs: Sequence[str] | None,
    font_files: Sequence[str] | None,
    skip_system_fonts: bool | None,
    use_bundled_fonts: bool,
    host_fonts: bool | None,
    check_glyphs: bool,
) -> "tuple[bytes, list[MissingGlyphs]]":
    """:func:`svg_to_png`, the text no loaded font can draw returned rather than warned
    about, so a caller rasterising many slides reports each face and script once."""
    # resvg and cairosvg both read a picture's samples as sRGB whatever profile it carries;
    # PowerPoint converts it.  See :mod:`pptx2svg.iccimages`.
    svg = srgb_images(svg)
    # An SVG picture's root size in units other than px (`width="72pt"`) makes resvg drop
    # the picture, and CairoSVG draws one with no viewBox unscaled; both get it restated in
    # pixels as PowerPoint reads it (:mod:`ooxml_common.svgimage`).
    # A metafile drawn as SVG (ooxml_common.metafile) goes inline: resvg draws no text in an
    # SVG it reads through <image>.
    svg = inline_metafile_images(svg)
    svg = with_pixel_sized_images(svg)

    chosen = raster.choose_backend(
        backend,
        none_installed="no SVG rasterizer installed; run `pip install pptx2svg[png]` for resvg-py "
        "(prebuilt wheels, no system dependencies) or `pip install pptx2svg[cairo]`",
    )

    if chosen == "resvg":
        svg, plan, font_files, skip_system_fonts, bundled = _resvg_fonts(
            svg, font_dirs=font_dirs, font_files=font_files, skip_system_fonts=skip_system_fonts,
            use_bundled_fonts=use_bundled_fonts, host_fonts=host_fonts)
        if plan is not None:
            missing = _missing(
                check_glyphs, svg,
                [*(font_files or ()), *(path for files in plan.passes for path in files)],
                list(font_dirs or ()) + bundled, not skip_system_fonts,
                GENERIC_FAMILY_DEFAULTS if bundled else None,
            )
            return _render_plan(
                svg,
                plan,
                width=width,
                height=height,
                scale=scale,
                background=background,
                font_dirs=list(font_dirs or ()) + bundled,
                font_files=list(font_files or ()),
                skip_system_fonts=skip_system_fonts,
                generic_families=GENERIC_FAMILY_DEFAULTS if bundled else None,
                own_dirs=list(font_dirs or ()),
            ), missing
        missing = _missing(check_glyphs, svg, list(font_files or ()),
                           list(font_dirs or ()) + bundled, not skip_system_fonts,
                           GENERIC_FAMILY_DEFAULTS if bundled else None)
        return _render_with_resvg(
            svg,
            width=width,
            height=height,
            scale=scale,
            background=background,
            font_dirs=list(font_dirs or ()) + bundled,
            font_files=font_files,
            skip_system_fonts=skip_system_fonts,
            generic_families=GENERIC_FAMILY_DEFAULTS if bundled else None,
        ), missing
    if chosen == "cairosvg":
        # cairosvg draws from the host's fontconfig, and from nothing else.
        missing = _missing(check_glyphs, svg, [], [], True, None)
        return _render_with_cairosvg(
            svg, width=width, height=height, scale=scale, background=background
        ), missing
    raise ValueError(f"unknown rasterizer backend: {backend!r}")  # pragma: no cover


def _resvg_fonts(svg: str, *, font_dirs, font_files, skip_system_fonts: bool | None, use_bundled_fonts: bool,
                 host_fonts: bool | None):
    """What resvg is handed for ``svg``: ``(svg, plan, font_files, skip_system_fonts,
    bundle folders)`` -- the SVG with the emoji its faces cannot draw named in the emoji
    face, the faces' plan of passes (:func:`pptx2svg.fonts.office.drawing_plan`, ``None``
    for one pass of what is handed over), the files and whether the system's fonts are
    left out.  :func:`_svg_to_png` draws with it, :func:`outline_svg` outlines with it."""
    bundled = bundled_font_dirs() if use_bundled_fonts else []
    if skip_system_fonts is None:
        # Reproducible by default, but only when we actually have fonts to be
        # reproducible *with*: an installation whose package data went missing must
        # fall back to the host rather than render every slide blank.
        skip_system_fonts = bool(bundled)
    if host_fonts is None:
        from .fonts import office

        host_fonts = office.available() or not skip_system_fonts
    plan = None
    if host_fonts:
        from .fonts import office

        # Emoji the run's face cannot draw, in macOS's colour emoji face, as
        # PowerPoint draws them (:func:`pptx2svg.fonts.office.with_emoji`).  Last
        # among the files, so it shadows nothing.
        svg, emoji = office.with_emoji(
            svg, supplied=office.supplied_faces(font_files, font_dirs)
        )
        font_files = [*(font_files or ()), *emoji]
        own = office.supplied_faces(font_files, font_dirs)
        plan = office.drawing_plan(svg, supplied=frozenset(own), own=own)
    elif font_dirs:
        from .fonts import office

        # The caller's folders alone: faces in them filed under one name (Aptos and
        # Aptos Display) are drawn apart, as the installed ones are.
        plan = office.drawing_plan(
            svg, own=office.supplied_faces(font_files, font_dirs), host=False
        )
    return svg, plan, font_files, skip_system_fonts, bundled


def outline_svg(
    svg: str,
    *,
    font_dirs: Sequence[str] | None = None,
    font_files: Sequence[str] | None = None,
    skip_system_fonts: bool | None = None,
    use_bundled_fonts: bool = True,
    host_fonts: bool | None = None,
    selectable: bool = False,
    glyphs: "dict[str, str] | None" = None,
):
    """``svg`` with its text drawn as glyph outlines: the faces, fallbacks, shaping and
    positions resvg draws it with when :func:`svg_to_png` is given the same arguments
    (:func:`ooxml_common.outlines.outline_text`), so the SVG needs no font to be viewed.
    ``selectable`` keeps each text, transparent, over its outlines; ``glyphs``, a dict
    shared by a deck's slides, takes the glyph definitions in place of the SVG's own
    ``<defs>`` (:func:`ooxml_common.outlines.outline_text`).  Returns an
    :class:`ooxml_common.outlines.Outlined`: the SVG and the texts kept as text, with why.

    ``font_dirs`` left out reads ``OOXML_FONT_DIRS``, as :func:`svg_to_png` does."""
    from ooxml_common.fonts.office import user_font_dirs
    from ooxml_common.outlines import FontSetup, outline_text

    # A metafile drawn as SVG goes inline, as svg_to_png draws it: its text is outlined too.
    svg = inline_metafile_images(svg)
    dirs = [str(path) for path in user_font_dirs(font_dirs)]
    svg, plan, files, skip, bundled = _resvg_fonts(
        svg, font_dirs=dirs, font_files=font_files, skip_system_fonts=skip_system_fonts,
        use_bundled_fonts=use_bundled_fonts, host_fonts=host_fonts)
    generic = GENERIC_FAMILY_DEFAULTS if bundled else None
    all_dirs = dirs + list(bundled)
    if plan is None:
        setup = FontSetup.of(list(files or ()), all_dirs, not skip, generic)
        return outline_text(svg, setup, selectable=selectable, glyphs=glyphs)
    setups = [FontSetup.of(*_plan_pass_fonts(plan, number, all_dirs, list(files or ()), skip, dirs), generic)
              for number in range(len(plan.passes))]
    import html

    stacks = {html.unescape(value): number for value, number in plan.stacks.items()}
    return outline_text(svg, lambda value: setups[stacks.get(value or "", 0)], selectable=selectable,
                        glyphs=glyphs)


def _missing(check: bool, svg: str, font_files: list, font_dirs: list,
             system_fonts: bool, generic_families) -> "list[MissingGlyphs]":
    """What of ``svg``'s text the rasteriser will not draw (:mod:`.glyphs`)."""
    if not check:
        return []
    return missing_glyphs(svg, font_files=[str(path) for path in font_files],
                          font_dirs=[str(path) for path in font_dirs],
                          system_fonts=system_fonts, generic_families=generic_families)


def _plan_pass_fonts(plan, number: int, font_dirs: list[str], font_files: list[str], skip_system_fonts: bool,
                     own_dirs: "list[str] | None" = None) -> "tuple[list[str], list[str], bool]":
    """``(font_files, font_dirs, system fonts loaded)`` of pass ``number`` of ``plan``:
    what :func:`_render_plan` hands resvg for it."""
    from .fonts import office

    system_dirs: list[str] = []
    skip = skip_system_fonts
    if plan.ahead_of_system and not skip_system_fonts:
        skip, system_dirs = True, office.resvg_system_dirs()
    own_dirs = list(own_dirs or ())
    files = plan.passes[number]
    dirs, extra = font_dirs, []
    withheld = set(plan.withheld[number]) if plan.withheld else set()
    if withheld and own_dirs and font_dirs[:len(own_dirs)] == own_dirs:
        from ooxml_common.fonts.rasterdb import load_order

        dirs = font_dirs[len(own_dirs):]
        extra = [path for path in load_order((), own_dirs) if path not in withheld]
    return font_files + files + extra, dirs + system_dirs, not skip


def _render_plan(
    svg: str,
    plan,
    *,
    width: int | None,
    height: int | None,
    scale: float | None,
    background: str | None,
    font_dirs: list[str],
    font_files: list[str],
    skip_system_fonts: bool,
    generic_families: dict[str, str] | None,
    own_dirs: "list[str] | None" = None,
) -> bytes:
    """Draw ``svg`` with the host faces ``plan`` names (:func:`pptx2svg.fonts.office.drawing_plan`).

    resvg keeps the first face it loads among those answering to the same name, style,
    weight and width, and loads ``font_files`` before ``font_dirs`` and the system's
    fonts before either -- so the faces go in as files right behind the caller's own,
    and where one of them shares its name with a system face (PowerPoint's Rockwell and
    Symbol are not macOS's) the system's folders are handed over as directories after
    them instead of being loaded first.  Faces that would shadow one another (Aptos
    Display and Aptos) are drawn in separate passes, each over the last
    (:func:`ooxml_common.raster.draw_in_passes`).  Where a pass leaves faces of the caller's
    folders out (:attr:`~pptx2svg.fonts.office.DrawingPlan.withheld`), those folders --
    ``own_dirs``, the first of ``font_dirs`` -- are handed over as their files, in the
    order resvg loads them, without those.
    """

    def draw(number: int, document: str) -> bytes:
        files, dirs, system = _plan_pass_fonts(plan, number, font_dirs, font_files, skip_system_fonts, own_dirs)
        return _render_with_resvg(
            document,
            width=width,
            height=height,
            scale=scale,
            background=background if number == 0 else None,
            font_dirs=dirs,
            font_files=files,
            skip_system_fonts=not system,
            generic_families=generic_families,
        )

    return raster.draw_in_passes(svg, plan.stacks, len(plan.passes), draw)


def _render_with_resvg(
    svg: str,
    *,
    width: int | None,
    height: int | None,
    scale: float | None,
    background: str | None,
    font_dirs: Sequence[str] | None,
    font_files: Sequence[str] | None,
    skip_system_fonts: bool,
    generic_families: dict[str, str] | None = None,
) -> bytes:
    """resvg's call (:func:`ooxml_common.raster.render_with_resvg`).  ``generic_families``
    points the generic keywords at the bundle: ``font-family="..., sans-serif"`` ends
    every stack this library emits, and resvg resolves the keywords against its own
    defaults -- "Arial", "Times New Roman", "Courier New" -- none of which we ship, so with
    system fonts off the keyword would resolve to nothing and the text silently not draw.
    Pointing them at the bundle turns "invisible" into "approximately right"."""
    return raster.render_with_resvg(
        svg, width=width, height=height, scale=scale, background=background,
        font_dirs=font_dirs, font_files=font_files, skip_system_fonts=skip_system_fonts,
        families=generic_families,
        not_installed="resvg-py is not installed; run `pip install pptx2svg[png]`",
    )


def _render_with_cairosvg(
    svg: str,
    *,
    width: int | None,
    height: int | None,
    scale: float | None,
    background: str | None,
) -> bytes:
    return raster.render_with_cairosvg(svg, width=width, height=height, scale=scale,
                                       background=background)
