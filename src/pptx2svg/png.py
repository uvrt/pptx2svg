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
"""

from __future__ import annotations

from typing import Iterable, Literal, Sequence

from .fonts import GENERIC_FAMILY_DEFAULTS, font_dirs as bundled_font_dirs

Backend = Literal["resvg", "cairosvg", "auto"]


class RasterizerNotAvailable(RuntimeError):
    """No SVG-to-PNG backend is installed."""


#: Backends in preference order.  resvg first: it takes the font arguments, so it is the
#: only one that renders reproducibly, and it ships prebuilt wheels with no system
#: library behind them.
_BACKEND_MODULES: tuple[tuple[str, str], ...] = (
    ("resvg", "resvg_py"),
    ("cairosvg", "cairosvg"),
)


def _importable(module: str) -> bool:
    """Whether *module* imports, treating **any** failure as "not available".

    Deliberately not ``except ImportError``.  cairosvg is a Python wheel in front of a
    *system* library, and with the wheel installed but libcairo missing -- the normal
    state of a Mac or a slim container that never ran `brew install cairo` -- importing
    it raises ``OSError('no library called "cairo-2" was found')``.  That is not an
    ``ImportError``, so it escaped this probe and took down every caller, including the
    ones that were about to choose resvg and never touch cairo at all.  A backend that
    cannot be imported is unavailable; why it cannot is not this function's business.
    """
    try:
        __import__(module)
    except Exception:
        return False
    return True


def available_backends() -> list[str]:
    """Which rasterisation backends can be imported right now, best first."""
    return [name for name, module in _BACKEND_MODULES if _importable(module)]


def _preferred_backend() -> str | None:
    """The best backend that imports, without importing the others.

    Separate from :func:`available_backends` because ``backend="auto"`` only needs the
    winner, and probing further costs an import of a library we are not going to use --
    which for cairosvg means loading a system library, the slowest and most fragile
    import of the two.
    """
    for name, module in _BACKEND_MODULES:
        if _importable(module):
            return name
    return None


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
) -> bytes:
    """Rasterise an SVG document to PNG bytes.

    ``width``/``height`` set the output size in pixels; give one and resvg keeps the
    aspect ratio.  ``scale`` is an alternative multiplier on the SVG's own size -- pass
    ``scale=2`` for a 2x render.  ``background`` paints a CSS colour behind the image
    (slides already paint their own background, so this is rarely needed).

    Fonts, in the order resvg searches them: ``font_dirs``/``font_files`` first, then the
    bundled families unless ``use_bundled_fonts=False``, then the host's own fonts unless
    ``skip_system_fonts`` is set.  ``skip_system_fonts`` defaults to *whether we have a
    bundle to be reproducible with*: on with ``pptx2svg-fonts`` installed, off without it,
    because skipping system fonts with no bundle would render every slide blank.  Set it
    to ``False`` explicitly to let bundled and installed fonts both take part, which is
    useful when a deck names a face the bundle does not carry.

    Only the resvg backend takes any of this; cairosvg reads the host's fontconfig and
    cannot be pointed at a directory, so it cannot render reproducibly.
    """
    chosen = backend
    if chosen == "auto":
        preferred = _preferred_backend()
        if preferred is None:
            raise RasterizerNotAvailable(
                "no SVG rasterizer installed; run `pip install pptx2svg[png]` for resvg-py "
                "(prebuilt wheels, no system dependencies) or `pip install pptx2svg[cairo]`"
            )
        chosen = preferred  # type: ignore[assignment]

    if chosen == "resvg":
        bundled = bundled_font_dirs() if use_bundled_fonts else []
        if skip_system_fonts is None:
            # Reproducible by default, but only when we actually have fonts to be
            # reproducible *with*: an installation whose package data went missing must
            # fall back to the host rather than render every slide blank.
            skip_system_fonts = bool(bundled)
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
        )
    if chosen == "cairosvg":
        return _render_with_cairosvg(
            svg, width=width, height=height, scale=scale, background=background
        )
    raise ValueError(f"unknown rasterizer backend: {backend!r}")


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
    try:
        import resvg_py
    except ImportError as error:  # pragma: no cover - exercised only without the extra
        raise RasterizerNotAvailable(
            "resvg-py is not installed; run `pip install pptx2svg[png]`"
        ) from error

    options: dict = {"svg_string": svg}
    if width is not None:
        options["width"] = int(width)
    if height is not None:
        options["height"] = int(height)
    if scale is not None:
        options["zoom"] = float(scale)
    if background is not None:
        options["background"] = background
    if font_dirs:
        options["font_dirs"] = list(font_dirs)
    if font_files:
        options["font_files"] = list(font_files)
    if skip_system_fonts:
        options["skip_system_fonts"] = True
    if generic_families:
        # `font-family="..., sans-serif"` is the end of every stack this library emits.
        # resvg resolves those keywords against its own defaults -- "Arial", "Times New
        # Roman", "Courier New" -- none of which we ship, so with system fonts off the
        # keyword resolves to nothing and the text silently does not draw.  Pointing them
        # at the bundle turns "invisible" into "approximately right".
        options.update(generic_families)

    result = resvg_py.svg_to_bytes(**options)
    return _as_bytes(result)


def _as_bytes(value) -> bytes:
    """resvg-py has returned both ``bytes`` and ``list[int]`` across releases."""
    if isinstance(value, (bytes, bytearray)):
        return bytes(value)
    if isinstance(value, Iterable):
        return bytes(value)
    raise TypeError(f"unexpected rasterizer result type: {type(value)!r}")


def _render_with_cairosvg(
    svg: str,
    *,
    width: int | None,
    height: int | None,
    scale: float | None,
    background: str | None,
) -> bytes:
    try:
        import cairosvg
    except ImportError as error:  # pragma: no cover - exercised only without the extra
        raise RasterizerNotAvailable("cairosvg is not installed") from error

    options: dict = {"bytestring": svg.encode("utf-8")}
    if width is not None:
        options["output_width"] = int(width)
    if height is not None:
        options["output_height"] = int(height)
    if scale is not None:
        options["scale"] = float(scale)
    if background is not None:
        options["background_color"] = background

    return cairosvg.svg2png(**options)
