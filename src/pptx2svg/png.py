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
"""

from __future__ import annotations

from typing import Iterable, Literal, Sequence

Backend = Literal["resvg", "cairosvg", "auto"]


class RasterizerNotAvailable(RuntimeError):
    """No SVG-to-PNG backend is installed."""


def available_backends() -> list[str]:
    """Which rasterisation backends can be imported right now."""
    found: list[str] = []
    try:
        import resvg_py  # noqa: F401

        found.append("resvg")
    except ImportError:
        pass
    try:
        import cairosvg  # noqa: F401

        found.append("cairosvg")
    except ImportError:
        pass
    return found


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
    skip_system_fonts: bool = False,
) -> bytes:
    """Rasterise an SVG document to PNG bytes.

    ``width``/``height`` set the output size in pixels; give one and resvg keeps the
    aspect ratio.  ``scale`` is an alternative multiplier on the SVG's own size -- pass
    ``scale=2`` for a 2x render.  ``background`` paints a CSS colour behind the image
    (slides already paint their own background, so this is rarely needed).

    ``font_dirs``/``font_files`` add fonts for text rendering; combine them with
    ``skip_system_fonts=True`` for output that does not depend on the host's font
    inventory.
    """
    chosen = backend
    if chosen == "auto":
        backends = available_backends()
        if not backends:
            raise RasterizerNotAvailable(
                "no SVG rasterizer installed; run `pip install pptx2svg[png]` for resvg-py "
                "(prebuilt wheels, no system dependencies) or `pip install pptx2svg[cairo]`"
            )
        chosen = backends[0]  # type: ignore[assignment]

    if chosen == "resvg":
        return _render_with_resvg(
            svg,
            width=width,
            height=height,
            scale=scale,
            background=background,
            font_dirs=font_dirs,
            font_files=font_files,
            skip_system_fonts=skip_system_fonts,
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
