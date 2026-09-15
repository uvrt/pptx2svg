"""Command-line interface.

Two entry points share one executable::

    pptx2svg deck.pptx -o out/        # render
    pptx2svg fonts --check deck.pptx  # will it render faithfully?

``fonts`` is spelled as a leading word rather than a real subparser because the render
form takes its input file positionally and has done since the first release; turning it
into ``pptx2svg render deck.pptx`` would break every existing caller.  So the first
argument is peeked at, and only the literal word ``fonts`` diverts.  A file actually
named ``fonts`` is still reachable as ``./fonts``.
"""

from __future__ import annotations

import argparse
import contextlib
import sys
import tempfile
from pathlib import Path

from . import ConvertOptions, __version__, _render
from .png import RasterizerNotAvailable, available_backends, svg_to_png


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx2svg",
        description="Render PowerPoint (.pptx) slides to SVG or PNG.",
    )
    parser.add_argument("input", type=Path, help="path to a .pptx file")
    parser.add_argument(
        "-o",
        "--output",
        type=Path,
        default=Path("."),
        help="output directory (default: current directory)",
    )
    parser.add_argument(
        "-f",
        "--format",
        choices=("svg", "png", "both"),
        default="svg",
        help="output format (default: svg)",
    )
    parser.add_argument(
        "-s",
        "--slides",
        help="slide numbers to render, e.g. '1', '1,3', '2-5' (default: all)",
    )
    parser.add_argument("--width", type=int, help="output width in pixels")
    parser.add_argument("--height", type=int, help="output height in pixels")
    parser.add_argument(
        "--backend",
        choices=("auto", "resvg", "cairosvg"),
        default="auto",
        help="PNG rasterizer backend (default: auto)",
    )
    parser.add_argument(
        "--font-dir",
        action="append",
        dest="font_dirs",
        metavar="DIR",
        help="directory of fonts for PNG rendering (repeatable)",
    )
    parser.add_argument(
        "--system-fonts",
        action="store_true",
        help=(
            "also use the host's installed fonts (default: bundled fonts only, so the "
            "same deck rasterises to the same bytes everywhere)"
        ),
    )
    parser.add_argument(
        "--no-embedded-fonts",
        action="store_true",
        help=(
            "ignore fonts the deck carries in <p:embeddedFontLst>; by default they are "
            "used for both measurement and drawing"
        ),
    )
    parser.add_argument(
        "--no-bundled-fonts",
        action="store_true",
        help="do not use the fonts shipped with pptx2svg",
    )
    parser.add_argument(
        "-q", "--quiet", action="store_true", help="suppress warnings about unsupported content"
    )
    parser.add_argument("--version", action="version", version=f"pptx2svg {__version__}")
    return parser


def parse_slide_selection(value: str | None) -> list[int] | None:
    """``"1,3,5-7"`` -> ``[1, 3, 5, 6, 7]``."""
    if not value:
        return None
    numbers: list[int] = []
    for part in value.split(","):
        part = part.strip()
        if not part:
            continue
        if "-" in part:
            start, _, end = part.partition("-")
            numbers.extend(range(int(start), int(end) + 1))
        else:
            numbers.append(int(part))
    return numbers or None


def main(argv: list[str] | None = None) -> int:
    argv = sys.argv[1:] if argv is None else list(argv)
    if argv and argv[0] == "fonts":
        return fonts_main(argv[1:])
    args = build_parser().parse_args(argv)

    if not args.input.is_file():
        print(f"pptx2svg: no such file: {args.input}", file=sys.stderr)
        return 1

    options = ConvertOptions(
        slide_numbers=parse_slide_selection(args.slides),
        width=args.width,
        height=args.height,
        use_embedded_fonts=not args.no_embedded_fonts,
    )

    try:
        documents, resolved = _render(args.input, options)
    except Exception as error:  # malformed package, unreadable XML, ...
        print(f"pptx2svg: {error}", file=sys.stderr)
        return 1

    if not documents:
        print("pptx2svg: no slides matched the selection", file=sys.stderr)
        return 1

    args.output.mkdir(parents=True, exist_ok=True)
    stem = args.input.stem
    numbers = options.slide_numbers or range(1, len(documents) + 1)

    wants_png = args.format in ("png", "both")
    if wants_png and not available_backends():
        print(
            "pptx2svg: PNG output needs a rasterizer; run `pip install pptx2svg[png]`",
            file=sys.stderr,
        )
        return 1

    # The extracted faces only need to exist while the rasteriser is indexing them; a
    # deck with no embedded fonts never creates the directory.
    with contextlib.ExitStack() as stack:
        embedded_files: list[str] = []
        if wants_png and resolved.embedded_fonts:
            directory = stack.enter_context(
                tempfile.TemporaryDirectory(prefix="pptx2svg-fonts-")
            )
            embedded_files = resolved.embedded_fonts.write(directory)
        return _write_outputs(args, options, documents, numbers, stem, wants_png, embedded_files)


def _write_outputs(args, options, documents, numbers, stem, wants_png, embedded_files) -> int:
    """Write the SVG and PNG files, then report warnings.

    Split out of :func:`main` only so the extracted embedded faces can live in a
    temporary directory that is open for the whole rasterisation and closed after it.
    """
    for number, document in zip(numbers, documents):
        if args.format in ("svg", "both"):
            path = args.output / f"{stem}-{number}.svg"
            path.write_text(document, encoding="utf-8")
            print(path)
        if wants_png:
            path = args.output / f"{stem}-{number}.png"
            try:
                path.write_bytes(
                    svg_to_png(
                        document,
                        backend=args.backend,
                        font_dirs=args.font_dirs,
                        font_files=embedded_files or None,
                        # None means "decide from the bundle": skip system fonts
                        # when there is a bundle to be deterministic with, keep them
                        # when there is not, because skipping both renders blank
                        # slides.  --system-fonts is explicit and overrides that.
                        skip_system_fonts=False if args.system_fonts else None,
                        use_bundled_fonts=not args.no_bundled_fonts,
                    )
                )
            except RasterizerNotAvailable as error:
                print(f"pptx2svg: {error}", file=sys.stderr)
                return 1
            print(path)

    if options.warnings and not args.quiet:
        print(f"\n{len(options.warnings)} warning(s):", file=sys.stderr)
        for warning in options.warnings:
            print(f"  {warning}", file=sys.stderr)

    return 0


# --------------------------------------------------------------------------------------
# `pptx2svg fonts`
# --------------------------------------------------------------------------------------


def build_fonts_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="pptx2svg fonts",
        description=(
            "Report which faces a deck asks for, which the bundle can draw, and which "
            "will be substituted.  Exits non-zero when a deck cannot be rendered "
            "faithfully, so it can gate a build."
        ),
    )
    parser.add_argument(
        "input",
        type=Path,
        nargs="?",
        help="a .pptx file; omit to report on the bundle itself",
    )
    parser.add_argument(
        "--check",
        action="store_true",
        help="exit 1 when any face cannot be drawn at the widths it was measured at",
    )
    parser.add_argument(
        "--system-fonts",
        action="store_true",
        help="note that rendering will also use the host's fonts, so verdicts are a floor",
    )
    return parser


#: Column width for the face name.  Wide enough for "Hiragino Kaku Gothic ProN".
_FACE_COLUMN = 26


def fonts_main(argv: list[str]) -> int:
    args = build_fonts_parser().parse_args(argv)
    from .fonts import BUNDLED_FAMILIES, INSTALL_HINT, bundle_dir, bundle_mode
    from .fonts.check import check_deck, check_families

    mode = bundle_mode()
    # Lead with the mode.  "Which faces are missing" is the second question; the first
    # is "is this render reproducible at all", and in system mode the answer is no
    # whatever the table below says.
    if mode == "bundled":
        print(f"mode:   bundled from {bundle_dir()}")
        print("        this machine's own fonts are ignored, so output is reproducible")
        print(f"        families: {', '.join(sorted(BUNDLED_FAMILIES))}")
    else:
        print("mode:   system; pptx2svg-fonts is NOT installed")
        print("        rendering uses this machine's fonts and is NOT reproducible")
        print(f"        fix: {INSTALL_HINT}")
    print()

    if args.input is None:
        # No deck: report on the bundle.  Listing the Office faces we claim to cover is
        # more useful than listing the families we ship, because the question people
        # actually have is "will my deck work", not "what is in the wheel".
        from .text.fontmap import SUBSTITUTIONS

        report = check_families(sorted(SUBSTITUTIONS), system_fonts=args.system_fonts)
    else:
        if not args.input.is_file():
            print(f"pptx2svg: no such file: {args.input}", file=sys.stderr)
            return 1
        try:
            report = check_deck(args.input, system_fonts=args.system_fonts)
        except Exception as error:  # malformed package, unreadable XML, ...
            print(f"pptx2svg: {error}", file=sys.stderr)
            return 1

    print(f"{'face':{_FACE_COLUMN}} {'verdict':12} {'drawn with':14} note")
    for face in report.faces:
        drawn = face.substitute or "-"
        print(f"{face.requested:{_FACE_COLUMN}} {face.verdict:12} {drawn:14} {face.reason}")
    if not report.faces:
        print("(no typefaces found)")

    print()
    if report.mode != "bundled":
        print(
            "every verdict above is provisional: without the bundle the rasteriser "
            "picks faces from this machine, and it does so silently."
        )
    unfaithful = [face for face in report.faces if not face.faithful]
    if unfaithful:
        print(
            f"{len(unfaithful)} of {len(report.faces)} faces will not be drawn at the "
            "widths they were measured at."
        )
    else:
        print(f"all {len(report.faces)} faces resolve to the metrics they were measured at.")
    if args.system_fonts:
        print(
            "note: --system-fonts was given, so the host's own fonts may cover some of "
            "the above -- and may differ from the next host's."
        )

    # Exit status is about the render being trustworthy, which needs both halves:
    # every face drawn at the widths it was measured at, *and* a bundle so that the
    # next machine gets the same answer.
    failed = bool(unfaithful) or report.mode != "bundled"
    return 1 if (args.check and failed) else 0


if __name__ == "__main__":
    raise SystemExit(main())
