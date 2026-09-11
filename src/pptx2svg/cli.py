"""Command-line interface: ``pptx2svg deck.pptx -o out/``."""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from . import ConvertOptions, __version__, convert_pptx_to_svg
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
        "--skip-system-fonts",
        action="store_true",
        help="ignore installed fonts; use only --font-dir (reproducible output)",
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
    args = build_parser().parse_args(argv)

    if not args.input.is_file():
        print(f"pptx2svg: no such file: {args.input}", file=sys.stderr)
        return 1

    options = ConvertOptions(
        slide_numbers=parse_slide_selection(args.slides),
        width=args.width,
        height=args.height,
    )

    try:
        documents = convert_pptx_to_svg(args.input, options)
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
                        skip_system_fonts=args.skip_system_fonts,
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


if __name__ == "__main__":
    raise SystemExit(main())
