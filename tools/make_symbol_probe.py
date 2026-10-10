#!/usr/bin/env python3
"""Build a deck that shows what PowerPoint draws for the symbol faces' common codes.

Each line is a label in Calibri followed by one character in a symbol face, at 40 pt so
the glyph can be compared with the Unicode character :mod:`ooxml_common.text.symbol_fonts`
draws in its place.  Every code is written twice: as the code itself (U+0020..U+00FF, as
PowerPoint writes a ``buChar`` and a typed run) and as its private-use alias
(U+F020..U+F0FF, as Word and some generators write it), and as a run and as a bullet, so
the export says whether PowerPoint reads all four spellings as the same glyph.

    python3 tools/make_symbol_probe.py ~/symbol-probe.pptx
    osascript tools/powerpoint_export_pdf.applescript ~/symbol-probe.pptx ~/symbol-probe.pdf

Read the export by eye (or with ``pypdfium2``'s text, which names each glyph's code).  The
deck and the export are throwaway and are not committed.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))
from make_field_probe import build_deck, text_box  # noqa: E402

#: (face, code) pairs: the bullets and marks of the Office bullet libraries.
PROBES = (
    ("Wingdings", 0x71), ("Wingdings", 0x6C), ("Wingdings", 0x6E), ("Wingdings", 0xA7),
    ("Wingdings", 0xD8), ("Wingdings", 0xFC), ("Wingdings", 0xFE), ("Wingdings", 0x76),
    ("Symbol", 0xB7), ("Wingdings 2", 0x50), ("Wingdings 3", 0x7D), ("Webdings", 0x34),
)


def _face(face: str) -> str:
    return f'<a:latin typeface="{face}" pitchFamily="2" charset="2"/><a:sym typeface="{face}" pitchFamily="2" charset="2"/>'


def line(label: str, face: str, char: str, bullet: bool) -> str:
    label_run = f'<a:r><a:rPr lang="en-US" sz="1400"/><a:t>{label} </a:t></a:r>'
    if bullet:
        return (
            f'<a:pPr marL="685800" indent="-685800"><a:buSzPct val="100000"/>'
            f'<a:buFont typeface="{face}" pitchFamily="2" charset="2"/><a:buChar char="{char}"/></a:pPr>'
            + label_run.replace('sz="1400"', 'sz="3200"')
        )
    return label_run + f'<a:r><a:rPr lang="en-US" sz="3200">{_face(face)}</a:rPr><a:t>{char}</a:t></a:r>'


def slides() -> list[str]:
    out = []
    for spelling, offset in (("code", 0), ("pua", 0xF000)):
        for bullet in (False, True):
            paragraphs = [
                line(f"{face} 0x{code:02X} {spelling}{' bullet' if bullet else ''}", face,
                     chr(code + offset), bullet)
                for face, code in PROBES
            ]
            out.append(text_box(10, "Probe", 457200, 100000, 11000000, 6600000, paragraphs))
    return out


def main() -> None:
    output = Path(sys.argv[1])
    output.write_bytes(build_deck(slides()))
    print(f"wrote {output}")


if __name__ == "__main__":
    main()
