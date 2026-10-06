#!/usr/bin/env python3
"""Read ``make_font_resolution_probe.py``'s exports: which face PowerPoint drew each run in.

For every probe box it collects the characters PowerPoint's PDF puts inside the box
(PyMuPDF's ``rawdict``) and the face each one is drawn in -- the ``/BaseFont`` with its
subset prefix removed.  For the embedded-font decks it also fits the glyph origins of the
Latin line against the advances of each candidate copy of the face -- the deck's own and
every installed one -- and prints each copy's largest error: the copy that agrees to a
hundredth of a point is the one PowerPoint laid the line out with.

Usage::

    python3 tools/read_font_resolution_probe.py ~/            # every ~/frp-*.pdf
    python3 tools/read_font_resolution_probe.py ~/ --json     # as the fixture records it

Needs PyMuPDF and fontTools.  Fonts are read where they are installed; only names and
numbers are printed.
"""

from __future__ import annotations

import io
import json
import re
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE.parent / "src"))
sys.path.insert(0, str(HERE))

import pymupdf  # noqa: E402

from make_font_resolution_probe import (  # noqa: E402
    BOX_H, BOX_W, EMBEDDED_DECKS, JAPANESE_DECKS, SIZE, deck_probes, geometry, place,
)


def base_font(name: str) -> str:
    return re.sub(r"^[A-Z]{6}\+", "", name)


def box_chars(pages, slide: int, x: float, y: float) -> list[tuple[float, str, str]]:
    """``(x origin, char, face)`` of every glyph whose origin is inside the box."""
    out = []
    for block in pages[slide]["blocks"]:
        for line in block.get("lines", []):
            for span in line["spans"]:
                for char in span["chars"]:
                    ox, oy = char["origin"]
                    if x - 1 <= ox <= x + BOX_W and y <= oy <= y + BOX_H + 6:
                        out.append((ox, char["c"], base_font(span["font"])))
    return sorted(out)


def faces_by_script(glyphs) -> dict[str, list[str]]:
    out: dict[str, list[str]] = {}
    for _, char, face in glyphs:
        if char.isspace():
            continue
        code = ord(char)
        kind = ("emoji" if code >= 0x1F000 or 0x2600 <= code <= 0x27BF else
                "east-asian" if 0x2E80 <= code <= 0x9FFF or 0xAC00 <= code <= 0xD7AF or 0xFF00 <= code <= 0xFFEF
                else "other")
        out.setdefault(kind, [])
        if face not in out[kind]:
            out[kind].append(face)
    return out


def candidates(deck_path: Path, family: str, bold: bool):
    """``label -> (cmap, advances, upem)`` for the deck's embedded face of ``family`` and
    every installed copy PowerPoint could find."""
    from fontTools.ttLib import TTFont

    from ooxml_common.fonts import office
    from ooxml_common.fonts.eot import decode_eot

    out = {}
    with zipfile.ZipFile(deck_path) as archive:
        for name in archive.namelist():
            style = "bold" if bold else "regular"
            if re.match(rf"ppt/fonts/{re.escape('Lato' if family == 'Latoprobe' else family)}-{style}\.fntdata$", name):
                face = TTFont(io.BytesIO(decode_eot(archive.read(name))))
                out["embedded " + face["name"].getDebugName(5)[:14]] = face
    for host in office.find(family):
        if host.bold == bold and not host.italic:
            face = TTFont(host.path, fontNumber=host.number, lazy=True)
            out[f"{host.location} " + face["name"].getDebugName(5)[:14]] = face
    return out


def max_error(glyphs, face) -> float | None:
    cmap = face.getBestCmap() or {}
    advances = face["hmtx"].metrics
    scale = SIZE / 100 / face["head"].unitsPerEm
    xs = [x for x, _, _ in glyphs]
    predicted, at = [], 0.0
    for _, char, _ in glyphs:
        predicted.append(at)
        glyph = cmap.get(ord(char))
        if glyph is None:
            return None
        at += advances[glyph][0] * scale
    return max(abs((x - xs[0]) - p) for x, p in zip(xs, predicted))


def read(directory: Path) -> dict:
    result: dict = {}
    for deck in [*JAPANESE_DECKS, *EMBEDDED_DECKS]:
        pdf = directory / f"frp-{deck}.pdf"
        if not pdf.exists():
            continue
        with pymupdf.open(str(pdf)) as document:
            pages = [page.get_text("rawdict") for page in document]
        per_column, row = geometry(deck)
        entries = {}
        for index, spec in enumerate(deck_probes(deck)):
            slide, x, y = place(index, per_column, row)
            glyphs = box_chars(pages, slide, x, y)
            entry = {"drawn": faces_by_script(glyphs)}
            if deck == "bold-sizes" and len(glyphs) > 1:
                # The pen's advance per kana, which synthetic bold widens.
                entry["advance_pt"] = round((glyphs[-1][0] - glyphs[0][0]) / (len(glyphs) - 1), 4)
            if deck in EMBEDDED_DECKS and spec["text"].startswith("Hamburgefonstiv"):
                latin = [g for g in glyphs if not g[1].isspace()]
                spaced = [g for g in glyphs]
                errors = {}
                for label, face in candidates(directory / f"frp-{deck}.pptx", spec["latin"], spec["bold"]).items():
                    error = max_error(spaced, face)
                    if error is not None:
                        errors[label] = round(error, 3)
                entry["max_error_pt"] = errors
                entry["glyphs"] = len(latin)
            entries[spec["key"]] = entry
        result[deck] = entries
    return result


def main() -> int:
    if len(sys.argv) < 2:
        print(__doc__)
        return 2
    result = read(Path(sys.argv[1]).expanduser())
    if "--json" in sys.argv:
        # One probe per line, so a re-measurement diffs probe by probe.
        decks = []
        for deck, entries in result.items():
            lines = ",\n".join(f"  {json.dumps(key, ensure_ascii=False)}: "
                               f"{json.dumps(entry, ensure_ascii=False, sort_keys=True)}"
                               for key, entry in entries.items())
            decks.append(f" {json.dumps(deck)}: {{\n{lines}\n }}")
        print("{\n" + ",\n".join(decks) + "\n}")
        return 0
    for deck, entries in result.items():
        print(f"== {deck}")
        for key, entry in entries.items():
            drawn = "; ".join(f"{k}: {', '.join(v)}" for k, v in entry["drawn"].items())
            extra = f"  errors {entry['max_error_pt']}" if "max_error_pt" in entry else ""
            print(f"  {key:40} {drawn}{extra}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
