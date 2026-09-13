#!/usr/bin/env python3
"""Regenerate ``src/pptx2svg/text/metrics.py`` from the fonts we actually ship.

The whole point of the metrics table is that layout is computed from the *same* advance
widths the rasteriser will draw with.  Keeping that true by hand does not work -- the
table this replaces was copied out of pptx-glimpse, and by the time it was checked its
Noto Sans JP column disagreed with the shipped Noto Sans JP for 186 of 191 characters.
So the table is generated, from the exact font files the ``pptx2svg-fonts``
distribution ships, and a test re-runs this and fails if the checked-in file drifts.

Aptos is the one entry with no font behind it.  Microsoft's new Office default is
proprietary and has no metric-compatible clone, so there is nothing we could ship that
draws it correctly.  Its advance widths are *measured* -- from the copy Office installs
on this machine, cross-checked against PowerPoint's own PDF export (see
``tools/measure_aptos.py``) -- and measurements are facts, not font software: no Aptos
outline, table or file is redistributed.  This is the same footing on which Carlito and
Arimo exist at all.

Usage::

    python3 tools/extract_font_metrics.py --check     # exit 1 if metrics.py is stale
    python3 tools/extract_font_metrics.py --write     # rewrite metrics.py

Needs fontTools (``pip install pptx2svg[measure]``).  Dev-only: the library itself never
reads a font file at runtime.
"""

from __future__ import annotations

import argparse
import os
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "src"))
# So a source checkout works without `pip install -e packages/pptx2svg-fonts`.
sys.path.insert(0, str(ROOT / "packages" / "pptx2svg-fonts" / "src"))

TARGET = ROOT / "src" / "pptx2svg" / "text" / "metrics.py"

BEGIN = "# --- BEGIN GENERATED METRICS (tools/extract_font_metrics.py) ---"
END = "# --- END GENERATED METRICS ---"

#: Faces we measure but never draw, resolved through the same local font profile the
#: fidelity harness uses (``tools/fidelity.py --write-profile``).  Going through the
#: profile rather than hard-coding paths matters twice over: Office keeps Aptos Display
#: only in an on-demand cloud-font cache under an opaque numeric filename, and the two
#: tools then agree by construction about which file a given face means.
#:
#: Read only.  Nothing from those directories is ever copied into the repository or a
#: wheel; see this module's docstring for why publishing the *measurements* is a
#: different question from redistributing the fonts.


# --------------------------------------------------------------------------------------
# Which characters get an entry
# --------------------------------------------------------------------------------------

def _sample_characters() -> list[str]:
    """The characters worth storing a real advance width for.

    Everything outside this set falls back to ``default_width`` (or ``cjk_width``), which
    is close enough for wrapping and wrong enough that widening the set is the cheapest
    fidelity win available.  Real decks are full of curly quotes, en dashes, ellipses and
    bullets, and those used to measure at the 0.6 em guess.
    """
    chars: list[str] = []
    chars += [chr(c) for c in range(0x20, 0x7F)]      # printable ASCII
    chars += [chr(c) for c in range(0xA0, 0x100)]     # Latin-1 supplement
    chars += [chr(c) for c in range(0x100, 0x180)]    # Latin Extended-A
    chars += [
        "ˆ", "˜",                            # circumflex, small tilde
        "–", "—", "―",                  # en/em dash, horizontal bar
        "‘", "’", "‚", "‛",        # single quotes
        "“", "”", "„",                  # double quotes
        "†", "‡", "•", "…",        # dagger, bullet, ellipsis
        "‰", "‹", "›", "⁄",        # per mille, guillemets, fraction
        "€", "™", "−",                  # euro, trademark, minus
        "■", "▪", "○", "●",        # square/circle bullets
        "◦", "⁃", "→",                  # hollow bullet, hyphen bullet, arrow
    ]
    seen: set[str] = set()
    ordered: list[str] = []
    for char in chars:
        if char not in seen:
            seen.add(char)
            ordered.append(char)
    return ordered


SAMPLE = _sample_characters()

#: Used for ``cjk_width`` and to check a CJK face really is full-width.
CJK_PROBE = "あ"  # HIRAGANA LETTER A


# --------------------------------------------------------------------------------------
# Reading one face
# --------------------------------------------------------------------------------------

def _open(path: Path, weight: int | None):
    """Open a face, instancing a variable font at ``weight`` when one is asked for.

    Arimo, Raleway and Noto Sans JP ship as single variable files.  resvg reads the
    weight axis correctly -- Arimo at ``wght=700`` renders pixel-for-pixel like static
    Liberation Sans Bold -- so the bold table has to be taken from the same instance the
    rasteriser will produce, not from the file's default instance.
    """
    from fontTools.ttLib import TTFont

    font = TTFont(os.fspath(path), fontNumber=0, lazy=True)
    if weight is None or "fvar" not in font:
        return font
    from fontTools.varLib.instancer import instantiateVariableFont

    return instantiateVariableFont(
        TTFont(os.fspath(path), fontNumber=0), {"wght": weight}, inplace=False
    )


def _widths(font) -> tuple[int, dict[str, int], int]:
    units_per_em = font["head"].unitsPerEm
    cmap = font.getBestCmap()
    hmtx = font["hmtx"]

    widths: dict[str, int] = {}
    for char in SAMPLE:
        glyph = cmap.get(ord(char))
        if glyph is not None:
            widths[char] = hmtx[glyph][0]

    cjk_glyph = cmap.get(ord(CJK_PROBE))
    cjk_width = hmtx[cjk_glyph][0] if cjk_glyph is not None else units_per_em
    return units_per_em, widths, cjk_width


def _default_width(widths: dict[str, int], units_per_em: int) -> int:
    """Stand-in for characters with no entry.

    The mean over the lower-case Latin alphabet, which beats the font's own
    ``advanceWidthMax`` or the width of a space -- both of which earlier tables used and
    both of which are far from typical.
    """
    alphabet = [widths[c] for c in "abcdefghijklmnopqrstuvwxyz" if c in widths]
    return round(sum(alphabet) / len(alphabet)) if alphabet else units_per_em // 2


def read_face(regular: Path, bold: Path, bold_weight: int | None) -> dict:
    """Everything one entry of the table needs, from the regular and bold files."""
    font = _open(regular, None)
    units_per_em, widths, cjk_width = _widths(font)
    hhea = font["hhea"]

    bold_font = _open(bold, bold_weight)
    bold_upm, bold_widths, bold_cjk = _widths(bold_font)
    if bold_upm != units_per_em:  # pragma: no cover - would mean a mismatched pair
        raise SystemExit(f"{regular.name} and {bold.name} disagree on unitsPerEm")

    return {
        "units_per_em": units_per_em,
        # hhea, not OS/2 sTypo: the first-baseline rule in text/measure.py was calibrated
        # against PowerPoint using hhea's descent, and the two differ (Arial: -434 vs
        # -431, Carlito: -550 vs -512).
        "ascender": hhea.ascender,
        "descender": hhea.descender,
        "default_width": _default_width(widths, units_per_em),
        "cjk_width": cjk_width,
        "widths": widths,
        "bold_default_width": _default_width(bold_widths, units_per_em),
        "bold_cjk_width": bold_cjk,
        "bold_widths": bold_widths,
    }


# --------------------------------------------------------------------------------------
# Which faces make up the table
# --------------------------------------------------------------------------------------

def source_faces() -> dict[str, tuple[Path, Path, int | None]]:
    """Metrics key -> (regular file, bold file, weight to instance the bold at)."""
    from pptx2svg.fonts import bundle_dir

    bundle = bundle_dir()
    if bundle is None:
        raise SystemExit(
            "the pptx2svg-fonts distribution is not importable, so there are no font\n"
            "files to measure.  Install it from this checkout with\n"
            "    pip install -e packages/pptx2svg-fonts"
        )
    return {
        "Carlito": (bundle / "Carlito-Regular.ttf", bundle / "Carlito-Bold.ttf", None),
        "Arimo": (bundle / "Arimo[wght].ttf", bundle / "Arimo[wght].ttf", 700),
        "Tinos": (bundle / "Tinos-Regular.ttf", bundle / "Tinos-Bold.ttf", None),
        "Cousine": (bundle / "Cousine-Regular.ttf", bundle / "Cousine-Bold.ttf", None),
        "Caladea": (bundle / "Caladea-Regular.ttf", bundle / "Caladea-Bold.ttf", None),
        "Noto Sans JP": (
            bundle / "NotoSansJP[wght].ttf", bundle / "NotoSansJP[wght].ttf", 700
        ),
        "Lato": (bundle / "Lato-Regular.ttf", bundle / "Lato-Bold.ttf", None),
        "Raleway": (bundle / "Raleway[wght].ttf", bundle / "Raleway[wght].ttf", 700),
    }


#: Faces we measure but never draw: Office-only, and no open font reproduces their
#: advance widths.  See the module docstring for why storing measurements is legitimate.
#:
#: Cambria is here for a reason worth recording.  Caladea is universally described as
#: "metric-compatible with Cambria", and for the line box it is -- but its advance widths
#: are not Cambria's: measured over four representative strings it runs 4.5% narrow
#: (ratios 0.9555, 0.9384, 0.9525, 0.9542).  Taking that claim on trust is how this
#: project ended up with a Noto Sans JP table that disagreed with Noto Sans JP.
#:
#: Aptos Display is a *cloud* font: Office does not install it, it downloads it on first
#: use into ~/Library/Group Containers/UBF8T346G9.Office/FontCache.  It took reading
#: /BaseFont out of PowerPoint's own PDF export to notice it was there at all -- the
#: export embeds "AptosDisplay", so PowerPoint had it even while every font directory
#: said otherwise.  Two of the corpus's seven decks use it, and they are the two that
#: scored worst.
MEASURED_ONLY = ("Cambria", "Aptos", "Aptos Display")


def measured_only_faces() -> dict[str, tuple[Path, Path, int | None]]:
    """Resolve :data:`MEASURED_ONLY` against the local profile, skipping what is absent."""
    sys.path.insert(0, str(HERE))
    import fidelity

    profile = fidelity.load_profile()
    if profile is None:
        return {}
    faces = profile.get("faces", {})
    resolved: dict[str, tuple[Path, Path, int | None]] = {}
    for family in MEASURED_ONLY:
        styles = faces.get(family)
        if not styles or "regular" not in styles:
            continue
        regular = Path(styles["regular"]["path"])
        # A face with no bold cut measures its bold from the upright, which the generated
        # table then reports as a 1.0 ratio.  That is a real property of the font, not a
        # gap: it is what a rasteriser will draw too.
        bold = Path(styles.get("bold", styles["regular"])["path"])
        resolved[family] = (regular, bold, None)
    return resolved


# --------------------------------------------------------------------------------------
# Emitting the table
# --------------------------------------------------------------------------------------

def _literal(char: str) -> str:
    """Source spelling for a dict key: readable for ASCII, escaped above it."""
    if char == '"':
        return '"\\""'
    if char == "\\":
        return '"\\\\"'
    if 0x20 <= ord(char) < 0x7F:
        return f'"{char}"'
    return f'"\\u{ord(char):04x}"'


def _width_rows(widths: dict[str, int]) -> list[str]:
    rows: list[str] = []
    row: list[str] = []
    for char, width in widths.items():
        row.append(f"{_literal(char)}: {width},")
        if len(row) == 6:
            rows.append("            " + " ".join(row))
            row = []
    if row:
        rows.append("            " + " ".join(row))
    return rows


def render_entry(key: str, face: dict, note: str) -> str:
    lines = [f'    "{key}": FontMetrics(']
    lines.append(f"        # {note}")
    lines.append(f'        units_per_em={face["units_per_em"]},')
    lines.append(f'        ascender={face["ascender"]},')
    lines.append(f'        descender={face["descender"]},')
    lines.append(f'        default_width={face["default_width"]},')
    lines.append(f'        cjk_width={face["cjk_width"]},')
    lines.append("        widths={")
    lines += _width_rows(face["widths"])
    lines.append("        },")
    lines.append(f'        bold_default_width={face["bold_default_width"]},')
    lines.append(f'        bold_cjk_width={face["bold_cjk_width"]},')
    lines.append("        bold_widths={")
    lines += _width_rows(face["bold_widths"])
    lines.append("        },")
    lines.append("    ),")
    return "\n".join(lines)


NOTES = {
    "Carlito": "metric-compatible with Calibri; shipped in pptx2svg-fonts",
    "Arimo": "metric-compatible with Arial and Helvetica; shipped in pptx2svg-fonts",
    "Tinos": "metric-compatible with Times New Roman; shipped in pptx2svg-fonts",
    "Cousine": "metric-compatible with Courier New; shipped in pptx2svg-fonts",
    "Caladea": "the closest serif we ship to Cambria; shipped in pptx2svg-fonts",
    "Noto Sans JP": "stands in for Japanese faces; shipped in pptx2svg-fonts",
    "Lato": "itself; shipped in pptx2svg-fonts",
    "Raleway": "itself; shipped in pptx2svg-fonts",
    "Cambria": "MEASURED ONLY -- proprietary; Caladea is 4.5% narrower, so it is not it",
    "Aptos": "MEASURED ONLY -- proprietary, no clone exists, drawn with a substitute",
    "Aptos Display": "MEASURED ONLY -- Office cloud font, no clone exists",
}


def build_block() -> str:
    parts = ["METRICS: dict[str, FontMetrics] = {"]
    for key, (regular, bold, weight) in source_faces().items():
        for path in (regular, bold):
            if not path.exists():
                raise SystemExit(f"missing bundled font: {path}")
        parts.append(render_entry(key, read_face(regular, bold, weight), NOTES[key]))
    local = measured_only_faces()
    for key in MEASURED_ONLY:
        entry = local.get(key)
        if entry is not None and entry[0].exists() and entry[1].exists():
            parts.append(render_entry(key, read_face(*entry), NOTES[key]))
        else:
            # No Office on this machine, or no font profile written yet.  Re-emit what is
            # already checked in rather than dropping the entry: ``--check`` has to stay
            # usable on CI, and its job is to catch drift between the *bundled* fonts and
            # their tables.  Measured-only faces can only be regenerated where the
            # licensed originals are installed.
            parts.append(render_entry(key, _checked_in(key), NOTES[key]))
    parts.append("}")
    return "\n".join(parts)


def _checked_in(key: str) -> dict:
    from pptx2svg.text.metrics import METRICS as CURRENT

    existing = CURRENT.get(key)
    if existing is None:
        raise SystemExit(
            f"{key} is neither installed locally nor already present in metrics.py; "
            "regenerate on a machine with Microsoft Office"
        )
    return {
        "units_per_em": existing.units_per_em,
        "ascender": existing.ascender,
        "descender": existing.descender,
        "default_width": existing.default_width,
        "cjk_width": existing.cjk_width,
        "widths": existing.widths,
        "bold_default_width": existing.bold_default_width,
        "bold_cjk_width": existing.bold_cjk_width,
        "bold_widths": existing.bold_widths,
    }


def splice(text: str, block: str) -> str:
    start = text.index(BEGIN) + len(BEGIN)
    end = text.index(END)
    return text[:start] + "\n\n" + block + "\n\n" + text[end:]


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument("--write", action="store_true", help="rewrite metrics.py in place")
    group.add_argument("--check", action="store_true", help="exit 1 if metrics.py is stale")
    args = parser.parse_args()

    current = TARGET.read_text(encoding="utf-8")
    updated = splice(current, build_block())

    if args.write:
        if updated == current:
            print("metrics.py already matches the bundled fonts")
            return 0
        TARGET.write_text(updated, encoding="utf-8")
        print(f"wrote {TARGET.relative_to(ROOT)}")
        return 0

    if updated != current:
        print(
            "metrics.py does not match the bundled fonts; "
            "run tools/extract_font_metrics.py --write",
            file=sys.stderr,
        )
        return 1
    print("metrics.py matches the bundled fonts")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
