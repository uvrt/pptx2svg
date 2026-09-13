#!/usr/bin/env python3
"""Score our render against PowerPoint's, in a way that tracks what actually broke.

The obvious metric -- what fraction of pixels differ by more than a threshold -- was
used here for a while and it is actively misleading.  Text is thin, high-contrast and
everywhere, so moving a baseline by one pixel lights up every glyph on the slide and
scores like a catastrophe, while filling a whole table header with the wrong colour
moves the number barely at all.  A correct fix to the table gridlines scored as a
regression that way, which is what prompted this file.

So there are two gates here instead, both borrowed from ``aiden0z/pptx-renderer``:

* **SSIM >= 0.95.**  Structural similarity compares local means, variances and
  covariance, so it answers "is the same thing in the same place" rather than "are these
  bytes equal".  Wrong geometry, missing elements and layout shifts move it; a half-pixel
  of anti-aliasing does not.
* **Colour histogram correlation >= 0.80.**  SSIM on greyscale is nearly blind to hue,
  and the bugs this project keeps hitting -- theme colour resolution, gradients, tint and
  shade -- are precisely hue bugs.  Correlating per-channel histograms catches those
  while ignoring position entirely, which makes it a good complement rather than a
  second opinion on the same thing.

Both are computed over **foreground pixels only** (grey < 245 in either image).  A slide
that is mostly white background otherwise scores near-perfect no matter what happens to
the content.  When foreground coverage is under 1.5% the slide is too sparse to say
anything useful and the score is defined as 1.0 rather than left to noise.

Foreground IoU was considered and deliberately left out: the upstream project dropped it
because thin-stroke shapes lose about half their IoU to one pixel of anti-aliasing even
when the geometry is exactly right, and table gridlines -- the thing that started this --
are exactly that shape of problem.

**Both sides draw with the real fonts.**  Comparing text metrics only means something
when both renderers draw the same outlines.  PowerPoint always draws with Microsoft's own
Calibri, Cambria and Aptos; if we draw with Carlito and Caladea instead, every score also
contains a glyph-shape term that has nothing to do with this library, and the two decks
that use Aptos -- for which no clone exists at all -- are not measurable at all.

So the harness renders *our* side with the same licensed faces PowerPoint used, read in
place from wherever Office installed them.  Those fonts are not ours to redistribute, so
their locations live in ``tests/font-profile.local.json``, which is written by
``--write-profile``, is specific to the machine that wrote it, and is gitignored.  With no
profile, or with a face the profile cannot supply, the deck is **skipped with a reason**
rather than scored against substitutes: a baseline taken under one font profile compared
against a run under another is exactly the apples-to-oranges number this file exists to
stop, which is what the profile hash in every baseline entry guards.

The bundled substitutes are what we *ship*; ``pptx2svg fonts --check`` and
``tests/test_fonts.py`` cover those.  They are deliberately not the reference here.

Dev-only.  This needs numpy, pillow, pypdfium2, fontTools and a rasteriser; the library
itself stays standard-library-only, which is why this lives in ``tools/`` and not in
``src/``.

Usage::

    python3 tools/fidelity.py --write-profile                        # once per machine
    python3 tools/fidelity.py --oracle ~/pptx2svg-oracle             # score and compare
    python3 tools/fidelity.py --oracle ~/pptx2svg-oracle --update    # rewrite baselines
"""

from __future__ import annotations

import argparse
import glob
import hashlib
import io
import json
import os
import re
import sys
import zipfile
from pathlib import Path

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
BASELINE_PATH = ROOT / "tests" / "fidelity-baselines.json"

#: Which checkout to import the library from.  ``--src`` repoints it at an older tree so
#: a before/after table can be produced without switching the working copy.
SOURCE_ROOT = str(ROOT / "src")

#: Render width for every comparison.  Large enough that a 1 pt rule is a few pixels,
#: small enough that a seven-deck pass is seconds rather than minutes.
WIDTH = 1280

#: Pass marks.  A slide below either of these is a failure, not a warning.
MIN_SSIM = 0.95
MIN_HISTOGRAM = 0.80

#: A stored score may drop by this much before it counts as a regression.  Rasterisers
#: differ in the last bit of anti-aliasing between versions, so an exact match is not a
#: reasonable thing to demand.
MAX_SSIM_DROP = 0.02

#: Pixels at least this bright in *both* images are background and take no part in the
#: score.  245 rather than 255 so the pale greys Office uses for card fills still count.
FOREGROUND_MAX = 245
#: Below this much foreground the slide is too sparse to score; treat it as a pass.
MIN_FOREGROUND = 0.015


# --------------------------------------------------------------------------------------
# Metrics
# --------------------------------------------------------------------------------------

def _gaussian_kernel(sigma: float = 1.5, radius: int = 5):
    import numpy as np

    x = np.arange(-radius, radius + 1, dtype=np.float64)
    k = np.exp(-(x**2) / (2 * sigma**2))
    return k / k.sum()


def _blur(image, kernel):
    """Separable Gaussian blur with edge padding, in plain numpy.

    SSIM is normally taken with scipy's ``uniform_filter`` or skimage; neither is a
    dependency here, and a separable convolution is a dozen lines.
    """
    import numpy as np

    radius = len(kernel) // 2
    padded = np.pad(image, ((radius, radius), (0, 0)), mode="edge")
    out = np.zeros_like(image, dtype=np.float64)
    for i, weight in enumerate(kernel):
        out += weight * padded[i : i + image.shape[0], :]
    padded = np.pad(out, ((0, 0), (radius, radius)), mode="edge")
    out = np.zeros_like(image, dtype=np.float64)
    for i, weight in enumerate(kernel):
        out += weight * padded[:, i : i + image.shape[1]]
    return out


def ssim_map(a, b, data_range: float = 255.0):
    """Per-pixel structural similarity between two greyscale arrays."""
    import numpy as np

    a = a.astype(np.float64)
    b = b.astype(np.float64)
    kernel = _gaussian_kernel()
    c1 = (0.01 * data_range) ** 2
    c2 = (0.03 * data_range) ** 2

    mu_a = _blur(a, kernel)
    mu_b = _blur(b, kernel)
    # Var(X) = E[X^2] - E[X]^2, clipped because the subtraction can go slightly negative
    # on flat regions and a negative variance makes the formula produce nonsense.
    var_a = np.maximum(_blur(a * a, kernel) - mu_a * mu_a, 0.0)
    var_b = np.maximum(_blur(b * b, kernel) - mu_b * mu_b, 0.0)
    cov = _blur(a * b, kernel) - mu_a * mu_b

    numerator = (2 * mu_a * mu_b + c1) * (2 * cov + c2)
    denominator = (mu_a**2 + mu_b**2 + c1) * (var_a + var_b + c2)
    return numerator / denominator


def foreground_mask(a_rgb, b_rgb):
    """Pixels that are content in either image."""
    import numpy as np

    a_gray = a_rgb.mean(axis=2)
    b_gray = b_rgb.mean(axis=2)
    return (a_gray < FOREGROUND_MAX) | (b_gray < FOREGROUND_MAX), np.float64(0)


def histogram_correlation(a_rgb, b_rgb, mask) -> float:
    """Pearson correlation of the two images' per-channel colour histograms.

    Position-blind on purpose: this is the gate for "did we resolve the right colours",
    and it must not also fail when the right colours are a pixel to the left.
    """
    import numpy as np

    bins = 64
    vectors = []
    for image in (a_rgb, b_rgb):
        parts = []
        for channel in range(3):
            values = image[:, :, channel][mask]
            counts, _ = np.histogram(values, bins=bins, range=(0, 256))
            parts.append(counts.astype(np.float64))
        vectors.append(np.concatenate(parts))

    first, second = vectors
    first = first - first.mean()
    second = second - second.mean()
    denominator = np.sqrt((first**2).sum() * (second**2).sum())
    if denominator == 0:
        # Both histograms are flat, which means both images are a single colour over the
        # mask.  Identical-and-featureless is a match, not a failure.
        return 1.0
    return float((first * second).sum() / denominator)


def score(ours_rgb, truth_rgb) -> dict:
    """SSIM, histogram correlation and the legacy pixel percentages for one slide."""
    import numpy as np

    mask, _ = foreground_mask(ours_rgb, truth_rgb)
    coverage = float(mask.mean())

    ours_gray = ours_rgb.mean(axis=2)
    truth_gray = truth_rgb.mean(axis=2)
    difference = np.abs(ours_rgb.astype(np.int16) - truth_rgb.astype(np.int16))

    result = {
        "coverage": round(coverage, 4),
        "over10": round(100 * float((difference.max(axis=2) > 10).mean()), 3),
        "over64": round(100 * float((difference.max(axis=2) > 64).mean()), 3),
        "mean_abs": round(float(difference.mean()), 3),
    }
    if coverage < MIN_FOREGROUND:
        # Too little content to say anything; scoring noise here would be worse than
        # declining to score at all.
        result["ssim"] = 1.0
        result["histogram"] = 1.0
        result["sparse"] = True
        return result

    result["ssim"] = round(float(ssim_map(ours_gray, truth_gray)[mask].mean()), 4)
    result["histogram"] = round(histogram_correlation(ours_rgb, truth_rgb, mask), 4)
    result["sparse"] = False
    return result


# --------------------------------------------------------------------------------------
# Font profile
# --------------------------------------------------------------------------------------

#: Where the licensed Microsoft faces live on this machine.  Office keeps its own copies
#: inside the application bundle -- that is the *point*: PowerPoint can use them and
#: nothing else can, and rendering our side without them is what made every score here a
#: measurement of font availability rather than of this library.
#:
#: Nothing is ever copied out of these directories.  ``--write-profile`` records paths and
#: hashes into ``tests/font-profile.local.json``, which is gitignored because it names a
#: particular machine, and because the fonts it points at are not ours to redistribute.
LICENSED_FONT_DIRECTORIES = (
    "/Applications/Microsoft PowerPoint.app/Contents/Resources/DFonts",
    # Office downloads some faces on demand instead of installing them: Aptos Display,
    # Lato, Raleway and Segoe UI live only here, under opaque numeric filenames.  Missing
    # this directory is why "Aptos Display" looked unavailable while PowerPoint was
    # happily embedding it in the exported PDF -- and why the two Aptos decks, the worst
    # scoring in the corpus, could not be measured at all.
    "~/Library/Group Containers/UBF8T346G9.Office/FontCache",
    "/Applications/Microsoft Word.app/Contents/Resources/DFonts",
    "/System/Library/Fonts/Supplemental",   # macOS ships Arial, Times New Roman, Courier
    "/Library/Fonts",
    "~/Library/Fonts",
    "/usr/share/fonts/truetype/msttcorefonts",
    "~/.local/share/fonts/ppviewer",        # tools/install-fonts-debian.sh puts them here
    "/usr/local/share/fonts",
)

PROFILE_PATH = ROOT / "tests" / "font-profile.local.json"


def _faces_in(path: str) -> list[tuple[set[str], str, int]]:
    """``(family names, style, weight)`` for every face in a font file.

    Style is one of ``regular``/``bold``/``italic``/``bolditalic``, taken from the OS/2
    selection flags rather than the subfamily string, which is localised and creative
    ("Negrita", "Fett", "Semibold" where the file is really the bold cut).

    Filenames lie -- ``calibril.ttf`` is "Calibri Light", ``YuGothR.ttc`` is "Yu Gothic"
    -- and matching a deck's ``typeface="Calibri Light"`` against a filename stem is how
    the old profile decided Calibri Light was missing while it sat right there.
    """
    from fontTools.ttLib import TTCollection, TTFont

    try:
        fonts = (
            TTCollection(path, lazy=True).fonts
            if path.lower().endswith((".ttc", ".otc"))
            else [TTFont(path, lazy=True, fontNumber=0)]
        )
    except Exception:  # unreadable, bitmap-only, or a format fontTools declines
        return []

    result: list[tuple[set[str], str, int]] = []
    for font in fonts:
        try:
            table = font["name"]
        except Exception:
            continue
        names: set[str] = set()
        # Every language record, not just the English one.  A Japanese deck asks for
        # "游ゴシック"; the file calls itself "Yu Gothic" in English and
        # "游ゴシック" in Japanese, and only reading both makes the two meet.
        # getDebugName() returns the English record alone, which is why the profile used
        # to report every Japanese face as missing while it sat in DFonts.
        # nameID 1 (the legacy family) only, deliberately.  nameID 16 -- the
        # typographic family -- groups every weight of a superfamily under one name, so
        # Aptos-Light.ttf and Aptos-Black.ttf both answer to "Aptos" there and the first
        # one scanned would become "Aptos regular".  nameID 1 keeps them apart as
        # "Aptos Light" and "Aptos Black", which is also how PowerPoint and fontdb match
        # a font-family string, so it is the name a deck is actually asking for.
        for record in table.names:
            if record.nameID != 1:
                continue
            try:
                value = record.toUnicode()
            except Exception:
                continue
            if value:
                names.add(value.strip())
        if not names:
            continue
        try:
            os2 = font["OS/2"]
            selection, weight = os2.fsSelection, os2.usWeightClass
        except Exception:
            selection, weight = 0x40, 400
        bold = bool(selection & 0x20)
        italic = bool(selection & 0x01)
        style = ("bold" if bold else "") + ("italic" if italic else "") or "regular"
        result.append((names, style, weight))
    return result


def scan_licensed_fonts() -> dict[str, dict]:
    """Family name -> ``{style: {path, sha256}}`` for every licensed face this machine has.

    Styles are kept apart because ``tools/extract_font_metrics.py`` needs the *bold* file
    of a measured-only face (Aptos, Cambria), and because a family whose bold cut is
    missing is a different thing from one that is fully present.

    The full file is hashed, not just its head: the baseline's whole purpose is to refuse
    a comparison across a font change, and a 64 kB prefix would miss a hinting update
    that moves glyphs.
    """
    found: dict[str, dict] = {}
    for directory in LICENSED_FONT_DIRECTORIES:
        base = os.path.expanduser(directory)
        if not os.path.isdir(base):
            continue
        for path in sorted(glob.glob(os.path.join(base, "**", "*"), recursive=True)):
            if not path.lower().endswith((".ttf", ".otf", ".ttc", ".otc")):
                continue
            faces = _faces_in(path)
            if not faces:
                continue
            digest = hashlib.sha256()
            with open(path, "rb") as handle:
                digest.update(handle.read())
            for families, style, weight in faces:
                entry = {
                    "path": path,
                    "sha256": digest.hexdigest()[:16],
                    "weight": weight,
                }
                for family in families:
                    styles = found.setdefault(family, {})
                    previous = styles.get(style)
                    # First directory wins -- PowerPoint's own DFonts are listed first,
                    # because they are the copies PowerPoint actually draws with -- but
                    # within a directory prefer the canonical weight, so a family that
                    # also ships Light and Black cuts does not hand us one of those as
                    # its "regular".
                    target = 700 if "bold" in style else 400
                    if previous is None or abs(weight - target) < abs(
                        previous["weight"] - target
                    ):
                        styles[style] = entry
    return found


def write_profile() -> dict:
    faces = scan_licensed_fonts()
    profile = {
        # resvg is pointed at whole directories rather than individual files: a face is
        # four files (regular, bold, italic, bold-italic) and the deck decides at render
        # time which it needs, so naming only the upright would silently drop bold.
        "directories": sorted(
            {
                os.path.dirname(entry["path"])
                for styles in faces.values()
                for entry in styles.values()
            }
        ),
        "faces": faces,
    }
    PROFILE_PATH.write_text(json.dumps(profile, indent=2, sort_keys=True) + "\n")
    return profile


def load_profile() -> dict | None:
    if not PROFILE_PATH.exists():
        return None
    return json.loads(PROFILE_PATH.read_text())


#: A theme's font scheme ends with a long ``<a:font script="Arab" typeface="..."/>`` list
#: -- forty-odd faces for scripts the deck never uses.  Counting those as "requested"
#: buries the two or three faces that actually get drawn, so they are skipped.
_SCRIPT_FONT = re.compile(rb"<a:font\s+script=")


def requested_faces(deck: Path) -> list[str]:
    """The faces a deck could actually draw with: slides, layouts, masters, theme.

    Only ``a:latin`` / ``a:ea`` / ``a:cs`` entries count.  Theme script fallbacks are
    excluded (see :data:`_SCRIPT_FONT`), as are ``+mj-lt``-style theme references, which
    are pointers rather than names.
    """
    names: set[str] = set()
    pattern = re.compile(rb'<a:(?:latin|ea|cs)\s+typeface="([^"]*)"')
    with zipfile.ZipFile(deck) as archive:
        for name in archive.namelist():
            if not name.startswith(
                ("ppt/slides/slide", "ppt/slideLayouts/", "ppt/slideMasters/", "ppt/theme/")
            ):
                continue
            body = _SCRIPT_FONT.split(archive.read(name))[0] if b"ppt/theme/" in name.encode() else archive.read(name)
            for match in pattern.findall(body):
                value = match.decode("utf-8", "replace")
                if value and not value.startswith("+"):
                    names.add(value)
    return sorted(names)


#: Why there is no "PowerPoint fell back to X, so we will too" table here.
#:
#: It is tempting: read ``/BaseFont`` out of the exported PDF, see that PowerPoint drew
#: Calibri where the deck asked for Noto Sans JP, and point our render at Calibri as
#: well.  It would make three more decks scoreable.  It would also be measuring the wrong
#: thing twice over -- our layout would still be computed from Noto Sans JP's advance
#: widths while drawing Calibri's, so the comparison would contain a deliberate
#: metrics/outline mismatch of our own making, and any conclusion drawn from it would be
#: about the fudge rather than about the renderer.
#:
#: A deck naming a face PowerPoint does not have is simply not a deck this corpus can
#: score.  The fix is to install the face where PowerPoint can see it and re-export, not
#: to guess around it.


def font_profile(deck: Path, profile: dict | None) -> dict:
    """Which of a deck's faces the licensed profile can supply, and a hash of that answer.

    ``missing`` is the field that decides whether a deck gets scored at all.  A deck we
    cannot draw with the same faces PowerPoint drew with is not a failing deck, it is an
    unmeasurable one, and scoring it anyway is how a font-availability difference gets
    recorded as a rendering regression.
    """
    faces = (profile or {}).get("faces", {})
    available: list[str] = []
    missing: list[str] = []
    for face in requested_faces(deck):
        (available if face in faces else missing).append(face)

    digest = hashlib.sha256()
    for face in available:
        for style, entry in sorted(faces[face].items()):
            digest.update(f"{face}:{style}:{entry['sha256']}\n".encode())
    return {
        "available": available,
        "missing": missing,
        "hash": digest.hexdigest()[:12],
    }


# --------------------------------------------------------------------------------------
# Running a corpus
# --------------------------------------------------------------------------------------

def _supported(target, **kwargs) -> dict:
    """Drop keyword arguments ``target`` does not accept."""
    import inspect

    accepted = set(inspect.signature(target).parameters)
    return {name: value for name, value in kwargs.items() if name in accepted}


def render_pair(deck: Path, pdf: Path, slide_index: int, profile: dict):
    """Our PNG and PowerPoint's, as equally sized RGB arrays.

    Our side is rendered with the *licensed* faces from the profile and nothing else --
    not the host's fonts, not the bundled substitutes.  That is the only configuration in
    which a difference between the two images is attributable to this library: PowerPoint
    drew with Microsoft's Calibri, so we draw with Microsoft's Calibri, and what is left
    over is layout.
    """
    import numpy as np
    import pypdfium2 as pdfium
    from PIL import Image

    sys.path.insert(0, SOURCE_ROOT)
    from pptx2svg import ConvertOptions, convert_pptx_to_svg
    from pptx2svg.png import svg_to_png

    # --src can point this at an older checkout for a before/after table, and older
    # checkouts do not have these keywords.  Filtering by signature keeps the comparison
    # possible instead of making it a TypeError.
    convert_kwargs = _supported(ConvertOptions, warn_on_font_substitution=False)
    raster_kwargs = _supported(
        svg_to_png,
        font_dirs=profile["directories"],
        skip_system_fonts=True,
        use_bundled_fonts=False,
    )

    svg = convert_pptx_to_svg(str(deck), ConvertOptions(width=WIDTH, **convert_kwargs))[
        slide_index
    ]
    ours = Image.open(
        io.BytesIO(svg_to_png(svg, backend="resvg", **raster_kwargs))
    ).convert("RGB")

    page = pdfium.PdfDocument(str(pdf))[slide_index]
    truth = page.render(scale=WIDTH / page.get_size()[0]).to_pil().convert("RGB")
    if truth.size != ours.size:
        truth = truth.resize(ours.size, Image.LANCZOS)
    return np.asarray(ours), np.asarray(truth)


def slide_count(deck: Path) -> int:
    with zipfile.ZipFile(deck) as archive:
        return sum(
            1
            for name in archive.namelist()
            if re.fullmatch(r"ppt/slides/slide\d+\.xml", name)
        )


def run(oracle_dir: Path, profile: dict) -> dict:
    """Score every deck in ``oracle_dir`` that has a matching exported PDF.

    A deck whose faces the profile cannot supply is recorded as ``skipped`` rather than
    scored.  Comparing it anyway would mean our render used substitute outlines while
    PowerPoint used Microsoft's, and every glyph would differ for a reason that has
    nothing to do with the renderer.
    """
    results: dict[str, dict] = {}
    for deck in sorted(oracle_dir.glob("*.pptx")):
        pdf = deck.with_suffix(".pdf")
        if not pdf.exists():
            continue
        fonts = font_profile(deck, profile)
        entry: dict = {"fonts": fonts, "slides": []}
        if fonts["missing"]:
            # PowerPoint did not have these faces either, so its PDF is already drawn
            # with substitutes of its own choosing.  Scoring against it would compare
            # our fallback to Microsoft's.
            entry["skipped"] = (
                "PowerPoint substituted too: no "
                + ", ".join(fonts["missing"])
                + " on this machine"
            )
            results[deck.stem] = entry
            continue
        for index in range(slide_count(deck)):
            ours, truth = render_pair(deck, pdf, index, profile)
            entry["slides"].append(score(ours, truth))
        entry["ssim"] = round(
            sum(s["ssim"] for s in entry["slides"]) / len(entry["slides"]), 4
        )
        entry["histogram"] = round(
            sum(s["histogram"] for s in entry["slides"]) / len(entry["slides"]), 4
        )
        results[deck.stem] = entry
    return results


def report(results: dict, baselines: dict | None = None) -> int:
    """Print a table and return the number of failures."""
    failures = 0
    print(f"{'deck':26} {'SSIM':>7} {'hist':>7} {'>10%':>7} {'fonts':>7}  verdict")
    for name, entry in results.items():
        fonts = entry["fonts"]
        supply = f"{len(fonts['available'])}/{len(fonts['available']) + len(fonts['missing'])}"
        if entry.get("skipped"):
            print(f"{name:26} {'-':>7} {'-':>7} {'-':>7} {supply:>7}  SKIPPED: {entry['skipped']}")
            continue
        over10 = sum(s["over10"] for s in entry["slides"]) / len(entry["slides"])
        notes = []
        if entry["ssim"] < MIN_SSIM:
            notes.append(f"SSIM<{MIN_SSIM}")
        if entry["histogram"] < MIN_HISTOGRAM:
            notes.append(f"hist<{MIN_HISTOGRAM}")
        if baselines and name in baselines:
            previous = baselines[name]
            if previous.get("skipped"):
                notes.append("baseline was skipped; nothing to compare")
            elif previous["fonts"]["hash"] != fonts["hash"]:
                # Different faces on the two runs.  The scores are both valid and they
                # are not comparable; saying so is the whole reason the hash is stored.
                notes.append("font profile changed; baseline not comparable")
            elif entry["ssim"] < previous["ssim"] - MAX_SSIM_DROP:
                notes.append(f"REGRESSED from {previous['ssim']}")
                failures += 1
        if any(note.startswith(("SSIM", "hist")) for note in notes):
            failures += 1
        verdict = "; ".join(notes) if notes else "ok"
        print(
            f"{name:26} {entry['ssim']:7.4f} {entry['histogram']:7.4f} "
            f"{over10:7.2f} {supply:>7}  {verdict}"
        )

    return failures


def main() -> int:
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument(
        "--oracle",
        default="~/pptx2svg-oracle",
        help="directory of deck.pptx / deck.pdf pairs exported through PowerPoint",
    )
    parser.add_argument("--update", action="store_true", help="rewrite the stored baselines")
    parser.add_argument(
        "--write-profile",
        action="store_true",
        help="record this machine's licensed fonts into tests/font-profile.local.json",
    )
    parser.add_argument("--json", action="store_true", help="dump raw scores instead of a table")
    parser.add_argument("--src", help="import pptx2svg from this tree instead of ./src")
    args = parser.parse_args()

    if args.write_profile:
        profile = write_profile()
        print(
            f"wrote {PROFILE_PATH.relative_to(ROOT)}: "
            f"{len(profile['faces'])} families across {len(profile['directories'])} "
            "directories"
        )
        for face in ("Calibri", "Calibri Light", "Cambria", "Aptos", "Arial",
                     "Times New Roman", "Courier New"):
            mark = "yes" if face in profile["faces"] else "NO"
            print(f"  {face:18} {mark}")
        return 0

    if args.src:
        global SOURCE_ROOT
        SOURCE_ROOT = os.path.abspath(os.path.expanduser(args.src))

    profile = load_profile()
    if profile is None:
        print(
            f"no font profile at {PROFILE_PATH}.\n"
            "The corpus is scored against PowerPoint's own export, which draws with "
            "Microsoft's fonts;\nrendering our side with anything else measures font "
            "availability rather than this library.\n"
            "On a machine with Office installed, run:\n"
            "    python3 tools/fidelity.py --write-profile",
            file=sys.stderr,
        )
        return 2

    oracle_dir = Path(os.path.expanduser(args.oracle))
    if not oracle_dir.is_dir():
        print(f"no such directory: {oracle_dir}", file=sys.stderr)
        return 2

    results = run(oracle_dir, profile)
    if not results:
        print(f"no deck.pptx/deck.pdf pairs in {oracle_dir}", file=sys.stderr)
        return 2

    if args.json:
        print(json.dumps(results, indent=2, sort_keys=True))
        return 0

    baselines = None
    if BASELINE_PATH.exists() and not args.update:
        baselines = json.loads(BASELINE_PATH.read_text())
    failures = report(results, baselines)

    if args.update:
        BASELINE_PATH.write_text(json.dumps(results, indent=2, sort_keys=True) + "\n")
        print(f"\nwrote {BASELINE_PATH.relative_to(ROOT)}")
        return 0
    return 1 if failures else 0


if __name__ == "__main__":
    raise SystemExit(main())
