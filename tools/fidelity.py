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

"The same faces" turned out to need asking three ways, not one, and the other two were
both learned from decks that this harness was happily scoring:

1. **A face the deck names that this machine lacks.**  The original check.  PowerPoint
   lacked it too, so its export is already drawn with a substitute of its own.
2. **Text that no face the deck names can draw.**  A name-based check cannot see this --
   there is no name to look up.  A theme saying ``<a:ea typeface=""/>`` over Japanese
   body text obliges *both* renderers to invent a face, and two independently-invented
   fallbacks are not a measurement of anything.  So coverage is read from the faces'
   own cmaps instead.
3. **A face PowerPoint resolved differently.**  Coverage is necessary and not
   sufficient.  ``sample.pptx`` names ＭＳ Ｐゴシック through its theme's ``Jpan`` script
   entry, this machine has it, and PowerPoint's export embeds ``MS-Gothic`` and
   ``MS-Mincho`` -- never ``MS-PGothic``.  Those are two different faces inside one
   ``.ttc``: MS Gothic advances every CJK glyph a full em, MS PGothic is proportional
   and runs 0.648 em to 1.0 for katakana.  Nothing this renderer does can close a gap
   that begins with different outlines at different widths.

Check 3 reads ``/BaseFont`` out of PowerPoint's PDF, which is the thing the comment above
:func:`font_profile` warns against doing -- for a different purpose.  It warns against
*adopting* PowerPoint's fallback as our own, and that still stands.  Using the same fact
to decide whether a comparison means anything is the argument of that comment, not an
exception to it.

The cost is real and worth stating: two decks that used to score, ``sample`` (0.06) and
``real-basic-theme`` (0.97), are now skipped.  The 0.97 was not wrong so much as not
about this library -- its Japanese runs were our MS PGothic against PowerPoint's MS
Gothic, and it passed because there is not much Japanese on those slides.  The way to get
them back is the one this file has always recommended: make the face PowerPoint resolves
match the face the deck names, then re-export.

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


def _faces_in(path: str) -> list[tuple[set[str], str, int, str]]:
    """``(family names, style, weight, PostScript name)`` for every face in a file.

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

    result: list[tuple[set[str], str, int, str]] = []
    for font in fonts:
        try:
            table = font["name"]
        except Exception:
            continue
        names: set[str] = set()
        postscript = ""
        # Every language record, not just the English one.  A Japanese deck asks for
        # "游ゴシック"; the file calls itself "Yu Gothic" in English and
        # "游ゴシック" in Japanese, and only reading both makes the two meet.
        # getDebugName() returns the English record alone, which is why the profile used
        # to report every Japanese face as missing while it sat in DFonts.
        # nameID 1 (the legacy family) only, deliberately.  nameID 16 -- the
        # typographic family -- groups every weight of a superfamily under one name, so
        # Aptos-Light.ttf and Aptos-Black.ttf both answer to "Aptos" there and the first
        # one scanned would become "Aptos regular".  nameID 1 keeps them apart as
        # "Aptos Light" and "Aptos Black", so it is the name a deck is actually asking
        # for.
        #
        # This used to add "and that is how fontdb matches a font-family string too".
        # That is false, and it cost `table-test` the gate: fontdb indexes a face under
        # nameID 16 *whenever the face has one*, and never under nameID 1.  Measured with
        # only Aptos-Light.ttf loaded, resvg draws nothing for font-family="Aptos Light"
        # and draws the Light face for font-family="Aptos".  So the name a deck asks for
        # and the name our rasteriser answers to are two different questions, and
        # :func:`addressable_font_files` is what reconciles them.
        for record in table.names:
            if record.nameID not in (1, 6):
                continue
            try:
                value = record.toUnicode()
            except Exception:
                continue
            if not value:
                continue
            if record.nameID == 6:
                # nameID 6 is what a PDF writes in /BaseFont, so it is the only name that
                # can be matched against an exported PDF without guessing.  It is also
                # what tells MS PGothic apart from MS Gothic, which share a file and
                # therefore share everything the path can say about them.
                postscript = postscript or value.strip()
            else:
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
        result.append((names, style, weight, postscript))
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
            for families, style, weight, postscript in faces:
                entry = {
                    "path": path,
                    "sha256": digest.hexdigest()[:16],
                    "weight": weight,
                    "postscript": postscript,
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


#: Every field :func:`scan_licensed_fonts` puts in a face entry, and the top-level keys
#: :func:`write_profile` writes.  Both are hashed into :data:`PROFILE_SCHEMA`.
#:
#: This exists because a profile that is merely *old* is far more dangerous than one that
#: is missing.  ``postscript`` was added to the face entry when the coverage-based skip
#: rule landed, and the ``/BaseFont`` comparison needs it to notice that PowerPoint drew
#: a face other than the one the deck named.  A profile written before that change has
#: the field nowhere, so the check could not fire -- and the harness scored two decks it
#: cannot measure as though they were fine, one of them at a passing 0.9674.  Same
#: commit, same decks, same PDFs, opposite verdicts, no error anywhere.  A wrong answer
#: in the direction of confidence is the worst kind, so this one is fatal, not a warning.
#:
#: Declared rather than derived from a sample so that the writer can *assert* against it:
#: add a field to a face entry without listing it here and ``--write-profile`` fails on
#: the spot, which is the moment the schema really changed.  Listing it then changes the
#: hash, which retires every profile written under the old shape.
_FACE_FIELDS = ("path", "postscript", "sha256", "weight")
_PROFILE_FIELDS = ("directories", "faces", "schema")

#: Fingerprint of the shape above.  Stamped into every profile and refused on mismatch.
PROFILE_SCHEMA = hashlib.sha256(
    ("|".join(_FACE_FIELDS) + "//" + "|".join(_PROFILE_FIELDS)).encode()
).hexdigest()[:12]


class StaleProfile(Exception):
    """A profile on disk that was written under a different schema."""


def write_profile() -> dict:
    faces = scan_licensed_fonts()
    for family, styles in faces.items():
        for style, entry in styles.items():
            if tuple(sorted(entry)) != tuple(sorted(_FACE_FIELDS)):
                raise AssertionError(
                    f"face entry for {family} {style} has fields {sorted(entry)}, but "
                    f"_FACE_FIELDS says {sorted(_FACE_FIELDS)}.  Update _FACE_FIELDS -- "
                    "that is what retires profiles written under the old shape."
                )
    profile = {
        "schema": PROFILE_SCHEMA,
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


#: Where name-normalised copies of superfamily faces are staged.  Gitignored, like the
#: profile that points at their originals, and for the same reason: they are derived from
#: licensed Office fonts and are not ours to redistribute.
SHADOW_DIR = ROOT / "tests" / "local-fonts"


def _typographic_name(path: str) -> str | None:
    """nameID 16 of the first face in ``path``, or ``None`` if it has none."""
    from fontTools.ttLib import TTCollection, TTFont

    try:
        fonts = (
            TTCollection(path, lazy=True).fonts
            if path.lower().endswith((".ttc", ".otc"))
            else [TTFont(path, lazy=True, fontNumber=0)]
        )
        for record in fonts[0]["name"].names:
            if record.nameID == 16:
                return (record.toUnicode() or "").strip() or None
    except Exception:
        return None
    return None


def _has_family(font, family: str) -> bool:
    """Whether ``font`` answers to ``family`` as its nameID 1, in any language."""
    for record in font["name"].names:
        if record.nameID != 1:
            continue
        try:
            if (record.toUnicode() or "").strip() == family:
                return True
        except Exception:
            continue
    return False


def addressable_font_files(profile: dict, names) -> list[str]:
    """Copies of ``names``' faces that resvg will answer to *by that name*.

    This file's premise is that both sides draw with the same faces.  The premise had a
    hole in it: a face being installed is not the same as the rasteriser being able to
    reach it, and for a superfamily member it usually is not.

    OpenType carries two family names.  nameID 1 is the four-style family a deck spells --
    "Aptos Display", "Calibri Light".  nameID 16 is the *typographic* family that gathers
    the whole superfamily -- "Aptos", "Calibri" -- with the distinguishing style in
    nameID 17.  fontdb, which is what resvg matches ``font-family`` against, files a face
    under nameID 16 whenever it has one and **never** under nameID 1.  Measured, not
    assumed: with only ``Aptos-Light.ttf`` loaded, resvg draws nothing at all for
    ``font-family="Aptos Light"`` and draws the Light face for ``font-family="Aptos"``.
    64 of the 580 families installed on this machine are unreachable that way.

    That is not a small error dressed up as a font question.  ``table-test``'s title is
    Aptos Display 44 pt; PowerPoint inked it 227 px wide and we inked it 249, because we
    measured with Aptos Display's real advance widths and resvg drew with its default
    sans-serif -- the request fell past *every* named face in the stack to the generic.
    Nothing in the score said "wrong face".  It said SSIM 0.9494, just under the gate.

    The deck's spelling is not negotiable and neither is fontdb's rule, so the fix goes
    between them: each affected face is copied once into :data:`SHADOW_DIR` with its
    nameID 16 and 17 records dropped, which leaves fontdb no choice but to index it under
    nameID 1 -- the name the deck asked for.  Nothing else is touched; same outlines, same
    ``hmtx``, same everything the comparison is about.

    Only faces a deck actually names are staged, and only where the name is genuinely
    unreachable.  Handing the copies to resvg as ``font_files`` *alongside* the untouched
    ``font_dirs`` is deliberate: the originals go on answering to their typographic name,
    so a slide using both Aptos Display and Aptos -- which ``table-test`` does -- gets
    each of them instead of one at the other's expense.
    """
    from fontTools.ttLib import TTCollection, TTFont

    wanted = {name.casefold() for name in names}
    staged: list[str] = []
    for family, styles in sorted(profile["faces"].items()):
        if family.casefold() not in wanted:
            continue
        for style, entry in sorted(styles.items()):
            source = entry["path"]
            typographic = _typographic_name(source)
            if not typographic or typographic == family:
                # Reachable already: either there is no typographic name for fontdb to
                # prefer, or the one it prefers is the name being asked for anyway.
                continue
            target = SHADOW_DIR / f"{entry['sha256']}-{style}.ttf"
            if not target.exists():
                SHADOW_DIR.mkdir(parents=True, exist_ok=True)
                if source.lower().endswith((".ttc", ".otc")):
                    collection = TTCollection(source).fonts
                    # One .ttc holds several unrelated faces -- msgothic.ttc holds both
                    # MS Gothic and MS PGothic -- so the right one is found by name, not
                    # by position.
                    font = next(
                        (f for f in collection if _has_family(f, family)), collection[0]
                    )
                else:
                    font = TTFont(source)
                font["name"].names = [
                    record
                    for record in font["name"].names
                    if record.nameID not in (16, 17)
                ]
                font.save(str(target))
            staged.append(str(target))
    return staged


def load_profile() -> dict | None:
    """The profile on disk, or ``None`` if there is none.

    Raises :class:`StaleProfile` for one written under a different schema rather than
    using it -- see :data:`PROFILE_SCHEMA` for what that silently cost.
    """
    if not PROFILE_PATH.exists():
        return None
    profile = json.loads(PROFILE_PATH.read_text())
    found = profile.get("schema")
    if found != PROFILE_SCHEMA:
        raise StaleProfile(
            f"{PROFILE_PATH.name} was written under schema {found or 'none'}, but this "
            f"checkout expects {PROFILE_SCHEMA}.\n"
            "The profile records one fact per face and the skip rules read all of them; "
            "an older one is missing\nfields those rules need, and the harness then "
            "scores decks it cannot measure as though they were fine.\n"
            "Regenerate it:\n"
            "    python3 tools/fidelity.py --write-profile"
        )
    return profile


#: A theme's font scheme ends with a long ``<a:font script="Arab" typeface="..."/>`` list
#: -- forty-odd faces for scripts the deck never uses.  Counting those as "requested"
#: buries the two or three faces that actually get drawn, so they are skipped.
#:
#: With one exception, added after this comment turned out to be wrong about a deck in
#: the corpus.  ``sample.pptx`` writes ``<a:ea typeface=""/>`` in both font collections
#: and sets Japanese body text anyway, so ``+mn-ea`` has nothing to expand to -- and the
#: face it actually resolves through is ``<a:font script="Jpan" typeface="ＭＳ Ｐゴシック"/>``
#: in that same skipped list.  For an East Asian script the list is not noise, it is the
#: resolution path, so those four entries are read separately by
#: :func:`script_faces` instead of being discarded with the rest.
#:
#: Matches the whole element so it can be *removed*.  This used to truncate the theme at
#: the first script font instead, which threw away everything after it -- including the
#: entire ``<a:minorFont>`` block, since the major scheme's script list sits between the
#: two.  That is why ``table-test`` reported only "Aptos Display" as requested while its
#: body text is Aptos, and why PowerPoint embedding Aptos looked like a substitution.
_SCRIPT_FONT = re.compile(rb"<a:font\s+script=\"[^\"]*\"\s+typeface=\"[^\"]*\"\s*/>")

#: Scripts whose theme fallback entry is a real resolution path for ``+mj-ea``/``+mn-ea``.
_EAST_ASIAN_SCRIPTS = ("Jpan", "Hans", "Hant", "Hang")

_THEME_PARTS = ("ppt/slides/slide", "ppt/slideLayouts/", "ppt/slideMasters/", "ppt/theme/")


def _deck_parts(deck: Path):
    with zipfile.ZipFile(deck) as archive:
        for name in archive.namelist():
            if name.startswith(_THEME_PARTS):
                yield name, archive.read(name)


def requested_faces(deck: Path) -> list[str]:
    """The faces a deck names outright: slides, layouts, masters, theme.

    ``a:latin`` / ``a:ea`` / ``a:cs``, plus ``a:buFont``, which names the face a bullet
    glyph is drawn in and is a face PowerPoint really does embed -- ``sample.pptx``'s
    bullets are Arial and its PDF says so.  Theme script fallbacks are excluded (see
    :data:`_SCRIPT_FONT` and :func:`script_faces`), as are ``+mj-lt``-style theme
    references, which are pointers rather than names.
    """
    names: set[str] = set()
    pattern = re.compile(rb'<a:(?:latin|ea|cs|buFont)\s+typeface="([^"]*)"')
    for name, body in _deck_parts(deck):
        if name.startswith("ppt/theme/"):
            body = _SCRIPT_FONT.sub(b"", body)
        for match in pattern.findall(body):
            value = match.decode("utf-8", "replace")
            if value and not value.startswith("+"):
                names.add(value)
    return sorted(names)


def script_faces(deck: Path) -> dict[str, str]:
    """The theme's ``a:font script="..."`` East Asian entries, keyed by script.

    Kept apart from :func:`requested_faces` because they are *conditional*: a stock
    Office theme carries all four whether or not the deck sets a word of East Asian text,
    so counting them as requested would skip a Latin-only deck for want of 宋体.  Which
    one is in play is decided by the characters -- see :func:`deck_script`.
    """
    found: dict[str, str] = {}
    pattern = re.compile(
        rb'<a:font\s+script="(' + b"|".join(s.encode() for s in _EAST_ASIAN_SCRIPTS)
        + rb')"\s+typeface="([^"]*)"'
    )
    for name, body in _deck_parts(deck):
        if not name.startswith("ppt/theme/"):
            continue
        for script, match in pattern.findall(body):
            value = match.decode("utf-8", "replace")
            if value:
                found.setdefault(script.decode(), value)
    return found


#: Characters that name their script outright, as (script, first, last) ranges.
#:
#: Han is deliberately absent: 概 is a Chinese, Japanese and Korean character at once and
#: says nothing about which face should draw it.  Kana and Hangul do say, and that is
#: enough to tell the two decks here apart from a Chinese one.
_SCRIPT_RANGES = (
    ("Jpan", 0x3040, 0x30FF),   # hiragana and katakana
    ("Hang", 0x1100, 0x11FF),   # hangul jamo
    ("Hang", 0x3130, 0x318F),   # hangul compatibility jamo
    ("Hang", 0xAC00, 0xD7A3),   # hangul syllables
)


def deck_script(text: str) -> str | None:
    """Which East Asian script this text is, when its characters are unambiguous."""
    for char in text:
        code = ord(char)
        for script, first, last in _SCRIPT_RANGES:
            if first <= code <= last:
                return script
    return None


def deck_text(deck: Path) -> str:
    """Every ``a:t`` run of text on the slides -- what actually has to be drawn."""
    pattern = re.compile(rb"<a:t>(.*?)</a:t>", re.S)
    out: list[str] = []
    for name, body in _deck_parts(deck):
        if not name.startswith("ppt/slides/slide"):
            continue
        for match in pattern.findall(body):
            out.append(match.decode("utf-8", "replace"))
    return "".join(out)


def _face_cmaps(faces: dict, names: list[str]) -> dict[str, set[int]]:
    """Code points each named face can draw, read from its own cmap.

    Per face rather than unioned, because the interesting question is not only "can
    anything draw this character" but "is there exactly one face that can" -- see
    :func:`font_profile`.
    """
    from fontTools.ttLib import TTCollection, TTFont

    result: dict[str, set[int]] = {}
    cache: dict[str, set[int]] = {}
    for name in names:
        covered: set[int] = set()
        for entry in (faces.get(name) or {}).values():
            path = entry["path"]
            if path not in cache:
                points: set[int] = set()
                try:
                    fonts = (
                        TTCollection(path, lazy=True).fonts
                        if path.lower().endswith((".ttc", ".otc"))
                        else [TTFont(path, lazy=True, fontNumber=0)]
                    )
                except Exception:
                    fonts = []
                for font in fonts:
                    try:
                        points.update(font.getBestCmap())
                    except Exception:
                        continue
                cache[path] = points
            covered |= cache[path]
        if covered:
            result[name] = covered
    return result


def _pdf_base_fonts(pdf: Path) -> set[str]:
    """The PostScript names in the export's ``/BaseFont`` entries.

    Read from the file rather than through a text API so that a face used for a single
    glyph still shows up.  The ``AAAAAE+`` subset prefix an embedded font carries is
    stripped; it is assigned per document and means nothing across exports.
    """
    data = pdf.read_bytes()
    found: set[str] = set()
    for match in re.finditer(rb"/BaseFont\s*/([A-Za-z0-9+#,.\-]+)", data):
        name = match.group(1).decode("latin-1")
        if len(name) > 7 and name[6] == "+":
            name = name[7:]
        found.add(name)
    return found


#: Why there is still no "PowerPoint fell back to X, so we will too" table here.
#:
#: It is tempting: read ``/BaseFont`` out of the exported PDF, see that PowerPoint drew
#: MS Gothic where the deck asked for MS PGothic, and point our render at MS Gothic as
#: well.  It would make two more decks scoreable.  It would also be measuring the wrong
#: thing twice over -- our layout would still be computed from MS PGothic's advance
#: widths while drawing MS Gothic's, so the comparison would contain a deliberate
#: metrics/outline mismatch of our own making, and any conclusion drawn from it would be
#: about the fudge rather than about the renderer.  Worse, it would bake one machine's
#: font matching into the library: PowerPoint on Windows, where MS PGothic is a
#: separately addressable family, would have drawn what the deck asked for.
#:
#: What *is* read out of the PDF is the far weaker question of whether the comparison is
#: meaningful at all -- see :func:`font_profile`.  That is the same instinct this comment
#: was always about: a deck PowerPoint drew differently is not a failing deck, it is an
#: unmeasurable one, and the honest thing is to say so rather than either score it or
#: quietly copy PowerPoint's choice.


def font_profile(deck: Path, profile: dict | None, pdf: Path | None = None) -> dict:
    """Whether this deck can be compared at all, and a hash of the answer.

    Three ways a comparison stops being about this renderer, all of which end in a
    skip-with-a-reason rather than a score:

    * **A named face this machine does not have.**  PowerPoint did not have it either,
      so its PDF is already drawn with a substitute of its own choosing, and scoring
      against it compares our fallback to Microsoft's.

    * **Text no named face can draw.**  The check has to be coverage-based because a
      name-based one cannot see this case: there is no name to look up.  A deck whose
      theme says ``<a:ea typeface=""/>`` and then sets Japanese body text obliges *both*
      renderers to invent a face, and two independently-invented fallbacks are not a
      measurement of anything.

    * **PowerPoint drew a face the deck never named.**  Cmap coverage is necessary and
      not sufficient: the deck can name a face, this machine can have it, and PowerPoint
      can still have resolved the name to something else.  ``sample.pptx`` asks for
      ＭＳ Ｐゴシック through its theme's ``Jpan`` entry, this machine has it, and
      PowerPoint's export embeds ``MS-Gothic`` and ``MS-Mincho`` -- never ``MS-PGothic``.
      Those are different faces, not different names for one: MS Gothic advances every
      CJK glyph a full em and MS PGothic is proportional, 0.648 em to 1.0 for katakana.
      Nothing our renderer does can close a gap that starts with different outlines at
      different widths, and reporting it as an SSIM failure says the opposite.
    """
    faces = (profile or {}).get("faces", {})
    named = requested_faces(deck)
    text = deck_text(deck)
    script = deck_script(text)
    # Only the script the deck actually writes in.  A theme offering Jpan, Hans, Hant and
    # Hang -- which is every stock Office theme -- otherwise hands four faces to a deck
    # that uses one, and four faces between them cover every CJK character, so nothing
    # ever looks load-bearing.
    conditional = [script_faces(deck)[script]] if script in script_faces(deck) else []
    available: list[str] = []
    missing: list[str] = []
    for face in named:
        (available if face in faces else missing).append(face)

    usable = available + [face for face in conditional if face in faces]
    uncovered = ""
    substituted: list[str] = []
    if not missing and profile:
        cmaps = _face_cmaps(faces, usable)
        wanted = {ord(char) for char in text if not char.isspace()}
        unmatched = sorted(
            code for code in wanted if not any(code in c for c in cmaps.values())
        )
        uncovered = "".join(chr(code) for code in unmatched[:8])

        if pdf is not None and pdf.exists() and not uncovered:
            # A face is *load-bearing* when some character in the deck can be drawn by
            # it and by nothing else the deck names.  Only those are worth checking
            # against the export: a Latin face the deck shares with two others tells us
            # nothing, and a face PowerPoint used for content we do not render at all --
            # Arial for a chart's axis labels, say -- is a missing feature of ours, not a
            # font-matching difference, and must not be dressed up as one.
            drawn = _pdf_base_fonts(pdf)
            for face, covered in cmaps.items():
                others = [c for name, c in cmaps.items() if name != face]
                exclusive = any(
                    code in covered and not any(code in other for other in others)
                    for code in wanted
                )
                if not exclusive:
                    continue
                postscript = {
                    entry.get("postscript")
                    for entry in faces[face].values()
                    if entry.get("postscript")
                }
                if postscript and not (postscript & drawn):
                    substituted.append(face)
            substituted.sort()

    digest = hashlib.sha256()
    for face in available:
        for style, entry in sorted(faces[face].items()):
            digest.update(f"{face}:{style}:{entry['sha256']}\n".encode())
    return {
        "available": available,
        "missing": missing,
        "conditional": conditional,
        "uncovered": uncovered,
        "substituted": substituted,
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
        # Superfamily members -- Aptos Display, Calibri Light -- are installed but not
        # addressable by the name the deck spells.  See :func:`addressable_font_files`.
        font_files=addressable_font_files(
            profile, requested_faces(deck) + list(script_faces(deck).values())
        ),
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


def skip_reason(fonts: dict) -> str | None:
    """Why this deck cannot be compared, or ``None`` if it can.

    Each of these is a difference between the two renderers' *inputs*, not their output.
    Scoring through one records a font difference as a rendering regression, which is the
    failure this whole file was built to stop.
    """
    if fonts["missing"]:
        # PowerPoint did not have these faces either, so its PDF is already drawn with
        # substitutes of its own choosing.  Scoring would compare our fallback to
        # Microsoft's.
        return (
            "PowerPoint substituted too: no "
            + ", ".join(fonts["missing"])
            + " on this machine"
        )
    if fonts["uncovered"]:
        # No name to check, so the name-based rule above cannot see this one.
        return (
            "no named face can draw "
            + fonts["uncovered"]
            + "; both renderers invented a fallback"
        )
    if fonts["substituted"]:
        return (
            "PowerPoint did not draw "
            + ", ".join(fonts["substituted"])
            + ", which nothing else here covers; it resolved the name differently"
        )
    return None


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
        fonts = font_profile(deck, profile, pdf)
        entry: dict = {"fonts": fonts, "slides": []}
        reason = skip_reason(fonts)
        if reason:
            entry["skipped"] = reason
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

    try:
        profile = load_profile()
    except StaleProfile as stale:
        print(stale, file=sys.stderr)
        return 2
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
