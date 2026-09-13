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

**Font profiles.**  Comparing text metrics only means something when both renderers draw
with the same outlines, and on this machine they largely do not: PowerPoint carries its
own copies of Calibri, Aptos and MS Gothic, while the rasteriser sees neither those nor
the Carlito/Arimo/Tinos substitutes, so it falls back to a generic sans that is much
wider.  A score computed across that mismatch is measuring font availability, not this
library.  :func:`font_profile` records which faces each deck asks for and which of them
the host can actually supply, and hashes the result into the baseline, so a stored score
is only ever compared against one taken with the same fonts.

Dev-only.  This needs numpy, pillow, pypdfium2 and a rasteriser; the library itself
stays standard-library-only, which is why this lives in ``tools/`` and not in ``src/``.

Usage::

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

#: Where a host keeps fonts a rasteriser will find.  Fonts bundled inside an application
#: -- PowerPoint carries Calibri, Aptos and MS Gothic in its own Resources -- are
#: deliberately not listed: PowerPoint can use them and nothing else can, which is the
#: mismatch this profile exists to record.
FONT_DIRECTORIES = (
    "/System/Library/Fonts",
    "/Library/Fonts",
    "~/Library/Fonts",
    "/usr/share/fonts",
    "/usr/local/share/fonts",
    "~/.fonts",
    "~/.local/share/fonts",
)


def installed_faces() -> dict[str, str]:
    """Font file stem -> sha1 of the file, for every font the host exposes."""
    found: dict[str, str] = {}
    for directory in FONT_DIRECTORIES:
        base = os.path.expanduser(directory)
        for path in glob.glob(os.path.join(base, "**", "*"), recursive=True):
            if not path.lower().endswith((".ttf", ".otf", ".ttc", ".dfont")):
                continue
            stem = os.path.splitext(os.path.basename(path))[0]
            key = re.sub(r"[^a-z0-9]", "", stem.lower())
            if key and key not in found:
                digest = hashlib.sha1()
                with open(path, "rb") as handle:
                    # The head of the file carries the tables that decide metrics; the
                    # whole file would make this pass noticeably slower for no gain.
                    digest.update(handle.read(65536))
                found[key] = digest.hexdigest()[:12]
    return found


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


def font_profile(deck: Path) -> dict:
    """Which of a deck's faces this host can actually draw, and a hash of that answer.

    ``missing`` is the interesting field.  A deck whose faces are all missing is not
    testing layout at all -- the rasteriser is substituting something arbitrary and the
    score mostly reflects how wide that substitute happens to be.
    """
    installed = installed_faces()
    available: list[str] = []
    missing: list[str] = []
    for face in requested_faces(deck):
        key = re.sub(r"[^a-z0-9]", "", face.lower())
        (available if key in installed else missing).append(face)

    digest = hashlib.sha1()
    for face in available:
        key = re.sub(r"[^a-z0-9]", "", face.lower())
        digest.update(f"{face}:{installed[key]}\n".encode())
    return {
        "available": available,
        "missing": missing,
        "hash": digest.hexdigest()[:12],
    }


# --------------------------------------------------------------------------------------
# Running a corpus
# --------------------------------------------------------------------------------------

def render_pair(deck: Path, pdf: Path, slide_index: int):
    """Our PNG and PowerPoint's, as equally sized RGB arrays."""
    import numpy as np
    import pypdfium2 as pdfium
    from PIL import Image

    sys.path.insert(0, SOURCE_ROOT)
    from pptx2svg import ConvertOptions, convert_pptx_to_svg
    from pptx2svg.png import svg_to_png

    svg = convert_pptx_to_svg(str(deck), ConvertOptions(width=WIDTH))[slide_index]
    ours = Image.open(io.BytesIO(svg_to_png(svg, backend="resvg"))).convert("RGB")

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


def run(oracle_dir: Path) -> dict:
    """Score every deck in ``oracle_dir`` that has a matching exported PDF."""
    results: dict[str, dict] = {}
    for deck in sorted(oracle_dir.glob("*.pptx")):
        pdf = deck.with_suffix(".pdf")
        if not pdf.exists():
            continue
        entry: dict = {"fonts": font_profile(deck), "slides": []}
        for index in range(slide_count(deck)):
            ours, truth = render_pair(deck, pdf, index)
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
    print(f"{'deck':26} {'SSIM':>7} {'hist':>7} {'>10%':>7} {'fonts':>16}  verdict")
    for name, entry in results.items():
        fonts = entry["fonts"]
        supply = f"{len(fonts['available'])}/{len(fonts['available']) + len(fonts['missing'])}"
        over10 = sum(s["over10"] for s in entry["slides"]) / len(entry["slides"])
        notes = []
        # A deck whose faces this host does not have is not a failing deck, it is an
        # unmeasurable one: the rasteriser substitutes a face of its own choosing and
        # every glyph lands somewhere else, which swamps whatever the renderer did.
        comparable = bool(fonts["available"]) and not fonts["missing"]
        if comparable:
            if entry["ssim"] < MIN_SSIM:
                notes.append(f"SSIM<{MIN_SSIM}")
            if entry["histogram"] < MIN_HISTOGRAM:
                notes.append(f"hist<{MIN_HISTOGRAM}")
        else:
            # The absolute gates are meaningless when the rasteriser is substituting
            # faces PowerPoint did not use: every glyph is a different shape in a
            # different place, and the score reflects the substitute, not the renderer.
            # The regression gate below still applies -- a run with the same fonts
            # missing is comparable to an earlier run with the same fonts missing.
            missing = ", ".join(fonts["missing"]) or "all faces"
            notes.append(f"no fonts: {missing}")
        if baselines and name in baselines:
            previous = baselines[name]
            if previous["fonts"]["hash"] != fonts["hash"]:
                notes.append("font profile changed; baseline not comparable")
            elif entry["ssim"] < previous["ssim"] - MAX_SSIM_DROP:
                notes.append(f"REGRESSED from {previous['ssim']}")
                failures += 1
        if comparable and len(notes) > 0 and any(n.startswith(("SSIM", "hist")) for n in notes):
            failures += 1
        verdict = "; ".join(notes) if notes else "ok"
        print(
            f"{name:26} {entry['ssim']:7.4f} {entry['histogram']:7.4f} "
            f"{over10:7.2f} {supply:>16}  {verdict}"
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
    parser.add_argument("--json", action="store_true", help="dump raw scores instead of a table")
    parser.add_argument("--src", help="import pptx2svg from this tree instead of ./src")
    args = parser.parse_args()

    if args.src:
        global SOURCE_ROOT
        SOURCE_ROOT = os.path.abspath(os.path.expanduser(args.src))

    oracle_dir = Path(os.path.expanduser(args.oracle))
    if not oracle_dir.is_dir():
        print(f"no such directory: {oracle_dir}", file=sys.stderr)
        return 2

    results = run(oracle_dir)
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
