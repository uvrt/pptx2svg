"""Tests for the fidelity harness in ``tools/fidelity.py``.

The harness itself needs PowerPoint, numpy, pillow and pypdfium2, none of which are
runtime dependencies and only one of which exists on a bare CI box.  So this splits in
two: the parts that read a ``.pptx`` with the standard library always run, and the
scoring maths runs only where numpy is importable.

Testing the *metrics* rather than the scores matters more than it looks.  A similarity
metric that silently returns 1.0 for everything passes every corpus run and tells you
nothing, which is a failure mode the old percent-of-pixels metric came close to.
"""

from __future__ import annotations

import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "tools"))

import fidelity  # noqa: E402

FIXTURES = ROOT / "tests" / "fixtures"


def _numpy():
    try:
        import numpy as np
    except ImportError:  # pragma: no cover - environment dependent
        pytest.skip("numpy is not installed")
    return np


def _profile():
    """The local licensed-font profile, or skip.

    The corpus is scored against PowerPoint's own export, and PowerPoint draws with
    Microsoft's fonts.  Rendering our side with anything else measures font availability
    rather than this library, so a machine without those fonts does not get a weaker
    comparison -- it gets no comparison.  Write one with::

        python3 tools/fidelity.py --write-profile
    """
    profile = fidelity.load_profile()
    if profile is None:
        pytest.skip(
            "no tests/font-profile.local.json; run `python3 tools/fidelity.py "
            "--write-profile` on a machine with Microsoft Office installed"
        )
    return profile


# --------------------------------------------------------------------------------------
# Deck inspection -- standard library only, so these always run
# --------------------------------------------------------------------------------------

def test_requested_faces_lists_only_real_typefaces():
    faces = fidelity.requested_faces(FIXTURES / "real-basic-theme.pptx")
    assert "Arial" in faces
    assert "Lato" in faces
    # "+mj-lt" is a pointer at the theme, not a face anyone can install.
    assert not any(face.startswith("+") for face in faces)


def test_requested_faces_drops_theme_script_fallbacks():
    """A theme names ~40 script fallbacks the deck never draws with; they are noise."""
    faces = fidelity.requested_faces(FIXTURES / "real-financial-report.pptx")
    assert "Angsana New" not in faces
    assert "Noto Sans JP" in faces
    assert len(faces) < 10


def test_requested_faces_keeps_the_minor_font_scheme():
    """The major scheme's script list sits between the two collections.

    Truncating the theme at the first `<a:font script=...>` threw the whole
    `<a:minorFont>` block away with it, which is why `table-test` reported only
    "Aptos Display" while its body text is Aptos.
    """
    faces = fidelity.requested_faces(FIXTURES / "authoring-integration.pptx")
    assert {"Aptos", "Aptos Display"} <= set(faces)


def test_the_east_asian_script_entry_is_read_separately():
    """`sample.pptx` names no `a:ea` face; its Japanese comes from the `Jpan` entry."""
    deck = FIXTURES / "sample.pptx"
    assert "ＭＳ Ｐゴシック" not in fidelity.requested_faces(deck)
    assert fidelity.script_faces(deck)["Jpan"] == "ＭＳ Ｐゴシック"


def test_deck_script_reads_the_characters_not_the_theme():
    """Every stock theme offers all four; only the text says which one is in play."""
    assert fidelity.deck_script("Markdownから") == "Jpan"
    assert fidelity.deck_script("한국어") == "Hang"
    # Han alone is Chinese, Japanese and Korean at once and decides nothing.
    assert fidelity.deck_script("概要") is None
    assert fidelity.deck_script("Latin only") is None


def test_slide_count_matches_the_package():
    assert fidelity.slide_count(FIXTURES / "real-basic-theme.pptx") == 2
    assert fidelity.slide_count(FIXTURES / "sample.pptx") == 6


def test_font_profile_partitions_faces_and_is_hashable():
    local = _profile()
    profile = fidelity.font_profile(FIXTURES / "real-basic-theme.pptx", local)
    assert set(profile) == {
        "available", "missing", "conditional", "uncovered", "substituted", "instead",
        "hash",
    }
    assert not set(profile["available"]) & set(profile["missing"])
    assert len(profile["hash"]) == 12
    # Stable across calls, or a baseline could never be matched to its inputs.
    assert (
        profile["hash"]
        == fidelity.font_profile(FIXTURES / "real-basic-theme.pptx", local)["hash"]
    )


def test_the_cjk_deck_is_exactly_what_its_generator_writes(tmp_path):
    """``sample-cjk.pptx`` is a derivation, and this is what keeps it one.

    The fixture's whole claim is that it is ``sample.pptx`` with two theme strings
    changed and nothing else -- which is what lets ``sample.pptx`` stay the file md-pptx
    generated while a scorable Japanese deck exists alongside it.  A claim like that is
    worth exactly as much as the check behind it: edit the derived deck by hand, or edit
    ``sample.pptx`` without regenerating, and the provenance in
    ``tests/fixtures/FIXTURES-README.md`` quietly stops being true.

    Byte-for-byte rather than "the themes match", because the point is that *nothing
    else* moved, and because the generator writes each entry back through its own
    ``ZipInfo`` precisely so that two runs agree.  Every entry -- name, metadata, bytes --
    is compared everywhere; the deflated bytes where deflate is zlib's own, since CPython
    3.14's Windows builds deflate with zlib-ng, whose bytes differ (``zip_content.py``).
    """
    import make_cjk_deck
    from zip_content import content_sha256, stock_deflate

    derived = tmp_path / "sample-cjk.pptx"
    replacements = make_cjk_deck.write_deck(FIXTURES / "sample.pptx", derived)
    assert replacements == 8, replacements
    written, committed = derived.read_bytes(), (FIXTURES / "sample-cjk.pptx").read_bytes()
    stale = ("tests/fixtures/sample-cjk.pptx is not what tools/make_cjk_deck.py writes; "
             "regenerate it rather than editing either deck by hand")
    assert content_sha256(written) == content_sha256(committed), stale
    if stock_deflate():
        assert written == committed, stale


def test_deflate_is_zlibs_unless_this_python_says_zlib_ng():
    """The deflated bytes go unchecked only where they cannot match, never by a probe gone
    wrong."""
    from zip_content import stock_deflate, zlib_ng

    assert stock_deflate() or zlib_ng()


def test_font_profile_hash_changes_when_a_face_changes():
    """The guard that stops a score from being compared across a font change."""
    local = _profile()
    deck = FIXTURES / "real-basic-theme.pptx"
    before = fidelity.font_profile(deck, local)
    if not before["available"]:
        pytest.skip("this machine supplies none of this deck's faces")

    tampered = {
        "directories": local["directories"],
        "faces": {
            family: {
                style: dict(entry, sha256="0" * 16)
                for style, entry in styles.items()
            }
            for family, styles in local["faces"].items()
        },
    }
    assert fidelity.font_profile(deck, tampered)["hash"] != before["hash"]


def test_profile_finds_the_faces_the_corpus_actually_needs():
    """Aptos Display in particular: it is a cloud font, not an installed one.

    Two of the seven corpus decks use it, they are the two that scored worst, and the
    profile reported it missing for as long as it only looked in font directories.
    PowerPoint's own export embeds it, so it is there -- in Office's on-demand cache.
    """
    faces = _profile()["faces"]
    for family in ("Calibri", "Calibri Light", "Cambria", "Aptos", "Aptos Display"):
        assert family in faces, family
        assert "regular" in faces[family], family


# --------------------------------------------------------------------------------------
# Metrics
# --------------------------------------------------------------------------------------

def test_ssim_of_an_image_with_itself_is_one():
    np = _numpy()
    rng = np.random.default_rng(0)
    image = rng.integers(0, 256, size=(64, 64)).astype(np.float64)
    assert fidelity.ssim_map(image, image).mean() == pytest.approx(1.0, abs=1e-9)


def test_ssim_falls_when_content_moves():
    np = _numpy()
    image = np.full((64, 64), 255.0)
    image[20:40, 20:40] = 0.0
    shifted = np.full((64, 64), 255.0)
    shifted[24:44, 20:40] = 0.0
    assert fidelity.ssim_map(image, shifted).mean() < 0.9


def test_ssim_is_insensitive_to_a_hairline_of_antialiasing():
    """The point of using SSIM at all: a faint edge must not read as a layout change."""
    np = _numpy()
    image = np.full((64, 64), 255.0)
    image[20:40, 20:40] = 0.0
    softened = image.copy()
    softened[19, 20:40] = 200.0
    softened[40, 20:40] = 200.0
    assert fidelity.ssim_map(image, softened).mean() > 0.98


def test_histogram_correlation_ignores_position_but_not_colour():
    np = _numpy()
    left = np.full((32, 32, 3), 255, dtype=np.uint8)
    left[:, :16] = (200, 30, 30)
    right = np.full((32, 32, 3), 255, dtype=np.uint8)
    right[:, 16:] = (200, 30, 30)          # same colours, other side
    wrong = np.full((32, 32, 3), 255, dtype=np.uint8)
    wrong[:, :16] = (30, 30, 200)          # blue where red belongs

    mask = np.ones((32, 32), dtype=bool)
    assert fidelity.histogram_correlation(left, right, mask) == pytest.approx(1.0)
    assert fidelity.histogram_correlation(left, wrong, mask) < 0.8


def test_score_declines_to_judge_a_nearly_empty_slide():
    np = _numpy()
    blank = np.full((100, 100, 3), 255, dtype=np.uint8)
    speck = blank.copy()
    speck[0:3, 0:3] = 0                    # 0.09% coverage, far below the floor
    result = fidelity.score(blank, speck)
    assert result["sparse"] is True
    assert result["ssim"] == 1.0


def test_score_judges_a_slide_with_real_content():
    np = _numpy()
    a = np.full((100, 100, 3), 255, dtype=np.uint8)
    a[10:60, 10:60] = 40
    b = np.full((100, 100, 3), 255, dtype=np.uint8)
    b[10:60, 10:60] = 40
    result = fidelity.score(a, b)
    assert result["sparse"] is False
    assert result["ssim"] == pytest.approx(1.0, abs=1e-3)
    assert result["histogram"] == pytest.approx(1.0, abs=1e-3)


def test_baselines_file_is_readable_and_carries_font_provenance():
    """A stored score without its font profile cannot be safely compared to a new one."""
    import json

    path = ROOT / "tests" / "fidelity-baselines.json"
    if not path.exists():
        pytest.skip("no baselines recorded yet")
    baselines = json.loads(path.read_text(encoding="utf-8"))
    assert baselines
    for name, entry in baselines.items():
        assert entry["fonts"]["hash"], name
        if entry.get("skipped"):
            # A deck the recording machine could not draw with PowerPoint's own faces.
            # It carries its font profile so a machine that *can* is able to tell.
            assert not entry["slides"], name
            continue
        assert "ssim" in entry, name
        assert "histogram" in entry, name
        # Which instrument took the score, or it cannot be compared with a new run: the
        # default truth's entry, with the old instrument's beside it.
        assert entry["truth"] == fidelity.DEFAULT_TRUTH, name
        assert entry["converter"].startswith("pymupdf-"), name
        assert len(entry["pdfium"]["slides"]) == len(entry["slides"]), name


def test_a_baseline_is_compared_only_under_the_truth_that_took_it():
    """A score against the converted page and one against pdfium's raster are two
    instruments' readings of one render; comparing across them would report the
    instrument's change as the renderer's."""
    fonts = {"hash": "abc"}
    baselines = {
        "deck": {"truth": "svg", "fonts": fonts, "ssim": 0.99, "slides": [{}],
                 "pdfium": {"ssim": 0.93, "histogram": 0.98, "slides": [{}]}},
        "old": {"fonts": fonts, "ssim": 0.9, "slides": [{}]},  # recorded before --truth existed
    }
    assert fidelity.baseline_for(baselines, "deck", "svg")["ssim"] == 0.99
    pdfium = fidelity.baseline_for(baselines, "deck", "pdfium")
    assert pdfium["ssim"] == 0.93 and pdfium["fonts"] is fonts
    assert fidelity.baseline_for(baselines, "old", "pdfium")["ssim"] == 0.9
    assert fidelity.baseline_for(baselines, "old", "svg") is None
    assert fidelity.baseline_for(baselines, "missing", "svg") is None
    assert fidelity.baseline_for(None, "deck", "svg") is None


def test_the_metric_clones_carry_the_originals_kern_pairs():
    """"Metric-compatible" turns out to cover the pair table too, and that is measured.

    The whole ``text/`` subsystem rests on measuring a Calibri deck with Carlito, and
    kerning was the one part of that claim nobody had checked -- a clone is free to kern
    however it likes.  Over the ~570 characters the metric tables store, **exactly one
    pair out of 9,409 disagrees by as much as 0.5/1000 em**: Calibri kerns ``",`` by
    -78/1000 em and Carlito does not.  Arimo matches Arial and Tinos matches Times New
    Roman pair for pair, and Cousine and Courier New both kern nothing at all.

    Two characters are excluded and it is worth saying which: the soft hyphen (U+00AD) and
    the no-break space (U+00A0), where the originals carry pairs the clones do not.  A
    soft hyphen is never laid out as ink; a no-break space is, and its pairs are the one
    known residue -- at most 55/1000 em on a single join against Arial.

    Local only: it reads the licensed originals in place through the same profile the
    fidelity harness uses, and nothing from them is copied anywhere.
    """
    pytest.importorskip("fontTools", reason="fontTools is not installed")
    profile = _profile()
    sys.path.insert(0, str(ROOT / "tools"))
    import extract_font_metrics as extractor
    from pptx2svg.fonts import bundle_dir

    bundle = bundle_dir()
    if bundle is None:
        pytest.skip("pptx2svg-fonts is not importable")
    characters = extractor.SAMPLE + extractor.CJK_SAMPLE
    ignorable = {"\u00ad", "\u00a0"}  # soft hyphen, no-break space: never laid out as ink

    for clone, original, filename in (
        ("Carlito", "Calibri", "Carlito-Regular.ttf"),
        ("Arimo", "Arial", "Arimo[wght].ttf"),
        ("Tinos", "Times New Roman", "Tinos-Regular.ttf"),
        ("Cousine", "Courier New", "Cousine-Regular.ttf"),
    ):
        styles = profile["faces"].get(original)
        if not styles or "regular" not in styles:
            continue
        office_path = Path(styles["regular"]["path"])
        clone_font = extractor._open(bundle / filename, None, 0)
        office_font = extractor._open(
            office_path, None, extractor.collection_index(office_path, original)
        )

        def per_mille(font):
            pairs = extractor.effective_kern(font, characters)
            pairs = pairs or extractor.legacy_kern(font, characters)
            upm = font["head"].unitsPerEm
            return {key: value * 1000 / upm for key, value in pairs.items()}

        ours, theirs = per_mille(clone_font), per_mille(office_font)
        shared = (set(ours) | set(theirs)) - {
            key for key in set(ours) | set(theirs) if ignorable & set(key)
        }
        disagree = {
            key: (ours.get(key, 0.0), theirs.get(key, 0.0))
            for key in shared
            if abs(ours.get(key, 0.0) - theirs.get(key, 0.0)) >= 0.5
        }
        disagree.pop(('"', ","), None)  # the one known divergence; see the docstring
        assert not disagree, f"{clone} vs {original}: {list(disagree.items())[:5]}"


# --------------------------------------------------------------------------------------
# SSIM cropped to the content: exact, or it is not done
# --------------------------------------------------------------------------------------

def _page(np, rows=90, cols=160):
    return np.full((rows, cols, 3), 255, dtype=np.uint8)


def _both_ways(a, b):
    cropped, full = fidelity.score(a, b), fidelity.score(a, b, crop=False)
    assert cropped == full, (cropped, full)
    return cropped


def test_the_crop_box_is_the_union_of_both_images_ink_plus_the_window():
    """Never one side's ink: what only PowerPoint drew is inside the box as surely as what
    only we drew, or the score would not see what we left out."""
    np = _numpy()
    ours, theirs = _page(np), _page(np)
    ours[30:40, 50:60] = 0           # ink only we drew
    theirs[60:70, 100:120] = 0       # ink only PowerPoint drew
    mask, _ = fidelity.foreground_mask(ours, theirs)
    rows, cols = fidelity.content_box(mask)
    radius = fidelity.SSIM_RADIUS
    assert radius == len(fidelity._gaussian_kernel()) // 2 == 5
    assert (rows.start, rows.stop) == (30 - radius, 70 + radius)
    assert (cols.start, cols.stop) == (50 - radius, 120 + radius)
    _both_ways(ours, theirs)


def test_cropped_ssim_is_the_full_pages_bit_for_bit():
    """Random content in random places, margins from none to wide: every score field is
    the same float either way, not merely close."""
    np = _numpy()
    rng = np.random.default_rng(7)
    for _trial in range(40):
        ours, theirs = _page(np), _page(np)
        top, left = int(rng.integers(0, 80)), int(rng.integers(0, 150))
        bottom = top + int(rng.integers(1, 90 - top + 1))
        right = left + int(rng.integers(1, 160 - left + 1))
        ours[top:bottom, left:right] = rng.integers(0, 256, size=(bottom - top, right - left, 3))
        shift = int(rng.integers(-4, 5))
        theirs[max(top + shift, 0):max(bottom + shift, 1), left:right] = 40
        _both_ways(ours, theirs)


def test_a_page_with_ink_at_its_edges_is_not_cropped_at_all():
    """Content to the edges, a page colour, a full-bleed picture: the box is the page, so
    there is nothing to gain and nothing is cut."""
    np = _numpy()
    ours, theirs = _page(np), _page(np)
    ours[0, 0] = 0
    theirs[-1, -1] = 0
    ours[40:50, 40:50] = 0
    mask, _ = fidelity.foreground_mask(ours, theirs)
    assert fidelity.content_box(mask) == (slice(0, 90), slice(0, 160))
    _both_ways(ours, theirs)

    tinted = np.full((90, 160, 3), 200, dtype=np.uint8)      # a page colour under 245
    tinted[20:30, 20:60] = 30
    mask, _ = fidelity.foreground_mask(tinted, tinted)
    assert fidelity.content_box(mask) == (slice(0, 90), slice(0, 160))
    assert _both_ways(tinted, tinted)["ssim"] == 1.0


def test_ink_within_the_window_of_an_edge_is_clamped_not_cut():
    """Content none to six pixels from the edges: the box is clamped to the page, and the
    page's own edge rows are what the blur's padding replicates either way."""
    np = _numpy()
    for gap in range(0, fidelity.SSIM_RADIUS + 2):
        ours, theirs = _page(np), _page(np)
        ours[gap:gap + 20, gap:gap + 30] = 0
        theirs[gap + 1:gap + 21, gap:gap + 30] = 0
        theirs[90 - gap - 12:90 - gap, 160 - gap - 25:160 - gap] = 90
        _both_ways(ours, theirs)


def _oracle_dir():
    oracle = Path.home() / "pptx2svg-oracle"
    if not any(oracle.glob("*.pdf")):
        pytest.skip("PowerPoint's exports are not in ~/pptx2svg-oracle here")
    return oracle


def _needs_the_harness(truth: str):
    for module in ("numpy", "PIL", "resvg_py", "fontTools", "pypdfium2"):
        pytest.importorskip(module)
    if truth == "svg":
        import pdf_svg

        if not pdf_svg.available():
            pytest.skip("PyMuPDF (the fidelity extra) is not installed here")


def _jobs() -> int:
    """Every core, unless this is already one of several xdist workers."""
    import os

    return 1 if os.environ.get("PYTEST_XDIST_WORKER") else fidelity.default_jobs()


@pytest.mark.parametrize("truth", ["svg", pytest.param("pdfium", marks=pytest.mark.pdfium)])
def test_cropped_ssim_equals_the_full_page_on_every_scored_slide(truth):
    """The corpus itself, every slide the harness scores: the run holds each slide's
    cropped row equal to its full-page row and raises on the first that is not.

    Rasters come from the harness's cache where it has them -- read-only, every entry's
    key components and pixel digest checked on the way in -- and are drawn otherwise;
    this test writes nothing.  Scores default to the svg truth, as the harness does;
    the pdfium truth runs under ``pytest --pdfium`` (its side is never cached)."""
    _needs_the_harness(truth)
    oracle, profile = _oracle_dir(), _profile()
    stats: dict = {}
    results = fidelity.run(oracle, profile, (truth,), _jobs(), cache_mode="read", stats=stats, crop_all=True)
    scored = sum(len(entry["slides"]) for entry in results[truth].values())
    assert scored and stats["slides"] == scored
    crops = [check for check in stats["checks"] if check.get("crop") == "equal"]
    assert len(crops) == scored, (len(crops), scored)


# --------------------------------------------------------------------------------------
# The raster cache: a stale entry can never be used quietly
# --------------------------------------------------------------------------------------

def _cache(tmp_path):
    _numpy()
    pytest.importorskip("PIL")
    import raster_cache

    return raster_cache, raster_cache.RasterCache(tmp_path / "rasters")


def test_the_cache_returns_what_it_stored_and_only_for_the_same_components(tmp_path):
    np = _numpy()
    raster_cache, cache = _cache(tmp_path)
    image = np.arange(4 * 5 * 3, dtype=np.uint8).reshape(4, 5, 3)
    components = {"svg": "abc", "fonts": "def", "width": 1280}
    assert cache.get("ours", components) is None
    cache.put("ours", components, image)
    assert np.array_equal(cache.get("ours", components), image)
    assert cache.get("ours", dict(components, fonts="xyz")) is None     # another key: a miss
    assert cache.get("truth", components) is None                        # another side
    assert cache.verify("ours", components, image) == "match"


def test_a_cached_raster_that_is_not_what_was_stored_is_refused(tmp_path):
    """Components that do not match the ones asked for, or pixels that are not the ones
    written, are a corrupted cache -- raised, not treated as a miss."""
    import json

    np = _numpy()
    raster_cache, cache = _cache(tmp_path)
    image = np.zeros((4, 5, 3), dtype=np.uint8)
    components = {"svg": "abc"}
    cache.put("ours", components, image)
    meta = cache.root / "ours" / f"{raster_cache.key(components)}.json"

    stored = json.loads(meta.read_text())
    meta.write_text(json.dumps(dict(stored, components={"svg": "other"})))
    with pytest.raises(raster_cache.CacheMismatch, match="components"):
        cache.get("ours", components)

    meta.write_text(json.dumps(dict(stored, pixels="0" * 64)))
    with pytest.raises(raster_cache.CacheMismatch, match="pixels"):
        cache.get("ours", components)


def test_a_fresh_raster_that_differs_from_the_cache_by_one_value_fails(tmp_path):
    np = _numpy()
    raster_cache, cache = _cache(tmp_path)
    image = np.full((4, 5, 3), 255, dtype=np.uint8)
    cache.put("truth", {"pdf": "p"}, image)
    fresh = image.copy()
    fresh[2, 3, 1] = 254
    with pytest.raises(raster_cache.CacheMismatch, match="1 pixels"):
        cache.verify("truth", {"pdf": "p"}, fresh)


def test_the_cache_refuses_to_live_inside_a_repository(tmp_path):
    """Rasters of PowerPoint's pages hold Microsoft's glyph shapes."""
    raster_cache, _ = _cache(tmp_path)
    with pytest.raises(RuntimeError, match="repository"):
        raster_cache.RasterCache(ROOT / "scratch" / "rasters", ROOT)
    with pytest.raises(RuntimeError, match="git checkout"):
        raster_cache.RasterCache(ROOT / "tests" / "rasters")


def test_the_spot_check_sample_is_spread_and_visits_every_page():
    import raster_cache

    pages = list(range(52))
    day = 739_000
    first = raster_cache.rotating_sample(pages, 6, day)
    assert len(first) == 6
    assert max(first) - min(first) >= 52 * 0.75       # spread through the corpus
    assert first == raster_cache.rotating_sample(pages, 6, day)
    stride = -(-52 // 6)
    seen: set = set()
    for offset in range(stride):
        seen |= set(raster_cache.rotating_sample(pages, 6, day + offset))
    assert seen == set(pages)


def test_a_font_changed_in_place_changes_the_key(tmp_path):
    """An Office update can change a face without changing its path, its name, its size
    or its date.  The key reads the file's contents."""
    import os

    _needs_the_harness("pdfium")
    fonts = tmp_path / "fonts"
    (fonts / "nested").mkdir(parents=True)
    face = fonts / "nested" / "Face.TTF"                # resvg takes either case
    face.write_bytes(b"one face")
    (fonts / "notes.txt").write_bytes(b"not a font")
    staged = tmp_path / "staged.ttf"
    staged.write_bytes(b"a staged face")
    profile = {"directories": [str(fonts)]}
    options = {"font_dirs": [str(fonts)], "font_files": [str(staged)], "skip_system_fonts": True,
               "use_bundled_fonts": False}

    before = fidelity.font_digests(profile, [staged])
    assert before["dir_files"] == 1
    key = fidelity.our_components("<svg/>", options, before)
    stamp = face.stat()
    face.write_bytes(b"two face")                        # same path, same size
    os.utime(face, ns=(stamp.st_atime_ns, stamp.st_mtime_ns))
    after = fidelity.font_digests(profile, [staged])
    assert fidelity.our_components("<svg/>", options, after) != key
    staged.write_bytes(b"another staged")
    restaged = fidelity.font_digests(profile, [staged])
    assert fidelity.our_components("<svg/>", options, restaged) != fidelity.our_components("<svg/>", options, after)
    assert fidelity.our_components("<svg />", options, after) != fidelity.our_components("<svg/>", options, after)


def _small_oracle(tmp_path):
    """A copy of one export (``table-test``, one slide) with its converted pages and
    verdicts, so the harness can run against a cache of its own.  The converted pages are
    font-derived: they stay under ``tmp_path`` and are removed by the caller."""
    import shutil

    import pdf_svg

    oracle = _oracle_dir()
    converted = pdf_svg.cache_dir(oracle / "table-test.pdf", oracle)
    if not (converted / f"validation-{fidelity.WIDTH}.json").exists():
        pytest.skip("table-test has not been validated here: python tools/pdf_svg.py --validate")
    small = tmp_path / "oracle"
    small.mkdir()
    for suffix in (".pptx", ".pdf"):
        shutil.copy2(oracle / f"table-test{suffix}", small)
    shutil.copytree(converted, pdf_svg.cache_dir(small / "table-test.pdf", small))
    return small


def test_a_stale_cached_raster_discards_the_cache_and_stops_the_run(tmp_path):
    """The detection the key cannot provide for itself.  An entry whose key is right and
    whose pixels are not -- exactly what a missing key input leaves behind -- is found by
    the spot-check, the whole cache is discarded, and the run raises instead of scoring."""
    import json
    import shutil

    _needs_the_harness("svg")
    import raster_cache

    profile = _profile()
    small = _small_oracle(tmp_path)
    try:
        stats: dict = {}
        first = fidelity.run(small, profile, ("svg",), 1, stats=stats)
        root = fidelity.raster_cache_root(small)
        entries = sorted((root / "ours").glob("*.json"))
        assert len(entries) == 1 and len(list((root / "truth").glob("*.json"))) == 1
        assert stats["sampled"] == ["table-test slide 1"]      # one slide: always sampled
        assert fidelity.run(small, profile, ("svg",), 1) == first

        # Stale but self-consistent: other pixels, recorded under the right key.
        stored = json.loads(entries[0].read_text())
        image = raster_cache._array(entries[0].with_suffix(".png").read_bytes()).copy()
        image[10, 10] = 255 - image[10, 10]
        entries[0].with_suffix(".png").write_bytes(raster_cache._png(image))
        entries[0].write_text(json.dumps(dict(stored, pixels=raster_cache.pixel_digest(image))))

        with pytest.raises(raster_cache.CacheMismatch, match="discarded"):
            fidelity.run(small, profile, ("svg",), 1)
        assert not root.exists()
        assert (small / "svg").exists()                        # the converted pages stay
        assert fidelity.run(small, profile, ("svg",), 1, cache_mode="verify") == first
    finally:
        shutil.rmtree(small, ignore_errors=True)


# --------------------------------------------------------------------------------------
# The pdfium baselines: kept, never deleted or changed by a run that did not score pdfium
# --------------------------------------------------------------------------------------

def _recorded(monkeypatch):
    import json

    import pdf_svg

    baselines = json.loads((ROOT / "tests" / "fidelity-baselines.json").read_text(encoding="utf-8"))
    converter = {entry["converter"] for entry in baselines.values()}
    assert len(converter) == 1
    monkeypatch.setattr(pdf_svg, "converter_version", lambda: next(iter(converter)))
    return baselines


def _svg_results(baselines):
    """The svg run that would reproduce ``baselines``' svg entries."""
    return {"svg": {name: {k: v for k, v in entry.items() if k not in ("truth", "converter", "pdfium")}
                    for name, entry in baselines.items()}}


def test_a_default_update_keeps_every_recorded_pdfium_entry_exactly(monkeypatch):
    """The default ``--update`` scores svg only; the pdfium entries it writes are the ones
    already recorded, and the file it writes is the same bytes."""
    import copy
    import json

    baselines = _recorded(monkeypatch)
    assert any("pdfium" in entry for entry in baselines.values())
    payload = fidelity.baselines_payload(_svg_results(baselines), copy.deepcopy(baselines))
    assert json.dumps(payload, indent=2, sort_keys=True) == json.dumps(baselines, indent=2, sort_keys=True)


def test_a_default_update_refuses_what_it_cannot_keep_honestly(monkeypatch):
    """A kept pdfium entry shares its deck's font profile and slides; where those moved,
    keeping it would misdescribe it, and deleting it is not a default run's to do."""
    import copy

    baselines = _recorded(monkeypatch)
    name = next(name for name, entry in baselines.items() if "pdfium" in entry)

    changed = _svg_results(baselines)
    changed["svg"][name]["fonts"] = dict(changed["svg"][name]["fonts"], hash="0" * 12)
    with pytest.raises(fidelity.KeptBaselineMismatch, match="font profile"):
        fidelity.baselines_payload(changed, copy.deepcopy(baselines))

    more = _svg_results(baselines)
    more["svg"][name]["slides"] = more["svg"][name]["slides"] + [{}]
    with pytest.raises(fidelity.KeptBaselineMismatch, match="slide count"):
        fidelity.baselines_payload(more, copy.deepcopy(baselines))

    skipped = _svg_results(baselines)
    skipped["svg"][name] = {"fonts": skipped["svg"][name]["fonts"], "slides": [], "skipped": "reason"}
    with pytest.raises(fidelity.KeptBaselineMismatch, match="skipped"):
        fidelity.baselines_payload(skipped, copy.deepcopy(baselines))

    gone = _svg_results(baselines)
    del gone["svg"][name]
    with pytest.raises(fidelity.KeptBaselineMismatch, match="not in this run"):
        fidelity.baselines_payload(gone, copy.deepcopy(baselines))


def test_pdfium_entries_are_rewritten_only_when_pdfium_was_scored(monkeypatch):
    import copy

    baselines = _recorded(monkeypatch)
    name = next(name for name, entry in baselines.items() if "pdfium" in entry)
    results = _svg_results(baselines)
    results["pdfium"] = {n: dict(e, ssim=0.5) for n, e in copy.deepcopy(results["svg"]).items()}
    payload = fidelity.baselines_payload(results, copy.deepcopy(baselines))
    assert payload[name]["pdfium"]["ssim"] == 0.5
    assert payload[name]["ssim"] == baselines[name]["ssim"]
