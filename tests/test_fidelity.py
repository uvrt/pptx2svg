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
    ``ZipInfo`` precisely so that two runs agree.
    """
    import make_cjk_deck

    derived = tmp_path / "sample-cjk.pptx"
    replacements = make_cjk_deck.write_deck(FIXTURES / "sample.pptx", derived)
    assert replacements == 8, replacements
    assert derived.read_bytes() == (FIXTURES / "sample-cjk.pptx").read_bytes(), (
        "tests/fixtures/sample-cjk.pptx is not what tools/make_cjk_deck.py writes; "
        "regenerate it rather than editing either deck by hand"
    )


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
