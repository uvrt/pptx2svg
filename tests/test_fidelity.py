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


def test_slide_count_matches_the_package():
    assert fidelity.slide_count(FIXTURES / "real-basic-theme.pptx") == 2
    assert fidelity.slide_count(FIXTURES / "sample.pptx") == 6


def test_font_profile_partitions_faces_and_is_hashable():
    profile = fidelity.font_profile(FIXTURES / "real-basic-theme.pptx")
    assert set(profile) == {"available", "missing", "hash"}
    assert not set(profile["available"]) & set(profile["missing"])
    assert len(profile["hash"]) == 12
    # Stable across calls, or a baseline could never be matched to its inputs.
    assert profile["hash"] == fidelity.font_profile(FIXTURES / "real-basic-theme.pptx")["hash"]


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
    baselines = json.loads(path.read_text())
    assert baselines
    for name, entry in baselines.items():
        assert "ssim" in entry, name
        assert "histogram" in entry, name
        assert entry["fonts"]["hash"], name
