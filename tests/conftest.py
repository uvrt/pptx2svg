from __future__ import annotations

import sys
from pathlib import Path

import pytest

FIXTURE_DIR = Path(__file__).parent / "fixtures"


def fixture_paths() -> list[Path]:
    return sorted(FIXTURE_DIR.glob("*.pptx"))


@pytest.fixture(params=fixture_paths(), ids=lambda path: path.stem)
def pptx_path(request) -> Path:
    return request.param


@pytest.fixture(scope="session")
def product_page() -> Path:
    return FIXTURE_DIR / "real-product-page.pptx"


@pytest.fixture(scope="session")
def basic_theme() -> Path:
    return FIXTURE_DIR / "real-basic-theme.pptx"


@pytest.fixture(scope="session")
def authoring() -> Path:
    return FIXTURE_DIR / "authoring-integration.pptx"


@pytest.fixture(scope="session")
def college_template() -> Path:
    return FIXTURE_DIR / "real-college-template.pptx"


# The font bundle is a sibling distribution that a source checkout has not pip-installed.
# Putting it on the path here makes the suite exercise the configuration users actually
# get from `pip install 'pptx2svg[fonts]'`, rather than silently testing the degraded
# system-fonts path and calling it a pass.  Tests that need the *other* mode monkeypatch
# pptx2svg.fonts.bundle_dir instead of relying on the environment.
_BUNDLE_SRC = Path(__file__).resolve().parents[1] / "packages" / "pptx2svg-fonts" / "src"
if _BUNDLE_SRC.is_dir() and str(_BUNDLE_SRC) not in sys.path:
    sys.path.insert(0, str(_BUNDLE_SRC))
