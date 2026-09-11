from __future__ import annotations

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
