from __future__ import annotations

import sys
from pathlib import Path

import pytest

FIXTURE_DIR = Path(__file__).parent / "fixtures"


def fixture_paths() -> list[Path]:
    return sorted(FIXTURE_DIR.glob("*.pptx"))


def pytest_addoption(parser) -> None:
    parser.addoption(
        "--update-snapshots",
        action="store_true",
        default=False,
        help=(
            "rewrite the committed SVG in tests/vrt/ from the current render instead of "
            "comparing against it. Read the diff before committing: every line of it is "
            "a change in what a user sees. See tests/vrt/README.md."
        ),
    )


@pytest.fixture(scope="session")
def update_snapshots(request) -> bool:
    """Whether ``--update-snapshots`` was passed; see :mod:`tests.test_vrt`."""
    return bool(request.config.getoption("--update-snapshots"))


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
def financial() -> Path:
    return FIXTURE_DIR / "real-financial-report.pptx"


@pytest.fixture(scope="session")
def chart_gallery() -> Path:
    """One chart type per slide -- see `tools/make_chart_gallery.py`, which writes it."""
    return FIXTURE_DIR / "chart-gallery.pptx"


#: Decks that are not ours to redistribute live outside the repository, in the gitignored
#: `scratch/` directory, so a checkout can still use one when the developer has a copy.
LOCAL_DIR = Path(__file__).resolve().parents[1] / "scratch"


@pytest.fixture(scope="session")
def college_template() -> Path:
    """Dickinson College's public sample deck -- **not committed**.

    It is a third-party document, so it is not in `tests/fixtures/`; see the licence note
    in README.md.  Tests that need it skip where it is absent, which is every machine but
    one.  Put a copy in `scratch/` to run them:

        https://www.dickinson.edu/download/downloads/id/1076/sample_powerpoint_slides.pptx
        sha256 ac7f2627645042190df3244cc25929f4b006d144fc2cac520e79ab376197bbbf
    """
    path = LOCAL_DIR / "real-college-template.pptx"
    if not path.is_file():
        pytest.skip(f"no {path}; see the fixture's docstring for where to get it")
    return path


# The font bundle is a sibling distribution that a source checkout has not pip-installed.
# Putting it on the path here makes the suite exercise the configuration users actually
# get from `pip install 'pptx2svg[fonts]'`, rather than silently testing the degraded
# system-fonts path and calling it a pass.  Tests that need the *other* mode monkeypatch
# pptx2svg.fonts.bundle_dir instead of relying on the environment.
_BUNDLE_SRC = Path(__file__).resolve().parents[1] / "packages" / "pptx2svg-fonts" / "src"
if _BUNDLE_SRC.is_dir() and str(_BUNDLE_SRC) not in sys.path:
    sys.path.insert(0, str(_BUNDLE_SRC))
