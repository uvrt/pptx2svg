from __future__ import annotations

import os
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
    parser.addoption(
        "--pdfium",
        action="store_true",
        default=False,
        help=(
            "also run the tests marked pdfium, which score against the fidelity harness's old "
            "instrument (PowerPoint's PDF drawn by pdfium).  Scores default to the svg truth, "
            "as tools/fidelity.py does; this is the suite's `--truth both`."
        ),
    )


#: The one xdist group every test that drives PowerPoint joins.  PowerPoint is a single
#: instance per machine, and two exports through it at once interleave in its one window.
POWERPOINT_GROUP = "powerpoint"
#: How the controller is distributing tests, handed to its workers (which xdist starts
#: with ``dist`` reset to ``"no"``) through the environment they inherit.
_DIST_ENV = "PPTX2SVG_XDIST_DIST"


def pytest_cmdline_main(config) -> None:
    """Make ``-n`` mean ``--dist loadgroup`` unless another mode was asked for.

    Runs after pytest-xdist's own (``tryfirst``) hook, which turns ``-n`` into ``load``.
    ``loadgroup`` is ``load`` for every test without an ``xdist_group``, so this changes
    nothing today; it is what keeps a future PowerPoint test on one worker without every
    caller having to remember the flag."""
    if not hasattr(config.option, "dist") or hasattr(config, "workerinput"):
        return
    if config.option.dist == "load":  # also an explicit --dist load: indistinguishable here
        config.option.dist = "loadgroup"
    os.environ[_DIST_ENV] = config.option.dist


@pytest.hookimpl(tryfirst=True)  # before xdist's own, which reads the groups into node ids
def pytest_collection_modifyitems(config, items) -> None:
    """Keep tests that drive PowerPoint off concurrent workers.

    Today no test does: the suite reads PowerPoint's *exports* (``~/pptx2svg-oracle``,
    read-only and ``cache=False`` in ``test_pdf_svg.py``) and never launches it; exports
    are made by hand with ``tools/powerpoint_export_pdf.applescript``.  This is the rail
    for the first one that does.  Mark it ``@pytest.mark.powerpoint`` and it joins one
    xdist group, which ``loadgroup`` (what ``-n`` means here: :func:`pytest_cmdline_main`)
    runs on a single worker, one test at a time.  Any other distribution ignores groups,
    so under one such a test fails rather than races.
    """
    marked = [item for item in items if item.get_closest_marker("powerpoint")]
    if not marked or not hasattr(config, "workerinput"):
        return
    for item in marked:
        item.add_marker(pytest.mark.xdist_group(POWERPOINT_GROUP))
    if os.environ.get(_DIST_ENV) == "loadgroup":
        # A worker re-parses the command line, so under a bare -n it does not know the
        # controller is grouping, and would leave the group out of the node ids that
        # xdist's scheduler reads it from.
        config.option.loadgroup = True


def pytest_runtest_setup(item) -> None:
    if item.get_closest_marker("pdfium") and not item.config.getoption("--pdfium"):
        pytest.skip("scores against the old pdfium instrument: run with `pytest --pdfium`")
    if (hasattr(item.config, "workerinput") and item.get_closest_marker("powerpoint")
            and os.environ.get(_DIST_ENV) != "loadgroup"):
        pytest.fail("this test drives PowerPoint, which is one instance per machine: run it serially, "
                    "or in parallel with --dist loadgroup")


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


@pytest.fixture(scope="session")
def feature_sweep() -> Path:
    """One unexercised feature per slide -- see `tools/make_feature_sweep.py`.

    Six slides are features the coverage sweep fixed and seven are features it did not.
    A baseline on a pinned slide records what this library does today and is not a claim
    that it is right; `tests/fixtures/FIXTURES-README.md` says which is which.
    """
    return FIXTURE_DIR / "feature-sweep.pptx"


#: Decks that are not ours to redistribute live outside the repository, in the gitignored
#: `scratch/` directory, so a checkout can still use one when the developer has a copy.
LOCAL_DIR = Path(__file__).resolve().parents[1] / "scratch"


@pytest.fixture(scope="session")
def college_template() -> Path:
    """Dickinson College's public sample deck -- **not committed**.

    It is a third-party document, so it is not in `tests/fixtures/`; see the licence note
    in docs/licensing.md.  Tests that need it skip where it is absent, which is every machine but
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
