"""pptx2svg's module state used from several threads at once.

With the GIL off (Python 3.15t) the picture cache in :mod:`pptx2svg.iccimages` raised
``KeyError`` reading back an entry another thread had just cleared, and PDFium, which is
not thread-safe, was entered by two threads at once.  Each test runs eight threads for at
most about a second and expects no error and the serial answer every time; under a GIL
build they pass either way, and quickly.
"""

from __future__ import annotations

import base64
import threading
import time

import pytest

from pptx2svg import iccimages
from pptx2svg.metafile.pdf import pdf_rasterizer_available, rasterise_pdf

THREADS = 8
SECONDS = 1.0
CALLS = 20_000

#: As tests/test_metafile.py's: one page, 72 x 36 pt, one red rectangle.
MINI_PDF = b"""%PDF-1.4
1 0 obj<</Type/Catalog/Pages 2 0 R>>endobj
2 0 obj<</Type/Pages/Kids[3 0 R]/Count 1>>endobj
3 0 obj<</Type/Page/Parent 2 0 R/MediaBox[0 0 72 36]/Contents 4 0 R/Resources<<>>>>endobj
4 0 obj<</Length 24>>stream
1 0 0 rg 0 0 72 36 re f
endstream
endobj
trailer<</Size 5/Root 1 0 R>>
%%EOF"""


def hammer(call, keys: int, calls: int = CALLS) -> list[str]:
    """Run ``call(key)`` from :data:`THREADS` threads over ``keys`` keys, each thread
    starting at a different key, until ``calls`` calls each or :data:`SECONDS`; the
    errors raised, as text."""
    errors: list[str] = []
    start = threading.Barrier(THREADS)

    def run(seed: int) -> None:
        start.wait()
        deadline = time.monotonic() + SECONDS
        for k in range(calls):
            if k % 8 == 0 and time.monotonic() > deadline:
                return
            try:
                call((seed * 5 + k) % keys)
            except Exception as error:  # noqa: BLE001 -- any error is the failure
                errors.append(repr(error))
                return

    threads = [threading.Thread(target=run, args=(seed,)) for seed in range(THREADS)]
    for thread in threads:
        thread.start()
    for thread in threads:
        thread.join()
    return errors


def test_icc_picture_cache_from_threads():
    pages = []
    for k in range(iccimages._CACHE_LIMIT + 6):
        png = iccimages.encode_png(k + 1, 1, bytes(3 * (k + 1)), 3)
        pages.append(f'<svg><image href="data:image/png;base64,{base64.b64encode(png).decode()}"/></svg>')
    expected = [iccimages.srgb_images(page) for page in pages]

    def call(key: int) -> None:
        assert iccimages.srgb_images(pages[key]) == expected[key]

    assert hammer(call, len(pages)) == []


@pytest.mark.skipif(not pdf_rasterizer_available(), reason="pypdfium2 not installed")
def test_pdf_previews_from_threads():
    widths = [914400 * (k + 1) // 4 for k in range(4)]
    expected = [rasterise_pdf(MINI_PDF, width_emu=width) for width in widths]
    assert all(expected)

    def call(key: int) -> None:
        assert rasterise_pdf(MINI_PDF, width_emu=widths[key]) == expected[key]

    assert hammer(call, len(widths), calls=200) == []


def test_table_rows_from_threads_in_a_fresh_interpreter():
    """The first ``table_row_heights`` calls of a process, from several threads at once: the
    table layout it reaches through ``pptx2svg.render.shape`` -- an alias module, which
    replaces itself in ``sys.modules`` with ooxml-common's -- is imported with pptx2svg,
    not by the first call, so no thread is handed the stand-in (``cannot import name
    '_row_heights'``, seen with six builders measuring tables at once)."""
    import subprocess
    import sys

    script = (
        "import sys, threading\n"
        "import pptx2svg\n"
        "assert 'pptx2svg.render.shape' in sys.modules\n"
        "from pptx2svg import model as m\n"
        "errors = []\n"
        "def rows():\n"
        "    try:\n"
        "        pptx2svg.table_row_heights(m.TableData(rows=[m.TableRow(height=100.0)]))\n"
        "    except BaseException as error:\n"
        "        errors.append(error)\n"
        "threads = [threading.Thread(target=rows) for _ in range(8)]\n"
        "[thread.start() for thread in threads]\n"
        "[thread.join() for thread in threads]\n"
        "assert not errors, errors\n"
    )
    subprocess.run([sys.executable, "-c", script], check=True, timeout=120)
