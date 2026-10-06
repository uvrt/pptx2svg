"""Converting works whatever imported ``ooxml_common.fonts``' submodules first.

``pptx2svg.fonts`` forwards every name it lacks to ``ooxml_common.fonts``.  Once
``ooxml_common.fonts.office`` was imported -- docx2svg imports it, so any process that
loaded docx2svg before pptx2svg did -- ``from .fonts import office`` found the *shared*
module through that forwarding instead of ``pptx2svg.fonts.office``, and the first
conversion raised ``AttributeError: module 'ooxml_common.fonts.office' has no attribute
'available'``.  Each case runs in a fresh interpreter, because import order is
process state.
"""

from __future__ import annotations

import subprocess
import sys
import textwrap
from pathlib import Path

import pytest

FIXTURE = Path(__file__).parent / "fixtures" / "table test.pptx"

FIRST = {
    "the shared office lookup": "import ooxml_common.fonts.office",
    "every shared font module": (
        "import ooxml_common.fonts.office, ooxml_common.fonts.check, "
        "ooxml_common.fonts.embedded, ooxml_common.fonts.eot, ooxml_common.fonts.sfnt"
    ),
    # docx2svg's own spelling, without needing docx2svg installed.
    "docx2svg's imports": (
        "from ooxml_common.fonts import office as _office\n"
        "from ooxml_common.fonts.office import Face, FontError"
    ),
}


@pytest.mark.parametrize("first", FIRST.values(), ids=FIRST.keys())
@pytest.mark.parametrize("host_fonts", [None, False])
def test_a_conversion_after_the_shared_modules_were_imported(first, host_fonts):
    script = textwrap.dedent(
        f"""
        {{first}}
        import pptx2svg
        from pptx2svg.fonts import check, office
        assert office.__name__ == "pptx2svg.fonts.office", office.__name__
        assert check.__name__ == "pptx2svg.fonts.check", check.__name__
        assert callable(office.available)
        svgs = pptx2svg.convert_pptx_to_svg({str(FIXTURE)!r}, pptx2svg.ConvertOptions(host_fonts={host_fonts!r}))
        assert svgs and svgs[0].startswith("<svg"), svgs[:1]
        """
    ).replace("{first}", first)
    done = subprocess.run([sys.executable, "-c", script], capture_output=True, text=True, timeout=300)
    assert done.returncode == 0, done.stderr[-2000:]


def test_the_shared_modules_stay_reachable_where_this_package_has_none():
    """The forwarding still answers for a shared submodule pptx2svg has no file for."""
    script = (
        "import ooxml_common.fonts.office, pptx2svg.fonts as f\n"
        "assert f.eot.__name__ == 'ooxml_common.fonts.eot'\n"
        "assert f.embedded.__name__ == 'ooxml_common.fonts.embedded'\n"
        "assert f.office.__name__ == 'pptx2svg.fonts.office'\n"
    )
    done = subprocess.run([sys.executable, "-c", script], capture_output=True, text=True, timeout=120)
    assert done.returncode == 0, done.stderr[-2000:]
