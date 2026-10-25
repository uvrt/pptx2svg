"""The names the sibling repositories import from pptx2svg: renaming one fails here.

Siblings' CI installs pptx2svg from ``main``, so a rename that breaks one shows only at its
next run, not on the pull request that made it.  ``USED_BY`` is that surface, each name with
the siblings that import it (in their ``src``, ``tests`` and ``tools``); docs/usage.md
("Used by pptx-agent") documents it, and ``test_the_documented_names_are_these`` holds the
two to the same names.

A sibling that starts importing a name: add it here and there.  Renaming one: keep the old
name as an alias for one release, with a comment saying who it is kept for (``OLD_NAMES``),
and move the sibling to the new one.
"""

from __future__ import annotations

import importlib
import re
from pathlib import Path

import pytest

DOC = Path(__file__).resolve().parent.parent / "docs/usage.md"
SECTION = "## Used by pptx-agent"

#: module -> {name: the siblings that import it}.
USED_BY: dict[str, dict[str, str]] = {
    "pptx2svg": {
        "ConvertOptions": "pptx-agent",
        "convert_pptx_to_agent_svg": "pptx-agent",
        "convert_pptx_to_model": "pptx-agent",
        "convert_pptx_to_png": "pptx-agent",
        "convert_pptx_to_svg": "pptx-agent",
        "png": "pptx-agent",
        "rasterise_svgs": "pptx-agent",
        "render_svgs": "pptx-agent",
        "table_row_heights": "pptx-agent",
        "user_font_dirs": "pptx-agent",
        "uses_host_fonts": "pptx-agent",
    },
    "pptx2svg.fonts": {
        "bundle_dir": "pptx-agent",
        "font_dirs": "pptx-agent",
    },
    "pptx2svg.fonts.check": {
        "resolved_families": "pptx-agent",
    },
    "pptx2svg.fonts.office": {
        "emoji_face": "pptx-agent",
        "search_dirs": "pptx-agent",
        "user_layout_metrics": "pptx-agent",
    },
    "pptx2svg.glyphs": {
        "system_font_dirs": "pptx-agent",
    },
    "pptx2svg.png": {
        "RasterizerNotAvailable": "pptx-agent",
        "svg_to_png": "pptx-agent",
    },
    "pptx2svg.resolve.fields": {
        "DATE_FORMATS": "pptx-agent",
        "NAMES": "pptx-agent",
        "date_format_index": "pptx-agent",
        "fallback_language": "pptx-agent",
        "format_date": "pptx-agent",
        "table_language": "pptx-agent",
    },
}

#: (module, old underscore name) -> the public name; aliases kept for one release.
OLD_NAMES: dict[tuple[str, str], str] = {
}


def _used() -> list[tuple[str, str]]:
    return [(module, name) for module, names in USED_BY.items() for name in names]


@pytest.mark.parametrize(("module", "name"), _used())
def test_a_name_a_sibling_imports_is_there(module, name):
    owner = importlib.import_module(module)
    if hasattr(owner, name):
        return
    try:
        importlib.import_module(f"{module}.{name}")     # a submodule not imported yet
    except ImportError:
        pytest.fail(f"{module}.{name} is gone, and {USED_BY[module][name]} import it: keep the "
                    "old name as an alias for one release and move them to the new one")


def test_siblings_import_no_private_name():
    private = [f"{module}.{name}" for module, name in _used() if name.startswith("_")]
    assert not private, f"give these a public name (keep the old one as an alias): {private}"


def test_an_old_name_is_the_public_one():
    for (module, old), new in OLD_NAMES.items():
        owner = importlib.import_module(module)
        assert getattr(owner, old) is getattr(owner, new), f"{module}.{old} is not {new}"
        assert new in USED_BY.get(module, {}), f"{module}.{new} is missing from USED_BY"


def _documented() -> set[tuple[str, str]]:
    lines = DOC.read_text(encoding="utf-8").splitlines()
    if SECTION is not None:
        start, depth = lines.index(SECTION) + 1, len(SECTION) - len(SECTION.lstrip("#"))
        end = next((i for i in range(start, len(lines)) if lines[i].startswith("#")
                    and len(lines[i]) - len(lines[i].lstrip("#")) <= depth), len(lines))
        lines = lines[start:end]
    found = set()
    for line in lines:
        if line.startswith("- `"):
            module, *names = re.findall(r"`([^`]+)`", line)
            found.update((module, name) for name in names)
    return found


def test_the_documented_names_are_these():
    documented, used = _documented(), set(_used())
    assert not used - documented, f"add to {DOC.name}: {sorted(used - documented)}"
    extra = sorted(documented - used)
    assert not extra, f"add to USED_BY, or take out of {DOC.name}: {extra}"
