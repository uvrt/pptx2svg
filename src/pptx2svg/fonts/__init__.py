"""Moved to :mod:`ooxml_common.fonts`, with its history; this path keeps working.

Unlike the other moved modules this one cannot simply *be* the shared module, because
it is still a package here: :mod:`pptx2svg.fonts.check` keeps the half of the font check
that reads a deck.  So every name the shared module defines is read from, written to and
deleted on the shared module.  Writes matter as much as reads -- a test that monkeypatches
``pptx2svg.fonts.bundle_dir`` has to reach the ``bundle_dir`` that ``bundle_mode`` and
``available_families`` actually call, which is the shared one.
"""

from __future__ import annotations

import sys
import types

from ooxml_common import fonts as _shared
from ooxml_common.fonts import __all__  # noqa: F401  -- `from pptx2svg.fonts import *`


def _is_shared(name: str) -> bool:
    # Submodules are the exception: the import system sets `pptx2svg.fonts.check` on
    # this package, and that must not overwrite `ooxml_common.fonts.check`.
    return (
        not name.startswith("__")
        and name in vars(_shared)
        and not isinstance(vars(_shared)[name], types.ModuleType)
    )


class _SharedFonts(types.ModuleType):
    def __getattr__(self, name: str):
        try:
            return getattr(_shared, name)
        except AttributeError:
            raise AttributeError(f"module {__name__!r} has no attribute {name!r}") from None

    def __setattr__(self, name: str, value) -> None:
        if _is_shared(name):
            setattr(_shared, name, value)
        else:
            super().__setattr__(name, value)

    def __delattr__(self, name: str) -> None:
        if _is_shared(name):
            delattr(_shared, name)
        else:
            super().__delattr__(name)

    def __dir__(self):
        return sorted(set(super().__dir__()) | set(dir(_shared)))


sys.modules[__name__].__class__ = _SharedFonts

# Loaded on `import pptx2svg` before the move, as a side effect of embedded.py's own
# imports; kept loaded so `pptx2svg.fonts.eot` and friends are there without an import.
# By importlib rather than `from . import`, which would find the shared modules through
# the forwarding above and never register these names.
import importlib as _importlib  # noqa: E402

for _name in ("eot", "mtx", "sfnt"):
    _importlib.import_module(f"{__name__}.{_name}")
