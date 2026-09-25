"""Moved to :mod:`ooxml_common.fonts.sfnt`, with its history; this path keeps working.

The module registered under this name *is* the shared one, not a copy of its names, so
state, identity and monkeypatching are the same through either path.
"""

import sys

from ooxml_common.fonts import sfnt as _shared

sys.modules[__name__] = _shared
