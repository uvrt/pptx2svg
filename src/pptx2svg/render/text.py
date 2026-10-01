"""Moved to :mod:`ooxml_common.drawingml.textbody`, with its history; this path keeps working.

The module registered under this name *is* the shared one, not a copy of its names, so
state, identity and monkeypatching are the same through either path.
"""

import sys

from ooxml_common.drawingml import textbody as _shared

sys.modules[__name__] = _shared
