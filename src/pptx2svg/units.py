"""Moved to :mod:`ooxml_common.units`, with its history; this path keeps working.

The module registered under this name *is* the shared one, not a copy of its names, so
state, identity and monkeypatching are the same through either path.
"""

import sys

from ooxml_common import units as _shared

sys.modules[__name__] = _shared
