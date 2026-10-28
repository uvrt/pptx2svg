"""``pptx2svg.fonts`` forwards the names it shares with ``ooxml_common.fonts``: a patch of one
through this package must leave the shared module whole once undone.  pytest 9.2's
``monkeypatch`` reads the old value from the target's ``__dict__``, finds none for a
forwarded name, and undoes the patch by deleting it -- which, forwarded, deleted
``ooxml_common.fonts.bundle_dir`` and broke every later font call (NameError)."""

from __future__ import annotations

import pytest
from ooxml_common import fonts as shared

import pptx2svg.fonts as fonts


def test_undoing_a_patch_of_a_forwarded_name_keeps_the_shared_one():
    original = shared.bundle_dir
    with pytest.MonkeyPatch.context() as patch:
        patch.setattr("pptx2svg.fonts.bundle_dir", lambda: None)
        assert shared.bundle_dir() is None and fonts.bundle_mode() == "system"
    assert shared.bundle_dir is original
    assert fonts.bundle_dir is original
    assert fonts.bundle_mode() in ("bundled", "system")       # no NameError


def test_deleting_a_forwarded_name_restores_the_shared_original():
    original = shared.bundle_dir
    shared.bundle_dir = lambda: None
    try:
        del fonts.bundle_dir
        assert shared.bundle_dir is original
    finally:
        shared.bundle_dir = original
