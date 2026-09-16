"""Snapshot VRT: render every committed fixture and compare against committed SVG.

What this is
------------

A regression net, and nothing more.  It answers exactly one question -- *did the output
change?* -- by rendering `tests/fixtures/*.pptx` and comparing byte for byte against the
SVG committed under `tests/vrt/`.  It runs on every platform and every Python version CI
covers, needs no PowerPoint, no rasteriser and no network, and it is the only check in
this repository that sees a change to, say, chart axis placement on a deck nobody thought
to write a unit test for.

**It cannot tell you a render is correct.**  A snapshot is a photograph of what this code
did on the day someone ran `--update-snapshots`, bugs included, and `tests/vrt/README.md`
lists the bugs that are in the photograph.  Correctness comes from the other layer:
`tools/fidelity.py` scores our render against PowerPoint's own export, needs PowerPoint,
and therefore runs on one Mac.  Two layers, two jobs -- snapshots catch *regressions*, the
oracle catches *being wrong in the first place*.  Never resolve a snapshot failure by
rebaselining until you know which of the two you are looking at.

Why SVG and not PNG
-------------------

The SVG is our own output, produced by standard-library code alone.  A PNG additionally
encodes resvg's version and its antialiasing, so bumping `resvg-py` would read as a
rendering regression on every fixture at once -- a gate that cries wolf on a dependency
bump is a gate people learn to rebaseline through.  It is also 3x the bytes (587 kB of
PNG against 195 kB of SVG for this corpus) and, being already compressed, those bytes
never shrink in the pack file or in history.

What makes the comparison legitimate
------------------------------------

Byte equality is only a fair test if the bytes are a function of the input alone.  Five
things could break that, and each is handled here rather than hoped for:

``set`` iteration order
    ``PYTHONHASHSEED`` is randomised per process, so output derived from a ``set`` of
    strings can differ *between two runs on the same machine* -- and never within one, so
    rendering twice in one process proves nothing.
    :func:`test_a_second_process_with_a_different_hash_seed_renders_the_same_bytes`
    renders in two subprocesses under two fixed, different seeds.

Line endings
    Git on Windows can rewrite LF to CRLF on checkout, which would fail every snapshot on
    four of the twelve CI legs.  `.gitattributes` pins ``tests/vrt/**.svg`` to ``eol=lf``,
    and :func:`read_snapshot` normalises anyway, so a checkout with the attribute ignored
    still passes.  That normalisation cannot mask a real change, because
    :func:`test_the_render_emits_no_carriage_returns` asserts the render itself has none.

Encoding
    Every read and write here names UTF-8 explicitly.  ``Path.read_text()`` without an
    encoding follows the locale -- cp1252 on a Windows CI runner -- and this repository
    has already lost a run to that, on an em dash.  These snapshots carry Japanese.

Host fonts
    They cannot leak in.  The SVG names a ``font-family`` and gets its widths from the
    generated tables in `pptx2svg.text.metrics`; no font file is read on the SVG path, so
    the output is identical with and without the `pptx2svg-fonts` bundle installed --
    measured, and then nailed down by
    :func:`test_the_render_does_not_depend_on_the_font_bundle`.  (The bundle governs
    *PNG*, where the rasteriser needs real files.  That is one more reason the snapshot
    layer is SVG.)

Float formatting, and the libm underneath it
    Numbers are rounded to three decimals and formatted by Python's own dtoa, which is
    platform-independent; ``sin``/``cos``/``log10`` are not.  Perturbing every one of them
    by an ulp changes exactly one snapshot, through one expression, and
    :func:`test_the_platform_agrees_about_log10_of_a_power_of_ten` is the tripwire for it.

Rebaselining
------------

    python -m pytest tests/test_vrt.py --update-snapshots

Then read the diff before committing it: every line of it is a change in what users see.
The flag deliberately cannot produce a green run -- it skips, not passes -- so a CI leg
that somehow acquired it cannot report success.

On a failure the actual render is written to `tests/vrt/diffs/` (gitignored) so it can be
opened in a browser next to the committed one.
"""

from __future__ import annotations

import math
import os
import subprocess
import sys
from pathlib import Path

import pytest

from pptx2svg import ConvertOptions, convert_pptx_to_svg

#: The decks, globbed here rather than imported from `conftest`: a module-level
#: ``from tests.conftest import ...`` resolves under ``python -m pytest`` (which puts the
#: repository root on the path) and raises ``ModuleNotFoundError`` under a bare ``pytest``,
#: and an ImportError at module level interrupts collection of the whole suite rather than
#: failing one test.  The rule is one line; keep it the same line as `conftest`'s.
FIXTURE_DIR = Path(__file__).parent / "fixtures"

SNAPSHOT_DIR = Path(__file__).parent / "vrt"
#: Where a failing render is dumped for eyeballing.  Gitignored; see .gitignore.
DIFF_DIR = SNAPSHOT_DIR / "diffs"

REBASELINE = "python -m pytest tests/test_vrt.py --update-snapshots"

#: The deck the cross-process determinism check renders.  The largest in the corpus, and
#: the one with charts and CJK text -- the two subsystems that iterate over sets.
HASH_SEED_DECK = "real-financial-report.pptx"


def render(deck: Path) -> list[str]:
    """One SVG per slide, with the defaults a caller gets from `convert_pptx_to_svg`."""
    return convert_pptx_to_svg(deck, ConvertOptions())


def snapshot_path(deck: Path, slide_number: int) -> Path:
    return SNAPSHOT_DIR / deck.stem / f"slide-{slide_number:02d}.svg"


def read_snapshot(path: Path) -> str:
    """UTF-8, and tolerant of a checkout that ignored `.gitattributes`.

    Bytes rather than :meth:`Path.read_text`: the latter takes its encoding from the
    locale, and `newline=` only arrived in 3.13 while this package supports 3.10.
    """
    return path.read_bytes().decode("utf-8").replace("\r\n", "\n")


def write_snapshot(path: Path, document: str) -> None:
    """Write LF and UTF-8 on every platform.

    Bytes rather than text, so nothing translates ``\\n`` to ``os.linesep`` on the way
    out; a snapshot written on Windows must be the same file as one written on Linux.
    """
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(document.encode("utf-8"))


def describe_difference(expected: str, actual: str) -> str:
    """Point at the first differing character.

    A slide is a single very long line, so `difflib`'s line diff would print the whole
    document twice and highlight nothing.  The offset and a window around it is the
    useful form.
    """
    limit = min(len(expected), len(actual))
    offset = next((i for i in range(limit) if expected[i] != actual[i]), limit)
    start = max(0, offset - 70)
    return (
        f"first difference at character {offset} "
        f"(snapshot {len(expected)} chars, render {len(actual)} chars)\n"
        # `ascii()` rather than `!r`: these documents carry Japanese, and a failure
        # message must survive a cp1252 console on the Windows legs.
        f"  snapshot: ...{ascii(expected[start:offset + 70])}\n"
        f"  render:   ...{ascii(actual[start:offset + 70])}"
    )


def test_svg_matches_the_committed_snapshot(pptx_path: Path, update_snapshots: bool):
    """The gate: every slide of every committed fixture, byte for byte.

    A failure means the render changed.  It does *not* say whether the change is an
    improvement -- read the diff, and if the deck is one `tools/fidelity.py` can score,
    let the oracle decide before rebaselining.
    """
    documents = render(pptx_path)

    if update_snapshots:
        directory = SNAPSHOT_DIR / pptx_path.stem
        for stale in sorted(directory.glob("slide-*.svg")):
            stale.unlink()
        for number, document in enumerate(documents, start=1):
            write_snapshot(snapshot_path(pptx_path, number), document)
        pytest.skip(f"--update-snapshots: rewrote {len(documents)} snapshot(s)")

    directory = SNAPSHOT_DIR / pptx_path.stem
    committed = sorted(directory.glob("slide-*.svg"))
    assert committed, (
        f"no snapshots in {directory}: the fixture is new, or the directory was lost. "
        f"Capture them with `{REBASELINE}` and read the result before committing it."
    )
    assert len(committed) == len(documents), (
        f"{pptx_path.name} now renders {len(documents)} slide(s), "
        f"{directory} holds {len(committed)}"
    )

    failures = []
    for number, document in enumerate(documents, start=1):
        path = snapshot_path(pptx_path, number)
        expected = read_snapshot(path)
        if document != expected:
            dump = DIFF_DIR / pptx_path.stem / path.name
            write_snapshot(dump, document)
            failures.append(f"{path}\n{describe_difference(expected, document)}\n  render written to {dump}")

    assert not failures, (
        f"{len(failures)} slide(s) of {pptx_path.name} no longer match their snapshot:\n\n"
        + "\n\n".join(failures)
        + f"\n\nIf the change is intended, rebaseline with `{REBASELINE}`."
    )


def test_every_committed_snapshot_has_a_fixture():
    """Catch a snapshot directory left behind by a renamed or deleted deck.

    Without this a stale directory is invisible: nothing renders it, so nothing compares
    it, and it sits in the repository looking like coverage.
    """
    expected = {path.stem for path in sorted(FIXTURE_DIR.glob("*.pptx"))}
    found = {
        directory.name
        for directory in SNAPSHOT_DIR.iterdir()
        if directory.is_dir() and directory.name != "diffs"
    }
    assert found == expected, (
        f"snapshot directories without a fixture: {sorted(found - expected)}; "
        f"fixtures without snapshots: {sorted(expected - found)}"
    )


def test_the_render_emits_no_carriage_returns(pptx_path: Path):
    """The premise :func:`read_snapshot`'s CRLF normalisation rests on.

    If a deck ever made the renderer emit a bare ``\\r``, normalising line endings on the
    way in would quietly erase that difference and the gate would stop being byte-exact.
    """
    assert not any("\r" in document for document in render(pptx_path))


def test_the_render_does_not_depend_on_the_font_bundle(monkeypatch, authoring, financial):
    """A snapshot must not encode whether `pptx2svg-fonts` happens to be installed.

    If it did, the committed file would be a picture of *this machine's* Helvetica, and
    the right answer would be to skip the suite wherever the bundle is absent rather than
    snapshot a host-font render.  It does not, and the reason is structural: the SVG path
    opens no font file.  Widths come from the generated tables in `pptx2svg.text.metrics`
    and the ``font-family`` is a name for a rasteriser to resolve later, so the bundle
    governs *PNG* and nothing here.  Measured across the whole corpus before these
    snapshots were captured -- with the bundle and without it, every one of the sixteen
    documents is identical to the byte; only `ConvertOptions.warnings` differ, which is
    why warnings are not snapshotted.

    Two decks rather than the corpus, to keep the check cheap: the one `tools/fidelity.py`
    scores, and the one whose text is mostly Japanese -- East Asian text being where a
    font lookup would creep back in.  The two renders are compared against each other
    rather than against the committed file, so an unrelated rendering change fails the
    gate above and not here.
    """
    for deck in (authoring, financial):
        with_bundle = render(deck)
        monkeypatch.setattr("pptx2svg.fonts.bundle_dir", lambda: None)
        without_bundle = render(deck)
        monkeypatch.undo()
        assert len(with_bundle) == len(without_bundle)
        for number, (bundled, bare) in enumerate(zip(with_bundle, without_bundle), 1):
            assert bundled == bare, (
                f"{deck.name} slide {number} renders differently without the font "
                f"bundle:\n{describe_difference(bundled, bare)}"
            )


def test_the_platform_agrees_about_log10_of_a_power_of_ten():
    """The one place a 1-ulp libm difference could rewrite a whole slide.

    Measured rather than reasoned about: perturbing *every* `math` transcendental in the
    render by one ulp -- systematically, in one direction, which is harsher than any real
    libm disagreement -- moves not a byte of fifteen of the sixteen snapshots.  The
    exception is `real-financial-report/slide-04.svg`, and perturbing one function at a
    time narrows it to `log10` alone, through

        unit = 10.0 ** math.floor(math.log10(span))     # resolve/chart.py

    That slide's radar spans exactly 0..100, so ``log10`` returns exactly 2.0 and the axis
    unit is 100.  One ulp low and ``floor`` gives 1, the unit becomes 10, and the chart
    re-lays out: the document grows from 29,659 characters to 34,639.  `floor` after a
    transcendental has no small errors -- it has no error at all, or a factor of ten.

    Every mainstream libm returns 2.0 here, because the true value is exactly
    representable and correct rounding therefore requires it; only a merely *faithful*
    implementation is allowed to return the value below.  This test is the tripwire, so
    that on a platform where that is not true the failure says which platform disagreed
    about what, instead of one unexplained snapshot mismatch on one CI leg.  The durable
    fix is in `resolve/chart.py`, not here: that expression amplifies an ulp into a
    different chart and should not.
    """
    for exponent in range(-3, 7):
        power = 10.0**exponent
        assert math.floor(math.log10(power)) == exponent, (
            f"this platform's log10({power!r}) is {math.log10(power)!r}, "
            f"not {float(exponent)!r}: chart axis units are computed from "
            f"floor(log10(span)) and will differ here"
        )


def test_a_second_process_with_a_different_hash_seed_renders_the_same_bytes():
    """The one hazard that a single process cannot see.

    `PYTHONHASHSEED` is randomised per process, so anything that derives output order
    from a ``set`` of strings is stable *within* a run and varies *between* runs.  Two
    subprocesses under two fixed, different seeds is the only honest test; rendering
    twice in this process would pass with the bug present.

    Compared against the committed snapshot too, so the check is not merely "both
    children agree with each other".
    """
    deck = FIXTURE_DIR / HASH_SEED_DECK
    script = (
        "import sys;"
        "from pptx2svg import ConvertOptions, convert_pptx_to_svg;"
        "sys.stdout.buffer.write("
        "chr(30).join(convert_pptx_to_svg(sys.argv[1], ConvertOptions())).encode('utf-8'))"
    )

    def run(seed: str) -> str:
        env = dict(os.environ, PYTHONHASHSEED=seed, PYTHONIOENCODING="utf-8")
        # The child needs the same import path this process has: a source checkout runs
        # from src/ via PYTHONPATH, and conftest.py adds the font bundle at runtime.
        env["PYTHONPATH"] = os.pathsep.join(path for path in sys.path if path)
        completed = subprocess.run(
            [sys.executable, "-c", script, str(deck)],
            check=False,
            capture_output=True,
            env=env,
        )
        assert completed.returncode == 0, (
            f"the child render died under PYTHONHASHSEED={seed}:\n"
            + completed.stderr.decode("utf-8", "replace")
        )
        return completed.stdout.decode("utf-8")

    first = run("0")
    second = run("524287")
    assert first == second, (
        "the render depends on PYTHONHASHSEED -- look for a set() of strings reaching "
        f"the output:\n{describe_difference(first, second)}"
    )

    committed = chr(30).join(
        read_snapshot(path)
        for path in sorted((SNAPSHOT_DIR / deck.stem).glob("slide-*.svg"))
    )
    assert first == committed, (
        "a subprocess render disagrees with the committed snapshot:\n"
        + describe_difference(committed, first)
    )
