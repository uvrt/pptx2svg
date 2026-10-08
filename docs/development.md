# Development

Back to the [README](../README.md); see also [CONTRIBUTING.md](../CONTRIBUTING.md).

```bash
pip install -e ../ooxml-common        # a checkout of https://github.com/uvrt/ooxml-common
pip install -e '.[dev]'
pip install -e packages/pptx2svg-fonts
pytest
```

A bare `pytest` runs serially, as CI does. To use every core, pass `-n` (pytest-xdist, in
the `dev` extra):

```bash
pytest -n auto
```

It collects, passes and skips exactly the tests the serial run does. Here `-n` means
`--dist loadgroup` (`tests/conftest.py`), which is `load` for every test but one marked
`@pytest.mark.powerpoint` -- one that drives PowerPoint, of which there are none today:
those all run on one worker, one at a time. Under any other `--dist` such a test fails
rather than races.

The format-neutral half of this library -- the OPC reader, units, the font modules and
the measured text metrics, and DrawingML: its value types, colour resolution, guides,
preset geometry, patterns, and the fill, outline, effect and geometry renderers; the
shape tree and text body readers, the shape, text body and group renderers, and the
charts, read and laid out -- lives in
[ooxml-common](https://github.com/uvrt/ooxml-common), extracted with its history so that
docx2svg can measure text with the same tables and draw DrawingML, charts and SmartArt
with the same code. It is a runtime dependency with no dependencies of its own, and not on
PyPI yet, so it is installed first. Every old import path (`pptx2svg.opc`,
`pptx2svg.text.metrics`, `pptx2svg.fonts`, `pptx2svg.render.fill`,
`pptx2svg.resolve.color`, `pptx2svg.resolve.chart`, `pptx2svg.render.text`, ...) still
works and returns the same module object, and every type in `pptx2svg.model` and
`pptx2svg.parse.source` that moved is the shared class.

`pptx2svg-fonts` is a sibling distribution in this repository and is not on PyPI yet, so
it is installed from the checkout rather than named as a dependency.

Tests run against real `.pptx` files in `tests/fixtures/` produced by PowerPoint, Google
Slides and python-pptx; `tests/fixtures/README.md` records where each came from.

Decks that are not ours to redistribute are not committed at all. `tests/conftest.py`
looks for those in the gitignored `scratch/`, and the tests that need one skip where it
is absent.

`tests/vrt/` holds a committed SVG of every slide of every fixture, and `tests/test_vrt.py`
fails when a render stops matching one. It is a regression net, not a correctness check --
it reports that the output *changed*, never that it is *right*, and the committed files
freeze today's bugs along with today's behaviour. Read `tests/vrt/README.md` before
trusting one, and rebaseline deliberately:

```bash
pytest tests/test_vrt.py --update-snapshots    # then read the diff
```

The metrics table, now `ooxml_common/text/metrics.py` in ooxml-common, is **generated**
from the font files in `packages/pptx2svg-fonts` by a tool that stayed here, and the tool
writes into whichever ooxml-common is installed -- the editable sibling checkout. Do not
edit the table by hand:

```bash
python3 tools/extract_font_metrics.py --check    # fails if the table has drifted
python3 tools/extract_font_metrics.py --write    # regenerate
```

A test runs `--check`, so a font update that is not accompanied by a regenerated table
fails the suite rather than silently making layout wrong. The same idea applies to the
two other generated tables: ooxml-common's `drawingml/preset_specs.py` and
`drawingml/presets.py` come from `tools/derive_preset_geometry.py`, and `parse/table_styles_builtin.py` from
`tools/derive_table_styles.py`, which measures the styles by rendering them through
PowerPoint. Edit the tool, not the table.

