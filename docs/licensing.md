# Licensing

Back to the [README](../README.md); the licence itself is [LICENSE](../LICENSE).

MIT. Originally ported from [pptx-glimpse](https://github.com/hirokisakabe/pptx-glimpse)
(MIT, © Hiroki Sakabe); since extended with geometry, charts, tables and text layout
checked against PowerPoint's own PDF export, PowerPoint's font-resolution and kerning rules,
the shared [ooxml-common](https://github.com/uvrt/ooxml-common) core, and an agent SVG view.

The `pptx2svg-fonts` distribution is a separate matter: its Python module is MIT, but the
font files it carries are each **SIL Open Font License 1.1**, redistributed byte-for-byte
as published on [Google Fonts](https://github.com/google/fonts) with their full licence
texts alongside them in `src/pptx2svg_fonts/licenses/`. Carlito, Caladea, Arimo, Tinos and
Cousine are © their respective Project Authors; Lato is © tyPoland Łukasz Dziedzic; Noto
Sans JP is © Adobe; Raleway is © the Raleway Project Authors. None has been subsetted,
renamed or otherwise modified, so the Reserved Font Name clauses are satisfied.

No Microsoft font is bundled, vendored or committed anywhere in this repository.

Neither is ECMA-376 itself. `pptx2svg/render/preset_specs.py` is *compiled* from the
standard's `presetShapeDefinitions.xml`, but that source is not redistributed here --
`tools/derive_preset_geometry.py` requires a copy you obtained yourself and verifies its
SHA-256 before reading it.

**The `.pptx` files in `tests/fixtures/` are not covered by the MIT grant above.** They are
third-party documents included as renderer inputs, and they remain their owners'. They came
from pptx-glimpse's own `shared-fixtures/` (MIT), so that rendering differences between the
two implementations can be compared directly; their provenance is recorded in
`tests/fixtures/FIXTURES-README.md`.

Decks that are *not* ours to redistribute are not committed at all -- not even as a test
input. `tests/conftest.py` looks for those in the gitignored `scratch/`, and the tests that
need one skip where it is absent. `real-college-template.pptx`, Dickinson College's public
sample deck, is the only such deck today; `tests/fixtures/README.md` says where to get it.

