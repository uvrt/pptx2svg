# Checking fidelity against PowerPoint

Local-only: this needs PowerPoint's own PDF exports and the Microsoft faces Office installed, so it never runs on CI. Back to the [README](../README.md).

`tools/fidelity.py` scores our render against PowerPoint's own PDF export on SSIM and
colour-histogram correlation. It renders **our** side with the same licensed Microsoft
faces PowerPoint used, read in place from wherever Office installed them, so a difference
between the two images is attributable to this library rather than to font availability:

```bash
python3 tools/fidelity.py --write-profile               # once, on a machine with Office
.venv/bin/python tools/fidelity.py --oracle ~/pptx2svg-oracle   # the .venv: see below
```

The profile records paths and hashes only; no licensed font is ever copied into the
repository, and `tests/font-profile.local.json` is gitignored. Without it the harness
refuses to run, and a deck naming a face PowerPoint did not have either is skipped rather
than scored against Microsoft's fallback.

Both sides are rasterised by the **same engine**, resvg: PowerPoint's PDF page is first
converted to SVG by `tools/pdf_svg.py` (glyphs redrawn unhinted from the fonts the PDF
embeds, and validated against the PDF page by page), so a score measures layout and
drawing rather than pdfium against resvg. `--truth pdfium`, the old instrument, stays
available, and runs only when asked for (`--truth pdfium`, or `--truth both`): the
baselines keep its recorded scores beside the svg ones, a default `--update` carries them
over untouched, and `--update --truth both` re-records them. The converter needs PyMuPDF, a development-only tool behind the `fidelity`
extra, which is **not** part of `dev` and is not installed by CI; install it into a
project virtual environment and run the harness from there:

```bash
python3 -m venv --system-site-packages .venv
.venv/bin/pip install -e '.[fidelity]'        # or just: .venv/bin/pip install pymupdf
.venv/bin/python tools/fidelity.py            # --truth pdfium / --truth both for the old instrument
.venv/bin/python tools/pdf_svg.py --validate  # hold the converter to every export
.venv/bin/python -m pytest                    # the converter's tests run here, and skip elsewhere
.venv/bin/python -m pytest --pdfium           # also the tests that score against pdfium
```

Both tools use every logical core (fewer if memory is short): the harness scores a slide
per process and the validation takes a page per process, and both reassemble the results
in order, so every score, baseline and printed line is the serial run's. `--jobs 1` is
the serial path. Scoring never launches PowerPoint; it reads the exports.

The harness caches both sides' rasters in `~/pptx2svg-oracle/svg/rasters/` (never in the
repository), keyed by every input that moves their pixels -- the SVG and PDF bytes, the
*contents* of every font file resvg is handed, the converter, resvg and Pillow versions
and the options -- and holds itself to them: each run re-draws a few slides, rotating
through the corpus by date, and compares them with the cache byte for byte. Any
difference discards the whole cache and stops the run. `--verify-cache` re-draws every
slide and compares; `--no-cache` bypasses the cache. SSIM is computed over the content's
bounding box only, which gives the full page's score bit for bit; the same sampled slides
are scored both ways each run and held equal.

PyMuPDF is AGPL-3.0, acceptable for a local tool that is never distributed with the
library; nothing under `src/` imports it, and without it the tests that need it skip.
Converted pages carry the glyph outlines PowerPoint embedded (Microsoft's fonts), so they
are cached only beside PowerPoint's exports, in `~/pptx2svg-oracle/svg/`, and
`pdf_svg.py` refuses to write one inside a repository.

If Microsoft PowerPoint is installed, `tools/powerpoint_export_pdf.applescript` exports a
deck through PowerPoint itself, giving authoritative ground truth to compare against:

```bash
osascript tools/powerpoint_export_pdf.applescript "$PWD/deck.pptx" "$HOME/gt.pdf"
python -c "import pypdfium2 as p; d=p.PdfDocument('$HOME/gt.pdf'); \
           d[0].render(scale=2).to_pil().save('gt.png')"
```

PowerPoint is sandboxed, and being under your home directory is not enough: both paths
must be in a directory PowerPoint has already been granted access to. A brand-new one is
refused, and the failure looks like a corrupt deck rather than a permissions problem.

