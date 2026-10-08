# Changelog

pptx2svg has no tagged release yet; every change so far is on the way to 0.1.0.
The full history is the [merged pull requests](https://github.com/uvrt/pptx2svg/pulls?q=is%3Apr+is%3Amerged).

## 0.1.0 -- unreleased (in development since 2026-09-11)

- First port of pptx-glimpse: PPTX slides to SVG and PNG in pure Python, with the full
  placeholder, colour and text inheritance cascades, all ECMA-376 preset geometries,
  tables with PowerPoint's 74 built-in styles, and charts laid out from measurements.
- Embedded fonts read and used for measurement and drawing; the `pptx2svg-fonts` bundle
  of SIL OFL faces for reproducible PNG output; `pptx2svg fonts --check`.
- The format-neutral half (OPC, units, fonts, text metrics, DrawingML, charts, SmartArt,
  text bodies) moved to [ooxml-common](https://github.com/uvrt/ooxml-common) (#1, #5, #6, #9).
- Fidelity harness scored with one rasteriser on both sides, using every core (#2, #3, #4).
- Measured against PowerPoint: axis tick marks (#8), theme style fills and outlines (#10),
  colour transforms and run placement (#11), autofit on open (#12), kerning and Office's
  own faces by default on a Mac (#13, #14), Japanese face resolution and colour emoji (#15).
- The agent view: a compact per-slide SVG for a model to read (#16).
