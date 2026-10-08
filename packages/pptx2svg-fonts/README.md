# pptx2svg-fonts

Font data for [pptx2svg](https://github.com/uvrt/pptx2svg).  Install it with

```sh
pip install 'pptx2svg[fonts]'
```

rather than by name — the extra pins a compatible version.

## What is in here

| Family | Stands in for | Licence |
| --- | --- | --- |
| Carlito | Calibri, Calibri Light (same advance widths) | SIL OFL 1.1 |
| Arimo | Arial, Helvetica, Liberation Sans (same advance widths) | SIL OFL 1.1 |
| Tinos | Times New Roman, Times, Liberation Serif (same advance widths) | SIL OFL 1.1 |
| Cousine | Courier New, Courier, Liberation Mono (same advance widths) | SIL OFL 1.1 |
| Caladea | Cambria (drawn only: its widths differ) | SIL OFL 1.1 |
| Noto Sans JP | Japanese and Chinese text: MS Gothic, Meiryo, Yu Gothic, MS Mincho, SimSun... (no Hangul) | SIL OFL 1.1 |
| Lato | itself | SIL OFL 1.1 |
| Raleway | itself | SIL OFL 1.1 |

pptx2svg draws its PNGs from these files and nothing else when this package is installed,
so the same deck gives the same pixels on every machine.  Without it the host's fonts are
used, and text no host font can draw is left out of the image: pptx2svg warns
(`glyphs-missing`) when that happens.  Which face stands in for which, and how well:
[FONTS.md](../../FONTS.md).

## Licences

The Python module is MIT, like pptx2svg.  **The font files are not.**  Each family is
licensed under the SIL Open Font License 1.1 and its full licence text ships alongside
it in `src/pptx2svg_fonts/licenses/` (`<Family>-OFL.txt`), with its copyright line --
among them:

* **Noto Sans JP** — Copyright 2014–2021 Adobe (<http://www.adobe.com/>), with Reserved
  Font Name 'Source'.
* **Carlito** — Copyright 2013 The Carlito Project Authors, with Reserved Font Name
  "Carlito".
* **Lato** — Copyright (c) 2010–2014 by tyPoland Łukasz Dziedzic
  (<team@latofonts.com>), with Reserved Font Name "Lato".
* **Raleway** — Copyright 2010 The Raleway Project Authors (<impallari@gmail.com>), with
  Reserved Font Name "Raleway".
* **Arimo, Tinos, Cousine** (The Arimo, Tinos and Cousine Project Authors) and
  **Caladea** (The Caladea Project Authors).

The files are redistributed byte-for-byte as published on
[Google Fonts](https://github.com/google/fonts); none has been subsetted, renamed or
otherwise modified, so the Reserved Font Name clauses are satisfied.

## Debian

The same faces are packaged as `fonts-noto-cjk`, `fonts-lato` and `fonts-raleway`.  Those
work, but they are whatever version the distribution shipped; only this distribution pins
the exact files, and only pinned files give byte-identical PNGs across machines.
