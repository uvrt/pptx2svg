# pptx2svg-fonts

Font data for [pptx2svg](https://github.com/uvrt/pptx2svg).  Install it with

```sh
pip install 'pptx2svg[fonts]'
```

rather than by name — the extra pins a compatible version.

## What is in here

| Family | Stands in for | Size | Licence |
| --- | --- | --- | --- |
| Noto Sans JP | MS Gothic, MS PGothic, Meiryo, Yu Gothic, MS Mincho, Yu Mincho | 9.6 MB | SIL OFL 1.1 |
| Lato | itself | 2.7 MB | SIL OFL 1.1 |
| Raleway | itself | 0.6 MB | SIL OFL 1.1 |

Everything else pptx2svg draws with — Carlito, Arimo, Tinos, Cousine, Caladea — is in the
main wheel.  These three are separate because Noto Sans JP alone is larger than that whole
bundle and most decks never touch Japanese text, and because Lato and Raleway are not
metric substitutes for anything: a deck that names Lato wants Lato.

## Licences

The Python module is MIT, like pptx2svg.  **The font files are not.**  Each family is
licensed under the SIL Open Font License 1.1 and its full licence text ships alongside
it in `src/pptx2svg_fonts/licenses/`:

* **Noto Sans JP** — `NotoSansJP-OFL.txt`.  Copyright 2014–2021 Adobe
  (<http://www.adobe.com/>), with Reserved Font Name 'Source'.
* **Lato** — `Lato-OFL.txt`.  Copyright (c) 2010–2014 by tyPoland Łukasz Dziedzic
  (<team@latofonts.com>), with Reserved Font Name "Lato".
* **Raleway** — `Raleway-OFL.txt`.  Copyright 2010 The Raleway Project Authors
  (<impallari@gmail.com>), with Reserved Font Name "Raleway".

The files are redistributed byte-for-byte as published on
[Google Fonts](https://github.com/google/fonts); none has been subsetted, renamed or
otherwise modified, so the Reserved Font Name clauses are satisfied.

## Debian

The same faces are packaged as `fonts-noto-cjk`, `fonts-lato` and `fonts-raleway`.  Those
work, but they are whatever version the distribution shipped; only this distribution pins
the exact files, and only pinned files give byte-identical PNGs across machines.
