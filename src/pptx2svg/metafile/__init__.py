"""Metafile (EMF/WMF) preview extraction.

PowerPoint stores pasted vector artwork -- Illustrator and Visio drawings, equations,
clip art -- as an Enhanced Metafile.  Rendering one properly means interpreting GDI
drawing records, which is a large piece of work.  It is almost always unnecessary:
Office writes a *ready-made preview* into the very same file, either an embedded PDF
(for artwork that came from a PostScript-flavoured source) or an embedded DIB bitmap.
Pulling that out needs no drawing-record interpretation at all, only a walk over the
record list, skipping each record by its declared size.

This package extracts that preview (:mod:`.emf_preview`, the record walk) and rasterises
a PDF one (:mod:`.pdf`).  The records themselves are drawn by
:mod:`ooxml_common.metafile`'s GDI player, which the resolver tries first (round 7,
``pptx2svg.resolve.view._metafile_preview``); :mod:`.dib` now lives there too.
"""

from .emf_preview import MetafilePreview, extract_metafile_preview

__all__ = ["MetafilePreview", "extract_metafile_preview"]
