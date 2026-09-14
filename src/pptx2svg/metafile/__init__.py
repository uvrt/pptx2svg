"""Metafile (EMF/WMF) preview extraction.

PowerPoint stores pasted vector artwork -- Illustrator and Visio drawings, equations,
clip art -- as an Enhanced Metafile.  Rendering one properly means interpreting GDI
drawing records, which is a large piece of work.  It is almost always unnecessary:
Office writes a *ready-made preview* into the very same file, either an embedded PDF
(for artwork that came from a PostScript-flavoured source) or an embedded DIB bitmap.
Pulling that out needs no drawing-record interpretation at all, only a walk over the
record list, skipping each record by its declared size.

So this package deliberately does not interpret geometry.  See :mod:`.emf_preview` for
the record walk and :mod:`.dib` for turning a device-independent bitmap into a PNG with
nothing but the standard library.
"""

from .emf_preview import MetafilePreview, extract_metafile_preview

__all__ = ["MetafilePreview", "extract_metafile_preview"]
