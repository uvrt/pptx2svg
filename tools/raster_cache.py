"""A cache of rasters for the fidelity harness that cannot hide a change.

``tools/fidelity.py`` draws two rasters a slide on every run -- ours, from the SVG this
library writes, and PowerPoint's, from its PDF converted by ``tools/pdf_svg.py`` -- and
most runs draw exactly what the last one drew.  Caching them is a large share of a run.
It is also the one optimisation here that could make the harness *lie*: a raster served
from a cache whose key missed an input is last week's picture scored as today's.  So this
cache is built the other way round from a usual one, and every rule below is there to make
a stale entry impossible to use quietly:

* **The key is every input that affects the pixels**, spelled out by the caller as a dict
  of *components* (:func:`key`): the source's bytes (as a digest), the *content* digests
  of the font files the rasteriser is handed (an Office update can change a face without
  changing its path), the rasteriser's version and binary, the decoder's version, the
  resolution and options, and a harness version constant for anything else.
* **The components are stored beside the raster and compared on read**, field by field,
  along with a digest of the pixels themselves.  An entry whose stored components are not
  the ones asked for, or whose pixels are not the ones written, is not a miss: it is a
  corrupted cache (:class:`CacheMismatch`).
* **Detection.**  The key can only be as complete as its author's knowledge of what moves
  pixels.  So every run re-draws a sample of cached pages (:func:`rotating_sample`: a few
  pages, spread across the corpus, rotating by date so every page is re-drawn within a
  few days) and compares them with the cache byte for byte (:meth:`RasterCache.verify`).
  **Any mismatch discards the whole cache** and is raised, never warned about: the caller
  stops without reporting or recording a score.
* **Only outside any repository** (:func:`refuse_repository`).  A raster of PowerPoint's
  page holds Microsoft's glyph shapes, like the converted pages it is drawn from.
* **Atomic writes**: every file is written aside and renamed into place, so the parallel
  harness's workers never read half a raster.

Nothing here knows about PowerPoint, so the sibling ``docx2svg`` can take the file as it
is; its harness supplies its own components.  Needs numpy and Pillow (the rasters are
arrays, stored as lossless PNG).
"""

from __future__ import annotations

import datetime
import hashlib
import io
import json
import os
import shutil
from pathlib import Path


class CacheMismatch(Exception):
    """A cached raster that is not what re-drawing it gives, or not what was stored."""


def digest_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


def digest_file(path) -> str:
    """The sha256 of a file's *content*, read whole (never its path, size or date)."""
    digest = hashlib.sha256()
    with open(path, "rb") as handle:
        while chunk := handle.read(1 << 20):
            digest.update(chunk)
    return digest.hexdigest()


def digest_files(paths, threads: int = 8) -> dict[str, str]:
    """``{path: content digest}`` for ``paths``, read on a few threads (hashlib releases
    the GIL, and a font directory runs to hundreds of megabytes)."""
    from concurrent.futures import ThreadPoolExecutor

    paths = sorted(set(str(p) for p in paths))
    with ThreadPoolExecutor(max_workers=threads) as pool:
        return dict(zip(paths, pool.map(digest_file, paths)))


def key(components: dict) -> str:
    """The cache key of ``components``: a digest of their canonical JSON."""
    return digest_bytes(json.dumps(components, sort_keys=True, separators=(",", ":")).encode())


def refuse_repository(path: Path, repository: Path | None = None) -> None:
    """Raise unless ``path`` is outside ``repository`` and outside any git checkout."""
    resolved = Path(path).resolve()
    if repository is not None and resolved.is_relative_to(Path(repository).resolve()):
        raise RuntimeError(f"refusing to cache a raster inside the repository: {path}")
    for parent in (resolved, *resolved.parents):
        if (parent / ".git").exists():
            raise RuntimeError(f"refusing to cache a raster inside a git checkout ({parent}): {path}")


def rotating_sample(items: list, count: int, day: int | None = None) -> list:
    """``count`` of ``items``, evenly spaced through the list and shifted by ``day``.

    With ``stride = ceil(len / count)`` the sample is ``items[offset::stride]`` for
    ``offset = day % stride``, so a list ordered by deck is sampled across decks, and
    ``stride`` consecutive days between them visit every item.  ``day`` defaults to
    today's ordinal: the sample is deterministic within a day and rotates across days."""
    if not items or count <= 0:
        return []
    if day is None:
        day = datetime.date.today().toordinal()
    stride = -(-len(items) // count)
    return items[day % stride::stride]


def _png(array) -> bytes:
    from PIL import Image

    buffer = io.BytesIO()
    Image.fromarray(array).save(buffer, "PNG", compress_level=1)
    return buffer.getvalue()


def _array(data: bytes):
    import numpy as np
    from PIL import Image

    return np.asarray(Image.open(io.BytesIO(data)).convert("RGB"))


def pixel_digest(array) -> str:
    """A digest of a raster's shape and every byte of its pixels."""
    import numpy as np

    array = np.ascontiguousarray(array)
    return digest_bytes(f"{array.shape}|{array.dtype}|".encode() + array.tobytes())


class RasterCache:
    """Rasters under ``root``, one pair of files per entry: ``<kind>/<key>.png`` and
    ``<kind>/<key>.json`` (the components, the pixels' digest).  ``kind`` separates the
    sides (``"ours"``, ``"truth"``) so a report can say which side went stale."""

    def __init__(self, root: Path, repository: Path | None = None):
        self.root = Path(root)
        refuse_repository(self.root, repository)

    def _paths(self, kind: str, components: dict) -> tuple[Path, Path]:
        name = key(components)
        return self.root / kind / f"{name}.png", self.root / kind / f"{name}.json"

    def get(self, kind: str, components: dict):
        """The cached raster for ``components``, or ``None`` if there is none.

        Raises :class:`CacheMismatch` for an entry that is there but is not what was
        stored: components that differ from ``components`` (recomputed by the caller on
        every read), or pixels whose digest is not the one recorded with them."""
        png, meta = self._paths(kind, components)
        if not meta.exists() or not png.exists():
            return None  # the JSON is written last: without it the entry is not there
        stored = json.loads(meta.read_text(encoding="utf-8"))
        if stored.get("components") != components:
            raise CacheMismatch(f"{kind} {meta.name}: the stored key components are not the ones asked for "
                                f"({_first_difference(stored.get('components'), components)})")
        array = _array(png.read_bytes())
        if pixel_digest(array) != stored.get("pixels"):
            raise CacheMismatch(f"{kind} {png.name}: the stored pixels are not the ones written")
        return array

    def put(self, kind: str, components: dict, array, label: str = "") -> None:
        """Store ``array`` under ``components``.  ``label`` (say, which slide) is written
        beside them for a reader of the cache; it is not part of the key."""
        png, meta = self._paths(kind, components)
        png.parent.mkdir(parents=True, exist_ok=True)
        _write_atomically(png, _png(array))
        _write_atomically(meta, json.dumps({"components": components, "pixels": pixel_digest(array),
                                            "label": label}, indent=1, sort_keys=True).encode())

    def verify(self, kind: str, components: dict, fresh, store: bool = True, label: str = "") -> str:
        """Compare a freshly drawn raster with the cached one, byte for byte:
        ``"match"``, or ``"stored"`` where nothing was cached yet (``fresh`` is stored,
        unless not ``store``: then ``"absent"``).  Raises :class:`CacheMismatch` on any
        difference."""
        import numpy as np

        cached = self.get(kind, components)
        if cached is None:
            if not store:
                return "absent"
            self.put(kind, components, fresh, label)
            return "stored"
        if cached.shape != fresh.shape or cached.dtype != fresh.dtype or not np.array_equal(cached, fresh):
            differing = int((cached != fresh).any(axis=-1).sum()) if cached.shape == fresh.shape else "all"
            raise CacheMismatch(f"{kind} {key(components)[:16]}: the cached raster differs from a fresh one "
                                f"({differing} pixels; cached {cached.shape}, fresh {fresh.shape})")
        return "match"

    def discard(self) -> None:
        """Remove every cached raster (and nothing else)."""
        shutil.rmtree(self.root, ignore_errors=True)

    def size(self) -> tuple[int, int]:
        """``(entries, bytes)`` on disk."""
        entries = total = 0
        if self.root.exists():
            for path in self.root.rglob("*"):
                if path.is_file():
                    total += path.stat().st_size
                    entries += path.suffix == ".png"
        return entries, total


def _first_difference(stored, wanted) -> str:
    if not isinstance(stored, dict):
        return "no components stored"
    for name in sorted(set(stored) | set(wanted)):
        if stored.get(name) != wanted.get(name):
            return f"{name}: stored {str(stored.get(name))[:40]!r}, now {str(wanted.get(name))[:40]!r}"
    return "none"


def _write_atomically(path: Path, data: bytes) -> None:
    """Write aside and rename into place: a concurrent reader sees all of it or none."""
    partial = path.with_name(f".{path.name}.{os.getpid()}.tmp")
    try:
        partial.write_bytes(data)
        os.replace(partial, path)
    finally:
        partial.unlink(missing_ok=True)
