"""Open Packaging Conventions reader: the ZIP container, content types and relationships.

A .pptx is a ZIP of XML "parts" wired together by relationship files.  Every part
``word/foo.xml`` has its relationships in the sibling ``word/_rels/foo.xml.rels``, and
references inside the XML are by relationship id (``r:embed="rId3"``) rather than by
path -- so resolving a picture means: part path -> rels file -> target -> normalise
against the part's directory.
"""

from __future__ import annotations

import posixpath
import zipfile
from dataclasses import dataclass, field
from typing import BinaryIO

from .xmlutil import attr, children, local_name, parse_xml

CONTENT_TYPES_PART = "[Content_Types].xml"

REL_SLIDE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
REL_SLIDE_LAYOUT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
)
REL_SLIDE_MASTER = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
)
REL_THEME = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
REL_IMAGE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
REL_CHART = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
REL_DIAGRAM_DRAWING = (
    "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
)
REL_FONT = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
REL_HYPERLINK = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
REL_TABLE_STYLES = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
)
REL_OFFICE_DOCUMENT = (
    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
)

#: Extension -> MIME type, used when [Content_Types].xml has no explicit override.
DEFAULT_MEDIA_TYPES: dict[str, str] = {
    "png": "image/png",
    "jpg": "image/jpeg",
    "jpeg": "image/jpeg",
    "jpe": "image/jpeg",
    "gif": "image/gif",
    "bmp": "image/bmp",
    "tif": "image/tiff",
    "tiff": "image/tiff",
    "svg": "image/svg+xml",
    "webp": "image/webp",
    "ico": "image/x-icon",
    "emf": "image/emf",
    "wmf": "image/wmf",
}


@dataclass(frozen=True)
class Relationship:
    id: str
    type: str
    target: str
    #: Normalised package path, or ``None`` for External targets (hyperlinks).
    target_part: str | None
    is_external: bool


@dataclass
class OpcPackage:
    """A loaded .pptx, addressed by normalised part path (no leading slash)."""

    parts: dict[str, bytes]
    content_types: dict[str, str] = field(default_factory=dict)
    default_types: dict[str, str] = field(default_factory=dict)
    _rels_cache: dict[str, dict[str, Relationship]] = field(default_factory=dict, repr=False)

    # -- loading ---------------------------------------------------------------

    @classmethod
    def open(cls, source: str | bytes | BinaryIO) -> "OpcPackage":
        if isinstance(source, bytes):
            import io

            source = io.BytesIO(source)
        with zipfile.ZipFile(source) as archive:
            parts = {
                normalize_part_path(info.filename): archive.read(info)
                for info in archive.infolist()
                if not info.is_dir()
            }
        package = cls(parts=parts)
        package._load_content_types()
        return package

    def _load_content_types(self) -> None:
        raw = self.parts.get(CONTENT_TYPES_PART)
        if raw is None:
            return
        root = parse_xml(raw)
        for node in children(root):
            name = local_name(node.tag)
            if name == "Default":
                extension = (attr(node, "Extension") or "").lower()
                content_type = attr(node, "ContentType")
                if extension and content_type:
                    self.default_types[extension] = content_type
            elif name == "Override":
                part_name = attr(node, "PartName")
                content_type = attr(node, "ContentType")
                if part_name and content_type:
                    self.content_types[normalize_part_path(part_name)] = content_type

    # -- part access -----------------------------------------------------------

    def has_part(self, path: str) -> bool:
        return normalize_part_path(path) in self.parts

    def read(self, path: str) -> bytes | None:
        return self.parts.get(normalize_part_path(path))

    def read_xml(self, path: str):
        raw = self.read(path)
        return None if raw is None else parse_xml(raw)

    def content_type(self, path: str) -> str | None:
        path = normalize_part_path(path)
        explicit = self.content_types.get(path)
        if explicit is not None:
            return explicit
        extension = posixpath.splitext(path)[1].lstrip(".").lower()
        return self.default_types.get(extension) or DEFAULT_MEDIA_TYPES.get(extension)

    # -- relationships ---------------------------------------------------------

    def relationships(self, part_path: str) -> dict[str, Relationship]:
        """Relationships declared by ``part_path``, keyed by relationship id."""
        part_path = normalize_part_path(part_path)
        cached = self._rels_cache.get(part_path)
        if cached is not None:
            return cached

        rels_path = rels_path_for(part_path)
        raw = self.parts.get(rels_path)
        result: dict[str, Relationship] = {}
        if raw is not None:
            base = posixpath.dirname(part_path)
            root = parse_xml(raw)
            for node in children(root, "Relationship"):
                rel_id = attr(node, "Id")
                rel_type = attr(node, "Type")
                target = attr(node, "Target")
                if rel_id is None or rel_type is None or target is None:
                    continue
                external = (attr(node, "TargetMode") or "") == "External"
                target_part = None if external else resolve_target(base, target)
                result[rel_id] = Relationship(
                    id=rel_id,
                    type=rel_type,
                    target=target,
                    target_part=target_part,
                    is_external=external,
                )
        self._rels_cache[part_path] = result
        return result

    def related_part(self, part_path: str, rel_id: str | None) -> str | None:
        """Resolve one relationship id declared by ``part_path`` to a package path."""
        if rel_id is None:
            return None
        relationship = self.relationships(part_path).get(rel_id)
        if relationship is None or relationship.is_external:
            return None
        return relationship.target_part

    def related_parts_of_type(self, part_path: str, rel_type: str) -> list[str]:
        return [
            rel.target_part
            for rel in self.relationships(part_path).values()
            if rel.type == rel_type and rel.target_part is not None
        ]

    def first_related_part(self, part_path: str, rel_type: str) -> str | None:
        parts = self.related_parts_of_type(part_path, rel_type)
        return parts[0] if parts else None

    # -- presentation entry point ---------------------------------------------

    def presentation_part(self) -> str | None:
        """Path of ``presentation.xml`` via the package root relationships."""
        root_rels = self.relationships("")
        for rel in root_rels.values():
            if rel.type == REL_OFFICE_DOCUMENT and rel.target_part is not None:
                return rel.target_part
        return "ppt/presentation.xml" if self.has_part("ppt/presentation.xml") else None


def normalize_part_path(path: str) -> str:
    path = path.replace("\\", "/").lstrip("/")
    return posixpath.normpath(path) if path else path


def rels_path_for(part_path: str) -> str:
    """``ppt/slides/slide1.xml`` -> ``ppt/slides/_rels/slide1.xml.rels``."""
    directory, filename = posixpath.split(part_path)
    return posixpath.join(directory, "_rels", f"{filename}.rels") if directory else f"_rels/{filename}.rels"


def resolve_target(base_dir: str, target: str) -> str:
    if target.startswith("/"):
        return normalize_part_path(target)
    return normalize_part_path(posixpath.join(base_dir, target))
