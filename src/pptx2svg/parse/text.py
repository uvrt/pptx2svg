"""Reader for ``p:txBody`` -- text bodies, paragraphs, runs and list styles.

Run order matters: ``a:r``, ``a:br`` and ``a:fld`` interleave inside a paragraph and the
visual result depends on their document order.  ElementTree hands children back in order,
so this reader just walks them, where the TypeScript original had to rebuild the sequence
from its XML parser's grouped output.

Run properties that are absent stay ``None`` -- they are filled in later from the
paragraph's ``defRPr``, the shape's ``lstStyle``, the master's ``txStyles`` and the
presentation's ``defaultTextStyle``, in that order.
"""

from __future__ import annotations

from xml.etree.ElementTree import Element

from ..model import (
    AutoNumBullet,
    BulletType,
    CharBullet,
    NoBullet,
    PercentSpacing,
    PointsSpacing,
    SpacingValue,
    TabStop,
)
from ..xmlutil import (
    attr,
    child,
    children,
    decode_char_refs,
    int_attr,
    is_true,
    local_name,
    ns_attr,
    num_attr,
)
from .drawing import parse_color
from .source import (
    SourceParagraph,
    SourceParagraphProperties,
    SourceRunProperties,
    SourceTextBody,
    SourceTextBodyProperties,
    SourceTextRun,
    SourceTextStyle,
)

ALIGN_MAP = {"l": "l", "ctr": "ctr", "r": "r", "just": "just", "justLow": "just", "dist": "just"}
ANCHOR_MAP = {"t": "t", "ctr": "ctr", "b": "b", "just": "t", "dist": "t"}
WRAP_VALUES = {"square", "none"}
VERTICAL_VALUES = {"horz", "vert", "vert270", "eaVert", "wordArtVert", "mongolianVert"}
AUTO_NUM_SCHEMES = {
    "arabicPeriod",
    "arabicParenR",
    "romanUcPeriod",
    "romanLcPeriod",
    "alphaUcPeriod",
    "alphaLcPeriod",
    "alphaLcParenR",
    "alphaUcParenR",
    "arabicPlain",
}


def parse_text_body(tx_body: Element | None) -> SourceTextBody | None:
    if tx_body is None:
        return None
    return SourceTextBody(
        paragraphs=[parse_paragraph(p) for p in children(tx_body, "p")],
        properties=parse_body_properties(child(tx_body, "bodyPr")),
        list_style=parse_text_style(child(tx_body, "lstStyle")),
    )


def parse_body_properties(body_pr: Element | None) -> SourceTextBodyProperties | None:
    if body_pr is None:
        return None

    anchor = ANCHOR_MAP.get(attr(body_pr, "anchor") or "")
    wrap = attr(body_pr, "wrap")
    vert = attr(body_pr, "vert")
    num_col = int_attr(body_pr, "numCol")

    norm_autofit = child(body_pr, "normAutofit")
    if norm_autofit is not None:
        auto_fit, font_scale, ln_spc_reduction = (
            "normAutofit",
            (num_attr(norm_autofit, "fontScale") or 100000) / 100000,
            (num_attr(norm_autofit, "lnSpcReduction") or 0) / 100000,
        )
    elif child(body_pr, "spAutoFit") is not None:
        auto_fit, font_scale, ln_spc_reduction = "spAutofit", 1.0, 0.0
    elif child(body_pr, "noAutofit") is not None:
        auto_fit, font_scale, ln_spc_reduction = "noAutofit", 1.0, 0.0
    else:
        auto_fit, font_scale, ln_spc_reduction = None, None, None

    properties = SourceTextBodyProperties(
        margin_left=num_attr(body_pr, "lIns"),
        margin_right=num_attr(body_pr, "rIns"),
        margin_top=num_attr(body_pr, "tIns"),
        margin_bottom=num_attr(body_pr, "bIns"),
        anchor=anchor,  # type: ignore[arg-type]
        wrap=wrap if wrap in WRAP_VALUES else None,  # type: ignore[arg-type]
        auto_fit=auto_fit,  # type: ignore[arg-type]
        font_scale=font_scale,
        ln_spc_reduction=ln_spc_reduction,
        num_col=max(1, num_col) if num_col is not None else None,
        vert=vert if vert in VERTICAL_VALUES else None,  # type: ignore[arg-type]
        rotation=num_attr(body_pr, "rot"),
    )
    return properties if _has_any(properties) else None


def parse_text_style(node: Element | None) -> SourceTextStyle | None:
    """Read ``a:lstStyle`` / ``p:titleStyle`` / ``p:bodyStyle`` / ``a:defaultTextStyle``."""
    if node is None:
        return None
    default_paragraph = parse_paragraph_properties(child(node, "defPPr"))
    levels = [parse_paragraph_properties(child(node, f"lvl{index + 1}pPr")) for index in range(9)]
    if default_paragraph is None and not any(levels):
        return None
    return SourceTextStyle(default_paragraph=default_paragraph, levels=levels)


def parse_paragraph(p: Element) -> SourceParagraph:
    return SourceParagraph(
        runs=parse_runs_in_order(p),
        properties=parse_paragraph_properties(child(p, "pPr")),
        end_para_run_properties=parse_run_properties(child(p, "endParaRPr")),
    )


def parse_runs_in_order(p: Element) -> list[SourceTextRun]:
    """Walk ``a:r`` / ``a:br`` / ``a:fld`` in document order.

    A ``a:br`` becomes a run holding a newline so the wrapper can treat forced breaks and
    text uniformly; a ``a:fld`` (slide number, date) renders whatever cached text it has.
    """
    runs: list[SourceTextRun] = []
    for node in p:
        name = local_name(node.tag)
        if name == "r":
            runs.append(
                SourceTextRun(
                    text=decode_char_refs(_run_text(node)),
                    properties=parse_run_properties(child(node, "rPr")),
                )
            )
        elif name == "fld":
            runs.append(
                SourceTextRun(
                    text=decode_char_refs(_run_text(node)),
                    properties=parse_run_properties(child(node, "rPr")),
                )
            )
        elif name == "br":
            runs.append(
                SourceTextRun(text="\n", properties=parse_run_properties(child(node, "rPr")))
            )
    return runs


def _run_text(node: Element) -> str:
    text_node = child(node, "t")
    if text_node is None:
        return ""
    return "".join(text_node.itertext())


def parse_paragraph_properties(p_pr: Element | None) -> SourceParagraphProperties | None:
    if p_pr is None:
        return None

    align = ALIGN_MAP.get(attr(p_pr, "algn") or "")
    properties = SourceParagraphProperties(
        align=align,  # type: ignore[arg-type]
        level=int_attr(p_pr, "lvl"),
        line_spacing=parse_spacing(child(p_pr, "lnSpc")),
        space_before=parse_spacing(child(p_pr, "spcBef")),
        space_after=parse_spacing(child(p_pr, "spcAft")),
        margin_left=num_attr(p_pr, "marL"),
        indent=num_attr(p_pr, "indent"),
        bullet=parse_bullet(p_pr),
        bullet_font=attr(child(p_pr, "buFont"), "typeface"),
        bullet_color=parse_color(child(p_pr, "buClr")),
        bullet_size_pct=num_attr(child(p_pr, "buSzPct"), "val"),
        tab_stops=parse_tab_stops(p_pr),
        default_run_properties=parse_run_properties(child(p_pr, "defRPr")),
    )
    return properties if _has_any(properties) else None


def parse_tab_stops(p_pr: Element | None) -> list[TabStop] | None:
    stops = [
        TabStop(
            position=num_attr(tab, "pos") or 0,
            alignment=attr(tab, "algn") or "l",  # type: ignore[arg-type]
        )
        for tab in children(child(p_pr, "tabLst"), "tab")
    ]
    return stops or None


def parse_spacing(node: Element | None) -> SpacingValue | None:
    points = num_attr(child(node, "spcPts"), "val")
    if points is not None:
        return PointsSpacing(value=points)
    percent = num_attr(child(node, "spcPct"), "val")
    if percent is not None:
        return PercentSpacing(value=percent)
    return None


def parse_bullet(p_pr: Element | None) -> BulletType | None:
    if p_pr is None:
        return None
    if child(p_pr, "buNone") is not None:
        return NoBullet()
    bu_char = child(p_pr, "buChar")
    if bu_char is not None:
        return CharBullet(char=decode_char_refs(attr(bu_char, "char") or "•"))
    bu_auto_num = child(p_pr, "buAutoNum")
    if bu_auto_num is not None:
        scheme = attr(bu_auto_num, "type") or "arabicPeriod"
        return AutoNumBullet(
            scheme=scheme if scheme in AUTO_NUM_SCHEMES else "arabicPeriod",  # type: ignore[arg-type]
            start_at=int_attr(bu_auto_num, "startAt") or 1,
        )
    return None


def parse_run_properties(r_pr: Element | None) -> SourceRunProperties | None:
    if r_pr is None:
        return None

    bold = attr(r_pr, "b")
    italic = attr(r_pr, "i")
    underline = attr(r_pr, "u")
    strike = attr(r_pr, "strike")
    baseline = num_attr(r_pr, "baseline")
    size = num_attr(r_pr, "sz")

    hyperlink = child(r_pr, "hlinkClick")
    has_hyperlink = hyperlink is not None

    outline = child(r_pr, "ln")
    color = parse_color(child(r_pr, "solidFill"))

    properties = SourceRunProperties(
        bold=is_true(bold) if bold is not None else None,
        italic=is_true(italic) if italic is not None else None,
        # A hyperlink is underlined unless the run says otherwise.
        underline=(underline != "none") if underline is not None else (True if has_hyperlink else None),
        strikethrough=(strike != "noStrike") if strike is not None else None,
        # `a:rPr@baseline` is a 1/1000 percent of the font size.
        baseline=baseline / 1000 if baseline is not None else None,
        # `a:rPr@sz` is in 1/100 pt.
        font_size=size / 100 if size is not None else None,
        typeface=attr(child(r_pr, "latin"), "typeface"),
        typeface_ea=attr(child(r_pr, "ea"), "typeface"),
        typeface_cs=attr(child(r_pr, "cs"), "typeface"),
        # Hyperlinks default to the theme's hlink colour when the run has none.
        color=color if color is not None else (_hlink_color() if has_hyperlink else None),
        highlight=parse_color(child(r_pr, "highlight")),
        outline_width=num_attr(outline, "w") if outline is not None else None,
        outline_color=parse_color(child(outline, "solidFill")) if outline is not None else None,
        hyperlink_rel_id=ns_attr(hyperlink, "id") if has_hyperlink else None,
        hyperlink_tooltip=attr(hyperlink, "tooltip") if has_hyperlink else None,
    )
    return properties if _has_any(properties) else None


def _hlink_color():
    from .source import SchemeColor

    return SchemeColor(scheme="hlink")


def _has_any(dataclass_instance) -> bool:
    return any(value is not None for value in vars(dataclass_instance).values())
