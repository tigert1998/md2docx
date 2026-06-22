from __future__ import annotations

import re
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .config import StyleConfig, list_level_layout


def _numbering_level(pattern: str, level: int) -> tuple[str, str]:
    placeholder = f"%{level + 1}"
    if "{cn}" in pattern:
        return "chineseCounting", pattern.replace("{cn}", placeholder)
    if "{n}" in pattern:
        return "decimal", pattern.replace("{n}", placeholder)
    if re.search(r"[一二三四五六七八九十]+", pattern):
        return (
            "chineseCounting",
            re.sub(r"[一二三四五六七八九十]+", placeholder, pattern, count=1),
        )
    if re.search(r"\d+", pattern):
        return "decimal", re.sub(r"\d+", placeholder, pattern, count=1)
    raise ValueError(f"numbering pattern {pattern!r} has no numeral placeholder")


def _next_id(numbering: Any, tag: str, attr: str) -> int:
    values = [
        int(value)
        for child in numbering.findall(qn(f"w:{tag}"))
        if (value := child.get(qn(f"w:{attr}"))) is not None and value.isdigit()
    ]
    return max(values, default=0) + 1


def _value(parent: Any, tag: str, value: str) -> Any:
    element = OxmlElement(f"w:{tag}")
    element.set(qn("w:val"), value)
    parent.append(element)
    return element


def _run_properties(style: StyleConfig) -> Any:
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    if style.latin_font is not None:
        for attr in ("w:ascii", "w:hAnsi", "w:cs"):
            fonts.set(qn(attr), style.latin_font)
    if style.chinese_font is not None:
        fonts.set(qn("w:eastAsia"), style.chinese_font)
    rpr.append(fonts)
    if style.color is not None:
        _value(rpr, "color", str(style.color))
    if style.size_pt is not None:
        half_points = str(round(style.size_pt * 2))
        _value(rpr, "sz", half_points)
        _value(rpr, "szCs", half_points)
    return rpr


def _new_abstract(document: Any, name: str, kind: str) -> tuple[Any, Any, int]:
    numbering = document.part.numbering_part.element
    abstract_id = _next_id(numbering, "abstractNum", "abstractNumId")
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    _value(abstract, "multiLevelType", kind)
    _value(abstract, "name", name)
    return numbering, abstract, abstract_id


def _append_level(
    abstract: Any,
    *,
    level: int,
    number_format: str,
    level_text: str,
    paragraph_style: str,
    text_style: StyleConfig,
    left_twips: int | None = None,
    hanging_twips: int | None = None,
) -> None:
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), str(level))
    _value(lvl, "start", "1")
    _value(lvl, "numFmt", number_format)
    _value(lvl, "lvlText", level_text)
    _value(lvl, "suff", "nothing")
    _value(lvl, "pStyle", paragraph_style)
    _value(lvl, "lvlJc", "left")
    if left_twips is not None:
        ppr = OxmlElement("w:pPr")
        indent = OxmlElement("w:ind")
        indent.set(qn("w:left"), str(left_twips))
        if hanging_twips is not None:
            indent.set(qn("w:hanging"), str(hanging_twips))
        ppr.append(indent)
        lvl.append(ppr)
    lvl.append(_run_properties(text_style))
    abstract.append(lvl)


def _install(numbering: Any, abstract: Any, abstract_id: int) -> int:
    numbering.append(abstract)
    num_id = _next_id(numbering, "num", "numId")
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    _value(num, "abstractNumId", str(abstract_id))
    numbering.append(num)
    return num_id


def restart_numbering(
    document: Any,
    template_num_id: int,
    *,
    level: int,
    start: int = 1,
) -> int:
    if level < 1:
        raise ValueError("numbering level must be at least 1")
    if start < 1:
        raise ValueError("numbering start must be at least 1")
    numbering = document.part.numbering_part.element
    template = next(
        (
            num
            for num in numbering.findall(qn("w:num"))
            if num.get(qn("w:numId")) == str(template_num_id)
        ),
        None,
    )
    if template is None:
        raise ValueError(f"numbering definition {template_num_id} does not exist")
    abstract = template.find(qn("w:abstractNumId"))
    if abstract is None:
        raise ValueError(f"numbering definition {template_num_id} has no abstractNumId")

    num_id = _next_id(numbering, "num", "numId")
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    _value(num, "abstractNumId", abstract.get(qn("w:val")))
    override = OxmlElement("w:lvlOverride")
    override.set(qn("w:ilvl"), str(level - 1))
    _value(override, "startOverride", str(start))
    num.append(override)
    numbering.append(num)
    return num_id


def install_heading_numbering(
    document: Any, heading_styles: list[tuple[int, StyleConfig]]
) -> int:
    numbering, abstract, abstract_id = _new_abstract(
        document, "md2docx heading numbering", "multilevel"
    )
    for level, style in heading_styles:
        fmt, text = (
            ("none", "")
            if style.numbering is None
            else _numbering_level(style.numbering, level - 1)
        )
        _append_level(
            abstract,
            level=level - 1,
            number_format=fmt,
            level_text=text,
            paragraph_style=style.name,
            text_style=style,
        )
    return _install(numbering, abstract, abstract_id)


def install_caption_numbering(document: Any, style: StyleConfig) -> int:
    if style.numbering is None:
        raise ValueError("image-caption.numbering must not be null")
    numbering, abstract, abstract_id = _new_abstract(
        document, "md2docx image caption numbering", "singleLevel"
    )
    fmt, text = _numbering_level(style.numbering, 0)
    _append_level(
        abstract,
        level=0,
        number_format=fmt,
        level_text=text,
        paragraph_style=style.name,
        text_style=style,
    )
    return _install(numbering, abstract, abstract_id)


def install_list_numbering(document: Any, style: StyleConfig, *, ordered: bool) -> int:
    if style.numbering is None:
        raise ValueError(f"{style.name}.numbering must not be null")
    numbering, abstract, abstract_id = _new_abstract(
        document, f"md2docx {style.name} numbering", "multilevel"
    )
    for level in range(9):
        if ordered:
            fmt, text = _numbering_level(style.numbering, level)
        else:
            fmt, text = "bullet", style.numbering
        layout = list_level_layout(style, level + 1)
        left_twips = round(layout.left_indent_pt * 20)
        hanging_twips = (
            None
            if layout.hanging_indent_pt is None
            else round(layout.hanging_indent_pt * 20)
        )
        _append_level(
            abstract,
            level=level,
            number_format=fmt,
            level_text=text,
            paragraph_style=f"{style.name}-{level + 1}",
            text_style=style,
            left_twips=left_twips,
            hanging_twips=hanging_twips,
        )
    return _install(numbering, abstract, abstract_id)


def set_style_numbering(style: Any, num_id: int, level: int) -> None:
    _set_numbering(style.element.get_or_add_pPr(), num_id, level)


def set_paragraph_numbering(paragraph: Any, num_id: int, level: int) -> None:
    _set_numbering(paragraph._p.get_or_add_pPr(), num_id, level)


def _set_numbering(ppr: Any, num_id: int, level: int) -> None:
    existing = ppr.find(qn("w:numPr"))
    if existing is not None:
        ppr.remove(existing)
    num_pr = OxmlElement("w:numPr")
    _value(num_pr, "ilvl", str(level - 1))
    _value(num_pr, "numId", str(num_id))
    ppr.append(num_pr)
