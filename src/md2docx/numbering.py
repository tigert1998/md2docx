from __future__ import annotations

import re
from typing import Any

from docx.oxml import OxmlElement
from docx.oxml.ns import qn

from .config import EnumeratedListStyle, TextStyle


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
    raise ValueError(
        f"heading numbering pattern {pattern!r} has no numeral placeholder"
    )


def _next_numbering_id(numbering: Any, tag: str, attr: str) -> int:
    values = []
    for child in numbering.findall(qn(f"w:{tag}")):
        value = child.get(qn(f"w:{attr}"))
        if value is not None and value.isdigit():
            values.append(int(value))
    return max(values, default=0) + 1


def _append_value(parent: Any, tag: str, value: str) -> Any:
    element = OxmlElement(f"w:{tag}")
    element.set(qn("w:val"), value)
    parent.append(element)
    return element


def _numbering_run_properties(style: TextStyle) -> Any:
    rpr = OxmlElement("w:rPr")
    fonts = OxmlElement("w:rFonts")
    for attribute, value in (
        ("w:ascii", style.latin_font),
        ("w:hAnsi", style.latin_font),
        ("w:eastAsia", style.chinese_font),
        ("w:cs", style.latin_font),
    ):
        fonts.set(qn(attribute), value)
    rpr.append(fonts)

    _append_value(rpr, "color", str(style.color))
    half_points = str(round(style.size_pt * 2))
    _append_value(rpr, "sz", half_points)
    _append_value(rpr, "szCs", half_points)
    for tag in ("b", "bCs", "i", "iCs"):
        _append_value(rpr, tag, "0")
    return rpr


def _new_abstract_numbering(
    document: Any, name: str, multi_level_type: str
) -> tuple[Any, Any, int]:
    numbering = document.part.numbering_part.element
    abstract_id = _next_numbering_id(numbering, "abstractNum", "abstractNumId")
    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    _append_value(abstract, "multiLevelType", multi_level_type)
    _append_value(abstract, "name", name)
    return numbering, abstract, abstract_id


def _append_level(
    abstract: Any,
    *,
    level: int,
    number_format: str,
    level_text: str,
    paragraph_style: str,
    text_style: TextStyle,
    restart: int | None = None,
    left_indent_twips: int | None = None,
) -> None:
    lvl = OxmlElement("w:lvl")
    lvl.set(qn("w:ilvl"), str(level))
    _append_value(lvl, "start", "1")
    if restart is not None:
        _append_value(lvl, "lvlRestart", str(restart))
    _append_value(lvl, "numFmt", number_format)
    _append_value(lvl, "lvlText", level_text)
    _append_value(lvl, "suff", "nothing")
    _append_value(lvl, "pStyle", paragraph_style)
    _append_value(lvl, "lvlJc", "left")
    if left_indent_twips is not None:
        paragraph_properties = OxmlElement("w:pPr")
        indent = OxmlElement("w:ind")
        indent.set(qn("w:left"), str(left_indent_twips))
        paragraph_properties.append(indent)
        lvl.append(paragraph_properties)
    lvl.append(_numbering_run_properties(text_style))
    abstract.append(lvl)


def _install_abstract_numbering(
    numbering: Any, abstract: Any, abstract_id: int
) -> int:
    numbering.append(abstract)
    num_id = _next_numbering_id(numbering, "num", "numId")
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    _append_value(num, "abstractNumId", str(abstract_id))
    numbering.append(num)
    return num_id


def install_heading_numbering(
    document: Any, heading_styles: list[tuple[int, TextStyle]]
) -> int:
    numbering, abstract, abstract_id = _new_abstract_numbering(
        document, "md2docx heading numbering", "multilevel"
    )
    for level, style in heading_styles:
        if style.numbering is None:
            number_format, level_text = "none", ""
        else:
            number_format, level_text = _numbering_level(style.numbering, level - 1)
        _append_level(
            abstract,
            level=level - 1,
            number_format=number_format,
            level_text=level_text,
            paragraph_style=f"Heading{level}",
            text_style=style,
            restart=level - 1 if level > 1 else None,
        )
    return _install_abstract_numbering(numbering, abstract, abstract_id)


def install_caption_numbering(document: Any, style: TextStyle) -> int:
    if style.numbering is None:
        raise ValueError("image-caption.numbering must not be null")
    numbering, abstract, abstract_id = _new_abstract_numbering(
        document, "md2docx image caption numbering", "singleLevel"
    )
    number_format, level_text = _numbering_level(style.numbering, 0)
    _append_level(
        abstract,
        level=0,
        number_format=number_format,
        level_text=level_text,
        paragraph_style="ImageCaption",
        text_style=style,
    )
    return _install_abstract_numbering(numbering, abstract, abstract_id)


def install_enumerated_list_numbering(
    document: Any, style: EnumeratedListStyle
) -> int:
    if style.numbering is None:
        raise ValueError("enumerated-list.numbering must not be null")
    numbering, abstract, abstract_id = _new_abstract_numbering(
        document, "md2docx enumerated list numbering", "multilevel"
    )
    for level in range(9):
        number_format, level_text = _numbering_level(style.numbering, level)
        left_twips = round(
            style.size_pt * style.indent_before_text_increment_em * level * 20
        )
        _append_level(
            abstract,
            level=level,
            number_format=number_format,
            level_text=level_text,
            paragraph_style=f"ListNumber{level + 1 if level else ''}",
            text_style=style,
            restart=level if level > 0 else None,
            left_indent_twips=left_twips,
        )
    return _install_abstract_numbering(numbering, abstract, abstract_id)


def set_style_numbering(
    style: Any, num_id: int, level: int, enabled: bool
) -> None:
    ppr = style.element.get_or_add_pPr()
    existing = ppr.find(qn("w:numPr"))
    if existing is not None:
        ppr.remove(existing)
    if not enabled:
        return
    num_pr = OxmlElement("w:numPr")
    _append_value(num_pr, "ilvl", str(level - 1))
    _append_value(num_pr, "numId", str(num_id))
    ppr.append(num_pr)
