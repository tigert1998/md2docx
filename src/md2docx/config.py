from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Any

import yaml
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Pt, RGBColor

TEXT_FIELDS = {
    "chinese-font",
    "latin-font",
    "size",
    "color",
    "space-before",
    "space-after",
    "line-spacing",
    "numbering",
    "hanging-indent",
    "first-line-indent",
    "indent-before-text",
    "align",
}
LIST_FIELDS = TEXT_FIELDS | {"indent-before-text-increment"}
REQUIRED_SECTIONS = {
    "title",
    "h1",
    "h2",
    "h3",
    "h4",
    "body",
    "image",
    "image-caption",
    "ordered-list",
    "unordered-list",
    "inline-math",
    "math-block",
    "inline-code",
    "code-block",
    "table-header",
    "table-body",
}
LIST_SECTIONS = {"ordered-list", "unordered-list"}
INLINE_SECTIONS = {"inline-math", "inline-code"}


@dataclass(frozen=True)
class Length:
    value: float
    unit: str

    def to_points(self, font_size_pt: float | None = None) -> float:
        if self.unit == "pt":
            return self.value
        if font_size_pt is None:
            raise ValueError("em length requires a configured font size")
        return self.value * font_size_pt


@dataclass(frozen=True)
class StyleConfig:
    name: str
    chinese_font: str | None
    latin_font: str | None
    size_pt: float | None
    color: RGBColor | None
    space_before: Length
    space_after: Length
    line_spacing: Length
    numbering: str | None
    hanging_indent: Length | None
    first_line_indent: Length | None
    indent_before_text: Length
    align: WD_ALIGN_PARAGRAPH
    indent_before_text_increment: Length | None = None


def _require_fields(section: str, data: dict[str, Any], required: set[str]) -> None:
    missing = sorted(required - data.keys())
    if missing:
        raise ValueError(
            f"{section} is missing required field(s): {', '.join(missing)}"
        )


def _length(value: Any, field: str, units: str = "pt|em") -> Length:
    match = re.fullmatch(
        rf"\s*(-?\d+(?:\.\d+)?)({units})\s*",
        str(value),
    )
    if not match:
        raise ValueError(f"{field} must use {units.replace('|', ' or ')} units")
    return Length(float(match.group(1)), match.group(2))


def _optional_length(value: Any, field: str) -> Length | None:
    if value is None:
        return None
    return _length(value, field)


def _points(value: Any, field: str) -> float:
    return _length(value, field, "pt").value


def _optional_text(value: Any, field: str) -> str | None:
    if value is None:
        return None
    if not isinstance(value, str) or not value:
        raise ValueError(f"{field} must be a non-empty string or null")
    return value


def _optional_points(value: Any, field: str) -> float | None:
    if value is None:
        return None
    return _points(value, field)


def _alignment(value: Any, field: str) -> WD_ALIGN_PARAGRAPH:
    choices = {
        "left": WD_ALIGN_PARAGRAPH.LEFT,
        "center": WD_ALIGN_PARAGRAPH.CENTER,
        "right": WD_ALIGN_PARAGRAPH.RIGHT,
        "justify": WD_ALIGN_PARAGRAPH.JUSTIFY,
    }
    try:
        return choices[str(value).lower()]
    except KeyError as exc:
        raise ValueError(f"{field} must be one of {', '.join(choices)}") from exc


def _optional_color(value: Any, field: str) -> RGBColor | None:
    if value is None:
        return None
    match = re.fullmatch(r"#?([0-9A-Fa-f]{6})", str(value))
    if not match:
        raise ValueError(f"{field} must be a six-digit hex color or null")
    return RGBColor.from_string(match.group(1).upper())


def _parse_style(name: str, data: Any) -> StyleConfig:
    if not isinstance(data, dict):
        raise ValueError(f"{name} must be a mapping")
    required = LIST_FIELDS if name in LIST_SECTIONS else TEXT_FIELDS
    _require_fields(name, data, required)
    numbering = data["numbering"]
    if numbering is not None and not isinstance(numbering, str):
        raise ValueError(f"{name}.numbering must be a string or null")
    hanging_indent = _optional_length(data["hanging-indent"], f"{name}.hanging-indent")
    first_line_indent = _optional_length(
        data["first-line-indent"], f"{name}.first-line-indent"
    )
    if hanging_indent is not None and first_line_indent is not None:
        raise ValueError(
            f"{name}.hanging-indent and {name}.first-line-indent are mutually exclusive"
        )
    increment = None
    if name in LIST_SECTIONS:
        increment = _length(
            data["indent-before-text-increment"],
            f"{name}.indent-before-text-increment",
        )
        if increment.value < 0:
            raise ValueError(
                f"{name}.indent-before-text-increment must be non-negative"
            )
    return StyleConfig(
        name=name,
        chinese_font=_optional_text(data["chinese-font"], f"{name}.chinese-font"),
        latin_font=_optional_text(data["latin-font"], f"{name}.latin-font"),
        size_pt=_optional_points(data["size"], f"{name}.size"),
        color=_optional_color(data["color"], f"{name}.color"),
        space_before=_length(data["space-before"], f"{name}.space-before"),
        space_after=_length(data["space-after"], f"{name}.space-after"),
        line_spacing=_length(data["line-spacing"], f"{name}.line-spacing"),
        numbering=numbering,
        hanging_indent=hanging_indent,
        first_line_indent=first_line_indent,
        indent_before_text=_length(
            data["indent-before-text"], f"{name}.indent-before-text"
        ),
        align=_alignment(data["align"], f"{name}.align"),
        indent_before_text_increment=increment,
    )


def load_config(path: str | Path) -> dict[str, StyleConfig]:
    with Path(path).open("r", encoding="utf-8") as stream:
        loaded = yaml.safe_load(stream)
    if not isinstance(loaded, dict):
        raise ValueError("configuration root must be a mapping")
    missing = sorted(REQUIRED_SECTIONS - loaded.keys())
    if missing:
        raise ValueError(
            "missing required configuration section(s): " + ", ".join(missing)
        )
    return {str(name): _parse_style(str(name), value) for name, value in loaded.items()}


def apply_config_to_style(word_style: Any, config: StyleConfig) -> None:
    if config.latin_font is not None:
        word_style.font.name = config.latin_font
    if config.size_pt is not None:
        word_style.font.size = Pt(config.size_pt)
    if config.color is not None:
        word_style.font.color.rgb = config.color
    if config.chinese_font is not None or config.latin_font is not None:
        rpr = word_style.element.get_or_add_rPr()
        rfonts = rpr.rFonts
        if rfonts is None:
            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn

            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        from docx.oxml.ns import qn

        if config.latin_font is not None:
            for attr in ("w:ascii", "w:hAnsi", "w:cs"):
                rfonts.set(qn(attr), config.latin_font)
        if config.chinese_font is not None:
            rfonts.set(qn("w:eastAsia"), config.chinese_font)

    if word_style.type == WD_STYLE_TYPE.PARAGRAPH:
        fmt = word_style.paragraph_format
        fmt.space_before = Pt(config.space_before.to_points(config.size_pt))
        fmt.space_after = Pt(config.space_after.to_points(config.size_pt))
        fmt.line_spacing = (
            Pt(config.line_spacing.value)
            if config.line_spacing.unit == "pt"
            else config.line_spacing.value
        )
        fmt.left_indent = Pt(config.indent_before_text.to_points(config.size_pt))
        if config.hanging_indent is not None:
            fmt.first_line_indent = Pt(-config.hanging_indent.to_points(config.size_pt))
        elif config.first_line_indent is None:
            fmt.first_line_indent = None
        else:
            fmt.first_line_indent = Pt(
                config.first_line_indent.to_points(config.size_pt)
            )
        fmt.alignment = config.align
