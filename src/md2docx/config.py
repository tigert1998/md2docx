from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path
import re
from typing import Any

import yaml
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
    "first-line-indent",
    "align",
}
NUMBERED_TEXT_FIELDS = TEXT_FIELDS | {"numbering"}
BLOCK_FIELDS = {"space-before", "space-after", "line-spacing", "align"}
REQUIRED_SECTIONS = {
    "title",
    "h1",
    "h2",
    "h3",
    "h4",
    "body",
    "image",
    "image-caption",
    "math-block",
}


@dataclass(frozen=True)
class LineSpacing:
    value: float
    unit: str


@dataclass(frozen=True)
class BlockStyle:
    space_before_pt: float
    space_after_pt: float
    line_spacing: LineSpacing
    align: WD_ALIGN_PARAGRAPH


@dataclass(frozen=True)
class TextStyle(BlockStyle):
    chinese_font: str
    latin_font: str
    size_pt: float
    color: RGBColor
    numbering: str | None
    first_line_indent_em: float | None


def _require_mapping(data: dict[str, Any], section: str) -> dict[str, Any]:
    if section not in data:
        raise ValueError(f"missing required configuration section: {section}")
    value = data[section]
    if not isinstance(value, dict):
        raise ValueError(f"{section} must be a mapping")
    return value


def _require_fields(section: str, data: dict[str, Any], required: set[str]) -> None:
    missing = sorted(required - data.keys())
    if missing:
        raise ValueError(
            f"{section} is missing required field(s): {', '.join(missing)}"
        )


def _points(value: Any, field: str) -> float:
    match = re.fullmatch(r"\s*(-?\d+(?:\.\d+)?)pt\s*", str(value))
    if not match:
        raise ValueError(f"{field} must use pt units, got {value!r}")
    return float(match.group(1))


def _line_spacing(value: Any, field: str) -> LineSpacing:
    match = re.fullmatch(r"\s*(-?\d+(?:\.\d+)?)(pt|em)\s*", str(value))
    if not match:
        raise ValueError(f"{field} must use pt or em units, got {value!r}")
    return LineSpacing(float(match.group(1)), match.group(2))


def _indent_em(value: Any, field: str) -> float | None:
    if value is None:
        return None
    match = re.fullmatch(r"\s*(-?\d+(?:\.\d+)?)em\s*", str(value))
    if not match:
        raise ValueError(f"{field} must use em units or null, got {value!r}")
    return float(match.group(1))


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


def _color(value: Any, field: str) -> RGBColor:
    match = re.fullmatch(r"#?([0-9A-Fa-f]{6})", str(value))
    if not match:
        raise ValueError(f"{field} must be a six-digit hex color, got {value!r}")
    return RGBColor.from_string(match.group(1).upper())


def _block_style(section: str, data: dict[str, Any]) -> BlockStyle:
    _require_fields(section, data, BLOCK_FIELDS)
    return BlockStyle(
        space_before_pt=_points(data["space-before"], f"{section}.space-before"),
        space_after_pt=_points(data["space-after"], f"{section}.space-after"),
        line_spacing=_line_spacing(data["line-spacing"], f"{section}.line-spacing"),
        align=_alignment(data["align"], f"{section}.align"),
    )


def _text_style(section: str, data: dict[str, Any], *, numbered: bool) -> TextStyle:
    _require_fields(
        section, data, NUMBERED_TEXT_FIELDS if numbered else TEXT_FIELDS
    )
    block = _block_style(section, data)
    numbering = data["numbering"] if numbered else None
    if numbering is not None and not isinstance(numbering, str):
        raise ValueError(f"{section}.numbering must be a string or null")
    return TextStyle(
        **block.__dict__,
        chinese_font=str(data["chinese-font"]),
        latin_font=str(data["latin-font"]),
        size_pt=_points(data["size"], f"{section}.size"),
        color=_color(data["color"], f"{section}.color"),
        numbering=numbering,
        first_line_indent_em=_indent_em(
            data["first-line-indent"], f"{section}.first-line-indent"
        ),
    )


def load_config(path: str | Path) -> dict[str, BlockStyle | TextStyle]:
    path = Path(path)
    with path.open("r", encoding="utf-8") as stream:
        loaded = yaml.safe_load(stream)
    if not isinstance(loaded, dict):
        raise ValueError("configuration root must be a mapping")

    missing_sections = sorted(REQUIRED_SECTIONS - loaded.keys())
    if missing_sections:
        raise ValueError(
            "missing required configuration section(s): "
            + ", ".join(missing_sections)
        )

    result: dict[str, BlockStyle | TextStyle] = {
        "title": _text_style("title", _require_mapping(loaded, "title"), numbered=True),
        "body": _text_style("body", _require_mapping(loaded, "body"), numbered=False),
        "image": _block_style("image", _require_mapping(loaded, "image")),
        "image-caption": _text_style(
            "image-caption",
            _require_mapping(loaded, "image-caption"),
            numbered=True,
        ),
        "math-block": _block_style(
            "math-block", _require_mapping(loaded, "math-block")
        ),
    }
    heading_names = sorted(
        (name for name in loaded if re.fullmatch(r"h[1-9]", str(name))),
        key=lambda name: int(name[1:]),
    )
    for name in heading_names:
        result[name] = _text_style(
            name, _require_mapping(loaded, name), numbered=True
        )
    return result


def apply_style_to_paragraph(
    paragraph: Any, style: BlockStyle, *, font_size_pt: float | None = None
) -> None:
    fmt = paragraph.paragraph_format
    fmt.space_before = Pt(style.space_before_pt)
    fmt.space_after = Pt(style.space_after_pt)
    if style.line_spacing.unit == "pt":
        fmt.line_spacing = Pt(style.line_spacing.value)
    else:
        fmt.line_spacing = style.line_spacing.value
    if isinstance(style, TextStyle):
        fmt.first_line_indent = (
            Pt(style.size_pt * style.first_line_indent_em)
            if style.first_line_indent_em is not None
            else None
        )
    fmt.alignment = style.align
