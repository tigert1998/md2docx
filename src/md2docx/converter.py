from __future__ import annotations

from io import BytesIO
from pathlib import Path
import re
from typing import Any, Iterable
from urllib.parse import urlparse
from urllib.request import urlopen

import mistune
import yaml
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt
from PIL import Image

from .config import (
    BlockStyle,
    EnumeratedListStyle,
    TextStyle,
    apply_style_to_paragraph,
    load_config,
)
from .math_render import render_latex
from .numbering import (
    install_caption_numbering,
    install_enumerated_list_numbering,
    install_heading_numbering,
    set_style_numbering,
)


def parse_frontmatter(source: str) -> tuple[dict[str, Any], str]:
    lines = source.splitlines(keepends=True)
    if not lines or lines[0].strip() != "---":
        raise ValueError(
            "Markdown must begin with YAML Frontmatter delimited by '---'"
        )
    end = next(
        (index for index, line in enumerate(lines[1:], start=1) if line.strip() == "---"),
        None,
    )
    if end is None:
        raise ValueError("Markdown Frontmatter is missing its closing '---'")
    try:
        metadata = yaml.safe_load("".join(lines[1:end]))
    except yaml.YAMLError as exc:
        raise ValueError(f"invalid Markdown Frontmatter YAML: {exc}") from exc
    if not isinstance(metadata, dict):
        raise ValueError("Markdown Frontmatter must be a mapping")
    title = metadata.get("title")
    if not isinstance(title, str) or not title.strip():
        raise ValueError("Markdown Frontmatter must contain a non-empty title")
    return metadata, "".join(lines[end + 1 :])


def _set_run_style(run: Any, style: TextStyle) -> None:
    run.font.name = style.latin_font
    run.font.size = Pt(style.size_pt)
    run.font.color.rgb = style.color
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:ascii"), style.latin_font)
    rfonts.set(qn("w:hAnsi"), style.latin_font)
    rfonts.set(qn("w:eastAsia"), style.chinese_font)
    rfonts.set(qn("w:cs"), style.latin_font)


def _reset_word_style_formatting(style: Any) -> None:
    element = style.element
    for tag in ("w:basedOn", "w:link", "w:pPr", "w:rPr"):
        child = element.find(qn(tag))
        if child is not None:
            element.remove(child)


def _shade_cell(cell: Any, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = tc_pr.find(qn("w:shd"))
    if shd is None:
        shd = OxmlElement("w:shd")
        tc_pr.append(shd)
    shd.set(qn("w:fill"), fill)


def _set_cell_margins(cell: Any, value: int = 100) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = tc_pr.first_child_found_in("w:tcMar")
    if tc_mar is None:
        tc_mar = OxmlElement("w:tcMar")
        tc_pr.append(tc_mar)
    for edge in ("top", "start", "bottom", "end"):
        node = tc_mar.find(qn(f"w:{edge}"))
        if node is None:
            node = OxmlElement(f"w:{edge}")
            tc_mar.append(node)
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _set_table_geometry(table: Any, widths: list[int]) -> None:
    total = sum(widths)
    tbl_pr = table._tbl.tblPr
    tbl_w = tbl_pr.find(qn("w:tblW"))
    if tbl_w is None:
        tbl_w = OxmlElement("w:tblW")
        tbl_pr.insert(0, tbl_w)
    tbl_w.set(qn("w:w"), str(total))
    tbl_w.set(qn("w:type"), "dxa")
    layout = tbl_pr.find(qn("w:tblLayout"))
    if layout is None:
        layout = OxmlElement("w:tblLayout")
        tbl_pr.append(layout)
    layout.set(qn("w:type"), "fixed")

    grid = table._tbl.tblGrid
    for child in list(grid):
        grid.remove(child)
    for width in widths:
        column = OxmlElement("w:gridCol")
        column.set(qn("w:w"), str(width))
        grid.append(column)
    for row in table.rows:
        for index, cell in enumerate(row.cells):
            tc_w = cell._tc.get_or_add_tcPr().get_or_add_tcW()
            tc_w.set(qn("w:w"), str(widths[index]))
            tc_w.set(qn("w:type"), "dxa")


class DocxBuilder:
    def __init__(
        self,
        styles: dict[str, BlockStyle | TextStyle | EnumeratedListStyle],
        source_dir: Path,
    ):
        self.styles = styles
        self.source_dir = source_dir
        self.document = Document()
        self._configure_document()

    def text_style(self, name: str) -> TextStyle:
        style = self.styles.get(name)
        if not isinstance(style, TextStyle):
            raise ValueError(f"configuration section {name} is not a text style")
        return style

    def block_style(self, name: str) -> BlockStyle:
        style = self.styles.get(name)
        if style is None:
            raise ValueError(f"missing required configuration section: {name}")
        return style

    def enumerated_list_style(self) -> EnumeratedListStyle:
        style = self.styles.get("enumerated-list")
        if not isinstance(style, EnumeratedListStyle):
            raise ValueError(
                "configuration section enumerated-list is not an enumerated-list style"
            )
        return style

    def _configure_document(self) -> None:
        section = self.document.sections[0]
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11)
        title_style = self.document.styles["Title"]
        _reset_word_style_formatting(title_style)
        self._configure_word_style(title_style, self.text_style("title"))
        title_style.font.bold = False
        title_style.font.italic = False

        self._configure_word_style(
            self.document.styles["Normal"], self.text_style("body")
        )
        heading_styles = sorted(
            (
                (int(name[1:]), style)
                for name, style in self.styles.items()
                if re.fullmatch(r"h[1-9]", name) and isinstance(style, TextStyle)
            ),
            key=lambda item: item[0],
        )
        num_id = install_heading_numbering(self.document, heading_styles)
        for level, config_style in heading_styles:
            word_style = self.document.styles[f"Heading {level}"]
            _reset_word_style_formatting(word_style)
            self._configure_word_style(word_style, config_style)
            word_style.font.bold = False
            word_style.font.italic = False
            set_style_numbering(
                word_style, num_id, level, config_style.numbering is not None
            )
        if "Image Caption" not in self.document.styles:
            self.document.styles.add_style(
                "Image Caption", WD_STYLE_TYPE.PARAGRAPH
            )
        self._configure_word_style(
            self.document.styles["Image Caption"],
            self.text_style("image-caption"),
        )
        caption_num_id = install_caption_numbering(
            self.document, self.text_style("image-caption")
        )
        set_style_numbering(
            self.document.styles["Image Caption"], caption_num_id, 1, True
        )
        enumerated_style = self.enumerated_list_style()
        enumerated_num_id = install_enumerated_list_numbering(
            self.document, enumerated_style
        )
        for level in range(9):
            style_name = self._list_number_style_name(level)
            if style_name not in self.document.styles:
                list_style = self.document.styles.add_style(
                    style_name, WD_STYLE_TYPE.PARAGRAPH
                )
                list_style.base_style = self.document.styles["List Number"]
            else:
                list_style = self.document.styles[style_name]
            self._configure_word_style(list_style, enumerated_style)
            list_style.font.bold = False
            list_style.font.italic = False
            list_style.paragraph_format.left_indent = Pt(
                enumerated_style.size_pt
                * enumerated_style.indent_before_text_increment_em
                * level
            )
            set_style_numbering(
                list_style, enumerated_num_id, level + 1, True
            )

    @staticmethod
    def _configure_word_style(word_style: Any, config: TextStyle) -> None:
        word_style.font.name = config.latin_font
        word_style.font.size = Pt(config.size_pt)
        word_style.font.color.rgb = config.color
        rpr = word_style.element.get_or_add_rPr()
        rfonts = rpr.rFonts
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:ascii"), config.latin_font)
        rfonts.set(qn("w:hAnsi"), config.latin_font)
        rfonts.set(qn("w:eastAsia"), config.chinese_font)
        rfonts.set(qn("w:cs"), config.latin_font)
        apply_style_to_paragraph(word_style, config)

    @staticmethod
    def _list_number_style_name(level: int) -> str:
        return "List Number" if level == 0 else f"List Number {level + 1}"

    def add_title(self, title: str) -> None:
        paragraph = self.document.add_paragraph(style="Title")
        paragraph.add_run(title.strip())

    def add_heading(self, token: dict[str, Any]) -> None:
        level = int(token["attrs"]["level"])
        key = f"h{level}"
        if key not in self.styles:
            raise ValueError(
                f"Markdown heading level {level} requires configuration section {key}"
            )
        style = self.text_style(key)
        paragraph = self.document.add_paragraph(style=f"Heading {level}")
        self.add_inline_nodes(
            paragraph, token.get("children", []), style, inherit_text_style=True
        )

    def add_paragraph(self, children: Iterable[dict[str, Any]]) -> None:
        children = list(children)
        if len(children) == 1 and children[0]["type"] == "image":
            self.add_figure(children[0])
            return
        style = self.text_style("body")
        paragraph = self.document.add_paragraph(style="Normal")
        apply_style_to_paragraph(paragraph, style)
        self.add_inline_nodes(paragraph, children, style)

    def add_inline_nodes(
        self,
        paragraph: Any,
        nodes: Iterable[dict[str, Any]],
        style: TextStyle,
        *,
        bold: bool = False,
        italic: bool = False,
        inherit_text_style: bool = False,
    ) -> None:
        for node in nodes:
            kind = node["type"]
            if kind == "text":
                run = paragraph.add_run(node.get("raw", ""))
                if bold:
                    run.bold = True
                if italic:
                    run.italic = True
                if not inherit_text_style:
                    _set_run_style(run, style)
            elif kind in {"strong", "emphasis"}:
                self.add_inline_nodes(
                    paragraph,
                    node.get("children", []),
                    style,
                    bold=bold or kind == "strong",
                    italic=italic or kind == "emphasis",
                    inherit_text_style=inherit_text_style,
                )
            elif kind == "codespan":
                run = paragraph.add_run(node.get("raw", ""))
                if not inherit_text_style:
                    _set_run_style(run, style)
                run.font.name = "Consolas"
                run.font.size = Pt(max(style.size_pt - 1, 8))
            elif kind == "inline_math":
                image = render_latex(node.get("raw", ""), font_size=style.size_pt)
                paragraph.add_run().add_picture(image)
            elif kind == "image":
                self._insert_image(paragraph.add_run(), node)
            elif kind == "link":
                text = self._plain_text(node.get("children", []))
                url = node.get("attrs", {}).get("url", "")
                run = paragraph.add_run(text or url)
                if not inherit_text_style:
                    _set_run_style(run, style)
                run.underline = True
            elif kind in {"linebreak", "softbreak"}:
                paragraph.add_run().add_break()
            elif kind == "strikethrough":
                run = paragraph.add_run(self._plain_text(node.get("children", [])))
                if not inherit_text_style:
                    _set_run_style(run, style)
                run.font.strike = True
            else:
                self.add_inline_nodes(
                    paragraph,
                    node.get("children", []),
                    style,
                    bold=bold,
                    italic=italic,
                    inherit_text_style=inherit_text_style,
                )

    def add_block_math(self, expression: str) -> None:
        block = self.block_style("math-block")
        paragraph = self.document.add_paragraph()
        apply_style_to_paragraph(paragraph, block)
        image = render_latex(
            expression, font_size=self.text_style("body").size_pt + 2
        )
        paragraph.add_run().add_picture(image)

    def _image_stream(self, url: str) -> BytesIO:
        parsed = urlparse(url)
        if parsed.scheme in {"http", "https"}:
            with urlopen(url, timeout=20) as response:
                return BytesIO(response.read())
        path = Path(url)
        if not path.is_absolute():
            path = self.source_dir / path
        if not path.is_file():
            raise FileNotFoundError(f"image not found: {path}")
        return BytesIO(path.read_bytes())

    def _insert_image(self, run: Any, token: dict[str, Any], *, block: bool = False) -> None:
        url = token.get("attrs", {}).get("url", "")
        stream = self._image_stream(url)
        max_width = 6.5 if block else 2.0
        with Image.open(stream) as image:
            width_px, height_px = image.size
            dpi = image.info.get("dpi", (96, 96))[0] or 96
            width_in = min(width_px / dpi, max_width)
            height_in = width_in * height_px / max(width_px, 1)
        stream.seek(0)
        run.add_picture(stream, width=Inches(width_in), height=Inches(height_in))

    def add_figure(self, token: dict[str, Any]) -> None:
        paragraph = self.document.add_paragraph()
        apply_style_to_paragraph(paragraph, self.block_style("image"))
        paragraph.paragraph_format.keep_with_next = True
        self._insert_image(paragraph.add_run(), token, block=True)

        caption = self._plain_text(token.get("children", []))
        if caption:
            style = self.text_style("image-caption")
            p = self.document.add_paragraph(style="Image Caption")
            apply_style_to_paragraph(p, style)
            run = p.add_run(caption)
            _set_run_style(run, style)

    def add_table(self, token: dict[str, Any]) -> None:
        rows: list[list[dict[str, Any]]] = []
        header = next(
            (child for child in token["children"] if child["type"] == "table_head"),
            None,
        )
        body = next(
            (child for child in token["children"] if child["type"] == "table_body"),
            None,
        )
        if header:
            rows.append(header.get("children", []))
        if body:
            rows.extend(row.get("children", []) for row in body.get("children", []))
        if not rows:
            return
        column_count = max(len(row) for row in rows)
        table = self.document.add_table(rows=len(rows), cols=column_count)
        table.style = "Table Grid"
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        table.autofit = False
        widths = [9360 // column_count] * column_count
        widths[-1] += 9360 - sum(widths)
        style = self.text_style("body")
        for row_index, row in enumerate(rows):
            for column_index, cell_token in enumerate(row):
                cell = table.cell(row_index, column_index)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
                _set_cell_margins(cell)
                paragraph = cell.paragraphs[0]
                apply_style_to_paragraph(paragraph, style)
                paragraph.paragraph_format.first_line_indent = None
                paragraph.paragraph_format.space_after = Pt(2)
                self.add_inline_nodes(paragraph, cell_token.get("children", []), style)
                if row_index == 0:
                    _shade_cell(cell, "E7E6E6")
                    for run in paragraph.runs:
                        run.bold = True
        _set_table_geometry(table, widths)

    def add_list(self, token: dict[str, Any], level: int = 0) -> None:
        ordered = bool(token.get("attrs", {}).get("ordered"))
        for item in token.get("children", []):
            for block in item.get("children", []):
                if block["type"] == "list":
                    self.add_list(block, level + 1)
                    continue
                if ordered:
                    style = self.enumerated_list_style()
                    paragraph = self.document.add_paragraph(
                        style=self._list_number_style_name(min(level, 8))
                    )
                else:
                    style_base = "List Bullet"
                    style_name = (
                        style_base
                        if level == 0
                        else f"{style_base} {min(level + 1, 3)}"
                    )
                    paragraph = self.document.add_paragraph(style=style_name)
                    style = self.text_style("body")
                if not ordered:
                    apply_style_to_paragraph(paragraph, style)
                    paragraph.paragraph_format.first_line_indent = None
                self.add_inline_nodes(paragraph, block.get("children", []), style)

    def add_code_block(self, token: dict[str, Any]) -> None:
        paragraph = self.document.add_paragraph()
        paragraph.paragraph_format.left_indent = Inches(0.3)
        paragraph.paragraph_format.right_indent = Inches(0.3)
        paragraph.paragraph_format.space_before = Pt(6)
        paragraph.paragraph_format.space_after = Pt(6)
        run = paragraph.add_run(token.get("raw", "").rstrip())
        run.font.name = "Consolas"
        run.font.size = Pt(10)
        run.font.color.rgb = self.text_style("body").color

    def add_block_quote(self, token: dict[str, Any]) -> None:
        style = self.text_style("body")
        for child in token.get("children", []):
            if child["type"] in {"paragraph", "block_text"}:
                paragraph = self.document.add_paragraph()
                apply_style_to_paragraph(paragraph, style)
                paragraph.paragraph_format.left_indent = Inches(0.35)
                paragraph.paragraph_format.first_line_indent = None
                self.add_inline_nodes(paragraph, child.get("children", []), style)

    @staticmethod
    def _plain_text(nodes: Iterable[dict[str, Any]]) -> str:
        result = []
        for node in nodes:
            if "raw" in node and node["type"] in {"text", "codespan"}:
                result.append(node["raw"])
            result.append(DocxBuilder._plain_text(node.get("children", [])))
        return "".join(result)

    def consume(self, tokens: Iterable[dict[str, Any]]) -> None:
        for token in tokens:
            kind = token["type"]
            if kind == "heading":
                self.add_heading(token)
            elif kind == "paragraph":
                self.add_paragraph(token.get("children", []))
            elif kind == "block_math":
                self.add_block_math(token.get("raw", ""))
            elif kind == "table":
                self.add_table(token)
            elif kind == "list":
                self.add_list(token)
            elif kind == "block_code":
                self.add_code_block(token)
            elif kind == "block_quote":
                self.add_block_quote(token)
            elif kind == "thematic_break":
                raise ValueError(
                    "Markdown thematic breaks are not rendered automatically"
                )


def convert_markdown(
    input_path: str | Path,
    output_path: str | Path,
    config_path: str | Path,
) -> Path:
    input_path = Path(input_path)
    output_path = Path(output_path)
    metadata, markdown_source = parse_frontmatter(
        input_path.read_text(encoding="utf-8")
    )
    markdown = mistune.create_markdown(
        renderer="ast",
        plugins=["table", "math", "strikethrough"],
    )
    tokens = markdown(markdown_source)
    builder = DocxBuilder(load_config(config_path), input_path.parent)
    builder.add_title(str(metadata["title"]))
    builder.consume(tokens)
    output_path.parent.mkdir(parents=True, exist_ok=True)
    builder.document.save(output_path)
    return output_path
