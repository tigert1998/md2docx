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

from .config import BlockStyle, TextStyle, apply_style_to_paragraph, load_config
from .math_render import render_latex


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


def _remove_paragraph_border(element: Any) -> None:
    ppr = element.get_or_add_pPr()
    border = ppr.find(qn("w:pBdr"))
    if border is not None:
        ppr.remove(border)


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


def _install_heading_numbering(
    document: Any, heading_styles: list[tuple[int, TextStyle]]
) -> int:
    numbering = document.part.numbering_part.element
    abstract_id = _next_numbering_id(numbering, "abstractNum", "abstractNumId")
    num_id = _next_numbering_id(numbering, "num", "numId")

    abstract = OxmlElement("w:abstractNum")
    abstract.set(qn("w:abstractNumId"), str(abstract_id))
    multi = OxmlElement("w:multiLevelType")
    multi.set(qn("w:val"), "multilevel")
    abstract.append(multi)
    name = OxmlElement("w:name")
    name.set(qn("w:val"), "md2docx heading numbering")
    abstract.append(name)

    for level, style in heading_styles:
        lvl = OxmlElement("w:lvl")
        lvl.set(qn("w:ilvl"), str(level - 1))
        start = OxmlElement("w:start")
        start.set(qn("w:val"), "1")
        lvl.append(start)
        if level > 1:
            restart = OxmlElement("w:lvlRestart")
            restart.set(qn("w:val"), str(level - 1))
            lvl.append(restart)

        if style.numbering is None:
            number_format, level_text = "none", ""
        else:
            number_format, level_text = _numbering_level(style.numbering, level - 1)
        num_fmt = OxmlElement("w:numFmt")
        num_fmt.set(qn("w:val"), number_format)
        lvl.append(num_fmt)
        text = OxmlElement("w:lvlText")
        text.set(qn("w:val"), level_text)
        lvl.append(text)
        suffix = OxmlElement("w:suff")
        suffix.set(qn("w:val"), "nothing")
        lvl.append(suffix)
        paragraph_style = OxmlElement("w:pStyle")
        paragraph_style.set(qn("w:val"), f"Heading{level}")
        lvl.append(paragraph_style)
        justification = OxmlElement("w:lvlJc")
        justification.set(qn("w:val"), "left")
        lvl.append(justification)
        abstract.append(lvl)

    numbering.append(abstract)
    num = OxmlElement("w:num")
    num.set(qn("w:numId"), str(num_id))
    abstract_ref = OxmlElement("w:abstractNumId")
    abstract_ref.set(qn("w:val"), str(abstract_id))
    num.append(abstract_ref)
    numbering.append(num)
    return num_id


def _set_style_numbering(style: Any, num_id: int, level: int, enabled: bool) -> None:
    ppr = style.element.get_or_add_pPr()
    existing = ppr.find(qn("w:numPr"))
    if existing is not None:
        ppr.remove(existing)
    if not enabled:
        return
    num_pr = OxmlElement("w:numPr")
    ilvl = OxmlElement("w:ilvl")
    ilvl.set(qn("w:val"), str(level - 1))
    num = OxmlElement("w:numId")
    num.set(qn("w:val"), str(num_id))
    num_pr.extend([ilvl, num])
    ppr.append(num_pr)


def _enable_field_updates(document: Any) -> None:
    settings = document.settings.element
    update = settings.find(qn("w:updateFields"))
    if update is None:
        update = OxmlElement("w:updateFields")
        settings.append(update)
    update.set(qn("w:val"), "true")


def _append_seq_field(paragraph: Any, name: str, cached_value: int, style: TextStyle) -> None:
    begin_run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    begin_run._r.append(begin)

    instr_run = paragraph.add_run()
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = f" SEQ {name} \\* ARABIC "
    instr_run._r.append(instr)

    separate_run = paragraph.add_run()
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    separate_run._r.append(separate)

    result_run = paragraph.add_run(str(cached_value))
    _set_run_style(result_run, style)

    end_run = paragraph.add_run()
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    end_run._r.append(end)


def _caption_pattern(pattern: str | None) -> tuple[str, str]:
    if pattern is None:
        return "", ""
    if "{n}" in pattern:
        return tuple(pattern.split("{n}", 1))  # type: ignore[return-value]
    match = re.search(r"\d+", pattern)
    if not match:
        raise ValueError(
            "image-caption.numbering must contain a decimal number or {n}"
        )
    return pattern[: match.start()], pattern[match.end() :]


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
        self, styles: dict[str, BlockStyle | TextStyle], source_dir: Path
    ):
        self.styles = styles
        self.source_dir = source_dir
        self.document = Document()
        self.figure_count = 0
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

    def _configure_document(self) -> None:
        section = self.document.sections[0]
        section.top_margin = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin = Inches(1)
        section.right_margin = Inches(1)
        section.page_width = Inches(8.5)
        section.page_height = Inches(11)
        _enable_field_updates(self.document)

        if "MD2DOCX Title" not in self.document.styles:
            title_style = self.document.styles.add_style(
                "MD2DOCX Title", WD_STYLE_TYPE.PARAGRAPH
            )
            title_style.base_style = self.document.styles["Normal"]
        else:
            title_style = self.document.styles["MD2DOCX Title"]
        self._configure_word_style(title_style, self.text_style("title"))
        _remove_paragraph_border(title_style.element)

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
        num_id = _install_heading_numbering(self.document, heading_styles)
        for level, config_style in heading_styles:
            word_style = self.document.styles[f"Heading {level}"]
            self._configure_word_style(word_style, config_style)
            _set_style_numbering(
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

    def add_title(self, title: str) -> None:
        style = self.text_style("title")
        paragraph = self.document.add_paragraph(style="MD2DOCX Title")
        apply_style_to_paragraph(paragraph, style)
        _remove_paragraph_border(paragraph._p)
        run = paragraph.add_run(title.strip())
        _set_run_style(run, style)

    def add_heading(self, token: dict[str, Any]) -> None:
        level = int(token["attrs"]["level"])
        key = f"h{level}"
        if key not in self.styles:
            raise ValueError(
                f"Markdown heading level {level} requires configuration section {key}"
            )
        style = self.text_style(key)
        paragraph = self.document.add_paragraph(style=f"Heading {level}")
        apply_style_to_paragraph(paragraph, style)
        self.add_inline_nodes(paragraph, token.get("children", []), style)

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
    ) -> None:
        for node in nodes:
            kind = node["type"]
            if kind == "text":
                run = paragraph.add_run(node.get("raw", ""))
                run.bold, run.italic = bold, italic
                _set_run_style(run, style)
            elif kind in {"strong", "emphasis"}:
                self.add_inline_nodes(
                    paragraph,
                    node.get("children", []),
                    style,
                    bold=bold or kind == "strong",
                    italic=italic or kind == "emphasis",
                )
            elif kind == "codespan":
                run = paragraph.add_run(node.get("raw", ""))
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
                _set_run_style(run, style)
                run.underline = True
            elif kind in {"linebreak", "softbreak"}:
                paragraph.add_run().add_break()
            elif kind == "strikethrough":
                run = paragraph.add_run(self._plain_text(node.get("children", [])))
                _set_run_style(run, style)
                run.font.strike = True
            else:
                self.add_inline_nodes(
                    paragraph, node.get("children", []), style, bold=bold, italic=italic
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
            self.figure_count += 1
            style = self.text_style("image-caption")
            p = self.document.add_paragraph(style="Image Caption")
            apply_style_to_paragraph(p, style)
            prefix, suffix = _caption_pattern(style.numbering)
            if prefix:
                run = p.add_run(prefix)
                _set_run_style(run, style)
            _append_seq_field(p, "Figure", self.figure_count, style)
            run = p.add_run(suffix + caption)
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
                style_base = "List Number" if ordered else "List Bullet"
                style_name = (
                    style_base if level == 0 else f"{style_base} {min(level + 1, 3)}"
                )
                paragraph = self.document.add_paragraph(style=style_name)
                style = self.text_style("body")
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
