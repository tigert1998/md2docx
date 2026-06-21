from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from PIL import Image

from md2docx.config import load_config
from md2docx.converter import convert_markdown, parse_frontmatter


CONFIG = """title:
  chinese-font: SimSun
  latin-font: Times New Roman
  size: 22pt
  color: "#000000"
  space-before: 0pt
  space-after: 8pt
  line-spacing: 28pt
  numbering: null
  first-line-indent: null
  align: center
h1:
  chinese-font: SimHei
  latin-font: Times New Roman
  size: 16pt
  color: "#000000"
  space-before: 0pt
  space-after: 0pt
  line-spacing: 28pt
  numbering: "一、"
  first-line-indent: 2em
  align: left
h2:
  chinese-font: KaiTi
  latin-font: Times New Roman
  size: 16pt
  color: "#000000"
  space-before: 0pt
  space-after: 0pt
  line-spacing: 28pt
  numbering: "（一）"
  first-line-indent: 2em
  align: left
h3:
  chinese-font: FangSong
  latin-font: Times New Roman
  size: 16pt
  color: "#000000"
  space-before: 0pt
  space-after: 0pt
  line-spacing: 28pt
  numbering: "1. "
  first-line-indent: 2em
  align: left
h4:
  chinese-font: FangSong
  latin-font: Times New Roman
  size: 16pt
  color: "#000000"
  space-before: 0pt
  space-after: 0pt
  line-spacing: 28pt
  numbering: "（1）"
  first-line-indent: 2em
  align: left
body:
  chinese-font: FangSong
  latin-font: Times New Roman
  size: 16pt
  color: "#000000"
  space-before: 0pt
  space-after: 0pt
  line-spacing: 28pt
  first-line-indent: 2em
  align: left
image:
  space-before: 1pt
  space-after: 2pt
  line-spacing: 1em
  align: center
image-caption:
  chinese-font: FangSong
  latin-font: Times New Roman
  size: 16pt
  color: "#000000"
  space-before: 0pt
  space-after: 0pt
  line-spacing: 28pt
  numbering: "图1 "
  first-line-indent: null
  align: center
math-block:
  space-before: 3pt
  space-after: 4pt
  line-spacing: 1em
  align: center
"""


def test_frontmatter_is_required() -> None:
    with pytest.raises(ValueError, match="Frontmatter"):
        parse_frontmatter("# 普通标题")
    metadata, body = parse_frontmatter("---\ntitle: 文档标题\n---\n# 一级标题\n")
    assert metadata["title"] == "文档标题"
    assert body.startswith("# 一级标题")


def test_config_is_strict(tmp_path: Path) -> None:
    config = tmp_path / "config.yaml"
    config.write_text("body:\n  size: 12pt\n", encoding="utf-8")
    with pytest.raises(ValueError, match="missing required configuration section"):
        load_config(config)

    config.write_text(CONFIG.replace('  color: "#000000"\n', "", 1), encoding="utf-8")
    with pytest.raises(ValueError, match="title is missing required field.*color"):
        load_config(config)


def test_end_to_end_uses_native_numbering_and_fields(tmp_path: Path) -> None:
    Image.new("RGB", (320, 160), "steelblue").save(tmp_path / "sample.png")
    markdown_path = tmp_path / "sample.md"
    markdown_path.write_text(
        """---
title: 测试文档
---

# 概述

正文含有 **粗体** 和行内公式 $E=mc^2$。

## 子标题

$$
\\frac{a}{b} = c
$$

![示例图片](sample.png)

| 名称 | 数值 |
| --- | ---: |
| 甲 | 10 |
""",
        encoding="utf-8",
    )
    config_path = tmp_path / "config.yaml"
    config_path.write_text(CONFIG, encoding="utf-8")
    output = tmp_path / "sample.docx"
    convert_markdown(markdown_path, output, config_path)

    document = Document(output)
    assert document.paragraphs[0].text == "测试文档"
    assert document.paragraphs[0].style.name == "MD2DOCX Title"
    assert any(p.text == "概述" and p.style.name == "Heading 1" for p in document.paragraphs)
    assert any(p.text == "子标题" and p.style.name == "Heading 2" for p in document.paragraphs)
    assert len(document.tables) == 1

    with ZipFile(output) as archive:
        document_xml = archive.read("word/document.xml").decode("utf-8")
        styles_xml = archive.read("word/styles.xml").decode("utf-8")
        numbering_xml = archive.read("word/numbering.xml").decode("utf-8")
        media = [name for name in archive.namelist() if name.startswith("word/media/")]

    assert "md2docx heading numbering" in numbering_xml
    assert 'w:val="chineseCounting"' in numbering_xml
    assert 'w:val="%1、"' in numbering_xml
    assert 'w:val="（%2）"' in numbering_xml
    assert "md2docx image caption numbering" in numbering_xml
    assert 'w:val="图%1 "' in numbering_xml
    assert "SEQ Figure" not in document_xml
    assert 'w:pStyle w:val="ImageCaption"' in numbering_xml
    assert 'w:eastAsia="SimHei"' in numbering_xml
    assert 'w:eastAsia="KaiTi"' in numbering_xml
    assert 'w:eastAsia="FangSong"' in numbering_xml
    assert 'w:ascii="Times New Roman"' in numbering_xml
    assert 'w:sz w:val="32"' in numbering_xml
    assert 'w:szCs w:val="32"' in numbering_xml
    assert "MD2DOCXTitle" in styles_xml
    assert "w:pBdr" not in document_xml.split("测试文档", 1)[0][-500:]
    assert 'w:color w:val="000000"' in styles_xml
    assert 'w:before="60" w:after="80" w:line="240"' in document_xml
    assert 'w:before="20" w:after="40" w:line="240"' in document_xml
    assert len(media) >= 3
