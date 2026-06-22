from pathlib import Path
from zipfile import ZipFile

import pytest
from docx import Document
from PIL import Image

from md2docx.config import REQUIRED_SECTIONS, load_config
from md2docx.converter import convert_markdown, parse_frontmatter
from md2docx.math_render import render_latex


PROJECT_ROOT = Path(__file__).parents[1]
CONFIG_PATH = PROJECT_ROOT / "config.yaml"
CONFIG = CONFIG_PATH.read_text(encoding="utf-8")


def test_frontmatter_is_required() -> None:
    with pytest.raises(ValueError, match="Frontmatter"):
        parse_frontmatter("# 普通标题")
    metadata, body = parse_frontmatter("---\ntitle: 文档标题\n---\n# 一级标题\n")
    assert metadata["title"] == "文档标题"
    assert body.startswith("# 一级标题")


def test_real_config_is_the_strict_schema(tmp_path: Path) -> None:
    styles = load_config(CONFIG_PATH)
    assert set(styles) == REQUIRED_SECTIONS
    assert styles["title"].first_line_indent is not None
    assert styles["title"].first_line_indent.unit == "pt"
    assert styles["ordered-list"].indent_before_text_increment is not None
    assert styles["unordered-list"].indent_before_text_increment is not None

    config = tmp_path / "config.yaml"
    config.write_text(CONFIG.replace("inline-code:", "missing-inline-code:", 1), encoding="utf-8")
    with pytest.raises(
        ValueError, match="missing required configuration section.*inline-code"
    ):
        load_config(config)

    config.write_text(
        CONFIG.replace('  color: "#000000"\n', "", 1),
        encoding="utf-8",
    )
    with pytest.raises(ValueError, match="title is missing required field.*color"):
        load_config(config)


def test_end_to_end_uses_yaml_named_styles_without_direct_formatting(
    tmp_path: Path,
) -> None:
    Image.new("RGB", (320, 160), "steelblue").save(tmp_path / "sample.png")
    markdown = tmp_path / "sample.md"
    markdown.write_text(
        """---
title: 测试文档
---

# 概述

正文含有 **粗体**、*斜体*、`code` 和行内公式 $E=mc^2$。

1. 有序一层
   1. 有序二层
      1. 有序三层

- 无序一层
  - 无序二层

```text
code block
```

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
    output = tmp_path / "sample.docx"
    convert_markdown(markdown, output, CONFIG_PATH)

    document = Document(output)
    assert document.paragraphs[0].style.name == "title"
    assert next(p for p in document.paragraphs if p.text == "概述").style.name == "h1"
    assert next(p for p in document.paragraphs if p.text == "code block").style.name == "code-block"

    ordered = [p for p in document.paragraphs if p.style.name.startswith("ordered-list-")]
    unordered = [
        p for p in document.paragraphs if p.style.name.startswith("unordered-list-")
    ]
    assert [p.style.name for p in ordered] == [
        "ordered-list-1",
        "ordered-list-2",
        "ordered-list-3",
    ]
    assert [p.style.name for p in unordered] == [
        "unordered-list-1",
        "unordered-list-2",
    ]
    assert [
        round(document.styles[f"ordered-list-{level}"].paragraph_format.left_indent.pt)
        for level in range(1, 4)
    ] == [0, 32, 64]
    assert [
        round(document.styles[f"unordered-list-{level}"].paragraph_format.left_indent.pt)
        for level in range(1, 3)
    ] == [0, 32]

    body = next(p for p in document.paragraphs if p.text.startswith("正文含有"))
    for run in body.runs:
        rpr = run._r.rPr
        if rpr is None:
            continue
        tags = {child.tag.rsplit("}", 1)[-1] for child in rpr}
        assert tags <= {"b", "bCs", "i", "iCs", "rStyle", "drawing"}
    for paragraph in document.paragraphs:
        if paragraph.text and paragraph.style.name not in {"body"}:
            ppr = paragraph._p.pPr
            if ppr is not None:
                tags = {child.tag.rsplit("}", 1)[-1] for child in ppr}
                assert tags <= {"pStyle"}

    with ZipFile(output) as archive:
        styles_xml = archive.read("word/styles.xml").decode("utf-8")
        numbering_xml = archive.read("word/numbering.xml").decode("utf-8")
        document_xml = archive.read("word/document.xml").decode("utf-8")

    for name in REQUIRED_SECTIONS:
        assert f'w:name w:val="{name}"' in styles_xml
    for level in range(1, 10):
        assert f'w:name w:val="ordered-list-{level}"' in styles_xml
        assert f'w:name w:val="unordered-list-{level}"' in styles_xml
    assert "md2docx ordered-list numbering" in numbering_xml
    assert "md2docx unordered-list numbering" in numbering_xml
    assert 'w:pStyle w:val="ordered-list-1"' in numbering_xml
    assert 'w:pStyle w:val="unordered-list-1"' in numbering_xml
    assert 'w:left="0"' in numbering_xml
    assert 'w:left="640"' in numbering_xml
    assert 'w:left="1280"' in numbering_xml
    assert "<w:rFonts" not in document_xml
    assert "<w:sz " not in document_xml
    assert "<w:color " not in document_xml


def test_local_image_path_supports_chinese_characters(tmp_path: Path) -> None:
    image_name = "结算包拓扑关系示意图.png"
    Image.new("RGB", (32, 16), "steelblue").save(tmp_path / image_name)
    markdown = tmp_path / "中文文档.md"
    markdown.write_text(
        f"---\ntitle: 中文图片路径测试\n---\n\n![拓扑关系]({image_name})\n",
        encoding="utf-8",
    )
    output = tmp_path / "输出文档.docx"
    convert_markdown(markdown, output, CONFIG_PATH)
    with ZipFile(output) as archive:
        assert any(name.startswith("word/media/") for name in archive.namelist())


def test_math_fontset_renders_without_configuring_font_family(
    monkeypatch: pytest.MonkeyPatch,
) -> None:
    configured_rc: dict[str, str] = {}
    original_rc_context = __import__("matplotlib").rc_context

    def capture_rc_context(rc: dict[str, str]):
        configured_rc.update(rc)
        return original_rc_context(rc)

    monkeypatch.setattr("md2docx.math_render.mpl.rc_context", capture_rc_context)
    image = render_latex("E=mc^2", fontset="cm")
    assert image.read(8) == b"\x89PNG\r\n\x1a\n"
    assert configured_rc == {"mathtext.fontset": "cm"}


def test_unknown_math_fontset_fails_instead_of_silently_falling_back() -> None:
    with pytest.raises(ValueError, match="unsupported math fontset"):
        render_latex("x", fontset="Definitely Missing Math Font")
