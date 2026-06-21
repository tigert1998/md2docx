# md2docx

一个基于 `mistune`、`python-docx` 和 `PyYAML` 的 Markdown 转 DOCX 工具。

## 功能

- 使用 YAML 配置标题、各级标题、正文和图片 caption 的字体、字号、颜色、段距、行距、首行缩进和对齐方式
- 使用 Word 原生多级编号生成 `一、`、`（一）`、`1. `、`（1）` 等标题编号
- 支持 `$...$` 行内公式和 `$$...$$` 块公式
- 支持本地图片和 HTTP(S) 图片；独占一段的图片会使用 alt 文本生成 caption
- 支持 Markdown 表格、粗体、斜体、删除线、链接、列表、引用和代码块

公式采用 Times 风格的 STIX 数学字体高分辨率渲染后嵌入，因此无需安装 LaTeX。MathText 支持常用 LaTeX 数学语法，但并非完整 TeX 引擎。

## 安装

```powershell
uv sync
```

## 使用

```powershell
uv run md2docx input.md -c config.yaml -o output.docx
```

也可以作为模块运行：

```powershell
uv run python -m md2docx input.md -c config.yaml -o output.docx
```

若省略 `-o`，输出文件与输入文件同名，仅扩展名改为 `.docx`。

## Markdown 示例

```markdown
---
title: 示例文档
---

# 第一部分

正文中的行内公式为 $E=mc^2$。

$$
\int_0^\infty e^{-x^2}\,dx = \frac{\sqrt{\pi}}{2}
$$

![系统结构图](images/architecture.png)

| 项目 | 数值 |
| --- | ---: |
| A | 10 |
| B | 20 |
```

文档标题必须来自文件顶部的 YAML Frontmatter。Markdown 的 `#`、`##`、`###` 会严格对应 `h1`、`h2`、`h3`，不再进行级别偏移。

## 配置说明

- 所有示例配置项均为必填项；缺少配置节或字段时会直接报错并指出名称
- 尺寸和段距使用 `pt`，例如 `16pt`
- 行距支持 `pt` 或 `em`
- 首行缩进使用 `em`，例如 `2em`
- `align` 可取 `left`、`center`、`right`、`justify`
- `color` 使用六位十六进制颜色，例如 `#000000`
- 标题 `numbering` 可使用示例模板，也可使用 `{n}` 和 `{cn}` 占位符；生成结果是可继续编辑的 Word 原生编号
- 图片图注编号使用绑定到 `Image Caption` 样式的 Word 原生单级编号，可在 Word 中继续插入并自动递增
- 有序列表使用 `enumerated-list` 配置修改 Word 内置 `List Number` 样式；第 N 层的左缩进为 `(N - 1) × indent-before-text-increment`
- 所有原生编号均不会自动添加空格或制表符；编号与正文的间距完全由 `numbering` 字段中的手动空格控制
