# md2docx

一个基于 `mistune`、`python-docx` 和 `PyYAML` 的 Markdown 转 DOCX 工具。

## 功能

- 使用 YAML 配置标题、各级标题、正文和图片 caption 的字体、字号、颜色、段距、行距、首行缩进和对齐方式
- 使用 Word 原生多级编号生成 `一、`、`（一）`、`1. `、`（1）` 等标题编号
- 支持 `$...$` 行内公式和 `$$...$$` 块公式
- 支持本地图片、HTTP(S) 图片和原生 SVG；独占一段的图片会使用 alt 文本生成 caption
- 支持 Markdown 表格、粗体、斜体、删除线、超链接、水平线、列表、引用和代码块

公式使用 `inline-math` 或 `math-block` 的 `latin-font` 作为 Matplotlib
`mathtext.fontset`，高分辨率渲染后嵌入，因此无需安装 LaTeX。MathText
支持常用 LaTeX 数学语法，但并非完整 TeX 引擎。

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

文件顶部的 YAML Frontmatter 是可选的；存在时必须包含非空 `title`，并使用
`title` 样式生成文档标题。没有 Frontmatter 时直接从 Markdown 正文开始转换。
Markdown 的 `#`、`##`、`###`、`####` 分别对应 `h1`、`h2`、`h3`、`h4`。

## 配置说明

- 所有示例配置项均为必填项；缺少配置节或字段时会直接报错并指出名称
- 字号使用 `pt`，例如 `16pt`
- 段前、段后、行距、首行缩进、悬挂缩进和文本前缩进支持 `pt` 或 `em`
- `align` 可取 `left`、`center`、`right`、`justify`
- `color` 使用六位十六进制颜色，例如 `#000000`
- 标题 `numbering` 可使用示例模板，也可使用 `{n}` 和 `{cn}` 占位符；生成结果是可继续编辑的 Word 原生编号
- 图片图注编号使用绑定到 `Image Caption` 样式的 Word 原生单级编号，可在 Word 中继续插入并自动递增
- 除 `ordered-list`、`unordered-list` 配置模板外，YAML 顶级键会创建同名 Word 样式；`inline-code`、`inline-math` 为字符样式，其余为段落样式
- 有序、无序列表会实例化为 `ordered-list-1`、`ordered-list-2`、`unordered-list-1` 等层级样式；第 N 层左缩进为 `indent-before-text + hanging-indent + (N - 1) × indent-before-text-increment`
- 生成正文时只引用命名样式，不直接覆写字体、字号、缩进或段距；粗体、斜体、删除线以及超链接的蓝色下划线允许直接格式化
- 表格表头和表体分别使用 `table-header`、`table-body`；Markdown 分割线指定的列对齐会覆盖对应单元格段落的 `align`
- 所有原生编号均不会自动添加空格或制表符；编号与正文的间距完全由 `numbering` 字段中的手动空格控制
