# AGENTS.md

本项目负责将 Markdown 转换为 `.docx`，排版样式由 `config.yaml` 单一驱动。

## 核心原则

**样式与内容分离**：`config.yaml` 是唯一事实来源（SSOT）。YAML 顶级键必须严格映射为 Word 同名样式。

> **层级实例化**：`ordered-list` 与 `unordered-list` 需实例化为 `*-1`、`*-2` 等层级样式。

**禁止硬编码**：生成时**不得**对段落（Paragraph）或文本块（Run）进行字体、字号、颜色等直接格式覆写。必须引用 YAML 中的命名样式。

> **例外**：粗体、斜体、删除线、链接允许直接格式化。

## 配置约束

`config.yaml` 兼具 Schema 定义功能，**禁止修改**。解析时如遇字段缺失，须立即抛出校验错误并终止执行。

**适用于列表的缩进规则**：
`indent-before-text-increment` 定义层级增量。第 $N$ 层实际左缩进计算公式：
$$
\text{indent-before-text}+(N - 1) \times \text{indent-before-text-increment}
$$
