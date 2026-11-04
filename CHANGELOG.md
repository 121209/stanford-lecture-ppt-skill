# 变更记录

合入 `main` 之后，若制作行为有变化，在这里追加一节。格式见 [迭代闭环](docs/ITERATION.md)。

## 2026-11-04

- 图优先与公式路径：算法、架构和多步流程页必须用原生示意图或明确「缺图」；公式禁止把 LaTeX 源码留在屏幕上，OMML 优先，高 DPI 数学图作可核验兜底。新增 `references/diagrams.md`、`references/formulas.md`、`scripts/diagram_helpers.py`、`scripts/inject_omml.py`。`audit_pptx.py` 增加 `algorithm_page_missing_diagram` 与 `raw_latex_on_screen`，仍只警告、不改变退出码。字数帽、论文封面和内容带下半部规则保持不变。
- 公式路径 1 改为 PowerPoint 原生 OMML：`mc:AlternateContent` / `mc:Choice Requires="a14"` / `a14:m` / `m:oMathPara` / `m:oMath`，并在同一形状区域放高 DPI 回退图。LibreOffice 里公式空白不再当作通过。审计对裸 OMML 或缺少回退图记 `office_math_missing_fallback`，仍只警告。缺图警告改为只看标题中的概念词，封面和已有表、图或图示的页不警告，模型名不再触发。菱形标签保持一行，双栏竖向箭头加长。每页 40 词、论文封面和内容带下半部规则不变。
