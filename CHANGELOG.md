# 变更记录

合入 `main` 之后，若制作行为有变化，在这里追加一节。格式见 [迭代闭环](docs/ITERATION.md)。

## 2026-11-09

- 字阶写死：封面和内容页标题 32pt，正文至少 20pt、硬底 18pt，图注和图表标签至少 16pt，来源 13pt。结果页标题写这一页的主张。
- 结构审计新增 `font_below_min`（页脚／来源带外的可见正文小于 18pt）和 `low_contrast_text`（相对白底低于 4.5:1，以及文字色 `#BFBFBF`）。两条只警告，退出码不变。
- 公式回退图把 `\frac` 画成 `\dfrac`，PNG 透明；送给 OMML 编译器的字符串不变。短符号用 `scripts/inline_math.py` 收成可编辑小框。屏幕上不写 `x_i` 这种 ASCII 下标。
- 公式回退在本机有 TeX 和 PNG 转换器时，改用真实 TeX 的透明 PNG（至少 300 DPI，`scripts/tex_fallback.py`）。没有 TeX 或 TeX 失败时仍用上面的 mathtext，不整页失败。可编辑 OMML 仍是主路径。pandoc 只在编译器抛错时再试。SVG 和 `.tex` 只作可选旁路，不进幻灯片。
- 用户明确要求推导动画、公式变形或三维直觉时，可以走可选的 ManimGX 路径：`manimgx inspect` 再 `manimgx render` 得到 MP4，`scripts/embed_motion.py` 把它嵌进 PPTX，标题和图注仍可编辑。ManimGX 不导出 PPTX，也不取代 OMML 或原生形状。本机没有 `manimgx`（需要 Python 3.13 或 3.14）时，`--check` 仍退出 0，页面退回可编辑公式或形状。见 [运动讲解](references/motion.md)。
- 数据图可以套 `assets/presentation.mplstyle`（院色，无上边框和右边框）。

## 2026-11-07

- 区分新稿、纯风格迁移和局部修改。保留现有内容的任务不再受拆页、字数上限或补图规则驱动；已明确授权的制作减少重复确认。
- 黄色批注作为内容选材依据，不默认转成幻灯片黄色强调；个人思考留白不代写。新增局部编辑短流程及九类行为复核情景。
- 补充原生公式的函数名正体、变量斜体、中文混排、跨图符号一致性及兼容回退同步规范，点击进入动画分别核对外观与时序。
- 验收区分原稿历史问题与本次修改，不为通过工具检查而增加透明对象或删改用户内容。维护者要求同步 GitHub 时，以远端提交核验和本机同步记录作为闭环。

## 2026-11-04

- 公式可以不靠 pandoc：LaTeX 经 `scripts/vendor/pptmaster_formula/`（ppt-master `4a68cdd`，MIT）编成块级 OMML，求和上下限为 `undOvr`，外层包装不变。回退图用同一条 LaTeX 的 mathtext（stixsans），并按图片自身比例放入占位框。不认识的命令会失败，不把源码写上幻灯片。标题含「方法」但已有带回退的 OMML 时，不再记 `algorithm_page_missing_diagram`。
- 结果对比可以写原生条形图：`add_bar_chart` 无图例、无标题、数据标签打开、零基线、隐藏数值轴，强调柱 `#C00000`、其余 `#BFBFBF`，并修正 python-pptx 的负数 `c:axId`。柱上的数由调用方传入，不在示例里编造。
- 图优先与公式路径：算法、架构和多步流程页必须用原生示意图或明确「缺图」；公式禁止把 LaTeX 源码留在屏幕上，OMML 优先，高 DPI 数学图作可核验兜底。新增 `references/diagrams.md`、`references/formulas.md`、`scripts/diagram_helpers.py`、`scripts/inject_omml.py`。`audit_pptx.py` 增加 `algorithm_page_missing_diagram` 与 `raw_latex_on_screen`，仍只警告、不改变退出码。字数帽、论文封面和内容带下半部规则保持不变。
- 公式路径 1 改为 PowerPoint 原生 OMML：`mc:AlternateContent` / `mc:Choice Requires="a14"` / `a14:m` / `m:oMathPara` / `m:oMath`，并在同一形状区域放高 DPI 回退图。LibreOffice 里公式空白不再当作通过。审计对裸 OMML 或缺少回退图记 `office_math_missing_fallback`，仍只警告。缺图警告改为只看标题中的概念词，封面和已有表、图或图示的页不警告，模型名不再触发。菱形标签保持一行，双栏竖向箭头加长。每页 40 词、论文封面和内容带下半部规则不变。
