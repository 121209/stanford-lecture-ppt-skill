# 变更记录

合入 `main` 之后，若制作行为有变化，在这里追加一节。格式见 [迭代闭环](docs/ITERATION.md)。

## 2026-11-11

- 屏幕字数上限从每页 40 词改为 56 词（中文仍按 2 字 = 1 词，封面、页码和 ≤15pt 来源行不计）。56 由 nature-paper2ppt 的「一行标题 + 最多 3 条短句 + 一句 24 字读解」换算，见 [屏幕密度](references/density.md)。结果页把关键数字、对照和一句读解放在屏幕上；备注改写讲解。字号硬底、配色、公式、示意图和内容带下半部不变。`audit_pptx.py` 的默认上限是 56，`--word-cap` 可改。cell-ppt 没有整页字数，只借用「短标注留在对象旁边」。
- `add_table` 的 56 词预算计入页上已有的标题。页码和 ≤15pt 来源行不计。只算单元格、表题和结论时，加上标题会超过审计上限。
- 新增 `scripts/doctor.py`。检查 Python 3.10+ 与 python-pptx；TeX、pdftocairo、LibreOffice、ManimGX 可选，缺失仍退出 0，并写明降级。`--write` 写出不含密钥、不含作者机器路径要求的 runtime profile。不把 PowerPoint、COM、WPS 或付费 API 当作必需项。
- 打开已有 PPTX 默认用 `scripts/safe_edit.py` 另存工作副本并记录源文件 sha256。只有明确传入 `--in-place` 才覆盖原文件。本来就写到新路径的助手保持原样，示例生成不会改写 `assets/demo.pptx`。
- 未点名的既有形状必须保留。`tests/test_preserve_existing.py` 用 `PREEXISTING_SENTINEL` 在临时目录里核对会往页上加内容的助手。没有 LibreOffice 时跳过渲染。

## 2026-11-10

- 风格里写过、脚本里没有的日常对象补成函数：原生表格、论文封面、文字双栏、页脚页码、图加可编辑图注、细线定义框、一行节分隔。见 `scripts/layout_helpers.py` 与 [版式构件](references/layouts.md)。公式仍是 OMML，算法和架构仍是原生形状。
- 原生图表在条形图之外增加折线和散点（`add_line_chart`、`add_scatter_chart`）。数字仍由调用方传入。样张里的整数是占位，不是测量。
- 用户要求单击出现或逐条进入时，用 `scripts/appear.py` 写原生进入动画（瞬间出现，或 0.5–0.7 秒淡入）。不调用就没有动画。不组合对象，也不取代 OMML 或 ManimGX 的影片。
- 匿名样张 `assets/layout-samples/basics.pptx`。

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
