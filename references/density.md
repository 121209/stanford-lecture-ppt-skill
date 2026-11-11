# 屏幕密度

这一页记录两个公开技能通常在一页上放多少内容，供组会和论文精读参考。本仓库不把这些数字当成每页上限。一页放什么，由论文内容和这一页要讲的主题决定。字号、配色、公式和示意图仍以各自参考为准。

下面若把汉字折成词，口径与 `scripts/audit_pptx.py` 的计数相同：中文 2 个汉字算 1 词，英文和数字按词计。页码和 ≤15pt 的来源行不计。这只是为了读来源，不是检查线。

## 量过什么

两个来源的公开仓库里都没有样张 PPTX，也没有渲染页。下面的数字来自它们写明的规则和检查阈值，不是从成品页统计出来的。查阅日期是 2026-11-11 的 `main`。

### cell-ppt

- [yrui-cmd/cell_figure_ppt](https://github.com/yrui-cmd/cell_figure_ppt) 的 `plugins/cell-ppt/skills/cell-ppt`（仓库原名 Cell_ppt）
- [yrui-cmd/Cell](https://github.com/yrui-cmd/Cell/tree/main/plugins/cell-ppt-edited) 的 `cell-ppt-edited`

读过 `SKILL.md`、`references/workflow.md`、`references/backends.md`、`references/platform-contract.json`、`tests/fixtures/text-manifest.json`，以及 Cell 仓库里插件目录和 `skills/cell-ppt-edited/` 下两份相同的 `SKILL.md`。没有 examples 目录，没有 PPTX。

量到的结果：

- 没有整页词数、条数或备注篇幅。
- 图上每个可见文字块单独做成一个可编辑文本框，位置和字号沿用原图。夹具里只有一条标签 `Cell_ppt`，`font_size` 为 0.08（相对坐标），不是讲义字号。
- `slideMarginPoints` 为 18。这是画进幻灯片时的边距，不是正文字号。
- `cell-ppt-edited` 只改已经打开的对象，可以多段混排，同样没有字数上限。

取用的做法：短标注留在它所指的对象旁边，各自一条，不收成一段，也不推进备注。

不取用：付费接口、密钥收集、`cell_no_ai`、Windows COM、WPS，以及这两处技能要求附在回复末尾的结束语。

### nature-paper2ppt

[Yuan1z0825/nature-skills 的 nature-paper2ppt](https://github.com/Yuan1z0825/nature-skills/tree/main/skills/nature-paper2ppt)。

读过 `references/design-and-layout.md`、`static/core/workflow.md`、`static/core/output-and-quality.md`、`references/self-review.md`、`references/figure-assets.md`、六个 paper-type 片段，以及 `scripts/audit_pptx_quality.py`。仓库里没有 PPTX。paper-type 片段没有另一套字数。

`design-and-layout.md` 写明的屏幕预算：

- 标题一行；只有竖向还够时才允许两行。
- 普通页 2–3 条，每条大约不超过 18 个汉字，或 8–10 个英文词。
- 结果页：一句短解读，再加最多 2 个紧凑标注。
- 卡片正文一句，通常不超过 24 个汉字。
- 指标标注尽量一行。来源短，不跟图抢。
- 讲解放进演讲者备注。超出就拆页，不靠缩小字号。
- 结果页构图：图约占画面 55–75%，窄解读条约占 20–30%，底下一句收束。用 1–3 个数字标注，或一条短解读。不要把每个 panel 都写成一段。
- 工作流第 6 步还写了「2–4 条短句、图注和一句 takeaway」，并要求服从上面这份预算。条数以 2–3 为准；多出来的那一条就是收束句，不再加第四段正文。

`audit_pptx_quality.py` 没有整页词数帽。单个文本框不少于 60 个字符、并且超过 80 字/平方英寸时，记一条中等的 `text_overload`。它不统计备注，也不规定字号。

该技能自己的字阶大约是标题 24–32pt、正文 12–16pt、图注和来源 7–9pt。本仓库不采用这档小字。

## 来源自己的一页大概多满

按 nature-paper2ppt 写明的最满一页，用「2 个汉字 = 1 词」折算，只作对照：

| 屏幕上的一块 | 来源写法 | 大约词数 |
| --- | --- | --- |
| 标题一行 | 本仓库标题若按 32pt，版心一行大约 24–28 个汉字 | 14 |
| 3 条短句 | 每条英文上限 10 词；18 个汉字只有 9 词 | 30 |
| 一句读解或收束 | 卡片正文 24 个汉字 | 12 |
| 合计 | | 约 56 |

结果页在它的写法里通常更紧：一行标题、最多 2 个标注、一句读解，大约 46 词。cell-ppt 没有这样的合计。这些都不是本仓库的上限。

## 本仓库因此怎么做

- 不设每页词数上限。一页一个主题。屏幕上放论文里看懂这个主题所需的关键数字、对照、各 panel 的读解，以及会改变含义的条件或定义。
- 备注写讲解：怎么带着看、推导、例子、局限的展开、过渡和完整出处。关键事实不进备注。
- 一个主题在字号硬底下装不下、文字会溢出或互相遮挡时，拆成两页。不靠删掉关键事实或把它们推进备注来留在一页。
- 字号不变：标题 32pt，正文至少 20pt、硬底 18pt，图注 20pt。不把字号降到硬底以下，不拉伸位图。内容带下半部仍不能整段留空。
- `scripts/audit_pptx.py` 仍会数屏幕字，默认不据此警告。只有传入 `--word-cap N` 才记 `onscreen_text_over_word_cap`，封面不计。不传就没有这条警告。
- `scripts/layout_helpers.py` 不再用整页词数拒绝表格或双栏。单元格和半栏里的一条仍要短到 20pt 放得下一行；放不下就拆页，不缩字。
