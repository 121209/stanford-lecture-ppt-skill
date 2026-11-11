# 版式构件

这一页是日常可编辑对象的做法。视觉规则仍以 [风格与版式](style.md) 为准。公式继续走 [公式](formulas.md)，算法和架构继续走 [示意图](diagrams.md)。用户明确要求的影片继续走 [运动讲解](motion.md)。

## 已经有的，和这里补上的

已经有脚本、不要在这里重做：

- 流程图、架构、判断菱形：`scripts/diagram_helpers.py`
- 原生条形图：`scripts/chart_helpers.py` 的 `add_bar_chart`
- 带图片回退的 OMML：`scripts/inject_omml.py`
- 可选的 ManimGX MP4：`scripts/embed_motion.py`

以前只写在风格或局部编辑里、没有函数的，用下面两个脚本：

- `scripts/layout_helpers.py`：封面、原生表格、文字双栏、页脚和页码、图加图注、定义框、节分隔
- `scripts/appear.py`：可选的单击出现。不调用就没有动画

样张是 `assets/layout-samples/basics.pptx`。里面的人名、会议和整数都是占位，不是论文结果，也不能当作 `assets/demo.pptx` 的科学内容。

```text
python scripts/layout_helpers.py --samples assets/layout-samples/basics.pptx
```

幻灯片必须是 16:9，12192000×6858000 EMU。`new_deck()` 会建成这个尺寸。

## 共同约束

- 白底，正文 `#222222`，强调 `#C00000`，辅助 `#606060`。不要阴影、渐变、大色带。
- 标题 32pt，正文和表内文字 20pt。来源 13pt，而且要带 `来源：` 或 `Source:`，否则审计会把它当成缩小的正文。
- 页码只写数字，例如 `4` 或 `04`。不要写 `3 / 9`。
- 表、图或定义框停在内容带上半部时，下半部放一行结论或读图说明。行高不要拉空。
- 这些函数画出的文字、单元格和图注都可以选中改。它们不生成整页 PNG。

## 封面

`add_cover(slide, title, presenter, date, kicker=None, authors=None, venue=None)`

论文讲解要传作者，以及会议或期刊和年份。这两行是 20pt，放在版心，不放进页脚。不是论文时可以不传作者和会议。署名和日期仍然要有。封面应是第 1 页。

## 表格

`add_table(slide, headers, rows, region, conclusion=None, caption=None, source=None, page=None)`

`region` 是 EMU 的 `(x, y, w, h)`。表是原生 `a:tbl`，表头和单元格都能改。行高按 20pt 固定，不会把表拉高去填满 `region`。表和表题都在上半部、又没给 `conclusion` 时，函数直接失败。

单元格保持短句，大约 6 词以内，方便 20pt 一行放下。关键数字留在单元格里。没有整页词数上限。超过一行的说明如果会把字号压下去，就拆页或放进备注里的讲解，不要把字号降下去，也不要把关键数字移出单元格。最多 6 列、5 行数据。数字由调用方传入。

## 文字双栏

`add_two_column(slide, title, left_title, left_points, right_title, right_points, conclusion, source=None, page=None)`

左右各 1–3 条，每条最多 10 词，这样 20pt 在半栏里仍是一行。比较的内容多于这个时拆成两页，不缩字。左栏是论文工作，右栏是笔记时，结论要写明右栏不是论文已证。这不是 `diagram_helpers.py` 的 `columns`：那边是架构节点，这边是文字。

## 页脚

`add_footer(slide, source=None, page=None)`

一页只加一次。`format_page(4)` 得到 `04`。

## 图和图注

`add_picture_caption(slide, image, region, caption)`

`image` 是路径或 PNG 字节。图注是 20pt 的一行。关键数字、对照和 panel 读解需要另写时，用单独的 20pt 文本框放在图旁，不写进备注。放不下或挡住图时拆页，不缩字。图如果还在上半部，图注会挪到下半部，充当那一句读图说明。论文原图可以这样嵌。不要用这张图代替原生表、原生图或 OMML。

## 定义框

`add_callout(slide, region, term, body, read=None)`

白底矩形，1.25pt 红线，没有阴影和渐变。术语是红色，解释是黑色。框在上半部时必须传 `read`。

## 节分隔

`add_section_divider(slide, title, kicker=None, source=None, page=None)`

只有一行 32pt 标题。标题框和字一样高，并跨在内容带中线上，不用拉高的空框假装填满。小标签写在标题下面，因为审计把最靠上的文字当成页标题。这里不做议程列表，也不做 SmartArt。

## 折线和散点

仍在 `scripts/chart_helpers.py`，也从 `diagram_helpers.py` 转一次。一条序列，没有图例和图表标题，坐标轴保留，网格线关掉，线和点是 `#C00000`。散点不连线。非负数据从 0 起；有负数时不强制 0。调用方传入每一个数。折线最多 8 个点，散点最多 12 个点。

```python
add_line_chart(slide, ["甲", "乙", "丙", "丁"], [2, 5, 9, 9], region)
add_scatter_chart(slide, [(1, 2), (2, 8), (3, 3), (4, 4)], region)
```

命令行 `--spec` 里用 `"kind": "line"` 或 `"kind": "scatter"`。缺省仍是条形图。示意流程不要改成折线。

## 单击出现

只有用户要求“单击后出现”或“逐条进入”时才用 `scripts/appear.py`。

```python
add_click_appear(slide, [shape_a, shape_b])  # 每次单击一行
add_click_appear(slide, shapes, together=True)  # 一次单击整组出现
add_click_appear(slide, shapes, effect="fade", duration_ms=600)
```

`appear` 是瞬间出现。`fade` 只能是 500–700 毫秒。函数不把对象打组，因此不把公式栅格化。已经有 `p:timing` 的页会拒绝覆盖。不调用这个函数，页面就没有进入动画。

这不是换页切换，也不是 ManimGX。推导动画和三维直觉仍按 [运动讲解](motion.md) 嵌 MP4。局部修改里“整图单击进入”的口径见 [局部编辑与动画](scoped-edits.md)。

## 不要做的事

- 不要用文本框拼成表格，也不要截一张表的图来代替原生表。
- 不要为了填下半部而缩小字号，或贴上一段讲解。一行结论或一句读图说明可以。
- 不要把样张里的整数写进用户的结果页。
- 不要默认打开单击动画，也不要用它替换 OMML 或原生形状。
