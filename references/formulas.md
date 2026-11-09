# 公式

屏幕上的公式要能被听众读懂，并且不能露出 LaTeX 源码。`python-pptx` 在 Linux 上写不好复杂公式，所以按下面的顺序做。验收口径见 [验收](qa.md)。

环境要求：内置 LaTeX 编译器使用 Python 3.10+ 的联合类型语法，不能在 Python 3.8 中运行。没有 TeX 时，公式图片回退用 matplotlib；TeX 渲染成功时这张回退图不依赖 matplotlib。完整验证还需 python-pptx 和 lxml。优先在虚拟环境中安装，具体命令见 [README](../README.md#检查工具)。真实 TeX 是可选工具，单元测试不要求安装。

## 数学排版和局部修改

用户说“像 LaTeX 那样好看”通常要的是排版后的公式，不是源码、代码框或整页截图。保持公式的数学含义，不因美化擅自增加等号、求和范围、变量或定义。变量与希腊字母用数学斜体，`tanh`、`softplus` 等函数名用正体；上下标用原生结构，减号、乘号与括号采用数学间距。

同一概念跨页、上下图和备注采用同一符号。用户要求把 g/f 换成 h/w 时，只替换对应的时间核与空间核，不全局改动其他函数或字母。先确认 κ、ξ、μ 的实际含义，不能把触发项与完整条件强度混称。

一句中文里的少量符号用短 LaTeX，放进旁边的小文本框，再经 `scripts/inline_math.py` 收成可编辑 OMML 小框。例如 `d_k`、`\lambda`、`\lambda(t)`、`\mathrm{softmax}`。框里只有 `{{MATH:id}}`，不要把整句中文塞进数学框，也不要把前后汉字和符号写在同一个形状里。整段公式仍用更大的独立 `{{MATH:}}` 框，走 `scripts/inject_omml.py`。中文保持中文字体，数学用 Cambria Math 或可用的同类字体。

屏幕上的数学字母不要写成 ASCII 下标（`x_i`、`d_k` 这种源码）。优先 latex→OMML。只有没有编译器时，才用 Unicode 数学字母顶一下，并且仍然不能把 LaTeX 源码留在屏幕上。

数学字母可能被 Office 保存为数学字母 Unicode（例如 𝑔、𝑓），查找不能只查 ASCII。修改 OMML 后同步更新兼容回退图和备注，避免 PowerPoint 显示新符号而其他软件仍显示旧符号。保留原生公式及有效的回退分支，不用回退图替换原生公式。

下列占位符脚本是独立公式框的一个实现，不是所有制作工具的强制 API。

## 选用顺序

| 顺序 | 做法 | LibreOffice 里怎么算通过 |
| --- | --- | --- |
| 1 | OMML。占位符 `{{MATH:id}}` 独占一个文本框里的一个 run，再用 `scripts/inject_omml.py` 写成 PowerPoint 原生结构：`mc:AlternateContent` → `mc:Choice Requires="a14"` → `a14:m` → `m:oMathPara` → `m:oMath`。`mc:Fallback` 是同一形状区域里的高 DPI 公式图：有 TeX 时用真实 TeX 透明 PNG，否则用 mathtext | LibreOffice 应显示这张回退图，而不是空白。裸 `m:oMath` 直接放进 `a:p` 通不过架构校验。包里不得残留 `{{MATH:` |
| 2 | 高 DPI 数学图。LaTeX 或 mathtext 导出 PNG 或 SVG，至少 300 DPI，最好 300–600 DPI。图旁用可编辑文字写符号说明 | 这是组会在 Linux 上默认可核验的路径。预览里公式模糊，算失败 |
| 3 | Cambria Math 或分段的可编辑数学文本。上下标用字体自身的基线。不要在屏幕上写 `x_i` 这种 ASCII | 只用于单个符号或很短的式子，而且是没有编译器时的权宜。分式、求和、积分升到 1 或 2 |

LaTeX 不必再等 pandoc。`scripts/vendor/pptmaster_formula/` 是一份 MIT 编译器（ppt-master，提交 `4a68cdd`），标准库就能把式子编成结构化 OMML。求和用块级编译，上下限是 `undOvr`；外层仍由 `inject_omml.py` 包进原来的 `m:oMathPara` 和图片回退。不认识的命令直接报错，源码不会落到幻灯片上。pandoc 只在编译器抛错、并且本机真有 pandoc 时才再试一次。这五份编译器文件保持原样，公式策略写在它们外面。

## 回退图的降级阶梯

可编辑 OMML 始终先做。LibreOffice 看见的是 `mc:Fallback` 里的图，不是 Choice 里的公式。缺工具就降到下一档，不因为没装 TeX 而整页失败，也不把 LaTeX 源码写上幻灯片。

1. **可编辑 OMML（始终尝试）。** 解析顺序是 `runs`，然后已有的 `omml`，然后 `latex`。`latex` 交给 vendored 编译器。编译器抛错时，只有本机有 pandoc 才再试一次。pandoc 不是图片格式，也不是主路径。
2. **TeX 透明 PNG（引擎和转换器都在时）。** 规格里有 `latex`，并且同时找得到 TeX 引擎和 PNG 转换器时，回退图走 `scripts/tex_fallback.py`。至少 300 DPI，默认 400。引擎顺序是 `xelatex`、`lualatex`、`pdflatex`。文档类是 `standalone`（`border=2pt`），导言区是 `lmodern`、`amsmath,amssymb,mathtools,bm` 和 `xcolor`，笔画颜色 `#222222`，和正文一致。这套 TeX Live 里 `standalone` 的默认裁切是水平盒子，`\[...\]` 进不了展示数学；加上 `varwidth` 之后又会在 `\textwidth` 处把长式子裁掉。所以展示公式写成 `$\displaystyle ...$`，数学样式与 `\[...\]` 相同，但不会被页宽截断。短符号（`inline: true`）用 `$...$`，并加一个共同的 `\vphantom`，避免裁掉下标后基线错开。短符号按 22pt 的真实宽度排，不塞进过窄的框；`softmax` 这种词会把框加宽，而不是把字压小。展示公式仍按占位框的字号渲染，再按自身比例放进框里。回退图用 `a:stretch/a:fillRect` 铺进这个框，避免 LibreOffice 按图片自身 DPI 把同一张图贴两次。编译参数是 `-interaction=nonstopmode -halt-on-error -file-line-error`，不打开 `-shell-escape`。PNG 转换的顺序是 `pdftocairo -png -transp`（纸色透明），然后 ghostscript 的 `pngalpha`，然后 ImageMagick（`magick` 或 `convert`）的 `-density`、`-trim`、`+repage`。白纸底会去掉，四角透明，按图片自身比例放进占位框。这张图只进 `mc:Fallback`，不替换 `mc:Choice` 里的 OMML。TeX 自己失败时改走第 3 档，并在标准错误打一行说明。
3. **mathtext（没有 TeX，或 TeX 失败）。** 这是降级后的第 6 轮回退，不是和 TeX 同一档。字体 stixsans。`\displaystyle` 这类 mathtext 画不出、去掉后符号不变的命令会先去掉。回退字符串里的 `\frac` 改成 `\dfrac`，分式按显示样式画；送给 OMML 编译器的仍是原来的字符串，结构还是 `m:f`。短符号固定 22pt，画在同一高度的画布上再按笔画裁左右，不按占位框去压。`softmax` 比 0.70 英寸宽时把框加宽，不把单词裁成残字。展示公式先按笔画裁掉多余空白，再按比例放进占位框，不拉伸。图透明，至少 300 DPI。`latex` 规格仍画不出时，注入失败。设 `STANFORD_FORMULA_DISABLE_TEX=1` 可在装了 TeX 的机器上强制走这一档。
4. **编译器失败时的 pandoc。** 只发生在第 1 步抛错之后。图片仍走第 2 或第 3 步。OMML 和图片都做不到时，命令失败，屏幕上不留源码。

同一形状里有多条带 `latex` 的公式时，每条单独渲染再竖向叠成一张 PNG。只要有一条没有 `latex`（只有 `runs` 或 `omml`），这一形状的回退图就用 mathtext，避免用猜出来的 TeX 源冒充原文。

可选旁路，不放进 pptx：`dvisvgm --pdf --no-fonts --exact`，否则 `pdf2svg`，写出 SVG。设置环境变量 `STANFORD_FORMULA_TEX_DIR` 时，注入会在该目录留下 `{id}.tex`。`.tex` 和 SVG 都不是幻灯片里的公式，审计要的回退仍是 PNG。

需要的命令：`xelatex`、`lualatex`、`pdflatex` 至少一个；`pdftocairo`、`gs`、`magick` 或 `convert` 至少一个。宏包：`standalone`、`amsmath`、`amssymb`、`mathtools`、`bm`、`xcolor`。SVG 旁路另需 `dvisvgm` 或 `pdf2svg`。例如 Debian/Ubuntu：`texlive-xetex texlive-latex-extra texlive-science poppler-utils`。没装这些时走 mathtext，测试保持绿色。

## 从 nature 学到的，以及故意不学的

nature-paper2ppt 本身多半是把论文图（含公式面板）裁成图片，没有 OMML 注入。高保真 TeX 资产在 nature-image2ppt 的 `formula_renderer.py`：真 TeX、`standalone`、amsmath/mathtools/bm、PDF 转至少 300 DPI 的 PNG（ImageMagick `-density`、`-trim`、`+repage`）或 SVG（`dvisvgm` / `pdf2svg`），并在资产旁边保留 `.tex`。那里的清单把公式标成 `editable: false`，视觉优先于能不能改。

本仓库把这套渲染拿来提高 LibreOffice 能看见的回退图。主路径仍是可编辑 OMML。故意不学的做法：

- 不把只出图、`editable: false` 当成公式主路径。
- 不在缺 TeX 或转换器时让整页失败；按上面的阶梯降级，并在本文写明 mathtext 是降级档。
- 不用生图模型画公式，也不在这一轮改示意图、图表或素材抠图。
- 不把 SVG 或 `.tex` 塞进幻灯片冒充可编辑公式。页面决策里「素材分离失败就整页卡住」的门，不套到公式上。

## 禁止

- 屏幕文字（含公式对象里的可见文本）出现原始 LaTeX：`\frac`、`\sum`、`\lambda`、`^{`、`_{`，以及其他带反斜杠的命令。`scripts/audit_pptx.py` 用 `raw_latex_on_screen` 记警告。
- 屏幕上用 ASCII 写下标，例如 `x_i`、`d_k`。短符号走上面的小 OMML 框。
- 用 ∑、∏、∈ 堆一长串冒充公式正文。单个符号配一句说明可以，整段假公式不行。
- 整页公式截图糊成一块，或没有来源说明。
- 因为公式难做，就把这一页改成整页生图。

Unicode 的 λ、α 可以出现在符号说明里。反斜杠加命令名不可以。

## formulas.json

`inject_omml.py` 读取一个对象，键是占位符 id（字母、数字、下划线、连字符）。

短式子可以写结构化 `runs`。带分式、求和或根号时，优先只写 `latex`，让编译器和回退图用同一条字符串。`runs` 长这样：

```json
{
  "rate": {
    "runs": [
      {"t": "λ("},
      {"t": "t"},
      {"t": ")="},
      {"t": "α"},
      {"sup": {"base": "e", "sup": "−βt"}}
    ]
  }
}
```

`runs` 里每一项只有一个键：

| 键 | 含义 |
| --- | --- |
| `t` | 一段普通文本 |
| `sub` | `{"base": "x", "sub": "i"}` |
| `sup` | `{"base": "e", "sup": "−βt"}` |
| `frac` | `{"num": "1", "den": "n"}` |

这些字符串里不得含 LaTeX 源码。分式套分式这类复杂式子，改用下面的 `omml`，或直接走高 DPI 图。

已经有 OMML 片段时：

```json
{
  "rate": {
    "omml": "<m:oMath xmlns:m=\"http://schemas.openxmlformats.org/officeDocument/2006/math\"><m:r><m:t>λ</m:t></m:r></m:oMath>"
  }
}
```

`"latex": "\\lambda(t)=\\alpha e^{-\\beta t}"` 会走上面的编译器。解析顺序是 `runs`，然后 `omml`，然后 `latex`。同一条规格里如果又有 `runs` 又有 `latex`，可编辑公式用 `runs`，回退图仍用 `latex`；要两边一致，就只写 `latex`。

幻灯片上的占位符必须独占一个文本 run，并且放在单独的文本框里，例如整段文字就是 `{{MATH:rate}}`。不要和标题或前后汉字写在同一个形状里。回退图会按自己的宽高比放在这个框里，不再拉满变形。

短符号用同一条占位符，但框要小，规格交给 `inline_math.py`。它只接受短 LaTeX：拒绝中文句子，也拒绝整段展示公式。

```text
python scripts/inline_math.py deck.pptx inline.json -o deck-inline.pptx
```

`inline.json` 的值可以是字符串，或 `{"latex": "d_k"}`。`d_k` 会变成下标结构，屏幕上不留下 `d_k` 或 `x_i`。

```text
python scripts/inject_omml.py deck.pptx formulas.json -o deck-math.pptx
```

脚本不覆盖输入文件。输出里若仍有 `{{MATH:`，命令失败。没有 TeX 时需要 matplotlib 画回退图；没有 matplotlib 且 TeX 也不可用时命令失败，不会把 LaTeX 源码写进幻灯片。TeX 渲染成功时，这一张回退图不经过 matplotlib。

## 高 DPI 图

路径 2 的图旁留下可编辑的符号说明，例如「λ：速率」。图注或页底写来源；示意式子写明不是论文结果。

`assets/diagram-samples/figure-first.pptx` 的第三页是这条路径的匿名样张：一张公式图，右边四行符号，下半部一句结论。它不是 OMML 页。OMML 是否注入，用审计里的 `office_math_count` 看；结构不完整时会有 `office_math_missing_fallback`。

## LibreOffice 验收

导出或打开预览后按这个口径，不要把三件事混在一起：

- 路径 1 在 LibreOffice 里整段空白：**失败**。回退图应该出现在原来的公式框里。再用 XML 确认没有残留 `{{MATH:`，`office_math_count` 大于 0，并且没有 `office_math_missing_fallback`。
- 残留 `{{MATH:`，或公式是直接放进 `a:p` 的裸 `m:oMath`：**失败**。
- 公式图模糊，或屏幕上能读到 `\frac`、`\lambda`、`^{`、`_{`：**失败**。
- 审计的 `raw_latex_on_screen`、`algorithm_page_missing_diagram` 和 `office_math_missing_fallback` 只记警告，退出码仍是 0。退出码 0 只说明包能读完，有这几条警告时不能交付。

备注里可以用文字讲推导。备注不是把源码藏起来的地方：屏幕本身不得出现源码。
