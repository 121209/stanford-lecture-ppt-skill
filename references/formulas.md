# 公式

屏幕上的公式要能被听众读懂，并且不能露出 LaTeX 源码。`python-pptx` 在 Linux 上写不好复杂公式，所以按下面的顺序做。验收口径见 [验收](qa.md)。

## 选用顺序

| 顺序 | 做法 | LibreOffice 里怎么算通过 |
| --- | --- | --- |
| 1 | OMML。占位符 `{{MATH:id}}` 独占一个文本框里的一个 run，再用 `scripts/inject_omml.py` 写成 PowerPoint 原生结构：`mc:AlternateContent` → `mc:Choice Requires="a14"` → `a14:m` → `m:oMathPara` → `m:oMath`。`mc:Fallback` 是同一形状区域里的高 DPI 公式图 | LibreOffice 应显示这张回退图，而不是空白。裸 `m:oMath` 直接放进 `a:p` 通不过架构校验。包里不得残留 `{{MATH:` |
| 2 | 高 DPI 数学图。LaTeX 或 mathtext 导出 PNG 或 SVG，至少 300 DPI，最好 300–600 DPI。图旁用可编辑文字写符号说明 | 这是组会在 Linux 上默认可核验的路径。预览里公式模糊，算失败 |
| 3 | Cambria Math 或分段的可编辑数学文本。上下标用字体自身的基线，不写 `x_i` 这种 ASCII | 只用于单个符号或很短的式子。分式、求和、积分升到 1 或 2 |

LaTeX 不必再等 pandoc。`scripts/vendor/pptmaster_formula/` 是一份 MIT 编译器（ppt-master，提交 `4a68cdd`），标准库就能把式子编成结构化 OMML。求和用块级编译，上下限是 `undOvr`；外层仍由 `inject_omml.py` 包进原来的 `m:oMathPara` 和图片回退。不认识的命令直接报错，源码不会落到幻灯片上。pandoc 只在编译器抛错、并且本机真有 pandoc 时才再试一次。

有 `latex` 时，回退图用这一条字符串，matplotlib mathtext 字体是 stixsans。`\displaystyle` 这类 mathtext 画不出、去掉后符号不变的命令会先去掉；剩下仍画不出时，注入失败。图按自身宽高比放进占位框。

## 禁止

- 屏幕文字（含公式对象里的可见文本）出现原始 LaTeX：`\frac`、`\sum`、`\lambda`、`^{`、`_{`，以及其他带反斜杠的命令。`scripts/audit_pptx.py` 用 `raw_latex_on_screen` 记警告。
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

```text
python scripts/inject_omml.py deck.pptx formulas.json -o deck-math.pptx
```

脚本不覆盖输入文件。输出里若仍有 `{{MATH:`，命令失败。需要本机有 matplotlib，用来画回退图。没有 matplotlib 时命令失败，不会把 LaTeX 源码写进幻灯片。

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
