# 运动讲解（可选）

只有用户明确要求推导动画、公式变形、几何运动或三维直觉时才走这一页。默认成品仍是可编辑 PPTX。公式继续用 [公式](formulas.md) 的 OMML 和 `scripts/tex_fallback.py`。算法、架构和多步流程继续用 [示意图](diagrams.md) 的原生形状。没有提出要动画时，不要调用 ManimGX，也不要为了好看把公式页改成视频。

ManimGX 是独立的动画引擎（[academa-labs/manimgx](https://github.com/academa-labs/manimgx)）。场景用 Manim Community 风格的 Python API，Rust + wgpu 画出画面。下面的命令和接口以它自己的文档为准：

- 安装与命令：[llms.txt](https://manimgx.academa.ai/llms.txt)
- 命令行：[Command line](https://manimgx.academa.ai/reference/rendering/command-line/)
- 从 Python 渲染：[From Python](https://manimgx.academa.ai/reference/rendering/from-python/)
- 背景和分辨率：[Configuration](https://manimgx.academa.ai/reference/rendering/configuration/)
- 公式：[Text and math](https://manimgx.academa.ai/user-guide/text-and-math/)
- 颜色：[Colors](https://manimgx.academa.ai/reference/mobjects/colors/)

## 它实际写出什么

| 命令 | 产物 | 不包含 |
| --- | --- | --- |
| `manimgx render` | MP4，以及旁边一份分镜说明 | PPTX、单独的 PNG/SVG 导出开关 |
| `manimgx inspect` | 分镜 PNG（`.storyboard.png`）、时间线和二维布局检查 | 视频 |
| `manimgx present` | 浏览器里按 `next_section` 播放的页面；`--pdf` 才另写一份讲义 PDF | 组会要交的 PPTX |
| `Scene.render(video=..., frames=...)` | 同样的 MP4；`frame.pixels()` 返回这一帧的 RGBA 字节 | 不是命令行上的 `--png` |

不要发明 `--png`、`--svg`、`-ql` 或透明底导出参数。命令行文档里的开关是 `-o`、`-r`、`--fps`（render），以及 inspect 的 `-t`、`-o`。`manimgx present` 的页面不是本技能的交付物。

分镜图带有播放序号、时间、源码位置，布局问题还会被框出来。它是检查图，不是干净的最后一帧。

## 什么时候用

用：用户要看一个式子怎么变成另一个式子、一个几何关系怎么动起来，或一个三维物体。画面放进 PPTX 的一页，周围的标题、图注和备注仍是可编辑文字。

不用：

- 屏幕上的公式要能改。那是 OMML，影片里的字形改不了。
- 算法、架构、流程。那是原生形状。
- 用户没要动画。探测到 ManimGX 也不等于要启用它。
- 用它代替论文原图，或用它的画面冒充实验数据。

同一份课件里，需要能改的式子仍在这一页或相邻页用 `scripts/inject_omml.py` 写入。本路径不生成 OMML。

大纲里这一页可以写 `visual_plan.kind: motion`、`engine: manimgx`。这是额外的一种，不替换 `equation` 或 `native_diagram`。

## 环境

ManimGX 需要 Python 3.13 或 3.14。文档给出的安装是：

```text
uv tool install --python 3.14 manimgx
manimgx --version
```

没有 uv 时，文档写的是 `uvx --python 3.14 manimgx`。不要 `import manim`，也不要为它再装 LaTeX 或 ffmpeg。公式由它自带的 Typst 排，`MathTex` 不接受任意 LaTeX 宏包和 `\def`。

Linux 上没有可用 GPU 时，轮子里的 lavapipe 用 CPU 画。报 `no GPU adapter` 是运行环境把显卡藏起来了，不是缺一个 PPTX 导出参数。

先探测，不要自动安装：

```text
python scripts/embed_motion.py --check
```

退出码是 0。JSON 里 `available` 为 false 就走下面的降级，不要在这一步装包。只有用户明确同意安装时才运行上面的安装命令。

## 场景

背景默认是黑的。讲义页改成白，并且在创建场景之前设置：

```python
import manimgx as m

m.config.background_color = m.WHITE


class Rearrange(m.Scene):
    def construct(self) -> None:
        before = m.MathTex("a^2", "+", "b^2", "=", "c^2", font_size=96)
        after = m.MathTex("a^2", "=", "c^2", "-", "b^2", font_size=96)
        before[2].set_color(m.GRAY_E)
        after[4].set_color(m.GRAY_E)
        self.play(m.Write(before))
        self.play(m.TransformMatchingTex(before, after))
        self.wait()
```

`m.GRAY_E` 是 `#222222`，和正文一致。颜色页写明十六进制可以代替颜色名，强调项可以用 `"#C00000"`。`Write` 和 `TransformMatchingTex` 的写法来自它的 Text and math 页。分辨率和帧率放在命令行里，不必在场景里再写一遍。

画面大约 14.2×8 个单位，内容留在 x ∈ [-7, 7]、y ∈ [-4, 4]。

## 先检查，再渲染

```text
python scripts/embed_motion.py --commands scene.py --scene Rearrange --video clip.mp4 --storyboard clip.storyboard.png
```

这只打印参数，不运行 ManimGX。推荐命令是：

```text
manimgx inspect -o clip.storyboard.png scene.py Rearrange
manimgx render -o clip.mp4 -r 1920x1080 --fps 30 scene.py Rearrange
```

文件里只有一个场景时，去掉末尾的类名。不传 `-r` 和 `--fps` 时，ManimGX 自己的默认是 1920×1080、60 fps。上面显式写成 30 fps，是为了组会文件小一些；两个开关都在它的命令行文档里。

`inspect` 发现布局问题会以退出码 1 结束。先改场景再查，通过后再 `render`。两条命令都会把写出的文件路径打出来。分镜路径用 inspect 的 `-o` 指定，不要猜 render 另外把分镜放在哪。

## 嵌进 PPTX

```text
python scripts/embed_motion.py -o deck.pptx --video clip.mp4 --poster clip.storyboard.png \
  --title "移项保持等式" \
  --caption "单击播放。可编辑式子仍以公式页为准。" \
  --notes "影片只演示变形。符号含义和适用条件写在这里，不把 LaTeX 源码留在屏幕上。"
```

已有文稿时把原文件放在前面，并换一个输出路径，脚本不会覆盖输入：

```text
python scripts/embed_motion.py current.pptx -o current-motion.pptx --video clip.mp4 --slide 4 --caption "单击播放。"
```

`--slide` 省略时追加一页。python-pptx 在加入视频时可能换掉这一页已有的时间线；有点击动画的页优先追加新页，不要叠到那一页上。

`--poster` 通常就是 inspect 写出的分镜。分镜是检查图，图注或备注里写明这一点。没有合适的静帧时可以不传 `--poster`：脚本放一张白底占位图，单击后播放 MP4。占位图不是 ManimGX 渲染的画面。只有分镜、还没有 MP4 时，可以只传 `--poster`，这一页就是一张普通图片。

视频和海报都嵌在标题和图注之间，标题 32pt `#C00000`，图注 20pt `#222222`。图注保持一行，避免压到字号硬底以下。变形前后的关键量留在屏幕上。这一页主题还需要的条件也留在屏幕上；20pt 装不下就拆页。推导写在备注。屏幕上仍然不能出现 `\frac`、`^{`、`_{` 这类源码。

从 Python 渲染时，`frame.pixels()` 在 `Scene.render(..., frames=sink)` 里返回 RGBA，而且必须在下一帧送来之前读取。那是 Python API，不是命令行导出。需要干净海报时由调用方自己把字节写成 PNG，再传给 `--poster`。

## 缺了 ManimGX

`--check` 里 `available` 为 false，或场景渲染失败时：

- 这一页改回 OMML 公式，或原生形状，并在备注写明「本机没有 ManimGX（需要 Python 3.13 或 3.14），动画没有渲染」。
- 不要整页改成生图，不要把 LaTeX 源码留在屏幕上，不要让整份课件失败。
- 单元测试不安装、不调用 ManimGX。

## 验收

静态预览只能看到海报或占位图，不能据此说已经看过播放。PowerPoint 里这一页应有视频对象：单击或播放控件能打开 MP4。只核对了 XML 时，不要说已经放映过。LibreOffice 可能只显示海报。

`scripts/audit_pptx.py` 会把视频计成一张图片。默认尺寸低于整页 75%，并且伸进内容带下半部。图注仍是可编辑文字。影片里的公式不算 OMML，`office_math_count` 不会因此增加。
