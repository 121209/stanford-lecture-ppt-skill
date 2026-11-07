# 流程来源与差异

流程思想参考 [codex-ppt-skill](https://github.com/ningzimu/codex-ppt-skill) 中的大纲确认、统一风格、样张审批、严格素材映射、逐页任务、执行状态、质检修复及演讲者备注。

本仓库是面向可编辑学术汇报的独立实现，未复制第三方生图API代码或私人课件。参考的流程在本项目中对应为：

| 来源阶段 | 本项目实现 |
|---|---|
| 材料与大纲 | 阅读材料、事实核对、逐页目的和素材映射 |
| 统一风格 | 白底黑红讲义系统，按内容换版式 |
| 后端选择 | 默认原生可编辑；整页图须用户明确接受 |
| 一个样张 | 内容样张先行；免用户确认不等于免内部检查 |
| 逐页任务 | 统一术语、完整局部上下文、素材角色与hash |
| 分工和结果 | 独立页／片段，主代理统一记录和组装 |
| 质检与讲稿 | 全页渲染、修复、备注写回、结构审计和状态验证 |

原codex-ppt主要用于整页生图组装；本项目不默认采用该实现，不把生图组装器用于原生可编辑分支。未安装codex-ppt也可运行这里的辅助脚本。

大纲和样张只在仍需用户决定内容范围或风格时合并确认；已明确授权的制作直接继续，局部修改使用短流程。已经明确的风格不再强迫多选。无并行能力时串行继续。辅助脚本只能核对记录和结构，不能证明科学正确性、视觉或编辑性通过。

`assets/demo.pptx` 为全新匿名示例；图表数据、内容和备注均为演示用途，不来自任何私人研究成果，也不是第三方论文实验结果。第 7 页是后加的图优先示例，不替换前 6 页的内容。

图优先和公式路径另外参考了下面这些公开技能里的规则。本仓库没有合并它们的代码、商业主题或投稿流程，也没有把 officecli、Go、TikZ 或某个生图接口写成硬依赖。

| 来源 | 取用的规则 |
| --- | --- |
| [Noi1r/powerpoint-skill](https://github.com/Noi1r/powerpoint-skill/blob/main/powerpoint-slides/SKILL.md) | 视觉优先（约十秒内不能只靠文字看懂就要有图；无视觉页目标不超过 30%）；结构用分层图、行为用 Mermaid、数学几何才考虑 TikZ；OMML 优先，禁止用 Unicode 堆假公式。该来源把 LibreOffice 里 OMML 空白当作预期；本仓库的注入脚本改为原生 OMML 加高 DPI 回退图，空白不再通过 |
| [PlutoLei/paperbanana-skill](https://github.com/PlutoLei/paperbanana-skill) | `editable` 失败即停，禁止静默退回整页 image 还声称可编辑 |
| [mermaid2pptx](https://github.com/serendipitynz/mermaid2pptx) 与 [pptx-diagram](https://github.com/thititongumpun/skills) | flowchart / sequence 可以落成可编辑的形状和连接线。组会默认用 python-pptx 手摆同一类节点表 |
| [nature-figure](https://github.com/Yuan1z0825/nature-skills/tree/main/skills/nature-figure) | AI 生图仅在用户明示时使用；产出是概念草稿，不是数据图；架构优先手绘矢量 |
| [nature-paper2ppt](https://github.com/Yuan1z0825/nature-skills/tree/main/skills/nature-paper2ppt) | 材料里有图时，结果页不做纯文字；裁图保留轴、图例、色条和 panel 标号 |

这些来源不是本项目的隶属方。CS329 的白底、红强调和黑灰正文保持不变。
