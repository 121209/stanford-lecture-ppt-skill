# 迭代闭环

做课件时以 [SKILL.md](../SKILL.md) 为准。本文件规定怎样把使用中的问题收进来，并改回仓库。

## 本机迭代与 GitHub 同步

维护者明确要求将本机迭代同步回 GitHub 时，本地文件更新只是中间状态。先定位仓库与最新分支，基于远端差异同步可公开的规则，保留公开版身份及本机私有覆盖；不得上传私人课件、批注、截图、账户信息或本机路径。基础校验和相关检查通过后，按授权直接提交，或使用 PR；不把创建了本地补丁当作远端同步完成。

发布后核验远端提交及分支，向用户返回提交或 PR 链接。只创建 PR 尚未合并时明确说明；受权限或检查阻塞时说明未完成的具体步骤。随后将已验证的远端提交记录到本机同步元数据，保留私有覆盖。未授权的本机试验不自动发布。

## 闭环

1. **使用。** 用本 skill 从材料做出或改好一版课件。
2. **记录。** 结果偏离预期时提交[使用反馈](https://github.com/121209/stanford-lecture-ppt-skill/issues/new?template=skill-feedback.yml)。已经知道要改哪条规则时提交[改进提议](https://github.com/121209/stanford-lecture-ppt-skill/issues/new?template=skill-improvement.yml)。
3. **分流。** 用下面的标签整理 issue，缺复现材料的先标 `needs-repro`。
4. **修改。** 一个改进对应一个 PR，只动这一件事。
5. **验证。** 改了脚本就跑单元测试；涉及质检或示例课件时，对样例跑 `audit_pptx.py`。
6. **合并。** 合入 `main` 后，若制作行为有变化，在 `CHANGELOG.md` 写一节。

合并只更新 GitHub 仓库，不会主动更新使用者已经安装的本地副本。维护者与使用者都应按 [README 的更新说明](../README.md#更新已安装的-skill) 同步自己的安装目录；保留私有素材与本地改动。

仓库不开放空白 issue，请用这两个表单。模板文件：

- [`.github/ISSUE_TEMPLATE/skill-feedback.yml`](../.github/ISSUE_TEMPLATE/skill-feedback.yml)
- [`.github/ISSUE_TEMPLATE/skill-improvement.yml`](../.github/ISSUE_TEMPLATE/skill-improvement.yml)
- [`.github/PULL_REQUEST_TEMPLATE.md`](../.github/PULL_REQUEST_TEMPLATE.md)

## 两种 issue

### 使用反馈

生成或修改课件之后填写。字段：

| 字段 | 写什么 |
| --- | --- |
| 当时的提问 / 主题 | 交给 skill 的要求：主题、页数、时长、新稿还是修改 |
| 出问题或不顺手的步骤 | 下拉项 1–7 对应 `SKILL.md`「制作流程」 |
| 预期 / 实际 | 这一步应按文档发生什么，实际发生了什么 |
| 附件 / 截图 | 把页面截图粘贴进文本框（GitHub 会上传图片），或留下可公开文件的链接 |
| 严重程度 | 低：能用但别扭；中：要返工；高：流程中断或不能拿去汇报 |

「制作流程」七步是：

1. 核对材料、任务类型、受众、时长、页数及未知状态。
2. 大纲包含逐页目的、主要信息、素材映射、`visual_plan`、个人收获与备注重点；尚有影响范围或风格的待选项时才合并确认。
3. 风格已有选择。检查可用的原生制作工具；示意图默认 python-pptx 形状。只有用户明确接受整页图片时才启用生图并固定后端。
4. 先做一页能展示正文、原生示意图或公式的内容样张并真实检查。需要用户选择方案时先确认；已获授权且风格确定时内部检查后继续。
5. `scripts/run_state.py` 保存门、方法、逐页任务、原图 hash、产物及检查记录。
6. 制作、渲染、逐页检查、修复、备注写回，再次导出。
7. 结构审计与状态验证，并人工核对内容、编辑性和视觉。

以后若改了这七步的表述，同一 PR 里同步更新反馈表单的下拉选项。

只提交可以公开的内容。私人论文、未发表图、密钥和本机路径先涂掉。

### 改进提议

从表单的五项里选一项：

| 范围 | 分流时追加的标签 |
| --- | --- |
| workflow｜制作流程 | `improve-workflow` |
| style rules｜风格规则 | 保持 `improve-skill` |
| QA & audit scripts｜质检与审计脚本 | `improve-qa` |
| demo & examples｜示例与 demo | `improve-demo` |
| docs｜文档 | 保持 `improve-skill` |

正文写三件事：现在的问题、提议改成什么样（点名 `SKILL.md` 的小节或具体文件）、怎样算改好。

## 标签

表单会带上 `feedback` 或 `improve-skill`。其余标签由维护者在分流时添加。

| 标签 | 何时使用 |
| --- | --- |
| `feedback` | 使用之后的原始反馈，尚未决定改哪一条规则 |
| `improve-skill` | 已形成改进提议，或反馈已被接受、准备改 skill |
| `improve-workflow` | 改「制作流程」、确认门、分工，或 [references/workflow.md](../references/workflow.md) |
| `improve-qa` | 改 [references/qa.md](../references/qa.md)、`scripts/audit_pptx.py` 或 `scripts/run_state.py` 的检查 |
| `improve-demo` | 改 `assets/demo.pptx`、`assets/previews/` 或示例所表达的版式 |
| `needs-repro` | 缺少提问、步骤、预期/实际或截图，还不能判断。补齐材料后去掉这个标签 |

一条反馈可以同时带着 `feedback` 和某个 `improve-*`。信息够了再去掉 `needs-repro`。

这些标签需要仓库里已经存在，表单上的默认标签才会贴上。创建命令见文末。

## 一个改进一个 PR

PR 使用 [检查清单](../.github/PULL_REQUEST_TEMPLATE.md)：

- 关联的 issue
- 改了 `SKILL.md` 的哪一节，或哪一份 reference、脚本、文档、示例
- 同一规则、同一页或同一命令的改前与改后
- 改了脚本时，`python -m unittest discover tests` 已通过
- 与成品质检或示例课件有关时，对样例跑过 `audit_pptx.py`

风格微调和审计脚本修复分成两个 PR，除非 issue 说明它们是同一处缺陷。

## 验证

改了 `scripts/` 或 `tests/`：

```bash
python -m unittest discover tests
```

改了质检标准、审计脚本或示例课件时，至少对匿名示例跑一次：

```bash
python scripts/audit_pptx.py assets/demo.pptx
```

有更贴近这次改动的样例时，对那一份再跑。结构审计看包内对象、备注和越界线索；页面好不好看、科学内容是否成立，仍按 [references/qa.md](../references/qa.md) 人工核对。

## 合并并写入变更记录

合入 `main` 之后，若这次改变了制作行为（`SKILL.md`、`references/`、脚本行为，或示例所示范的规则），在仓库根目录 `CHANGELOG.md` 增加一节：

```markdown
## YYYY-MM-DD

- #123：改了 `SKILL.md`「制作流程」第 4 步。改前：豁免样张确认后直接铺开。改后：豁免确认时仍把内部样张检查写入逐页 QA。
```

只改 issue 模板或本说明、制作行为保持不变时，不必写行为变更。

## 可选：Darwin 式打分

改进落在 `SKILL.md` 的措辞、步骤或边界上时，可以加一轮打分，确认文字变具体了。打分对象是 `SKILL.md`。某一维要靠它指向的 reference 才能变清楚时，允许同一轮只改那一份 reference，复评时仍看 `SKILL.md` 是否把人带到这条规则。

九个维度，每维 1–5 分。1 分表示几乎没写或会误导；3 分表示能照着做，但仍要猜；5 分表示换一个代理也能做对，并且写明了禁止事项。

| 维度 | 在 SKILL.md 里看什么 |
| --- | --- |
| Frontmatter | `name` / `description` 是否覆盖该触发的场合：斯坦福讲义风、CS329/CS329Z、组会、论文阅读、科研学习汇报；仅整理语料或仅审阅时不启动 PPT 生产 |
| Workflow | 「制作流程」七步是否按顺序写清，并能对应到 `references/workflow.md` |
| Failure handling | 缺证据、缺原生工具、公式无法做成可编辑对象时，是否停在受影响页并说明限制；用户未接受整页图片时保持可编辑路径 |
| Checkpoints | 大纲与样张只在仍需用户决策时合并确认；已授权任务直接执行内部质量检查；局部修改不强制整稿门流程 |
| Specificity | 白底、黑灰正文、红色强调、标题 30–32pt、正文 20–22pt、一页一个实质问题；文字与公式默认可编辑，复杂证据图保留原图 |
| Resources | 按需指向 style、workflow、content、qa、provenance；写明脚本不生成 PPT，`assets/demo.pptx` 只供版式参考且不是科研证据 |
| Architecture | 主代理独占最终 PPTX 和全局状态；能并行时一人一页，不能并行时串行，并继续留下逐页记录 |
| Tests | `scripts/` 的行为有 `tests/` 覆盖；改脚本时要跑 `python -m unittest discover tests` |
| Blacklist | 不外发私人材料，不读取密钥，不使用未授权的付费 API，不编造阅读日期、复现经历和实验结果，不把铺满的截图当成可编辑页，不把 codex-ppt 的整页图片组装器用于可编辑分支 |

每一轮只改一个维度对应的文字。

**成对复评。** 把改前和改后的 `SKILL.md` 放在一起，用同一张表、同一套判据，九维都重新打分。比较的是这一对的前后分数。目标维度至少高 1 分，另外八维没有下降，才保留这次 diff。任一其他维度下降，就放弃这次修改。

可贴进 PR 的记分表：

```text
维度                改前  改后
Frontmatter
Workflow
Failure handling
Checkpoints
Specificity
Resources
Architecture
Tests
Blacklist
本轮维度：
保留：是 / 否
```

**合并前的人工检查点。** 分数用来决定这版文字值不值得提交。合并的人还要看三件事：diff 是否只动了所声称的那一个维度；前后分数是否和文字对得上；Frontmatter 的触发范围、不编造、默认可编辑、不外发私人材料、不私自更换生图路径，这些边界是否还在。检查点没通过就不合并。

## 维护者：创建标签

标签要先在仓库里存在。在本机有 `issues: write` 权限时执行：

```bash
gh label create "feedback" --description "生成课件后的使用反馈：哪一步偏离预期" --color "0075CA"
gh label create "improve-skill" --description "对 skill 的改进提议：问题、改法与验证" --color "0E8A16"
gh label create "improve-workflow" --description "制作流程、确认门或分工方式的改进" --color "6F42C1"
gh label create "improve-qa" --description "验收标准、审计或状态检查的改进" --color "C4520A"
gh label create "improve-demo" --description "匿名示例课件、预览或示例文案的改进" --color "1B7C83"
gh label create "needs-repro" --description "缺少可复现的提问、材料或截图" --color "D93F0B"
```

颜色是标签背景色，不含 `#`。
