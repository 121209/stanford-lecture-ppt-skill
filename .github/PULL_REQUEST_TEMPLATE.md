## 关联 issue

- [ ] 已关联反馈或改进 issue（例如 `Closes #123`）

## 改了哪里

- [ ] 已写明本次改动的 `SKILL.md` 小节，以及／或 reference、脚本、文档、示例文件

文件与小节：

## 改前 / 改后

- [ ] 下面有同一规则、同一页或同一命令的改前与改后

改前：

改后：

## 验证

- [ ] 若改了 `scripts/` 或 `tests/`：`python -m unittest discover tests` 已通过（未改则在此注明「未改脚本」）
- [ ] 若与成品质检、审计脚本或示例课件有关：已对样例运行 `python scripts/audit_pptx.py <sample.pptx>`（可用 `assets/demo.pptx`；无关则注明「与审计无关」）

一次 PR 只处理一个改进。改变制作行为时，在 `CHANGELOG.md` 增加一节，格式见 [迭代闭环](https://github.com/121209/stanford-lecture-ppt-skill/blob/main/docs/ITERATION.md)。
