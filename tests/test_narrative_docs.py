"""Paper-structuring rules stay in the docs and do not become a word cap."""
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]


def read(relative):
    return (ROOT / relative).read_text(encoding="utf-8")


class NarrativeDocsTest(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.narrative = read("references/narrative.md")
        cls.skill = read("SKILL.md")
        cls.content = read("references/content.md")
        cls.workflow = read("references/workflow.md")
        cls.density = read("references/density.md")
        cls.provenance = read("references/provenance.md")
        cls.qa = read("references/qa.md")
        cls.changelog = read("CHANGELOG.md")

    def test_spine_beats_and_types(self):
        text = self.narrative
        self.assertIn("科学论证", text)
        for beat in ("为什么重要", "缺口或瓶颈", "作者做了什么", "关键证据", "凭什么相信", "边界"):
            self.assertIn(beat, text)
        self.assertIn("局限和边界不能只写进备注", text)
        for name in (
            "发现／机制",
            "方法／模型／工具",
            "资源／数据集／图谱／基准",
            "临床／人群／干预",
            "材料／化学／物理／工程",
            "综述／观点／评论／元分析",
        ):
            self.assertIn(name, text)
        self.assertIn("领域知情的神经点过程", text)
        for arc in ("问题到证据", "问题到方案", "工作流到验证", "设计到推断", "证据地图"):
            self.assertIn(arc, text)
        self.assertIn("直接做", text)
        self.assertIn("起步角色，按这篇论文删改", text)

    def test_figures_titles_and_evidence(self):
        text = self.narrative
        self.assertIn("Fig. 6 可见历史", text)
        self.assertIn("「约」", text)
        self.assertIn("18pt", text)
        self.assertIn("拆页是为了装下", text)
        self.assertIn("不写个人或实验室的复现", text)
        self.assertIn("8–12", text)
        self.assertIn("起点", text)

    def test_narrative_does_not_add_a_word_cap_or_smaller_type(self):
        text = self.narrative
        self.assertIn("不设每页词数上限", text)
        self.assertIn("不采用", text)
        self.assertIn("12–16pt", text)
        self.assertIn("7–9pt", text)
        self.assertNotIn("每页不超过", text)
        self.assertNotIn("词数上限为", text)
        self.assertIn("不授权重做", text)
        self.assertIn("纯风格迁移", text)

    def test_outline_procedure_is_linked(self):
        self.assertEqual(self.skill.count("references/narrative.md"), 4)
        self.assertIn("论文类型", self.skill)
        self.assertIn("不设固定字数上限", self.skill)
        self.assertIn("硬底 18pt", self.skill)
        self.assertIn("narrative.md", self.content)
        self.assertIn("科学论证", self.content)
        self.assertIn("标“约”", self.content)
        self.assertIn("中心论证", self.workflow)
        self.assertIn("arc_role", self.workflow)
        self.assertIn("不机械逐章搬运", self.workflow)
        self.assertIn("不设词数上限", self.workflow)
        self.assertIn("8–12页只是起点", self.workflow)
        self.assertIn("narrative.md", self.qa)

    def test_density_keeps_word_cap_opt_in_and_rejects_small_type(self):
        self.assertIn("narrative.md", self.density)
        self.assertIn("不设每页词数上限", self.density)
        self.assertIn("12–16pt", self.density)
        self.assertIn("7–9pt", self.density)
        self.assertIn("--word-cap", self.density)
        self.assertIn("叙述结构另记在 [文章梳理](narrative.md)", self.density)
        nature = self.provenance.split("nature-paper2ppt", 1)[1].split("\n", 1)[0]
        self.assertIn("narrative.md", nature)
        self.assertIn("12–16pt", nature)
        self.assertIn("屏幕字数预算", nature)

    def test_changelog_records_the_boundary(self):
        section = self.changelog.split("## 2026-11-12", 1)[1].split("## 2026-11-11", 1)[0]
        self.assertIn("文章梳理", section)
        self.assertIn("不采用", section)
        self.assertIn("12–16pt", section)
        self.assertIn("--word-cap", section)
        self.assertIn("8–12", section)
