"""高级设置面板数据层测试：分组结构、术语清洗、人话取值、默认对照。

面板 IA 依据 design/ia-regroup.md（卡片分类），全部断言不依赖 Tk。
"""
import re
import unittest

import ui_icons
from app import (
    SETTING_GROUPS,
    cfg_get,
    fmt_number,
    humanize_value,
    raw_equal,
    spec_default,
)

BANNED_CHARS = "\u2014\u2013\u2192!\uff01"


def find_spec(key):
    for group in SETTING_GROUPS:
        for spec in group["rows"]:
            if tuple(spec["key"]) == tuple(key):
                return group, spec
    return None, None


class GroupStructureTests(unittest.TestCase):
    def test_eleven_groups_in_expected_order(self):
        titles = [group["title"] for group in SETTING_GROUPS]
        self.assertEqual(titles, [
            "页面", "页码", "正文",
            "章标题（第1章）", "节标题（1.1）", "条标题（1.1.1）", "四级标题（1.1.1.1）",
            "图片", "图题与说明", "表格", "参考文献",
        ])

    def test_total_rows_cover_all_settings(self):
        total = sum(len(group["rows"]) for group in SETTING_GROUPS)
        self.assertEqual(total, 85)
        keys = [".".join(spec["key"]) for group in SETTING_GROUPS for spec in group["rows"]]
        self.assertEqual(len(keys), len(set(keys)), "设置键重复")

    def test_old_panel_sections_are_regrouped_not_dropped(self):
        """旧面板全部字段仍可设置：边距、页码、正文、标题四级、图、表、参考文献。"""
        must_have = [
            ("page", "margins", "top"), ("page", "margins", "bottom"),
            ("page", "margins", "left"), ("page", "margins", "right"),
            ("pageNumber", "enabled"), ("pageNumber", "onlyUpdateExisting"),
            ("pageNumber", "align"), ("pageNumber", "size"),
            ("pageNumber", "font"), ("pageNumber", "asciiFont"),
            ("body", "font"), ("body", "asciiFont"), ("body", "size"),
            ("body", "lineSpacing"), ("body", "firstLineIndentChars"),
            ("body", "align"), ("body", "spaceAfterPt"),
            ("figure", "sizeMode"), ("figure", "maxWidthCm"), ("figure", "maxHeightCm"),
            ("figure", "fixedWidthCm"), ("figure", "fixedHeightCm"),
            ("figure", "keepAspectRatio"), ("figure", "clampToTextWidth"),
            ("figure", "skipInTables"), ("figure", "keepWithCaption"),
            ("figure", "captionPosition"), ("figure", "captionSize"), ("figure", "captionBold"),
            ("figure", "captionAlign"), ("figure", "captionLineSpacing"),
            ("figure", "noteSize"), ("figure", "noteAlign"), ("figure", "prefix"),
            ("table", "captionPosition"), ("table", "captionSize"), ("table", "captionBold"),
            ("table", "captionAlign"), ("table", "captionLineSpacing"),
            ("table", "noteSize"), ("table", "noteAlign"),
            ("table", "cellSize"), ("table", "cellAlign"), ("table", "cellVAlign"),
            ("table", "cellLineSpacing"), ("table", "headerBold"), ("table", "prefix"),
            ("references", "font"), ("references", "asciiFont"), ("references", "size"),
            ("references", "lineSpacing"), ("references", "align"),
            ("references", "hangingIndentPt"),
        ]
        for level in range(1, 5):
            for field in ("font", "asciiFont", "size", "lineSpacing", "bold",
                          "spaceBeforePt", "spaceAfterPt", "align"):
                must_have.append(("headings", "h" + str(level), field))
        for key in must_have:
            self.assertIsNotNone(find_spec(key), f"缺设置键: {'.'.join(key)}")

    def test_sample_doc_group_is_not_a_setting(self):
        for group in SETTING_GROUPS:
            for spec in group["rows"]:
                self.assertNotIn("output", spec["key"], "示例文档不是设置，不得留在面板")

    def test_every_group_icon_exists(self):
        for group in SETTING_GROUPS:
            self.assertIn(group["icon"], ui_icons.ICONS, f"组 {group['title']} 图标缺失")


class TerminologyTests(unittest.TestCase):
    def test_group_titles_and_row_names_have_no_latin_dev_terms(self):
        for group in SETTING_GROUPS:
            self.assertIsNone(re.search(r"[A-Za-z]", group["title"]),
                              f"组名混入英文: {group['title']}")
            for spec in group["rows"]:
                self.assertIsNone(re.search(r"[A-Za-z]", spec["name"]),
                                  f"行名混入英文: {spec['name']}")

    def test_no_banned_marks_in_names(self):
        for group in SETTING_GROUPS:
            texts = [group["title"], group.get("about", "")]
            texts += [spec["name"] for spec in group["rows"]]
            for text in texts:
                for ch in BANNED_CHARS:
                    self.assertNotIn(ch, text)

    def test_units_are_spoken_not_abbreviated(self):
        for group in SETTING_GROUPS:
            for spec in group["rows"]:
                if spec["kind"] == "number":
                    self.assertIn(spec["unit"], ("厘米", "磅", "倍", "字"),
                                  f"{spec['name']} 单位不是人话: {spec['unit']}")


class HumanizeTests(unittest.TestCase):
    def test_number_value_carries_unit(self):
        spec = {"kind": "number", "unit": "厘米", "lo": 0, "hi": 15, "step": 0.5}
        self.assertEqual(humanize_value(spec, 2.54), "2.54 厘米")
        self.assertEqual(humanize_value(spec, 3), "3 厘米")

    def test_size_value_carries_pt_note(self):
        spec = {"kind": "size", "options": ["小四"]}
        self.assertEqual(humanize_value(spec, "小四"), "小四 · 12 磅")
        self.assertEqual(humanize_value(spec, "五号"), "五号 · 10.5 磅")
        self.assertEqual(humanize_value(spec, 12), "12")

    def test_bool_and_enum_values(self):
        self.assertEqual(humanize_value({"kind": "bool"}, True), "开")
        self.assertEqual(humanize_value({"kind": "bool"}, False), "关")
        spec = {"kind": "enum", "options": ["left", "center"]}
        self.assertEqual(humanize_value(spec, "center"), "居中")
        self.assertEqual(humanize_value(spec, "smart"), "智能识别")


class BaselineCompareTests(unittest.TestCase):
    def test_raw_equal_numbers_by_value(self):
        spec = {"kind": "number", "unit": "厘米"}
        self.assertTrue(raw_equal(spec, 2.5, 2.50))
        self.assertTrue(raw_equal(spec, "2.5", 2.5))
        self.assertFalse(raw_equal(spec, 2.5, 3))

    def test_raw_equal_bool_and_enum(self):
        self.assertTrue(raw_equal({"kind": "bool"}, 1, True))
        self.assertTrue(raw_equal({"kind": "enum"}, "center", "center"))
        self.assertFalse(raw_equal({"kind": "enum"}, "center", "left"))

    def test_spec_defaults(self):
        self.assertIs(spec_default({"kind": "bool"}), False)
        self.assertEqual(spec_default({"kind": "number", "lo": 4}), 4)
        self.assertEqual(spec_default({"kind": "enum", "options": ["a", "b"]}), "a")

    def test_fmt_number_integral(self):
        self.assertEqual(fmt_number(3.0), "3")
        self.assertEqual(fmt_number(1.5), "1.5")

    def test_cfg_get_walks_nested_chain(self):
        cfg = {"page": {"margins": {"top": 2.54}}}
        self.assertEqual(cfg_get(cfg, ("page", "margins", "top")), 2.54)
        self.assertIsNone(cfg_get(cfg, ("page", "margins", "missing")))
        self.assertEqual(cfg_get(cfg, ("nope",), "x"), "x")


if __name__ == "__main__":
    unittest.main()
