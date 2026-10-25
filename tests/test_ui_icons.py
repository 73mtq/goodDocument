"""图标系统测试：码点合法性、字体检测顺序、纯文字回退。"""
import unittest

import ui_icons


class IconGlyphTests(unittest.TestCase):
    def test_all_glyphs_in_private_use_area(self):
        for name, char in ui_icons.ICONS.items():
            code = ord(char)
            self.assertTrue(ui_icons.PUA_RANGE[0] <= code <= ui_icons.PUA_RANGE[1],
                            f"{name} 的码点 {hex(code)} 不在私用区")

    def test_glyph_codes_unique(self):
        values = list(ui_icons.ICONS.values())
        self.assertEqual(len(values), len(set(values)), "图标码点重复")

    def test_unknown_glyph_returns_none(self):
        self.assertIsNone(ui_icons.glyph("不存在的图标名"))

    def test_required_ui_icons_exist(self):
        required = {
            "page", "dialpad", "font", "bold", "photo", "caption", "table", "library",
            "document", "save", "search", "settings", "undo", "play", "view", "edit",
            "warning", "list", "folder", "check", "chevron_right", "clear",
        }
        self.assertTrue(required.issubset(ui_icons.ICONS))


class IconFamilyTests(unittest.TestCase):
    def test_prefers_fluent_over_mdl2(self):
        family = ui_icons.select_icon_family(["Segoe MDL2 Assets", "Segoe Fluent Icons"])
        self.assertEqual(family, "Segoe Fluent Icons")

    def test_falls_back_to_mdl2(self):
        family = ui_icons.select_icon_family(["Segoe MDL2 Assets", "宋体"])
        self.assertEqual(family, "Segoe MDL2 Assets")

    def test_text_fallback_when_no_icon_font(self):
        self.assertIsNone(ui_icons.select_icon_family(["宋体", "Microsoft YaHei UI"]))


if __name__ == "__main__":
    unittest.main()
