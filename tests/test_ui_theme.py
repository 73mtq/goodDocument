import unittest

import ui_theme


class ThemeTokenTests(unittest.TestCase):
    def test_required_color_tokens_exist(self):
        required = {
            "bg",
            "surface",
            "surface_alt",
            "surface_hover",
            "surface_pressed",
            "border",
            "border_strong",
            "ink",
            "ink_muted",
            "ink_disabled",
            "accent",
            "accent_hover",
            "accent_pressed",
            "accent_soft",
            "accent_disabled_bg",
            "accent_disabled_fg",
            "success",
            "warning",
            "danger",
        }
        self.assertTrue(required.issubset(ui_theme.COLORS))

    def test_text_contrast_meets_wcag_aa(self):
        """界面里每一组实际用到的文字/背景组合，正文对比度不低于 4.5:1。

        数值来源与量色依据见 design/apple-measurements.md；禁用态组合按惯例不检查。
        """
        pairs = [
            ("ink", "bg"),
            ("ink", "surface"),
            ("ink", "surface_alt"),
            ("ink", "surface_hover"),
            ("ink", "surface_pressed"),
            ("ink_muted", "bg"),
            ("ink_muted", "surface"),
            ("ink_muted", "surface_alt"),
            ("ink", "accent_soft"),
            ("surface", "accent"),
            ("surface", "accent_hover"),
            ("surface", "accent_pressed"),
            ("accent", "surface"),
            ("success", "bg"),
            ("success", "surface"),
            ("warning", "surface"),
            ("danger", "surface"),
            ("danger", "bg"),
        ]
        for foreground, background in pairs:
            ratio = ui_theme.contrast_ratio(ui_theme.COLORS[foreground], ui_theme.COLORS[background])
            self.assertGreaterEqual(round(ratio, 2), 4.5, f"{foreground}/{background} = {ratio:.2f}")


class ValueMappingTests(unittest.TestCase):
    def test_label_mapping_round_trip(self):
        for value, label in ui_theme.VALUE_LABELS.items():
            self.assertEqual(ui_theme.LABEL_VALUES[label], value)

    def test_option_values_translate_align_options(self):
        self.assertEqual(
            ui_theme.option_values(["left", "center", "right", "justify"]),
            ["左", "居中", "右", "两端"],
        )

    def test_display_label_passthrough_unknown_value(self):
        self.assertEqual(ui_theme.display_label("宋体"), "宋体")


class AppModuleTests(unittest.TestCase):
    def test_app_module_exposes_app_class(self):
        import app

        self.assertTrue(callable(app.main))
        self.assertTrue(hasattr(app, "App"))


if __name__ == "__main__":
    unittest.main()
