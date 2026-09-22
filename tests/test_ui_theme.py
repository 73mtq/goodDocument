import unittest

import ui_theme


class ThemeTokenTests(unittest.TestCase):
    def test_required_color_tokens_exist(self):
        required = {
            "bg", "surface", "surface_alt", "border", "border_strong", "ink", "ink_muted",
            "accent", "accent_hover", "accent_pressed", "accent_soft", "success", "warning", "danger",
        }
        self.assertTrue(required.issubset(ui_theme.COLORS))

    def test_text_contrast_meets_wcag_aa(self):
        pairs = [
            ("ink", "bg"),
            ("ink_muted", "bg"),
            ("accent", "bg"),
            ("surface", "accent"),
        ]
        for foreground, background in pairs:
            ratio = ui_theme.contrast_ratio(ui_theme.COLORS[foreground], ui_theme.COLORS[background])
            self.assertGreaterEqual(round(ratio, 2), 4.5,
                                    f"{foreground}/{background} = {ratio:.2f}")


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
