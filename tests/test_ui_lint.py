"""界面机械 lint：文案零 emoji / 零感叹号 / 零长破折号 / 零箭头 / 零字符图形当图标，
间距与行高全部来自令牌表，私用区字形只准出现在 ui_icons.py。

不依赖观感：直接扫描 app.py / ui_theme.py / ui_icons.py 的字符串字面量与源码模式。
"""

import ast
import re
import unittest
from pathlib import Path

import ui_icons
import ui_theme

ROOT = Path(__file__).resolve().parents[1]
UI_FILES = ("app.py", "settings_model.py", "ui_theme.py")

# em-dash U+2014、en-dash U+2013、箭头 U+2192、半角/全角感叹号
BANNED_CHARS = "\u2014\u2013\u2192!\uff01"
# 字符图形当图标的禁区：几何图形（▸▾▲▼■□●…）、箭头区、杂项符号/emoji 区、变体选择符
ART_RANGES = (
    (0x25A0, 0x25FF),
    (0x2190, 0x21FF),
    (0x2600, 0x27BF),
    (0x2B00, 0x2BFF),
    (0x1F000, 0x1FAFF),
    (0xFE0F, 0xFE0F),
)


def string_literals(path):
    tree = ast.parse(path.read_text(encoding="utf-8"))
    for node in ast.walk(tree):
        if isinstance(node, ast.Constant) and isinstance(node.value, str):
            yield node.value


def art_problems(text):
    return [hex(ord(ch)) for ch in text if any(lo <= ord(ch) <= hi for lo, hi in ART_RANGES)]


class CopyLintTests(unittest.TestCase):
    def test_ui_copy_has_no_banned_marks(self):
        problems = []
        for name in UI_FILES:
            for text in string_literals(ROOT / name):
                if not re.search(r"[\u4e00-\u9fff]", text):
                    continue  # 只检查含中文的用户可见文案，跳过标识符/扩展名
                for ch in BANNED_CHARS:
                    if ch in text:
                        problems.append((name, ch, text[:40]))
        self.assertEqual(problems, [])

    def test_ui_copy_has_no_character_art_as_icon(self):
        """几何/箭头/emoji 区字符一律不许进 UI 文案（字符画图标禁令）。"""
        problems = []
        for name in UI_FILES:
            for text in string_literals(ROOT / name):
                if not re.search(r"[\u4e00-\u9fff]", text):
                    continue
                hits = art_problems(text)
                if hits:
                    problems.append((name, hits, text[:40]))
        self.assertEqual(problems, [])

    def test_private_use_glyphs_only_in_icon_module(self):
        """图标字形（私用区 E000-F8FF）只准出现在 ui_icons.py，别处不得硬编码。"""
        problems = []
        for name in UI_FILES:
            for text in string_literals(ROOT / name):
                hits = [hex(ord(ch)) for ch in text if ui_icons.PUA_RANGE[0] <= ord(ch) <= ui_icons.PUA_RANGE[1]]
                if hits:
                    problems.append((name, hits, text[:40]))
        self.assertEqual(problems, [])


class SourcePatternLintTests(unittest.TestCase):
    def test_no_global_wheel_hijack(self):
        source = (ROOT / "app.py").read_text(encoding="utf-8")
        self.assertNotIn(".bind_all(", source, "滚轮不得全局接管，应逐控件局部绑定")

    def test_no_dashed_canvas_art(self):
        source = (ROOT / "app.py").read_text(encoding="utf-8")
        self.assertNotIn("dash=", source, "拖放区不得用虚线矩形字符画，用图标字体")


class SpacingLintTests(unittest.TestCase):
    def test_spacing_values_come_from_token_table(self):
        allowed = set(ui_theme.SPACING.values()) | {0}
        pattern = re.compile(r"\b(?:padx|pady|padding|spacing1|spacing3)\s*=\s*(?:\(([^)]*)\)|([-0-9]+))")
        problems = []
        for name in UI_FILES:
            source = (ROOT / name).read_text(encoding="utf-8")
            for match in pattern.finditer(source):
                value = match.group(1) if match.group(1) is not None else match.group(2)
                for number in re.findall(r"\d+", value):
                    if int(number) not in allowed:
                        problems.append((name, match.group(0)))
        self.assertEqual(problems, [])

    def test_row_heights_and_wraps_come_from_token_table(self):
        """控件 height / wraplength 关键字实参必须来自令牌表，禁止裸数字。"""
        allowed = set(ui_theme.MEASUREMENTS.values()) | set(ui_theme.SPACING.values()) | {0, 1, 2}
        problems = []
        for name in UI_FILES:
            tree = ast.parse((ROOT / name).read_text(encoding="utf-8"))
            for node in ast.walk(tree):
                if not isinstance(node, ast.Call):
                    continue
                for keyword in node.keywords:
                    if keyword.arg not in ("height", "wraplength"):
                        continue
                    value = keyword.value
                    if isinstance(value, ast.Constant) and isinstance(value.value, int) and value.value not in allowed:
                        problems.append((name, keyword.arg, value.value))
        self.assertEqual(problems, [])

    def test_spacing_tokens_follow_four_px_grid(self):
        for key, value in ui_theme.SPACING.items():
            self.assertEqual(value % 4, 0, f"间距令牌 {key}={value} 不在 4px 网格")

    def test_measurement_tokens_follow_four_px_grid(self):
        for key, value in ui_theme.MEASUREMENTS.items():
            self.assertEqual(value % 4, 0, f"结构令牌 {key}={value} 不在 4px 网格")


if __name__ == "__main__":
    unittest.main()
