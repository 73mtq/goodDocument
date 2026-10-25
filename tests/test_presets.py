"""排版预设（config/templates）与新增配置项校验测试。"""
import copy
import json
import tempfile
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]

import config_model  # noqa: E402
from config_model import list_presets, load_preset, validate_config  # noqa: E402


def base_config():
    return json.loads((ROOT / "config.json").read_text(encoding="utf-8"))


class PresetLoadingTests(unittest.TestCase):
    def test_repo_preset_discovered(self):
        presets = list_presets()
        names = [p["name"] for p in presets]
        self.assertIn("成都理工大学学士论文", names)

    def test_preset_validates_and_carries_spec_values(self):
        preset = load_preset("成都理工大学学士论文")
        cfg = preset["config"]
        self.assertTrue(validate_config(cfg))

        # 页面：A4、上下左右 2.54、页眉 1.5、页脚 1.75
        margins = cfg["page"]["margins"]
        self.assertEqual(cfg["page"]["size"], "A4")
        self.assertEqual((margins["top"], margins["bottom"], margins["left"], margins["right"]),
                         (2.54, 2.54, 2.54, 2.54))
        self.assertEqual((margins["header"], margins["footer"]), (1.5, 1.75))

        # 正文：宋体小四 1.5 倍行距 首行缩进 2 字符
        self.assertEqual((cfg["body"]["font"], cfg["body"]["size"],
                          cfg["body"]["lineSpacing"], cfg["body"]["firstLineIndentChars"]),
                         ("宋体", "小四", 1.5, 2))

        # 页码小五居中；页眉宋体小五居中
        self.assertEqual(cfg["pageNumber"]["size"], "小五")
        self.assertEqual(cfg["pageNumber"]["align"], "center")
        self.assertEqual((cfg["header"]["font"], cfg["header"]["size"], cfg["header"]["align"]),
                         ("宋体", "小五", "center"))

        # 标题：章小二黑体居中段前 0.5 行；节小三；条小四，段前后 0.5 行
        h1, h2, h3 = cfg["headings"]["h1"], cfg["headings"]["h2"], cfg["headings"]["h3"]
        self.assertEqual((h1["font"], h1["size"], h1["align"], h1["spaceBeforeLines"]),
                         ("黑体", "小二", "center", 0.5))
        self.assertEqual((h2["font"], h2["size"]), ("黑体", "小三"))
        self.assertEqual((h3["font"], h3["size"], h3["spaceBeforeLines"], h3["spaceAfterLines"]),
                         ("黑体", "小四", 0.5, 0.5))

        # 表：表题上方五号黑体；内容五号宋体；智能对齐
        table = cfg["table"]
        self.assertEqual((table["captionPosition"], table["captionText"], table["captionSize"]),
                         ("above", "黑体", "五号"))
        self.assertEqual((table["cellFont"], table["cellSize"]), ("宋体", "五号"))
        self.assertEqual(table["cellAlign"], "smart")
        self.assertEqual(table["smartAlignThresholdChars"], 6)

        # 图：图题下方五号黑体
        figure = cfg["figure"]
        self.assertEqual((figure["captionPosition"], figure["captionText"], figure["captionSize"]),
                         ("below", "黑体", "五号"))

        # 参考文献：五号宋体 GB/T 7714 悬挂缩进
        self.assertEqual((cfg["references"]["size"], cfg["references"]["font"]),
                         ("五号", "宋体"))

    def test_load_preset_unknown_name_raises(self):
        with self.assertRaises(ValueError):
            load_preset("不存在的预设-xyz-9999")

    def test_invalid_template_is_skipped(self):
        with tempfile.TemporaryDirectory() as tmp:
            bad = Path(tmp) / "坏模板.json"
            bad.write_text(json.dumps({"page": 1}, ensure_ascii=False), encoding="utf-8")
            original = config_model.templates_dir

            def fake_dir():
                return tmp
            config_model.templates_dir = fake_dir
            try:
                self.assertEqual(list_presets(), [])
            finally:
                config_model.templates_dir = original


class NewConfigValidationTests(unittest.TestCase):
    def _with_table(self, key, value):
        cfg = base_config()
        cfg["table"][key] = value
        return cfg

    def test_cell_accepts_smart(self):
        self.assertTrue(validate_config(self._with_table("cellAlign", "smart")))

    def test_cell_align_rejects_unknown(self):
        with self.assertRaises(ValueError):
            validate_config(self._with_table("cellAlign", "diagonal"))

    def test_smart_threshold_validated(self):
        self.assertTrue(validate_config(self._with_table("smartAlignThresholdChars", 6)))
        with self.assertRaises(ValueError):
            validate_config(self._with_table("smartAlignThresholdChars", 999))
        with self.assertRaises(ValueError):
            validate_config(self._with_table("smartAlignThresholdChars", "很多"))

    def test_header_footer_margins_validated(self):
        cfg = base_config()
        cfg["page"]["margins"]["header"] = 1.5
        cfg["page"]["margins"]["footer"] = 1.75
        self.assertTrue(validate_config(cfg))
        cfg["page"]["margins"]["header"] = 99
        with self.assertRaises(ValueError):
            validate_config(cfg)

    def test_heading_lines_validated(self):
        cfg = base_config()
        cfg["headings"]["h1"]["spaceBeforeLines"] = 0.5
        cfg["headings"]["h1"]["spaceAfterLines"] = 0.5
        self.assertTrue(validate_config(cfg))
        cfg["headings"]["h1"]["spaceBeforeLines"] = 10
        with self.assertRaises(ValueError):
            validate_config(cfg)

    def test_header_section_optional_but_validated(self):
        cfg = base_config()
        self.assertNotIn("header", cfg)  # 默认配置无 header 段，不影响校验
        cfg["header"] = {"font": "宋体", "size": "小五", "align": "center"}
        self.assertTrue(validate_config(cfg))
        cfg["header"]["align"] = "sideways"
        with self.assertRaises(ValueError):
            validate_config(cfg)

    def test_default_config_still_validates(self):
        self.assertTrue(validate_config(base_config()))

    def test_validate_does_not_mutate_input(self):
        cfg = base_config()
        snapshot = copy.deepcopy(cfg)
        validate_config(cfg)
        self.assertEqual(cfg, snapshot)


if __name__ == "__main__":
    unittest.main()
