import json
import tempfile
import unittest
from pathlib import Path

from config_model import load_config_file, save_config_file, validate_config

ROOT = Path(__file__).resolve().parents[1]


def load_config():
    return json.loads((ROOT / "config.json").read_text(encoding="utf-8"))


class ConfigModelTests(unittest.TestCase):
    def test_load_save_and_validate_config(self):
        cfg = {
            "page": {"margins": {"top": 2.5, "bottom": 2.5, "left": 3, "right": 2.5}},
            "pageNumber": {"enabled": True},
            "body": {"font": "宋体", "asciiFont": "Times New Roman", "size": "小四"},
            "headings": {"h1": {}, "h2": {}, "h3": {}, "h4": {}},
            "figure": {},
            "table": {},
            "references": {},
            "output": {"filename": "demo.docx"},
        }

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "config.json"
            save_config_file(cfg, path)
            loaded = load_config_file(path)

        self.assertEqual(loaded["output"]["filename"], "demo.docx")
        validate_config(loaded)

    def test_validate_config_rejects_missing_required_sections(self):
        with self.assertRaises(ValueError):
            validate_config({"page": {}})

    def test_validate_config_accepts_both_figure_size_modes(self):
        base = load_config()
        for mode in ("fit", "fixed", "FIXED"):
            config = dict(base)
            config["figure"] = dict(base["figure"], sizeMode=mode)
            self.assertTrue(validate_config(config))

    def test_validate_config_rejects_unknown_figure_size_mode(self):
        base = load_config()
        config = dict(base)
        config["figure"] = dict(base["figure"], sizeMode="统一")
        with self.assertRaises(ValueError) as ctx:
            validate_config(config)
        self.assertIn("sizeMode", str(ctx.exception))


if __name__ == "__main__":
    unittest.main()
