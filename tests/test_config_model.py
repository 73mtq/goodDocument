import tempfile
import unittest
from pathlib import Path

from config_model import load_config_file, save_config_file, validate_config


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


if __name__ == "__main__":
    unittest.main()
