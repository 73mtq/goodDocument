import json
import re
import unittest
from pathlib import Path

import version
from formatter import SIZE_NAME_TO_PT


class VersionConsistencyTests(unittest.TestCase):
    def test_version_matches_package_json(self):
        root = Path(__file__).resolve().parents[1]
        pkg = json.loads((root / "package.json").read_text(encoding="utf-8"))
        self.assertEqual(version.__version__, pkg["version"])

    def test_size_table_matches_formatter_js(self):
        root = Path(__file__).resolve().parents[1]
        text = (root / "formatter.js").read_text(encoding="utf-8")
        block = re.search(r"SIZE_NAME_TO_PT\s*=\s*\{([^}]+)\}", text, re.S)
        self.assertIsNotNone(block, "formatter.js 缺少 SIZE_NAME_TO_PT")
        pairs = re.findall(r'"([^"]+)"\s*:\s*([0-9.]+)', block.group(1))
        js_table = {name: float(value) for name, value in pairs}
        self.assertEqual(js_table, {k: float(v) for k, v in SIZE_NAME_TO_PT.items()})


if __name__ == "__main__":
    unittest.main()
