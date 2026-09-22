import json
import unittest
from pathlib import Path

import version


class VersionConsistencyTests(unittest.TestCase):
    def test_version_matches_package_json(self):
        root = Path(__file__).resolve().parents[1]
        pkg = json.loads((root / "package.json").read_text(encoding="utf-8"))
        self.assertEqual(version.__version__, pkg["version"])


if __name__ == "__main__":
    unittest.main()
