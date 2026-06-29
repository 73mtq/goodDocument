import unittest
from pathlib import Path

from paths import make_normalized_output_path


class AppPathTests(unittest.TestCase):
    def test_make_normalized_output_path_uses_same_folder_and_suffix(self):
        source = Path(r"D:\docs\报告.v1.docx")

        output = make_normalized_output_path(source)

        self.assertEqual(output, str(Path(r"D:\docs\报告.v1_规范化.docx")))


if __name__ == "__main__":
    unittest.main()
