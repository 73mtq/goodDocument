import json
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

from formatter import (
    Formatter,
    generate,
    resolve_align,
    resolve_page_size,
    resolve_size,
    resolve_table_align,
    set_zero_indent,
)

ROOT = Path(__file__).resolve().parents[1]


def load_config():
    return json.loads((ROOT / "config.json").read_text(encoding="utf-8"))


class ResolveHelpersTests(unittest.TestCase):
    def test_resolve_size_accepts_chinese_name_and_number(self):
        self.assertEqual(resolve_size("小四"), 12.0)
        self.assertEqual(resolve_size("三号"), 16.0)
        self.assertEqual(resolve_size(14), 14.0)
        self.assertEqual(resolve_size(None), 12.0)

    def test_resolve_page_size_defaults_to_a4_and_supports_letter(self):
        self.assertEqual(resolve_page_size({}), (21.0, 29.7))
        self.assertAlmostEqual(resolve_page_size({"page": {"size": "Letter"}})[0], 21.59, places=2)

    def test_resolve_table_align_maps_values(self):
        self.assertEqual(resolve_table_align("left"), WD_TABLE_ALIGNMENT.LEFT)
        self.assertEqual(resolve_table_align("right"), WD_TABLE_ALIGNMENT.RIGHT)
        self.assertEqual(resolve_table_align(None), WD_TABLE_ALIGNMENT.CENTER)


class FigureSizeTests(unittest.TestCase):
    """生成端的图片尺寸策略（figure.sizeMode）。"""

    def _formatter(self, **overrides):
        cfg = load_config()
        cfg["figure"].update(overrides)
        return Formatter(cfg)

    def test_fit_mode_caps_width_and_height(self):
        width, height = self._formatter()._figure_size(str(ROOT / "assets" / "figure1.png"))
        self.assertAlmostEqual(width, 15.5, delta=0.01)
        # figure1.png 为 3:2，宽度受限后高度应为 2/3 宽
        self.assertAlmostEqual(height, 15.5 * 2 / 3, delta=0.01)

    def test_fixed_mode_forced_size_ignores_ratio(self):
        width, height = self._formatter(
            sizeMode="fixed", fixedWidthCm=12, fixedHeightCm=9, keepAspectRatio=False,
        )._figure_size(str(ROOT / "assets" / "figure1.png"))
        self.assertAlmostEqual(width, 12, delta=0.01)
        self.assertAlmostEqual(height, 9, delta=0.01)

    def test_fixed_mode_uniform_width_keeps_ratio(self):
        width, height = self._formatter(
            sizeMode="fixed", fixedWidthCm=10, fixedHeightCm=None, keepAspectRatio=True,
        )._figure_size(str(ROOT / "assets" / "figure1.png"))
        self.assertAlmostEqual(width, 10, delta=0.01)
        self.assertAlmostEqual(height, 10 * 2 / 3, delta=0.01)

    def test_fixed_mode_width_clamped_to_text_width(self):
        width, _ = self._formatter(
            sizeMode="fixed", fixedWidthCm=20, fixedHeightCm=None, keepAspectRatio=True,
        )._figure_size(str(ROOT / "assets" / "figure1.png"))
        self.assertAlmostEqual(width, 15.5, delta=0.01)


class GenerateTests(unittest.TestCase):
    def test_generate_sample_document_applies_config(self):
        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "sample.docx"
            generate(cfg, out, assets_dir=ROOT / "assets")
            doc = Document(out)
            self.assertEqual(len(doc.tables), 1)
            self.assertEqual(len(doc.inline_shapes), 1)
            section = doc.sections[0]
            self.assertAlmostEqual(section.left_margin.cm, cfg["page"]["margins"]["left"], places=1)
            body_para = next(p for p in doc.paragraphs if p.text.startswith("随着信息技术"))
            self.assertEqual(body_para.paragraph_format.first_line_indent, Pt(24))
            self.assertEqual(body_para.alignment, resolve_align(cfg["body"]["align"]))

    def test_generate_respects_letter_page_size(self):
        cfg = load_config()
        cfg["page"]["size"] = "Letter"
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "letter.docx"
            generate(cfg, out, assets_dir=ROOT / "assets")
            doc = Document(out)
            self.assertAlmostEqual(doc.sections[0].page_width.cm, 21.59, places=1)

    def test_generate_respects_table_align(self):
        cfg = load_config()
        cfg["table"]["align"] = "left"
        with tempfile.TemporaryDirectory() as tmp:
            out = Path(tmp) / "left_table.docx"
            generate(cfg, out, assets_dir=ROOT / "assets")
            doc = Document(out)
            self.assertEqual(doc.tables[0].alignment, WD_TABLE_ALIGNMENT.LEFT)


class SetZeroIndentTests(unittest.TestCase):
    def test_set_zero_indent_clears_char_units(self):
        cfg = load_config()
        F = Formatter(cfg)
        para = F.body_para("文字")
        pPr = para._element.get_or_add_pPr()
        ind = pPr.find(qn("w:ind"))
        if ind is None:
            ind = OxmlElement("w:ind")
            pPr.append(ind)
        ind.set(qn("w:firstLine"), "480")
        ind.set(qn("w:firstLineChars"), "200")
        ind.set(qn("w:left"), "0")
        ind.set(qn("w:leftChars"), "200")
        ind.set(qn("w:right"), "0")

        set_zero_indent(para)
        ind = pPr.find(qn("w:ind"))
        self.assertIsNone(ind.get(qn("w:firstLineChars")))
        self.assertIsNone(ind.get(qn("w:leftChars")))
        self.assertEqual(ind.get(qn("w:firstLine")), "0")
        self.assertEqual(ind.get(qn("w:left")), "0")
        self.assertEqual(ind.get(qn("w:right")), "0")


if __name__ == "__main__":
    unittest.main()
