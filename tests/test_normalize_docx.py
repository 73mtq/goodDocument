import json
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.shared import Cm, Pt


ROOT = Path(__file__).resolve().parents[1]


def load_config():
    return json.loads((ROOT / "config.json").read_text(encoding="utf-8"))


class NormalizeDocxTests(unittest.TestCase):
    def test_inspect_docx_reports_structure(self):
        from normalizer import inspect_docx

        source = ROOT / "规范文档示例.docx"

        info = inspect_docx(source)

        self.assertGreater(info.paragraph_count, 0)
        self.assertEqual(info.table_count, 1)
        self.assertEqual(info.image_count, 1)
        self.assertIn(1, info.heading_levels)
        self.assertIn(2, info.heading_levels)
        self.assertTrue(info.has_header_footer)
        self.assertFalse(info.has_complex_objects)

    def test_normalize_docx_preserves_structure_and_applies_core_formatting(self):
        from normalizer import normalize_docx

        cfg = load_config()
        source = ROOT / "规范文档示例.docx"
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "规范文档示例_规范化.docx"

            original = Document(source)
            normalize_docx(cfg, source, output)
            normalized = Document(output)

            self.assertTrue(output.exists())
            self.assertNotEqual(source.resolve(), output.resolve())
            self.assertEqual(len(original.paragraphs), len(normalized.paragraphs))
            self.assertEqual(len(original.tables), len(normalized.tables))
            self.assertEqual(original.paragraphs[0].text, normalized.paragraphs[0].text)

            margins = cfg["page"]["margins"]
            section = normalized.sections[0]
            self.assertAlmostEqual(section.top_margin.cm, margins["top"], places=1)
            self.assertAlmostEqual(section.left_margin.cm, margins["left"], places=1)
            self.assertIn("PAGE", section.footer._element.xml)

            body = cfg["body"]
            body_para = normalized.paragraphs[1]
            self.assertEqual(body_para.alignment, 3)  # WD_ALIGN_PARAGRAPH.JUSTIFY
            self.assertEqual(body_para.paragraph_format.line_spacing, body["lineSpacing"])
            self.assertEqual(body_para.paragraph_format.first_line_indent, Pt(24))

            table_cfg = cfg["table"]
            cell_para = normalized.tables[0].rows[0].cells[0].paragraphs[0]
            self.assertEqual(cell_para.paragraph_format.first_line_indent, Pt(0))
            self.assertEqual(cell_para.paragraph_format.line_spacing, table_cfg["cellLineSpacing"])

    def test_normalize_docx_rejects_invalid_paths_and_overwrite(self):
        from normalizer import normalize_docx

        cfg = load_config()
        source = ROOT / "规范文档示例.docx"

        with self.assertRaises(ValueError):
            normalize_docx(cfg, source, source)

        with tempfile.TemporaryDirectory() as tmp:
            bad_input = Path(tmp) / "not-word.txt"
            bad_input.write_text("plain text", encoding="utf-8")
            output = Path(tmp) / "out.docx"
            with self.assertRaises(ValueError):
                normalize_docx(cfg, bad_input, output)

    def test_normalize_docx_handles_no_heading_no_table_document(self):
        from normalizer import inspect_docx, normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "plain.docx"
            output = Path(tmp) / "plain_规范化.docx"
            doc = Document()
            doc.add_paragraph("这是一份没有标题和表格的普通文档。")
            doc.add_paragraph("第二段正文用于检查段落数量。")
            doc.save(source)

            info = inspect_docx(source)
            normalize_docx(cfg, source, output)
            normalized = Document(output)

            self.assertEqual(info.table_count, 0)
            self.assertEqual(info.heading_levels, ())
            self.assertEqual(len(normalized.paragraphs), 2)
            self.assertEqual(normalized.paragraphs[0].paragraph_format.line_spacing, cfg["body"]["lineSpacing"])

    def test_normalize_docx_handles_multilevel_headings_and_margin_changes(self):
        from normalizer import inspect_docx, normalize_docx

        cfg = load_config()
        cfg["page"]["margins"]["left"] = 2.2
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "headings.docx"
            output = Path(tmp) / "headings_规范化.docx"
            doc = Document()
            doc.add_paragraph("一、一级标题")
            doc.add_paragraph("（一）二级标题")
            doc.add_paragraph("1. 三级标题")
            doc.add_paragraph("（1）四级标题")
            doc.save(source)

            info = inspect_docx(source)
            normalize_docx(cfg, source, output)
            normalized = Document(output)

            self.assertEqual(info.heading_levels, (1, 2, 3, 4))
            self.assertAlmostEqual(normalized.sections[0].left_margin.cm, 2.2, places=1)
            self.assertEqual(normalized.paragraphs[0].style.name, "Heading 1")


if __name__ == "__main__":
    unittest.main()
