import json
import os
import shutil
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.style import WD_STYLE_TYPE
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt

ROOT = Path(__file__).resolve().parents[1]

from tests._helpers import picture_sizes  # noqa: E402

_fixture_dir = None
SAMPLE_DOCX = None


def load_config():
    return json.loads((ROOT / "config.json").read_text(encoding="utf-8"))


def setUpModule():
    """运行时生成样例文档，避免依赖仓库中的二进制 fixture。"""
    global _fixture_dir, SAMPLE_DOCX
    from formatter import generate

    _fixture_dir = tempfile.mkdtemp(prefix="gooddocument-sample-")
    SAMPLE_DOCX = Path(_fixture_dir) / "规范文档示例.docx"
    generate(load_config(), SAMPLE_DOCX, assets_dir=ROOT / "assets")


def tearDownModule():
    global _fixture_dir
    if _fixture_dir and os.path.isdir(_fixture_dir):
        shutil.rmtree(_fixture_dir, ignore_errors=True)
    _fixture_dir = None


class NormalizeDocxTests(unittest.TestCase):
    def test_inspect_docx_reports_structure(self):
        from normalizer import inspect_docx

        source = SAMPLE_DOCX

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
        source = SAMPLE_DOCX
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
        from normalizer import NormalizeError, normalize_docx

        cfg = load_config()
        source = SAMPLE_DOCX

        with self.assertRaises(NormalizeError):
            normalize_docx(cfg, source, source)

        with tempfile.TemporaryDirectory() as tmp:
            bad_input = Path(tmp) / "not-word.txt"
            bad_input.write_text("plain text", encoding="utf-8")
            output = Path(tmp) / "out.docx"
            with self.assertRaises(NormalizeError):
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

    def test_normalize_docx_clears_first_line_chars_indent_in_table(self):
        """回归测试：单元格段落若有 OOXML 字符单位缩进（firstLineChars 等），
        规范化后必须清掉——否则 Word 仍按字符数渲染首行缩进。
        """
        from normalizer import normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "table_with_firstlinechars.docx"
            output = Path(tmp) / "table_with_firstlinechars_规范化.docx"

            # 手工构造：单元格段落带 firstLine=480（24pt）+ firstLineChars=200（2字符）
            doc = Document()
            tbl = doc.add_table(rows=1, cols=1)
            tbl.style = "Table Grid"
            p = tbl.rows[0].cells[0].paragraphs[0]
            p.add_run("带字符单位缩进的单元格文字")
            pPr = p._element.get_or_add_pPr()
            ind = OxmlElement("w:ind")
            ind.set(qn("w:left"), "0")
            ind.set(qn("w:leftChars"), "0")
            ind.set(qn("w:firstLine"), "480")
            ind.set(qn("w:firstLineChars"), "200")
            ind.set(qn("w:right"), "0")
            pPr.append(ind)
            doc.save(source)

            normalize_docx(cfg, source, output)
            normalized = Document(output)
            cell_p = normalized.tables[0].rows[0].cells[0].paragraphs[0]
            pPr = cell_p._element.get_or_add_pPr()
            ind = pPr.find(qn("w:ind"))
            # 三个数值单位必须为 0
            self.assertEqual(ind.get(qn("w:firstLine")), "0")
            self.assertEqual(ind.get(qn("w:left")), "0")
            self.assertEqual(ind.get(qn("w:right")), "0")
            # 三个字符单位属性必须被清掉（OOXML 中字符单位优先于数值单位，残留会导致 bug）
            self.assertIsNone(ind.get(qn("w:firstLineChars")), "firstLineChars 残留会导致 Word 仍按字符数缩进")
            self.assertIsNone(ind.get(qn("w:leftChars")), "leftChars 残留会导致 Word 仍按字符数缩进")
            self.assertIsNone(ind.get(qn("w:rightChars")), "rightChars 残留会导致 Word 仍按字符数缩进")

    def test_normalize_docx_clears_first_line_chars_in_normal_style(self):
        """回归测试：Normal 样式里的 firstLineChars 必须清掉。
        表格单元格段落没有显式 firstLineChars 时会继承 Normal 样式，
        若 Normal 样式 firstLineChars=200（2 字符），Word 仍按 2 字符渲染。
        """
        from normalizer import normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "normal_with_firstlinechars.docx"
            output = Path(tmp) / "normal_with_firstlinechars_规范化.docx"

            # 手工构造：Normal 样式带 firstLineChars=200
            doc = Document()
            normal = doc.styles["Normal"]
            pPr = normal.element.find(qn("w:pPr"))
            if pPr is None:
                pPr = OxmlElement("w:pPr")
                normal.element.append(pPr)
            ind = OxmlElement("w:ind")
            ind.set(qn("w:firstLine"), "420")
            ind.set(qn("w:firstLineChars"), "200")
            pPr.append(ind)
            # 加一个表格 + 普通段落，模拟真实场景
            doc.add_paragraph("这是普通正文段落。")
            tbl = doc.add_table(rows=1, cols=1)
            tbl.rows[0].cells[0].paragraphs[0].add_run("单元格文字")
            doc.save(source)

            normalize_docx(cfg, source, output)
            normalized = Document(output)
            normal = normalized.styles["Normal"]
            pPr = normal.element.find(qn("w:pPr"))
            ind = pPr.find(qn("w:ind")) if pPr is not None else None
            # 关键断言：Normal 样式里 firstLineChars 必须清掉
            if ind is not None:
                self.assertIsNone(
                    ind.get(qn("w:firstLineChars")), "Normal 样式的 firstLineChars 必须清掉，否则表格段落会继承"
                )


def _add_page_field(paragraph):
    """在段落中插入 PAGE 域。"""
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instr = OxmlElement("w:instrText")
    instr.set(qn("xml:space"), "preserve")
    instr.text = "PAGE"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instr)
    run._r.append(end)


def _build_cover_toc_document(path):
    """构造一份含封面（无页码 section）+ 目录条目 + 正文的测试文档。

    结构：
      section 0（封面/前置页，无页码）：封面标题、文档信息、封面表格、分节符
      section 1（正文，有页码）：目录标题、TOC 条目（toc 1/2 样式）、一级标题、正文
    """
    doc = Document()
    try:
        doc.styles.add_style("toc 1", WD_STYLE_TYPE.PARAGRAPH)
        doc.styles.add_style("toc 2", WD_STYLE_TYPE.PARAGRAPH)
    except Exception:
        pass

    # 封面 section（section 0）
    doc.add_paragraph("封面标题")
    doc.add_paragraph("文档信息")
    cover_tbl = doc.add_table(rows=2, cols=2)
    cover_tbl.rows[0].cells[0].text = "项目"
    cover_tbl.rows[0].cells[1].text = "编号"
    cover_tbl.rows[1].cells[0].text = "名称"
    cover_tbl.rows[1].cells[1].text = "示例"
    doc.add_paragraph()  # 承载分节符的空段落
    doc.add_section(WD_SECTION.NEW_PAGE)

    # 正文 section（section 1）
    doc.add_paragraph("目录")  # 目录标题
    p_toc1 = doc.add_paragraph("1. 引言\t1")
    try:
        p_toc1.style = "toc 1"
    except Exception:
        pass
    p_toc2 = doc.add_paragraph("1.1 目的\t1")
    try:
        p_toc2.style = "toc 2"
    except Exception:
        pass
    doc.add_paragraph("一、引言")  # 一级标题
    doc.add_paragraph("正文内容。")  # 正文

    # 仅给正文 section 加页码，封面 section 不加
    sec1 = doc.sections[1]
    sec1.footer.is_linked_to_previous = False
    footer_p = sec1.footer.paragraphs[0]
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    _add_page_field(footer_p)

    doc.save(path)
    return doc


class NormalizeCoverAndTocTests(unittest.TestCase):
    """验证封面保留与目录错乱修复。"""

    def test_normalize_preserves_cover_and_toc(self):
        from normalizer import normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "with_cover.docx"
            output = Path(tmp) / "with_cover_规范化.docx"
            _build_cover_toc_document(source)

            normalize_docx(cfg, source, output)
            result = Document(output)

            # 段落索引（add_section 会插入一个空段作为分节符载体）：
            # 0 封面标题 / 1 文档信息 / 2 空段 / 3 分节符空段 / 4 目录 / 5 toc1 / 6 toc2 / 7 H1 / 8 正文
            # 1) 目录标题保留原格式，未被套正文格式（正文会加 Pt(24) 首行缩进）
            title_para = result.paragraphs[4]
            self.assertEqual((title_para.text or "").strip(), "目录")
            self.assertNotEqual(title_para.paragraph_format.first_line_indent, Pt(24))

            # 2) TOC 条目样式保留，未被改成 Heading
            toc1 = result.paragraphs[5]
            toc2 = result.paragraphs[6]
            self.assertTrue(toc1.style.name.lower().startswith("toc"))
            self.assertTrue(toc2.style.name.lower().startswith("toc"))

            # 3) 封面段落保留原格式，未被套正文格式
            cover_para = result.paragraphs[1]
            self.assertEqual((cover_para.text or "").strip(), "文档信息")
            self.assertNotEqual(cover_para.paragraph_format.first_line_indent, Pt(24))

            # 4) 封面 section 无页码，正文 section 有页码
            self.assertNotIn("PAGE", result.sections[0].footer._element.xml)
            self.assertIn("PAGE", result.sections[1].footer._element.xml)

            # 5) 正文标题与正文段落正常规范化
            h1_para = result.paragraphs[7]
            self.assertEqual(h1_para.style.name, "Heading 1")
            body_para = result.paragraphs[8]
            self.assertEqual(body_para.alignment, 3)  # JUSTIFY
            self.assertEqual(body_para.paragraph_format.first_line_indent, Pt(24))

    def test_normalize_preserves_cover_table(self):
        from normalizer import normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "with_cover.docx"
            output = Path(tmp) / "with_cover_规范化.docx"
            _build_cover_toc_document(source)

            normalize_docx(cfg, source, output)
            result = Document(output)

            # 封面表格保留原版式：单元格段落对齐未被改为 center（正文表格默认会居中）
            cover_cell = result.tables[0].rows[0].cells[0]
            self.assertNotEqual(cover_cell.paragraphs[0].alignment, WD_ALIGN_PARAGRAPH.CENTER)

    def test_mark_toc_fields_dirty(self):
        from normalizer.normalize import _mark_toc_fields_dirty

        doc = Document()
        p = doc.add_paragraph()
        r1 = p.add_run()
        begin = OxmlElement("w:fldChar")
        begin.set(qn("w:fldCharType"), "begin")
        r1._r.append(begin)
        r2 = p.add_run()
        instr = OxmlElement("w:instrText")
        instr.set(qn("xml:space"), "preserve")
        instr.text = 'TOC \\o "1-3" \\h \\u '
        r2._r.append(instr)
        r3 = p.add_run()
        sep = OxmlElement("w:fldChar")
        sep.set(qn("w:fldCharType"), "separate")
        r3._r.append(sep)
        r4 = p.add_run()
        end = OxmlElement("w:fldChar")
        end.set(qn("w:fldCharType"), "end")
        r4._r.append(end)

        _mark_toc_fields_dirty(doc)
        self.assertEqual(begin.get(qn("w:dirty")), "true")

    def test_helpers_section_and_toc_detection(self):
        from normalizer.normalize import (
            _build_body_section_maps,
            _collect_cover_section_indices,
            _find_toc_title_index,
            _is_toc_style,
            _section_has_page_number,
        )

        with tempfile.TemporaryDirectory() as tmp:
            path = Path(tmp) / "with_cover.docx"
            _build_cover_toc_document(path)
            doc = Document(path)

            # section 0 无页码（封面），section 1 有页码（正文）
            self.assertFalse(_section_has_page_number(doc.sections[0]))
            self.assertTrue(_section_has_page_number(doc.sections[1]))

            cover = _collect_cover_section_indices(doc)
            self.assertEqual(cover, {0})

            para_to_section, table_to_section = _build_body_section_maps(doc)
            # 封面段落属于 section 0，正文段落属于 section 1
            self.assertEqual(para_to_section[0], 0)
            self.assertEqual(para_to_section[4], 1)
            # 封面表格属于 section 0
            self.assertEqual(table_to_section[0], 0)

            # 目录标题索引应为 4（首个 toc 条目 5 之前最近的非空非 toc 段）
            self.assertEqual(_find_toc_title_index(doc), 4)

            # toc 样式识别
            self.assertTrue(_is_toc_style(doc.paragraphs[5]))
            self.assertFalse(_is_toc_style(doc.paragraphs[7]))


class NormalizeErrorTests(unittest.TestCase):
    def test_normalize_error_includes_message_cause_hint(self):
        from normalizer import NormalizeError

        e = NormalizeError("找不到文件", cause=FileNotFoundError("x"), hint="检查路径")
        s = str(e)
        self.assertIn("找不到文件", s)
        self.assertIn("FileNotFoundError", s)
        self.assertIn("检查路径", s)

    def test_input_not_found_is_normalize_error(self):
        from normalizer import InputNotFoundError, NormalizeError

        self.assertTrue(issubclass(InputNotFoundError, NormalizeError))

    def test_each_subclass_exists(self):
        from normalizer import (
            CorruptDocxError,
            InputNotFoundError,
            InvalidFileTypeError,
            NormalizeError,
            OutputNotWritableError,
            SameInputOutputError,
        )

        for cls in (
            InputNotFoundError,
            InvalidFileTypeError,
            CorruptDocxError,
            OutputNotWritableError,
            SameInputOutputError,
        ):
            self.assertTrue(issubclass(cls, NormalizeError))


class ValidatePathsTests(unittest.TestCase):
    def test_raises_input_not_found(self):
        from normalizer import InputNotFoundError
        from normalizer.paths import _validate_paths

        with self.assertRaises(InputNotFoundError) as ctx:
            _validate_paths("Z:/nonexistent-xyz-9999.docx", "Z:/out.docx")
        self.assertIn("找不到文件", str(ctx.exception))

    def test_raises_invalid_file_type(self):
        from normalizer import InvalidFileTypeError
        from normalizer.paths import _validate_paths

        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False, mode="w") as f:
            f.write("hi")
            p = f.name
        try:
            with self.assertRaises(InvalidFileTypeError) as ctx:
                _validate_paths(p, p + ".out.docx")
            self.assertIn("不是有效的 docx", str(ctx.exception))
        finally:
            os.remove(p)

    def test_raises_same_input_output(self):
        from normalizer import SameInputOutputError
        from normalizer.paths import _validate_paths

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            p = f.name
        try:
            with self.assertRaises(SameInputOutputError):
                _validate_paths(p, p)
        finally:
            os.remove(p)


class BackupSourceTests(unittest.TestCase):
    def test_creates_bak_file(self):
        from normalizer.paths import _backup_source

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            f.write(b"x")
            p = f.name
        try:
            bak = _backup_source(p)
            self.assertIsNotNone(bak)
            self.assertTrue(bak.endswith(".bak"))
            self.assertTrue(os.path.exists(bak))
        finally:
            os.remove(p)
            if bak and os.path.exists(bak):
                os.remove(bak)

    def test_uses_incremented_name_when_bak_exists(self):
        from normalizer.paths import _backup_source

        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            f.write(b"x")
            p = f.name
        bak1 = p + ".bak"
        try:
            open(bak1, "wb").close()
            bak2 = _backup_source(p)
            self.assertEqual(bak2, p + ".bak.1")
        finally:
            os.remove(p)
            for b in (bak1, p + ".bak.1"):
                if os.path.exists(b):
                    os.remove(b)

    def test_returns_none_on_missing_input(self):
        from normalizer.paths import _backup_source

        result = _backup_source("Z:/nonexistent-xyz-qq-9999/doc.docx")
        self.assertIsNone(result)


class NormalizeDocxNewSignatureTests(unittest.TestCase):
    def test_default_returns_str_backward_compat(self):
        from normalizer import normalize_docx

        cfg = load_config()
        source = SAMPLE_DOCX
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "out.docx"
            result = normalize_docx(cfg, source, output)
            self.assertIsInstance(result, str)
            self.assertEqual(result, str(output))

    def test_return_result_returns_normalize_result(self):
        from normalizer import NormalizeResult, normalize_docx

        cfg = load_config()
        source = SAMPLE_DOCX
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "out.docx"
            result = normalize_docx(cfg, source, output, return_result=True)
            self.assertIsInstance(result, NormalizeResult)
            self.assertEqual(result.input_path, str(source))

    def test_input_not_found_raises_typed_error(self):
        from normalizer import InputNotFoundError, normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(InputNotFoundError):
                normalize_docx(cfg, Path(tmp) / "nope.docx", Path(tmp) / "out.docx")

    def test_dry_run_does_not_write(self):
        from normalizer import normalize_docx

        cfg = load_config()
        source = SAMPLE_DOCX
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "out.docx"
            result = normalize_docx(cfg, source, output, dry_run=True, return_result=True)
            self.assertTrue(result.dry_run)
            self.assertFalse(output.exists())
            self.assertGreater(result.paragraphs_processed, 0)


class NormalizeErrorRaiseTests(unittest.TestCase):
    def test_input_not_found(self):
        from normalizer import InputNotFoundError, normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(InputNotFoundError):
                normalize_docx(load_config(), Path(tmp) / "missing-xyz-9999.docx", Path(tmp) / "out.docx")

    def test_invalid_file_type_for_txt(self):
        from normalizer import InvalidFileTypeError, normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            txt = Path(tmp) / "fake.txt"
            txt.write_text("hi", encoding="utf-8")
            with self.assertRaises(InvalidFileTypeError):
                normalize_docx(load_config(), txt, Path(tmp) / "out.docx")

    def test_same_input_output(self):
        from normalizer import SameInputOutputError, normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "same.docx"
            Document().save(src)
            with self.assertRaises(SameInputOutputError):
                normalize_docx(load_config(), src, src)

    def test_corrupt_docx_garbage(self):
        from normalizer import CorruptDocxError, normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "garbage.docx"
            from tests._helpers import make_garbage_docx

            make_garbage_docx(str(src))
            with self.assertRaises(CorruptDocxError):
                normalize_docx(load_config(), src, Path(tmp) / "out.docx")

    def test_corrupt_docx_truncated(self):
        from normalizer import CorruptDocxError, normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "truncated.docx"
            from tests._helpers import make_corrupt_docx

            make_corrupt_docx(str(src))
            with self.assertRaises(CorruptDocxError):
                normalize_docx(load_config(), src, Path(tmp) / "out.docx")

    def test_output_not_writable(self):
        from normalizer import OutputNotWritableError, normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "ok.docx"
            Document().save(src)
            with self.assertRaises(OutputNotWritableError):
                normalize_docx(load_config(), src, "Z:/nonexistent-zzz-9999/out.docx")


class BackupAndDryRunTests(unittest.TestCase):
    def test_default_no_backup_does_not_create_bak(self):
        """0.2.1+ 默认 backup=False,不产生 .bak。"""
        from normalizer import normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "doc.docx"
            Document().save(src)
            normalize_docx(load_config(), src, Path(tmp) / "out.docx")
            self.assertFalse((Path(tmp) / "doc.docx.bak").exists())

    def test_backup_true_creates_bak_file(self):
        from normalizer import normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "doc.docx"
            Document().save(src)
            normalize_docx(load_config(), src, Path(tmp) / "out.docx", backup=True)
            self.assertTrue((Path(tmp) / "doc.docx.bak").exists())

    def test_backup_true_increments_when_exists(self):
        from normalizer import normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "doc.docx"
            Document().save(src)
            (Path(tmp) / "doc.docx.bak").write_bytes(b"x")
            normalize_docx(load_config(), src, Path(tmp) / "out.docx", backup=True)
            self.assertTrue((Path(tmp) / "doc.docx.bak.1").exists())

    def test_dry_run_does_not_write_output(self):
        from normalizer import normalize_docx

        cfg = load_config()
        source = SAMPLE_DOCX
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "out.docx"
            result = normalize_docx(cfg, source, output, dry_run=True, return_result=True)
            self.assertTrue(result.dry_run)
            self.assertFalse(output.exists())

    def test_dry_run_returns_normalize_result(self):
        from normalizer import NormalizeResult, normalize_docx

        cfg = load_config()
        source = SAMPLE_DOCX
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "out.docx"
            result = normalize_docx(cfg, source, output, dry_run=True, return_result=True)
            self.assertIsInstance(result, NormalizeResult)
            self.assertGreater(result.paragraphs_processed, 0)

    def test_backup_false_explicit_skips_backup(self):
        from normalizer import normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "doc.docx"
            Document().save(src)
            normalize_docx(load_config(), src, Path(tmp) / "out.docx", backup=False)
            self.assertFalse((Path(tmp) / "doc.docx.bak").exists())


class EdgeCaseGracefulDegradationTests(unittest.TestCase):
    def test_continues_with_unrecognized_paragraph(self):
        from normalizer import normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            from tests._helpers import make_docx_with_unrecognized_paragraph

            src = Path(tmp) / "doc.docx"
            make_docx_with_unrecognized_paragraph(str(src))
            result = normalize_docx(load_config(), src, Path(tmp) / "out.docx", return_result=True)
            self.assertTrue((Path(tmp) / "out.docx").exists())
            self.assertGreater(result.paragraphs_processed, 0)

    def test_handles_nested_table(self):
        from normalizer import normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            from tests._helpers import make_docx_with_nested_table

            src = Path(tmp) / "nested.docx"
            make_docx_with_nested_table(str(src))
            normalize_docx(load_config(), src, Path(tmp) / "out.docx")
            self.assertTrue((Path(tmp) / "out.docx").exists())

    def test_handles_empty_table(self):
        from normalizer import normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            from tests._helpers import make_docx_with_empty_table

            src = Path(tmp) / "empty_tbl.docx"
            make_docx_with_empty_table(str(src))
            normalize_docx(load_config(), src, Path(tmp) / "out.docx")
            self.assertTrue((Path(tmp) / "out.docx").exists())

    def test_message_chinese_human_readable(self):
        from normalizer import NormalizeError, normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            try:
                normalize_docx(load_config(), Path(tmp) / "missing-9999.docx", Path(tmp) / "out.docx")
            except NormalizeError as e:
                s = str(e)
                self.assertIn("找不到文件", s)
                self.assertIn("检查路径", s)


class DetectionRegressionTests(unittest.TestCase):
    """题注、标题、参考文献识别不再误判正文。"""

    def _normalize(self, doc, cfg=None):
        from normalizer import normalize_docx

        cfg = cfg or load_config()
        tmp = tempfile.TemporaryDirectory()
        self.addCleanup(tmp.cleanup)
        source = Path(tmp.name) / "in.docx"
        output = Path(tmp.name) / "out.docx"
        doc.save(source)
        normalize_docx(cfg, source, output)
        return Document(output)

    def test_body_paragraph_starting_with_figure_or_table_word_is_not_caption(self):
        doc = Document()
        doc.add_paragraph("图片统一宽度等于正文宽度，高度按宽高比缩放。")
        doc.add_paragraph("表题置于表上方，单元格水平居中。")
        result = self._normalize(doc)
        for para in (result.paragraphs[0], result.paragraphs[1]):
            self.assertEqual(para.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
            self.assertFalse(para.runs[0].font.bold)

    def test_numbered_captions_still_detected(self):
        doc = Document()
        doc.add_paragraph("图1  各季度销售额对比")
        doc.add_paragraph("表2  主要指标年度对比")
        result = self._normalize(doc)
        for para in (result.paragraphs[0], result.paragraphs[1]):
            self.assertEqual(para.alignment, WD_ALIGN_PARAGRAPH.CENTER)
            self.assertTrue(para.runs[0].font.bold)

    def test_decimal_prefix_is_not_heading(self):
        doc = Document()
        doc.add_paragraph("1.5 倍行距是最常见的选择。")
        result = self._normalize(doc)
        self.assertNotEqual(result.paragraphs[0].style.name, "Heading 3")

    def test_reference_mention_in_body_does_not_switch_mode(self):
        doc = Document()
        doc.add_paragraph("本文参考文献如下所示。")
        doc.add_paragraph("[1] 此段仍是正文，不是参考文献条目。")
        result = self._normalize(doc)
        second = result.paragraphs[1]
        self.assertEqual(second.alignment, WD_ALIGN_PARAGRAPH.JUSTIFY)
        self.assertEqual(second.paragraph_format.first_line_indent, Pt(24))


class PageNumberFooterTests(unittest.TestCase):
    def test_existing_page_field_paragraph_is_rewritten_in_place(self):
        from normalizer import normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "footer_two_paras.docx"
            output = Path(tmp) / "out.docx"
            doc = Document()
            doc.add_paragraph("正文")
            footer = doc.sections[0].footer
            footer.paragraphs[0].text = "页脚说明文字"
            _add_page_field(footer.add_paragraph())
            doc.save(source)

            normalize_docx(cfg, source, output)
            normalized = Document(output)
            footer = normalized.sections[0].footer
            self.assertEqual((footer.paragraphs[0].text or "").strip(), "页脚说明文字")
            page_fields = [el.text for el in footer._element.iter(qn("w:instrText")) if el.text and "PAGE" in el.text]
            self.assertEqual(len(page_fields), 1)


class PageSizeSupportTests(unittest.TestCase):
    def test_letter_page_size_applied(self):
        from normalizer import normalize_docx

        cfg = load_config()
        cfg["page"]["size"] = "Letter"
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "plain.docx"
            output = Path(tmp) / "out.docx"
            doc = Document()
            doc.add_paragraph("正文")
            doc.save(source)
            normalize_docx(cfg, source, output)
            normalized = Document(output)
            self.assertAlmostEqual(normalized.sections[0].page_width.cm, 21.59, places=1)
            self.assertAlmostEqual(normalized.sections[0].page_height.cm, 27.94, places=1)


class HeaderFooterDetectionTests(unittest.TestCase):
    def test_empty_document_has_no_header_footer(self):
        from normalizer import inspect_docx

        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "empty.docx"
            Document().save(source)
            self.assertFalse(inspect_docx(source).has_header_footer)

    def test_sample_document_has_header_footer(self):
        from normalizer import inspect_docx

        self.assertTrue(inspect_docx(SAMPLE_DOCX).has_header_footer)


class InspectDocxErrorTests(unittest.TestCase):
    def test_inspect_docx_rejects_non_docx(self):
        from normalizer import InvalidFileTypeError, inspect_docx

        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False, mode="w") as f:
            f.write("hi")
            p = f.name
        try:
            with self.assertRaises(InvalidFileTypeError):
                inspect_docx(p)
        finally:
            os.remove(p)


class ChangesCollectionTests(unittest.TestCase):
    @staticmethod
    def _build_scruffy_doc(path):
        """未规范化的原始文档：默认样式、无题注格式、无表格格式。"""
        doc = Document()
        doc.add_paragraph("一、绪论")
        doc.add_paragraph("正文内容，默认样式。")
        doc.add_paragraph("图1  示意图")
        doc.add_paragraph("表1  数据表")
        doc.add_paragraph("参考文献")
        doc.add_paragraph("[1] 张三. 某主题研究[J]. 某学报, 2020, 3: 1-9.")
        tbl = doc.add_table(rows=2, cols=2)
        tbl.rows[0].cells[0].text = "指标"
        tbl.rows[0].cells[1].text = "数值"
        tbl.rows[1].cells[0].text = "长度"
        tbl.rows[1].cells[1].text = "12"
        doc.save(path)

    def test_sample_changes_cover_headings_captions_references_tables(self):
        from normalizer import normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "scruffy.docx"
            output = Path(tmp) / "out.docx"
            self._build_scruffy_doc(source)
            result = normalize_docx(cfg, source, output, return_result=True)
            texts = "\n".join(result.changes)
            self.assertIn("级标题", texts)
            self.assertIn("图题注", texts)
            self.assertIn("表题注", texts)
            self.assertIn("参考文献条目", texts)
            self.assertIn("表格 0", texts)
            self.assertGreater(result.paragraphs_processed, 0)

    def test_counts_and_changes_exclude_cover_sections(self):
        from normalizer import normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "with_cover.docx"
            output = Path(tmp) / "out.docx"
            _build_cover_toc_document(source)
            result = normalize_docx(cfg, source, output, return_result=True)
            self.assertEqual(result.paragraphs_processed, 2)
            self.assertEqual(result.tables_processed, 0)
            self.assertTrue(any("1 级标题" in c for c in result.changes))
            self.assertFalse(any(c.startswith("段落 0") or c.startswith("段落 1") for c in result.changes))

    def test_dry_run_without_result_returns_none(self):
        from normalizer import normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "out.docx"
            self.assertIsNone(normalize_docx(cfg, SAMPLE_DOCX, output, dry_run=True))


class WarningCollectionTests(unittest.TestCase):
    def test_invalid_image_size_produces_warning_and_callback(self):
        from normalizer import normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "zero_image.docx"
            output = Path(tmp) / "out.docx"
            doc = Document()
            para = doc.add_paragraph()
            para.add_run().add_picture(str(ROOT / "assets" / "figure1.png"))
            extent = para._element.findall(".//" + qn("wp:extent"))[0]
            extent.set("cx", "0")
            extent.set("cy", "0")
            doc.save(source)

            received = []
            result = normalize_docx(
                cfg,
                source,
                output,
                return_result=True,
                on_warning=lambda msg, location=None: received.append((msg, location)),
            )
            self.assertTrue(any("尺寸无效" in msg for msg in result.warnings))
            self.assertTrue(any("尺寸无效" in msg for msg, _ in received))


class ImageSizingTests(unittest.TestCase):
    """图片尺寸策略：fit（等比上限）/ fixed（统一尺寸，等效 Word VBA）。"""

    EMU_PER_CM = 360000
    DELTA_CM = 0.01  # 页面尺寸经 twips 往返换算，厘米值有极小误差

    def cm(self, emu):
        return emu / self.EMU_PER_CM

    def assertCm(self, emu, expected_cm):
        self.assertAlmostEqual(self.cm(emu), expected_cm, delta=self.DELTA_CM)

    def _normalize(self, tmp, source, **figure_overrides):
        from normalizer import normalize_docx

        cfg = load_config()
        cfg["figure"].update(figure_overrides)
        output = Path(tmp) / "out.docx"
        warnings = []
        result = normalize_docx(
            cfg,
            source,
            output,
            return_result=True,
            on_warning=lambda msg, location=None: warnings.append(msg),
        )
        return result, picture_sizes(output), warnings

    def test_fit_mode_only_shrinks_oversized_image(self):
        """默认 fit：大图等比缩到上限内，小图原样不动。"""
        from tests._helpers import make_docx_with_images

        with tempfile.TemporaryDirectory() as tmp:
            source = make_docx_with_images(Path(tmp) / "in.docx", [(18, 12), (3, 2)])
            result, sizes, _ = self._normalize(tmp, source)

            self.assertEqual(result.images_processed, 1)
            self.assertCm(sizes[0][1], 15.5)
            self.assertCm(sizes[0][2], 12 * 15.5 / 18)
            self.assertCm(sizes[1][1], 3)
            self.assertCm(sizes[1][2], 2)

    def test_fixed_mode_stretches_both_dimensions(self):
        """fixed + 不保持比例：所有图片强制拉伸到同一尺寸（等同 VBA 解锁长宽比）。"""
        from tests._helpers import make_docx_with_images

        with tempfile.TemporaryDirectory() as tmp:
            source = make_docx_with_images(Path(tmp) / "in.docx", [(18, 12), (3, 2)])
            result, sizes, _ = self._normalize(
                tmp,
                source,
                sizeMode="fixed",
                fixedWidthCm=12,
                fixedHeightCm=9,
                keepAspectRatio=False,
            )

            self.assertEqual(result.images_processed, 2)
            for _, cx, cy, _ in sizes:
                self.assertCm(cx, 12)
                self.assertCm(cy, 9)

    def test_fixed_mode_uniform_width_keeps_ratio(self):
        """fixed + 保持比例 + 只设宽度：所有图片统一宽度、高度各自等比。"""
        from tests._helpers import make_docx_with_images

        with tempfile.TemporaryDirectory() as tmp:
            source = make_docx_with_images(Path(tmp) / "in.docx", [(18, 12), (3, 2)])
            result, sizes, _ = self._normalize(
                tmp,
                source,
                sizeMode="fixed",
                fixedWidthCm=10,
                fixedHeightCm=None,
                keepAspectRatio=True,
            )

            self.assertEqual(result.images_processed, 2)
            self.assertCm(sizes[0][1], 10)
            self.assertCm(sizes[1][1], 10)
            self.assertCm(sizes[0][2], 12 * 10 / 18)
            self.assertCm(sizes[1][2], 2 * 10 / 3)

    def test_target_width_clamped_to_text_width_with_warning(self):
        """目标宽度超过正文可用宽度时按可用宽度收窄，并给出警告。"""
        from tests._helpers import make_docx_with_images

        with tempfile.TemporaryDirectory() as tmp:
            source = make_docx_with_images(Path(tmp) / "in.docx", [(18, 12)])
            result, sizes, warnings = self._normalize(
                tmp,
                source,
                sizeMode="fixed",
                fixedWidthCm=20,
                fixedHeightCm=None,
                keepAspectRatio=True,
                clampToTextWidth=True,
            )

            self.assertCm(sizes[0][1], 15.5)
            self.assertTrue(any("超过正文可用宽度" in msg for msg in warnings))
            self.assertTrue(any("超过正文可用宽度" in msg for msg in result.warnings))

    def test_clamp_can_be_disabled(self):
        """关闭安全阀后允许按给定厘米数写死（与 VBA 行为一致）。"""
        from tests._helpers import make_docx_with_images

        with tempfile.TemporaryDirectory() as tmp:
            source = make_docx_with_images(Path(tmp) / "in.docx", [(18, 12)])
            _, sizes, _ = self._normalize(
                tmp,
                source,
                sizeMode="fixed",
                fixedWidthCm=20,
                fixedHeightCm=None,
                keepAspectRatio=True,
                clampToTextWidth=False,
            )

            self.assertCm(sizes[0][1], 20)

    def test_both_extent_elements_stay_in_sync(self):
        """wp:extent 与 a:xfrm/a:ext 必须同步，否则 WPS / LibreOffice 显示尺寸不一致。"""
        from tests._helpers import make_docx_with_images

        with tempfile.TemporaryDirectory() as tmp:
            source = make_docx_with_images(Path(tmp) / "in.docx", [(18, 12)])
            _, sizes, _ = self._normalize(
                tmp,
                source,
                sizeMode="fixed",
                fixedWidthCm=12,
                fixedHeightCm=9,
                keepAspectRatio=False,
            )

            for _, cx, _, xfrm_cx in sizes:
                self.assertEqual(cx, xfrm_cx)

    def test_floating_image_is_resized(self):
        """浮动图片（wp:anchor）不在 doc.inline_shapes 中，同样要被处理。"""
        from tests._helpers import make_docx_with_floating_image

        with tempfile.TemporaryDirectory() as tmp:
            source = make_docx_with_floating_image(Path(tmp) / "float.docx", 18, 12)
            result, sizes, _ = self._normalize(tmp, source)

            self.assertEqual(len(sizes), 1)
            self.assertEqual(sizes[0][0], "anchor")
            self.assertEqual(result.images_processed, 1)
            self.assertCm(sizes[0][1], 15.5)

    def test_inspect_docx_counts_floating_images(self):
        """文档摘要的图片数要把浮动图片算进去。"""
        from normalizer import inspect_docx
        from tests._helpers import make_docx_with_floating_image

        with tempfile.TemporaryDirectory() as tmp:
            source = make_docx_with_floating_image(Path(tmp) / "float.docx", 18, 12)
            self.assertEqual(inspect_docx(source).image_count, 1)

    def test_table_image_skipped_by_default(self):
        """表格单元格里的图标默认跳过（单元格远窄于正文，强制统一会被拉爆）。"""
        from tests._helpers import make_docx_with_images

        with tempfile.TemporaryDirectory() as tmp:
            source = make_docx_with_images(Path(tmp) / "in.docx", [(18, 12)], table_image=True)
            result, sizes, _ = self._normalize(
                tmp,
                source,
                sizeMode="fixed",
                fixedWidthCm=10,
                fixedHeightCm=None,
                keepAspectRatio=True,
            )

            self.assertEqual(result.images_processed, 1)
            self.assertCm(sizes[-1][1], 2)
            self.assertTrue(any("位于表格内" in item for item in result.changes))

    def test_table_image_processed_when_skip_disabled(self):
        from tests._helpers import make_docx_with_images

        with tempfile.TemporaryDirectory() as tmp:
            source = make_docx_with_images(Path(tmp) / "in.docx", [(18, 12)], table_image=True)
            result, sizes, _ = self._normalize(
                tmp,
                source,
                sizeMode="fixed",
                fixedWidthCm=10,
                fixedHeightCm=None,
                keepAspectRatio=True,
                skipInTables=False,
            )

            self.assertEqual(result.images_processed, 2)
            self.assertCm(sizes[-1][1], 10)


class FigureCaptionPaginationTests(unittest.TestCase):
    """图与图题不应被分页拆开（w:keepNext）。"""

    @staticmethod
    def _keep_next_paragraphs(path):
        from docx.oxml.ns import qn

        doc = Document(str(path))
        result = []
        for index, para in enumerate(doc.paragraphs):
            ppr = para._element.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:keepNext")) is not None:
                result.append(index)
        return result

    def _run(self, tmp, caption_position, keep=True):
        from normalizer import normalize_docx
        from tests._helpers import make_docx_with_images

        cfg = load_config()
        cfg["figure"]["captionPosition"] = caption_position
        cfg["figure"]["keepWithCaption"] = keep
        source = make_docx_with_images(Path(tmp) / "in.docx", [(10, 7)], caption_position=caption_position)
        output = Path(tmp) / "out.docx"
        normalize_docx(cfg, source, output)
        return self._keep_next_paragraphs(output)

    def test_caption_below_keeps_image_with_caption(self):
        with tempfile.TemporaryDirectory() as tmp:
            # 段落顺序：0 正文、1 图、2 图题 → 图段（1）带 keepNext
            self.assertEqual(self._run(tmp, "below"), [1])

    def test_caption_above_keeps_caption_with_image(self):
        with tempfile.TemporaryDirectory() as tmp:
            # 段落顺序：0 正文、1 图题、2 图 → 图题段（1）带 keepNext
            self.assertEqual(self._run(tmp, "above"), [1])

    def test_disabled_adds_no_keep_next(self):
        with tempfile.TemporaryDirectory() as tmp:
            self.assertEqual(self._run(tmp, "below", keep=False), [])


class ChapterHeadingDetectionTests(unittest.TestCase):
    """「第X章 / X.1 / X.1.1 / 未编号章」体系识别，且不误判正文小数。"""

    @staticmethod
    def _levels(*texts):
        from normalizer.normalize import _detect_heading_level

        doc = Document()
        for text in texts:
            doc.add_paragraph(text)
        return [_detect_heading_level(p) for p in doc.paragraphs]

    def test_chapter_section_clause_numbering(self):
        self.assertEqual(self._levels("第1章 前言")[0], 1)
        self.assertEqual(self._levels("第12章 结论与展望")[0], 1)
        self.assertEqual(self._levels("第二章 相关技术")[0], 1)
        self.assertEqual(self._levels("1.1 绪论")[0], 2)
        self.assertEqual(self._levels("1.1.1 背景")[0], 3)
        self.assertEqual(self._levels("1.1.1.1 细分要点")[0], 4)
        self.assertEqual(self._levels("1.1.1背景")[0], 3)  # 规格范本里编号与标题可无空格

    def test_unnumbered_chapter_titles(self):
        for text in ("结论", "致谢", "参考文献", "附录", "摘要", "摘  要", "Abstract"):
            self.assertEqual(self._levels(text)[0], 1, text)
        # 带正文的句子不是未编号章标题
        self.assertEqual(self._levels("结论表明该方法有效。")[0], 0)

    def test_body_decimals_not_misjudged(self):
        for text in (
            "1.1 是小数",
            "1.5 倍行距是最常见的选择。",
            "3.14 是圆周率的近似值。",
            "第1章已经完成。",
            "2023 年度报告",
            "见第1章。",
        ):
            self.assertEqual(self._levels(text)[0], 0, text)

    def test_legacy_numbering_still_works(self):
        levels = self._levels("一、一级", "（一）二级", "1. 三级", "（1）四级")
        self.assertEqual(levels, [1, 2, 3, 4])

    def test_chapter_headings_applied_end_to_end(self):
        from normalizer import normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "thesis.docx"
            output = Path(tmp) / "thesis_规范化.docx"
            doc = Document()
            for text in ("第1章  前言", "正文内容。", "1.1  研究背景", "1.1.1  研究意义", "结论"):
                doc.add_paragraph(text)
            doc.save(source)

            normalize_docx(cfg, source, output)
            normalized = Document(output)
            styles = [p.style.name for p in normalized.paragraphs]
            self.assertEqual(styles, ["Heading 1", "Normal", "Heading 2", "Heading 3", "Heading 1"])


class NormalizeIdempotencyTests(unittest.TestCase):
    """同一文件连续规范化两次，第二次「变更」计数为 0。"""

    def test_second_normalize_reports_zero_changes(self):
        from normalizer import normalize_docx

        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            out1 = Path(tmp) / "out1.docx"
            out2 = Path(tmp) / "out2.docx"
            first = normalize_docx(cfg, SAMPLE_DOCX, out1, return_result=True)
            self.assertTrue(first.changes)

            second = normalize_docx(cfg, out1, out2, return_result=True)
            self.assertEqual(
                second.changes, [], "第二次规范化变更应为空，实际: " + json.dumps(second.changes, ensure_ascii=False)
            )

    def test_dry_run_lists_alignment_decisions_with_preset(self):
        from config_model import load_preset
        from normalizer import normalize_docx

        cfg = load_preset("成都理工大学学士论文")["config"]
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "in.docx"
            output = Path(tmp) / "out.docx"
            doc = Document()
            doc.add_paragraph("表1  实验数据")
            tbl = doc.add_table(rows=2, cols=3)
            for r, row in enumerate([["指标", "2024年", "备注"], ["长度", "12", "保持稳定"]]):
                for c, text in enumerate(row):
                    tbl.rows[r].cells[c].text = text
            doc.save(source)

            result = normalize_docx(cfg, source, output, dry_run=True, return_result=True)
            self.assertFalse(output.exists())
            self.assertTrue(any(c.startswith("表格 0：") and "横+纵表头" in c for c in result.changes))


if __name__ == "__main__":
    unittest.main()
