"""表格智能对齐与表格独立性测试。

覆盖：内容分类器、四形态检测（横/纵/横+纵/无表头）、对齐优先级逐格断言
（含合并单元格、「—」占位、千分位金额、百分比、日期、状态）、headerBold
只加表头行、决策报告文案、幂等（第二次规范化变更数为 0 且对齐不变）、
表格独立于正文（Normal/自定义样式缩进清零、底纹清除、边框恢复黑色、五号宋体）。
"""

import json
import tempfile
import unittest
from pathlib import Path

from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

ROOT = Path(__file__).resolve().parents[1]

from config_model import load_preset  # noqa: E402


def preset_config():
    return load_preset("成都理工大学学士论文")["config"]


def tc_alignment(tc):
    """单元格首个段落的 w:jc 值（normalizer 会统一写入每个段落）。"""
    for p in tc.findall(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        if pPr is not None:
            jc = pPr.find(qn("w:jc"))
            if jc is not None:
                return jc.get(qn("w:val"))
    return None


def grid_of(doc, table_index=0):
    from normalizer.table_align import table_grid

    grid = table_grid(doc.tables[table_index])
    return [[(info.text, tc_alignment(info.tc)) for info in row] for row in grid]


def normalize_with_preset(tmp, doc, filename="in.docx"):
    from normalizer import normalize_docx

    source = Path(tmp) / "scruffy_source.docx"
    output = Path(tmp) / filename
    doc.save(source)
    result = normalize_docx(preset_config(), source, output, return_result=True)
    return output, result


# ----- 内容分类器 -----


class ClassifierTests(unittest.TestCase):
    def test_numeric_forms(self):
        from normalizer.table_align import is_numeric

        for text in ("123", "1,234.56", "89%", "12.5万元", "-3.14", "0.50", "2024", "100 kg"):
            self.assertTrue(is_numeric(text), text)
        for text in ("2024-05-01", "2024年", "甲方案", "1.1 是小数", "—", ""):
            self.assertFalse(is_numeric(text), text)

    def test_date_forms(self):
        from normalizer.table_align import is_date

        for text in ("2024-05-01", "2024/5/1", "2025年3月15日", "2025年3月", "2024-05", "2024年"):
            self.assertTrue(is_date(text), text)
        for text in ("2024-13-45x", "12.5万元", "2023 年度报告"):
            self.assertFalse(is_date(text), text)

    def test_status_and_placeholder(self):
        from normalizer.table_align import is_placeholder, is_status

        self.assertTrue(is_status("已完成"))
        self.assertTrue(is_status("进行中"))
        self.assertFalse(is_status("甲方案"))
        for text in ("—", "-", "/", "无", "N/A"):
            self.assertTrue(is_placeholder(text), text)
        self.assertFalse(is_placeholder("—备注"))

    def test_display_width(self):
        from normalizer.table_align import display_width

        self.assertEqual(display_width("指标名称"), 4.0)  # 全中文
        self.assertEqual(display_width("ABC"), 1.5)  # 半角
        self.assertEqual(display_width("图1"), 1.5)  # 中英混合


# ----- 对齐优先级（横+纵表头 + 合并 + 占位 + 日期） -----


class SmartAlignmentPriorityTests(unittest.TestCase):
    def _build(self):
        doc = Document()
        tbl = doc.add_table(rows=4, cols=5)
        rows = [
            ["指标", "2023年", "2024年", "同比增长", "备注"],
            ["营业收入（万元）", "1,200", "1,500", "25.0%", "总体保持增长态势"],
            ["净利润（万元）", "180", "225", "25.0%", "—"],
            ["执行情况", "全面达标", "", "", "2024-05-01"],
        ]
        for r, row in enumerate(rows):
            for c, text in enumerate(row):
                if text:
                    tbl.rows[r].cells[c].text = text
        # 第 4 行第 2-4 列横向合并（「全面达标」跨 3 列）
        tbl.rows[3].cells[1].merge(tbl.rows[3].cells[3])
        return doc

    def test_priority_rules_cell_by_cell(self):
        with tempfile.TemporaryDirectory() as tmp:
            output, _ = normalize_with_preset(tmp, self._build())
            doc = Document(output)
            grid = grid_of(doc)

            # 第 0 行为横表头（含左上角「指标」）→ 全部居中
            self.assertEqual([a for _, a in grid[0]], ["center"] * 5)
            # 首列为纵表头（左上角是标签词「指标」）→ 数据行首列左对齐
            self.assertEqual(grid[1][0], ("营业收入（万元）", "left"))
            self.assertEqual(grid[2][0], ("净利润（万元）", "left"))
            self.assertEqual(grid[3][0], ("执行情况", "left"))
            # 数字列（2023/2024/同比）按列投票 → 右对齐
            self.assertEqual(grid[1][1], ("1,200", "right"))
            self.assertEqual(grid[1][2], ("1,500", "right"))
            self.assertEqual(grid[1][3], ("25.0%", "right"))
            self.assertEqual(grid[2][1], ("180", "right"))
            self.assertEqual(grid[2][3], ("25.0%", "right"))
            # 长文本左对齐；「—」占位跟随列；日期居中（先于长度规则）
            self.assertEqual(grid[1][4], ("总体保持增长态势", "left"))
            self.assertEqual(grid[2][4], ("—", "left"))
            # 第 4 行合并后为 3 个物理单元格：[执行情况][全面达标(跨3列)][日期]
            self.assertEqual(grid[3][2], ("2024-05-01", "center"))
            # 合并单元格按自身文本归类：跨 3 列的「全面达标」宽度 4 ≤ 6 → 居中
            self.assertEqual(grid[3][1], ("全面达标", "center"))

    def test_decision_report(self):
        with tempfile.TemporaryDirectory() as tmp:
            _, result = normalize_with_preset(tmp, self._build())
            table_changes = [c for c in result.changes if c.startswith("表格 0")]
            self.assertEqual(len(table_changes), 1)
            self.assertIn("横+纵表头", table_changes[0])
            self.assertIn("首列左对齐", table_changes[0])
            self.assertIn("数字 3 列右对齐", table_changes[0])


class SmartAlignmentRowHeaderOnlyTests(unittest.TestCase):
    """横表头（无纵表头）：序号列按数字列右对齐，短文本居中，状态居中。"""

    def _build(self):
        doc = Document()
        tbl = doc.add_table(rows=4, cols=4)
        rows = [
            ["序号", "方案名称", "金额（元）", "状态"],
            ["1", "甲方案", "1,234.56", "已完成"],
            ["2", "乙方案", "2,468.00", "进行中"],
            ["3", "丙方案", "3,702.50", "待定"],
        ]
        for r, row in enumerate(rows):
            for c, text in enumerate(row):
                tbl.rows[r].cells[c].text = text
        return doc

    def test_alignment(self):
        with tempfile.TemporaryDirectory() as tmp:
            output, result = normalize_with_preset(tmp, self._build())
            grid = grid_of(Document(output))

            self.assertEqual([a for _, a in grid[0]], ["center"] * 4)
            for r in (1, 2, 3):
                self.assertEqual(grid[r][0], (str(r), "right"))  # 序号列数字右对齐
                self.assertEqual(grid[r][1][1], "center")  # 短文本（3 字）居中
                self.assertEqual(grid[r][2][1], "right")  # 千分位金额右对齐
                self.assertEqual(grid[r][3][1], "center")  # 状态标签居中
            report = next(c for c in result.changes if c.startswith("表格 0"))
            self.assertIn("横表头", report)
            self.assertIn("数字 2 列右对齐", report)

    def test_header_bold_only_on_header_row(self):
        with tempfile.TemporaryDirectory() as tmp:
            output, _ = normalize_with_preset(tmp, self._build())
            tbl = Document(output).tables[0]
            self.assertTrue(tbl.rows[0].cells[0].paragraphs[0].runs[0].font.bold)
            for r in (1, 2, 3):
                for c in range(4):
                    run = tbl.rows[r].cells[c].paragraphs[0].runs[0]
                    self.assertFalse(run.font.bold, f"({r},{c})")


class SmartAlignmentNoHeaderTests(unittest.TestCase):
    """首行全为数字 → 无表头（Excel 粘贴形态），全表数字右对齐。"""

    def test_no_header(self):
        doc = Document()
        tbl = doc.add_table(rows=2, cols=3)
        for r, row in enumerate([["1", "2", "3"], ["4", "5", "6"]]):
            for c, text in enumerate(row):
                tbl.rows[r].cells[c].text = text
        with tempfile.TemporaryDirectory() as tmp:
            output, result = normalize_with_preset(tmp, doc)
            grid = grid_of(Document(output))
            for row in grid:
                self.assertEqual([a for _, a in row], ["right"] * 3)
            report = next(c for c in result.changes if c.startswith("表格 0"))
            self.assertIn("无表头", report)


class SmartAlignmentColHeaderOnlyTests(unittest.TestCase):
    """仅纵表头：首行全数字（无横表头）+ 首列加粗 → 首列左对齐（含左上角格）。"""

    def _build(self):
        doc = Document()
        tbl = doc.add_table(rows=3, cols=3)
        rows = [["1", "100", "1,000"], ["2", "200", "2,000"], ["3", "300", "3,000"]]
        for r, row in enumerate(rows):
            for c, text in enumerate(row):
                tbl.rows[r].cells[c].text = text
                if c == 0:
                    for run in tbl.rows[r].cells[c].paragraphs[0].runs:
                        run.font.bold = True
        return doc

    def test_col_header_only(self):
        with tempfile.TemporaryDirectory() as tmp:
            output, result = normalize_with_preset(tmp, self._build())
            grid = grid_of(Document(output))
            for r in range(3):
                self.assertEqual(grid[r][0][1], "left")
                self.assertEqual(grid[r][1][1], "right")
                self.assertEqual(grid[r][2][1], "right")
            report = next(c for c in result.changes if c.startswith("表格 0"))
            self.assertIn("纵表头", report)
            self.assertIn("首列左对齐", report)
            self.assertIn("数字 2 列右对齐", report)  # 汇总不计纵表头列


class TblHeaderFlagTests(unittest.TestCase):
    """行属性带 w:tblHeader 时即使首行全数字也判定为表头。"""

    def test_tbl_header_row_is_header(self):
        doc = Document()
        tbl = doc.add_table(rows=2, cols=2)
        for r, row in enumerate([["1", "2"], ["3", "4"]]):
            for c, text in enumerate(row):
                tbl.rows[r].cells[c].text = text
        trPr = OxmlElement("w:trPr")
        trPr.append(OxmlElement("w:tblHeader"))
        tbl.rows[0]._tr.insert(0, trPr)
        with tempfile.TemporaryDirectory() as tmp:
            output, result = normalize_with_preset(tmp, doc)
            grid = grid_of(Document(output))
            self.assertEqual([a for _, a in grid[0]], ["center"] * 2)
            self.assertEqual([a for _, a in grid[1]], ["right"] * 2)
            report = next(c for c in result.changes if c.startswith("表格 0"))
            self.assertIn("横表头", report)


# ----- 幂等：同一表格跑两次结果一致；第二次变更数为 0 -----


class SmartAlignIdempotencyTests(unittest.TestCase):
    def _doc_with_everything(self):
        doc = Document()
        doc.add_paragraph("第1章  前言")
        doc.add_paragraph("正文第一段，默认格式。")
        doc.add_paragraph("1.1  研究背景")
        doc.add_paragraph("正文第二段。")
        doc.add_paragraph("表1  实验数据")
        tbl = doc.add_table(rows=3, cols=4)
        rows = [
            ["指标", "2024年", "同比", "备注"],
            ["样本数（个）", "1,200", "25.0%", "—"],
            ["合格率", "99.5%", "1.5%", "保持稳定"],
        ]
        for r, row in enumerate(rows):
            for c, text in enumerate(row):
                tbl.rows[r].cells[c].text = text
        return doc

    def test_second_run_reports_no_changes_and_keeps_alignment(self):
        from normalizer import normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "in.docx"
            out1 = Path(tmp) / "out1.docx"
            out2 = Path(tmp) / "out2.docx"
            self._doc_with_everything().save(source)

            first = normalize_docx(preset_config(), source, out1, return_result=True)
            self.assertTrue(first.changes, "首次规范化应有变更记录")

            second = normalize_docx(preset_config(), out1, out2, return_result=True)
            self.assertEqual(
                second.changes,
                [],
                "第二次规范化「变更」计数应为 0，实际: " + json.dumps(second.changes, ensure_ascii=False),
            )
            self.assertEqual(
                [row[:] for row in grid_of(Document(out1))],
                [row[:] for row in grid_of(Document(out2))],
                "同一表格跑两次，对齐结果必须一致",
            )


# ----- 表格独立于正文（样式级修复） -----


class TableIndependenceTests(unittest.TestCase):
    def _build_colored_indented(self):
        """Normal 样式 + 自定义「表格文字」样式都带字符缩进；单元格带彩色底纹；
        表格边框为红色；单元格字号 14pt。"""
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

        custom = doc.styles.add_style("表格文字", WD_STYLE_TYPE.PARAGRAPH)
        cpPr = OxmlElement("w:pPr")
        cInd = OxmlElement("w:ind")
        cInd.set(qn("w:firstLine"), "480")
        cInd.set(qn("w:firstLineChars"), "100")
        cpPr.append(cInd)
        custom.element.append(cpPr)

        tbl = doc.add_table(rows=2, cols=2)
        for r, row in enumerate([["指标", "数值"], ["长度", "12"]]):
            for c, text in enumerate(row):
                cell = tbl.rows[r].cells[c]
                cell.text = text
                para = cell.paragraphs[0]
                para.style = custom
                tcPr = cell._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:fill"), "FF0000")
                tcPr.append(shd)
                for run in para.runs:
                    run.font.size = None
                    run.font.bold = True
                    rpr = run._element.get_or_add_rPr()
                    sz = OxmlElement("w:sz")
                    sz.set(qn("w:val"), "28")
                    rpr.append(sz)

        tblPr = tbl._tbl.tblPr
        borders = OxmlElement("w:tblBorders")
        for side in ("top", "left", "bottom", "right", "insideH", "insideV"):
            el = OxmlElement("w:" + side)
            el.set(qn("w:val"), "single")
            el.set(qn("w:sz"), "4")
            el.set(qn("w:color"), "FF0000")
            borders.append(el)
        tblPr.append(borders)
        return doc

    def test_indent_shading_border_font_all_normalized(self):
        with tempfile.TemporaryDirectory() as tmp:
            output, _ = normalize_with_preset(tmp, self._build_colored_indented())
            doc = Document(output)
            tbl = doc.tables[0]

            for row in tbl.rows:
                for cell in row.cells:
                    # 段落级缩进清零且无字符单位属性
                    pPr = cell.paragraphs[0]._element.find(qn("w:pPr"))
                    ind = pPr.find(qn("w:ind")) if pPr is not None else None
                    if ind is not None:
                        self.assertIsNone(ind.get(qn("w:firstLineChars")))
                        self.assertEqual(ind.get(qn("w:firstLine")), "0")
                    # 底纹清除：显式无填充
                    tcPr = cell._tc.find(qn("w:tcPr"))
                    shd = tcPr.find(qn("w:shd"))
                    self.assertIsNotNone(shd, "应显式写入无填充以压过表格样式底色")
                    self.assertEqual(shd.get(qn("w:val")), "clear")
                    self.assertEqual(shd.get(qn("w:fill")), "auto")
                    # 字号五号（10.5pt），独立于正文
                    size = cell.paragraphs[0].runs[0].font.size
                    self.assertAlmostEqual(size.pt, 10.5)

            # 自定义样式（含 basedOn 链）样式级缩进清零
            custom = doc.styles["表格文字"]
            cpPr = custom.element.find(qn("w:pPr"))
            cInd = cpPr.find(qn("w:ind")) if cpPr is not None else None
            if cInd is not None:
                self.assertIsNone(cInd.get(qn("w:firstLineChars")))
                self.assertEqual(cInd.get(qn("w:firstLine")), "0")

            # Normal 样式同样清零
            normal = doc.styles["Normal"]
            npPr = normal.element.find(qn("w:pPr"))
            nInd = npPr.find(qn("w:ind")) if npPr is not None else None
            if nInd is not None:
                self.assertIsNone(nInd.get(qn("w:firstLineChars")))

            # 表格边框恢复黑色
            tblPr = tbl._tbl.tblPr
            borders = tblPr.find(qn("w:tblBorders"))
            for side in borders:
                self.assertEqual(side.get(qn("w:color")), "000000")

    def test_idempotent_after_color_strip(self):
        from normalizer import normalize_docx

        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "in.docx"
            out1 = Path(tmp) / "out1.docx"
            out2 = Path(tmp) / "out2.docx"
            self._build_colored_indented().save(source)
            normalize_docx(preset_config(), source, out1, return_result=True)
            second = normalize_docx(preset_config(), out1, out2, return_result=True)
            self.assertEqual(second.changes, [])


class FullTableShadingNotColHeaderTests(unittest.TestCase):
    """整表统一底纹不算「首列带底纹」：序号列仍按数字列右对齐。"""

    def test_uniform_shading_does_not_mark_first_column(self):
        doc = Document()
        tbl = doc.add_table(rows=3, cols=3)
        rows = [["序号", "项目", "结果"], ["1", "采样", "0.95"], ["2", "分析", "0.88"]]
        for r, row in enumerate(rows):
            for c, text in enumerate(row):
                tbl.rows[r].cells[c].text = text
                tcPr = tbl.rows[r].cells[c]._tc.get_or_add_tcPr()
                shd = OxmlElement("w:shd")
                shd.set(qn("w:val"), "clear")
                shd.set(qn("w:fill"), "FFF2CC")
                tcPr.append(shd)
        with tempfile.TemporaryDirectory() as tmp:
            output, result = normalize_with_preset(tmp, doc)
            grid = grid_of(Document(output))
            # 表头行居中；序号数字列右对齐；「采样/分析」短文本居中
            self.assertEqual([a for _, a in grid[0]], ["center"] * 3)
            self.assertEqual(grid[1][0], ("1", "right"))
            self.assertEqual(grid[2][0], ("2", "right"))
            report = next(c for c in result.changes if c.startswith("表格 0"))
            self.assertIn("横表头", report)


class CellFontDefaultTests(unittest.TestCase):
    """未配置 cellFont/cellSize 时默认五号宋体，不回退到正文字体字号。"""

    def test_default_cell_font_is_independent_of_body(self):
        from normalizer import normalize_docx

        cfg = json.loads((ROOT / "config.json").read_text(encoding="utf-8"))
        del cfg["table"]["cellSize"]  # 强制走默认路径
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "in.docx"
            output = Path(tmp) / "out.docx"
            doc = Document()
            doc.add_paragraph("正文。")
            tbl = doc.add_table(rows=1, cols=1)
            tbl.rows[0].cells[0].text = "内容"
            doc.save(source)
            normalize_docx(cfg, source, output)
            run = Document(output).tables[0].rows[0].cells[0].paragraphs[0].runs[0]
            self.assertAlmostEqual(run.font.size.pt, 10.5)  # 五号，而非正文小四
            rpr = run._element.get_or_add_rPr()
            rfonts = rpr.find(qn("w:rFonts"))
            self.assertEqual(rfonts.get(qn("w:eastAsia")), "宋体")


if __name__ == "__main__":
    unittest.main()
