"""表格智能对齐：自动检测表头形态，并按单元格内容类型决定水平对齐。

呼应《成都理工大学学士论文》规格「表格内数字须上下对齐」——数字列整列
右对齐实现纵向对齐；表头居中、行名列左对齐、长短文本分治。

两轴检测（组合出 仅横 / 仅纵 / 横纵都有 / 无表头 四种形态）：
- 横轴（首行是否表头）：行属性带 w:tblHeader → 是；否则默认是（大多数表只有
  横表头）；例外——首行全部为纯数字/日期形态 → 否（Excel 粘贴的无表头数据表）。
- 纵轴（首列是否行名）：保守检测，未命中即视为普通数据列。命中信号（任一）：
  ① 左上角单元格为空，或为「项目/名称/类别/指标/内容/科目」类标签词；
  ② 首列（表头行除外）全为非数字短文本，且其余列至少一列多数为数字；
  ③ 首列带加粗或底纹（必须在清除底纹之前读取）。

对齐优先级（先命中先赢）：
  1. 表头行（含左上角格，横轴成立时）→ 居中
  2. 纵表头列（纵轴成立时）→ 左对齐；仅纵表头时左上角格也左对齐
  3. 纯数字/金额/计量 → 右对齐，按列投票：多数为数字则整列右对齐，
     「—」与空占位符跟随该列
  4. 日期（2024-05-01、2025年3月 等）与状态标签 → 居中（先于长度规则）
  5. 显示宽度 ≤ 阈值（中文计 1、半角计 0.5）→ 居中
  6. 其余长文本 → 左对齐

合并单元格按自身文本归类（不参与列规则）；空单元格跟随所在列。

TODO(小数点对齐)：Word 支持按小数制表位对齐（w:tab w:val="decimal"），可让
"12.5 / 13.75 / 120" 的小数点严格对齐成竖线；本期不做，数字列统一右对齐，
引入时在 _decide_cell 的数字列分支扩展。
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass, field

from docx.oxml.ns import qn

# ----- 内容类型识别 -----

# 数字/金额/计量：1,234.56、89%、12.5万元、-3.14 等（年份月日不在此列，归日期）
_NUM_RE = re.compile(
    r"^[+-]?\d[\d,]*(?:\.\d+)?"
    r"(?:\s*(?:%|％|万元|亿元|万|亿|元|吨|千克|克|公斤|kg|g|km|m|cm|mm|"
    r"台|套|件|个|人|次|页|张|辆|度|平方米|m2|m²))?$"
)

# 日期：2024-05-01、2024/5/1、2025年3月15日、2025年3月、2024-05、2024年
_DATE_RE = re.compile(
    r"^(?:\d{4}[-/.]\d{1,2}[-/.]\d{1,2}"
    r"|\d{4}年\d{1,2}月\d{1,2}日"
    r"|\d{4}年\d{1,2}月"
    r"|\d{4}[-/.]\d{1,2}"
    r"|\d{4}年)$"
)

# 空占位符：「—」「-」「/」「无」「N/A」等
_PLACEHOLDER_RE = re.compile(r"^(?:[-—―–~～/\\]+|N/A|n/a|无|不详)$")

# 纵表头左上角标签词（保守集合，宁缺勿滥）
_LABEL_WORDS = {"项目", "名称", "类别", "指标", "内容", "科目"}

# 状态标签（居中展示）
_STATUS_WORDS = {
    "是", "否", "已完成", "未完成", "进行中", "已开始", "未开始", "已结束",
    "已通过", "未通过", "已批准", "待审批", "已审批", "合格", "不合格",
    "正常", "异常", "通过", "不通过", "成功", "失败", "已发布", "待发布",
    "已提交", "待提交", "已归档", "待定", "优秀", "良好", "中等", "及格", "不及格",
}


def _clean(text):
    return (text or "").strip()


def is_numeric(text):
    return bool(_NUM_RE.match(_clean(text)))


def is_date(text):
    return bool(_DATE_RE.match(_clean(text)))


def is_placeholder(text):
    return bool(_PLACEHOLDER_RE.match(_clean(text)))


def is_status(text):
    return _clean(text) in _STATUS_WORDS


def display_width(text):
    """显示宽度：中文/全角计 1，半角计 0.5。"""
    width = 0.0
    for ch in _clean(text):
        width += 1.0 if unicodedata.east_asian_width(ch) in ("W", "F", "A") else 0.5
    return width


def is_label_word(text):
    return _clean(text) in _LABEL_WORDS


# ----- 表格网格（直接读 XML，正确处理合并单元格） -----

@dataclass
class CellInfo:
    tc: object          # lxml 元素 w:tc
    row: int            # 行号（0 起）
    col: int            # 网格列号（0 起，跨列时为起始列）
    span: int           # 横向跨列数
    vmerge: str | None  # None=普通 / "restart"=纵向合并锚 / "continue"=延续格
    text: str


def _grid_span(tc):
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return 1
    gs = tcPr.find(qn("w:gridSpan"))
    if gs is None:
        return 1
    try:
        return max(1, int(gs.get(qn("w:val"))))
    except (TypeError, ValueError):
        return 1


def _vmerge_state(tc):
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return None
    vm = tcPr.find(qn("w:vMerge"))
    if vm is None:
        return None
    return vm.get(qn("w:val")) or "continue"


def cell_text(tc):
    return _clean("".join(t.text or "" for t in tc.iter(qn("w:t"))))


def table_grid(tbl):
    """产出每行的物理单元格信息；跨列格占一个槽位、col 指向起始网格列。"""
    grid = []
    for r, tr in enumerate(tbl._tbl.findall(qn("w:tr"))):
        row_cells, col = [], 0
        for tc in tr.findall(qn("w:tc")):
            span = _grid_span(tc)
            row_cells.append(CellInfo(tc=tc, row=r, col=col, span=span,
                                      vmerge=_vmerge_state(tc), text=cell_text(tc)))
            col += span
        grid.append(row_cells)
    return grid


def tc_has_bold(tc):
    for b in tc.iter(qn("w:b")):
        if b.get(qn("w:val")) not in ("0", "false", "none"):
            return True
    return False


def tc_has_shading(tc):
    """是否有可见底纹（白/无填充不算）——必须在清除底纹之前调用。"""
    tcPr = tc.find(qn("w:tcPr"))
    if tcPr is None:
        return False
    shd = tcPr.find(qn("w:shd"))
    if shd is None:
        return False
    fill = (shd.get(qn("w:fill")) or "").strip()
    return fill not in ("", "auto", "none", "FFFFFF", "ffffff")


# ----- 分析与决策 -----

@dataclass
class TableAlignmentPlan:
    header_rows: tuple = ()
    col_header: bool = False
    numeric_cols: frozenset = frozenset()
    shape_label: str = "无表头"
    summary: str = ""
    grid: list = field(default_factory=list)
    # (行, 起始列) -> "left"/"center"/"right"。lxml 元素代理对象每次访问可能
    # 重建，不能以 id(tc) 为键；用网格坐标定位每个物理单元格。
    decisions: dict = field(default_factory=dict)

    def align_for(self, row, col):
        return self.decisions.get((row, col), "center")


def _leading_header_rows(grid):
    """行属性带 w:tblHeader 的起始连续行为表头行。"""
    rows = []
    for row in grid:
        tr = row[0].tc.getparent() if row else None
        if tr is None:
            break
        trPr = tr.find(qn("w:trPr"))
        if trPr is None or trPr.find(qn("w:tblHeader")) is None:
            break
        rows.append(row[0].row)
    return tuple(rows)


def _first_row_is_data(grid):
    """首行全部为纯数字/日期形态 → 无横表头（Excel 粘贴的无表头数据表）。

    严格判定：任一格为空或非数字形态（如交叉表左上角空格）即视为表头。
    """
    if not grid:
        return False
    cells = [c for c in grid[0] if c.span == 1]
    if not cells:
        return False
    return all(c.text and (is_numeric(c.text) or is_date(c.text)) for c in cells)


def _numeric_columns(grid, header_rows, n_cols):
    """按列投票：数据格多数为数字的列 → 数字列（占位符与空格不计票）。"""
    votes = {c: [0, 0] for c in range(n_cols)}
    for row in grid:
        if not row or row[0].row in header_rows:
            continue
        for cell in row:
            if cell.span != 1 or cell.vmerge == "continue":
                continue
            if not cell.text or is_placeholder(cell.text):
                continue
            if is_numeric(cell.text):
                votes[cell.col][0] += 1
            votes[cell.col][1] += 1
    return frozenset(c for c, (num, total) in votes.items() if total and num * 2 > total)


def _detect_col_header(grid, header_rows, numeric_cols, n_cols, threshold):
    """纵轴（首列行名）保守检测：任一信号命中才算。"""
    if n_cols < 2 or not grid or not grid[0]:
        return False
    # ① 左上角为空或标签词
    top_left = grid[0][0]
    if top_left.text == "" or is_label_word(top_left.text):
        return True
    # ② 首列（表头行除外）全为非数字短文本，且其余列至少一列多数为数字
    first_col = [c for row in grid if row and row[0].row not in header_rows
                 for c in (row[0],) if c.span == 1 and c.vmerge != "continue" and c.text]
    others_numeric = any(c in numeric_cols for c in range(1, n_cols))
    if (first_col and others_numeric
            and all(not is_numeric(c.text) and display_width(c.text) <= threshold * 2
                    for c in first_col)):
        return True
    # ③ 首列带加粗或底纹（须在清除底纹之前读取）；强调须区别于同行其他列，
    #    整表统一底纹/全部加粗不算首列标记
    for row in grid:
        if not row or row[0].row in header_rows:
            continue
        first = row[0]
        others = [c for c in row[1:] if c.span == 1]
        if not others:
            continue
        bold_marked = tc_has_bold(first.tc) and not any(tc_has_bold(c.tc) for c in others)
        shd_marked = tc_has_shading(first.tc) and not any(tc_has_shading(c.tc) for c in others)
        if bold_marked or shd_marked:
            return True
    return False


def analyze_table(tbl, threshold_chars=6):
    """检测表格形态并给出每个物理单元格的对齐决策与汇总说明。"""
    grid = table_grid(tbl)
    n_cols = max((sum(c.span for c in row) for row in grid), default=0)
    if not grid or n_cols == 0:
        return TableAlignmentPlan(summary="空表")

    header_rows = _leading_header_rows(grid)
    if not header_rows and not _first_row_is_data(grid):
        header_rows = (0,)  # 默认首行为表头
    numeric_cols = _numeric_columns(grid, set(header_rows), n_cols)
    col_header = _detect_col_header(grid, set(header_rows), numeric_cols, n_cols, threshold_chars)

    if header_rows and col_header:
        shape = "横+纵表头"
    elif header_rows:
        shape = "横表头"
    elif col_header:
        shape = "纵表头"
    else:
        shape = "无表头"

    # 列文本多数（用于空/占位格跟随所在列）
    col_prefers_center = {}
    for c in range(n_cols):
        centers = lefts = 0
        for row in grid:
            for cell in row:
                if (cell.col == c and cell.span == 1 and cell.vmerge != "continue"
                        and cell.text and cell.row not in header_rows
                        and not is_numeric(cell.text) and not is_date(cell.text)
                        and not is_status(cell.text) and not is_placeholder(cell.text)):
                    if display_width(cell.text) <= threshold_chars:
                        centers += 1
                    else:
                        lefts += 1
        col_prefers_center[c] = centers >= lefts

    decisions = {}
    stats = {"date_status": 0, "by_width": 0}
    for row in grid:
        for cell in row:
            align = _decide_cell(cell, header_rows, col_header, numeric_cols,
                                 threshold_chars, col_prefers_center, stats)
            decisions[(cell.row, cell.col)] = align

    display_numeric = numeric_cols - {0} if col_header else numeric_cols
    parts = []
    if col_header:
        parts.append("首列左对齐")
    if display_numeric:
        parts.append("数字 {0} 列右对齐".format(len(display_numeric)))
    if stats["date_status"]:
        parts.append("日期/状态 {0} 格居中".format(stats["date_status"]))
    if stats["by_width"]:
        parts.append("其余按文本长度")
    summary = shape if not parts else shape + "；" + "、".join(parts)

    return TableAlignmentPlan(header_rows=header_rows, col_header=col_header,
                              numeric_cols=numeric_cols, shape_label=shape,
                              summary=summary, grid=grid, decisions=decisions)


def _decide_cell(cell, header_rows, col_header, numeric_cols, threshold, col_prefers_center, stats):
    # 1. 表头行（含左上角格，横轴成立时）→ 居中
    if cell.row in header_rows:
        return "center"
    # 合并单元格按自身文本归类，不参与列规则
    if cell.span > 1:
        return _classify_text(cell.text, threshold, stats)
    # 2. 纵表头列 → 左对齐（仅纵表头时首行数据格同样走这里，左上角格在上方已处理）
    if col_header and cell.col == 0:
        return "left"
    # 3. 数字列整列右对齐（「—」与空占位符跟随该列）
    if cell.col in numeric_cols:
        return "right"
    # 空单元格 / 占位符跟随所在列的多数取向
    if not cell.text or is_placeholder(cell.text):
        return "center" if col_prefers_center.get(cell.col, True) else "left"
    # 4. 日期与状态标签 → 居中（先于长度规则）
    if is_date(cell.text) or is_status(cell.text):
        stats["date_status"] += 1
        return "center"
    # 5/6. 短文本居中，长文本左对齐
    stats["by_width"] += 1
    return "center" if display_width(cell.text) <= threshold else "left"


def _classify_text(text, threshold, stats):
    """合并单元格等独立格：按自身文本归类。"""
    if not text or is_placeholder(text):
        return "center"
    if is_date(text) or is_status(text):
        return "center"
    if is_numeric(text):
        return "right"
    return "center" if display_width(text) <= threshold else "left"
