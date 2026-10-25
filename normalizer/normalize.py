"""normalizer 核心规范化：样式应用、段落/表格/图片处理、TOC、检测。"""
from __future__ import annotations

import re
import zipfile
from dataclasses import dataclass

import lxml.etree
from docx import Document
from docx.enum.text import WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Cm, Pt
from docx.text.paragraph import Paragraph

from formatter import (
    add_page_number_field,
    clear_style_indent_chars,
    resolve_align,
    resolve_page_size,
    resolve_size,
    resolve_table_align,
    set_indent_values,
    set_para_spacing,
    set_run_font,
    set_run_font_element,
    set_zero_indent,
)

from .errors import CorruptDocxError, InvalidFileTypeError
from .paths import _is_docx_path
from .table_align import analyze_table, table_grid

# ----- 常量 -----
EMU_PER_CM = 360000

_CN_NUM = "〇零一二三四五六七八九十百千"

# 未编号章级标题：整段仅此词（内部可含空格，如「摘  要」）。
# 对应规格中不加章号、按章标题排版的结论/致谢/参考文献/摘要等。
UNNUMBERED_H1_TEXTS = {"结论", "结语", "致谢", "谢辞", "参考文献", "附录", "摘要", "abstract"}

# 「N.M 标题」/「N.M.K 标题」体系。两道防线避免把正文小数误判为标题：
# 1) 编号后紧跟谓词/量词（是/为/约/倍…）→ 是小数作主语的句子，不是标题；
# 2) 余文含句末标点（。！？…）→ 是完整句子。
# 另要求编号后不能再跟数字/点（排除 1.1.1、3.14）。
_DOTTED_GUARD = r"(?![.\d])(?!\s*(?:是|不是|为|等于|约|超过|接近|倍))(?![\s\S]*[。！？!?…])"

HEADING_PATTERNS = {
    1: re.compile(r"^[一二三四五六七八九十百]+、|^第[" + _CN_NUM + r"\d]+章(?=\s|$)"),
    2: re.compile(r"^（[一二三四五六七八九十百]+）|^\([一二三四五六七八九十百]+\)"
                  r"|^\d+\.\d+" + _DOTTED_GUARD),
    3: re.compile(r"^\d+[.、](?!\d)|^\d+\.\d+\.\d+" + _DOTTED_GUARD),
    4: re.compile(r"^（\d+）|^\(\d+\)|^\d+\.\d+\.\d+\.\d+(?![.\d])"),
}


# ----- 段落/标题检测 -----
def _paragraph_text(para) -> str:
    return (para.text or "").strip()


def _shorten(text, limit=24) -> str:
    clean = " ".join((text or "").split())
    return clean if len(clean) <= limit else clean[:limit] + "…"


def _detect_heading_level(para) -> int:
    style_name = ""
    try:
        style_name = para.style.name or ""
    except (AttributeError, KeyError, ValueError):
        style_name = ""
    match = re.search(r"Heading\s+([1-4])", style_name, re.I)
    if match:
        return int(match.group(1))

    text = _paragraph_text(para)
    # 未编号章级标题（结论/致谢/参考文献/摘要…，允许「摘  要」内部空格）
    compact = re.sub(r"\s+", "", text)
    if compact and compact.lower() in UNNUMBERED_H1_TEXTS:
        return 1
    for level, pattern in HEADING_PATTERNS.items():
        if pattern.search(text):
            return level
    return 0


def _is_toc_style(para) -> bool:
    """检测段落样式是否为目录条目样式（toc 1/2/3...）。"""
    try:
        name = para.style.name or ""
    except (AttributeError, KeyError, ValueError):
        return False
    return name.lower().startswith("toc")


def _build_caption_pattern(prefixes):
    """题注匹配规则：前缀后紧跟编号（如图1、表 2、Figure3），避免误判正文。"""
    alts = "|".join(re.escape(p) for p in prefixes if p)
    if not alts:
        return None
    return re.compile(r"^(?:" + alts + r")\s*\d+")


def _is_caption_para(text, pattern) -> bool:
    if pattern is None:
        return False
    return bool(pattern.match((text or "").strip()))


def _is_reference_title(text, level) -> bool:
    """参考文献区段的起始标题：单独的"参考文献/参考资料"，或含该词的标题段。"""
    clean = (text or "").strip().rstrip(":：")
    if clean in ("参考文献", "参考资料"):
        return True
    return bool(level) and ("参考文献" in clean or "参考资料" in clean)


def _paragraph_has_drawing(para) -> bool:
    return bool(para._element.findall(".//" + qn("w:drawing")))


def _is_blank_paragraph(para) -> bool:
    if (para.text or "").strip():
        return False
    if para.runs:
        return False
    return not _paragraph_has_drawing(para)


def _find_toc_title_index(doc):
    """找到首个 TOC 条目段落，向前回溯最近的非空非 TOC 段落索引（即"目录"标题段）。"""
    paragraphs = doc.paragraphs
    first_toc = None
    for i, p in enumerate(paragraphs):
        if _is_toc_style(p):
            first_toc = i
            break
    if first_toc is None:
        return None
    for j in range(first_toc - 1, -1, -1):
        p = paragraphs[j]
        if _is_toc_style(p):
            continue
        if _paragraph_text(p):
            return j
    return None


# ----- Section 辅助 -----
def _section_has_page_number(section) -> bool:
    """检测 section 的页脚（含继承的页脚定义）中是否已存在 PAGE 域。"""
    for instr in section.footer._element.iter(qn("w:instrText")):
        if instr.text and "PAGE" in instr.text:
            return True
    return False


def _find_page_number_paragraph(section):
    """返回页脚（含继承）中承载 PAGE 域的段落，找不到返回 None。"""
    for para in section.footer.paragraphs:
        for instr in para._element.iter(qn("w:instrText")):
            if instr.text and "PAGE" in instr.text:
                return para
    return None


def _build_body_section_maps(doc):
    """遍历 body 子元素，构建段落索引→section 索引、表格索引→section 索引映射。

    OOXML 中，分节符以段落 pPr 内的 sectPr 标记该段为当前节的最后一段；
    body 末尾的 sectPr 为最后一节的属性。
    """
    para_to_section = {}
    table_to_section = {}
    p_idx = t_idx = sec_idx = 0
    for child in doc.element.body.iterchildren():
        tag = child.tag.split("}", 1)[-1]
        if tag == "p":
            para_to_section[p_idx] = sec_idx
            ppr = child.find(qn("w:pPr"))
            if ppr is not None and ppr.find(qn("w:sectPr")) is not None:
                sec_idx += 1
            p_idx += 1
        elif tag == "tbl":
            table_to_section[t_idx] = sec_idx
            t_idx += 1
    return para_to_section, table_to_section


def _collect_cover_section_indices(doc) -> set:
    """返回封面/前置页 section 索引集合（需保留原版式）。

    判定：从首个 section 起连续无页码的 section 视为封面/前置页，直到遇到首个
    有页码的 section 为止；若全文均无页码，则不构成"封面/正文"区分，返回空集
    （按普通文档处理，正常规范化全部内容）。
    """
    has_pn = [_section_has_page_number(sec) for sec in doc.sections]
    if not any(has_pn):
        return set()
    cover = set()
    for i, has in enumerate(has_pn):
        if has:
            break
        cover.add(i)
    return cover


# ----- 样式应用 -----
def _apply_run_defaults(para, zh_font, en_font, size_pt, bold=None, color=None):
    runs = para.runs
    if not runs and para.text:
        runs = [para.add_run()]
    for run in runs:
        set_run_font(run, zh_font, en_font, size_pt=size_pt, bold=bold, color=color)


def _apply_page_settings(doc, cfg):
    page = cfg.get("page", {})
    margins = page.get("margins", {})
    page_w, page_h = resolve_page_size(cfg)
    pn = cfg.get("pageNumber", {})
    only_existing = pn.get("onlyUpdateExisting", True)

    for sec in doc.sections:
        sec.page_width = Cm(page_w)
        sec.page_height = Cm(page_h)
        if "top" in margins:
            sec.top_margin = Cm(margins["top"])
        if "bottom" in margins:
            sec.bottom_margin = Cm(margins["bottom"])
        if "left" in margins:
            sec.left_margin = Cm(margins["left"])
        if "right" in margins:
            sec.right_margin = Cm(margins["right"])
        if "header" in margins:
            sec.header_distance = Cm(margins["header"])
        if "footer" in margins:
            sec.footer_distance = Cm(margins["footer"])

        if not pn.get("enabled", True):
            continue
        if only_existing and not _section_has_page_number(sec):
            continue

        footer = sec.footer
        # 定位到真正承载 PAGE 域的段落（可能是继承来的页脚定义），
        # 避免清空页脚首段导致内容丢失或出现两个页码。
        p = _find_page_number_paragraph(sec)
        if p is None:
            blanks = [q for q in footer.paragraphs if _is_blank_paragraph(q)]
            p = blanks[0] if blanks else footer.add_paragraph()
        p.clear()
        p.alignment = resolve_align(pn.get("align", "center"))
        set_para_spacing(p, line_spacing_mult=1.0, before_pt=0, after_pt=0)
        set_zero_indent(p)
        add_page_number_field(p, lambda r: set_run_font(
            r, pn.get("font", "宋体"), pn.get("asciiFont", "Times New Roman"),
            size_pt=resolve_size(pn.get("size", "小五"))))


def _format_headers(doc, cfg, cover_sections):
    """格式化已有页眉（不创建/不改文字）：宋体小五居中，从正文开始。

    仅当配置含 header 区段时生效（学位论文预设启用；默认配置无此段，行为不变）。
    """
    h = cfg.get("header")
    if not isinstance(h, dict):
        return
    size_pt = resolve_size(h.get("size", "小五"))
    for idx, sec in enumerate(doc.sections):
        if idx in cover_sections:  # 页眉从正文开始，封面/前置节保持原样
            continue
        try:
            if sec.header.is_linked_to_previous:
                continue
        except (AttributeError, ValueError):
            continue
        for para in sec.header.paragraphs:
            if not _paragraph_text(para):
                continue
            para.alignment = resolve_align(h.get("align", "center"))
            set_para_spacing(para, line_spacing_mult=1.0, before_pt=0, after_pt=0)
            set_zero_indent(para)
            _apply_run_defaults(para, h.get("font", "宋体"),
                                h.get("asciiFont", "Times New Roman"), size_pt)


def _apply_document_styles(doc, cfg):
    normal = doc.styles["Normal"]
    body = cfg.get("body", {})
    set_run_font_element(normal.element, body.get("font", "宋体"), body.get("asciiFont", "Times New Roman"),
                         resolve_size(body.get("size", "小四")), body.get("color", "000000"))
    normal.font.size = Pt(resolve_size(body.get("size", "小四")))
    npf = normal.paragraph_format
    npf.line_spacing = 1.0
    npf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
    npf.first_line_indent = Pt(0)
    npf.space_before = Pt(0)
    npf.space_after = Pt(0)
    # 关键：清掉 Normal 样式 <w:ind> 里的字符单位属性（firstLineChars 等），
    # 否则单元格段落继承后会按字符数渲染首行缩进。
    clear_style_indent_chars(normal)

    for lv in range(1, 5):
        hc = cfg.get("headings", {}).get("h" + str(lv))
        if not hc:
            continue
        try:
            style = doc.styles["Heading " + str(lv)]
        except KeyError:
            continue
        set_run_font_element(style.element, hc.get("font", "黑体"), hc.get("asciiFont", "Times New Roman"),
                             resolve_size(hc.get("size", "小四")), hc.get("color", "000000"))
        style.font.size = Pt(resolve_size(hc.get("size", "小四")))
        style.font.bold = hc.get("bold", True)
        pf = style.paragraph_format
        pf.line_spacing = hc.get("lineSpacing", 1.5)
        pf.line_spacing_rule = WD_LINE_SPACING.MULTIPLE
        pf.space_before = Pt(hc.get("spaceBeforePt", 0))
        pf.space_after = Pt(hc.get("spaceAfterPt", 0))


def _apply_heading_format(para, cfg, level):
    hc = cfg["headings"].get("h" + str(level), {})
    try:
        para.style = "Heading " + str(level)
    except (KeyError, ValueError, AttributeError):
        pass
    para.alignment = resolve_align(hc.get("align", "left"))
    set_para_spacing(para, line_spacing_mult=hc.get("lineSpacing", 1.5),
                     before_pt=hc.get("spaceBeforePt", 0), after_pt=hc.get("spaceAfterPt", 0),
                     before_lines=hc.get("spaceBeforeLines"), after_lines=hc.get("spaceAfterLines"))
    set_zero_indent(para)
    _apply_run_defaults(
        para,
        hc.get("font", "黑体"),
        hc.get("asciiFont", cfg["body"].get("asciiFont", "Times New Roman")),
        resolve_size(hc.get("size", cfg["body"].get("size", "小四"))),
        bold=hc.get("bold", True),
        color=hc.get("color", "000000"),
    )


def _apply_body_format(para, cfg):
    body = cfg["body"]
    para.alignment = resolve_align(body.get("align", "justify"))
    set_para_spacing(para, line_spacing_mult=body.get("lineSpacing", 1.5),
                     before_pt=0, after_pt=body.get("spaceAfterPt", 0))
    indent_chars = body.get("firstLineIndentChars", 0)
    first_line_pt = resolve_size(body.get("size", "小四")) * indent_chars if indent_chars > 0 else 0
    set_indent_values(para, first_line_pt=first_line_pt, left_pt=0, right_pt=0)
    _apply_run_defaults(
        para,
        body.get("font", "宋体"),
        body.get("asciiFont", "Times New Roman"),
        resolve_size(body.get("size", "小四")),
        color=body.get("color", "000000"),
    )


def _apply_caption_format(para, cfg, kind):
    c = cfg[kind]
    para.alignment = resolve_align(c.get("captionAlign", "center"))
    set_para_spacing(para, line_spacing_mult=c.get("captionLineSpacing", 1.0), before_pt=0, after_pt=2)
    set_zero_indent(para)
    _apply_run_defaults(
        para,
        c.get("captionText", cfg["body"].get("font", "宋体")),
        c.get("captionAsciiFont", cfg["body"].get("asciiFont", "Times New Roman")),
        resolve_size(c.get("captionSize", "五号")),
        bold=c.get("captionBold", True),
    )


def _apply_reference_format(para, cfg):
    r = cfg.get("references", {})
    para.alignment = resolve_align(r.get("align", "justify"))
    set_para_spacing(para, line_spacing_mult=r.get("lineSpacing", 1.0), before_pt=0, after_pt=3)
    hang_pt = r.get("hangingIndentPt", 24)
    para.paragraph_format.left_indent = Pt(hang_pt)
    para.paragraph_format.first_line_indent = Pt(-hang_pt)
    _apply_run_defaults(
        para,
        r.get("font", "宋体"),
        r.get("asciiFont", "Times New Roman"),
        resolve_size(r.get("size", "小五")),
    )


# ----- 规范化主流程 -----
def _classify_paragraph(text, level, fig_re, table_re, in_refs):
    if level:
        return "heading"
    if _is_caption_para(text, fig_re):
        return "figure_caption"
    if _is_caption_para(text, table_re):
        return "table_caption"
    if in_refs and re.match(r"^\[\d+\]", text):
        return "reference"
    return "body"


def _handle_drawing_paragraph(para, cfg, fig_keep, fig_caption_below):
    fig_cfg = cfg.get("figure", {})
    para.alignment = resolve_align(fig_cfg.get("align", "center"))
    set_para_spacing(para, line_spacing_mult=1.0, before_pt=0, after_pt=0)
    set_zero_indent(para)
    if fig_keep and fig_caption_below:
        _set_keep_with_next(para)
    return 1


def _fingerprint(*elements):
    """元素序列化指纹——用于幂等变更检测（第二次规范化不再重复上报）。"""
    return b"".join(lxml.etree.tostring(el) for el in elements if el is not None)


def _normalize_paragraphs(doc, cfg, para_to_section, cover_sections, on_warning=None, on_change=None):
    fig_cfg = cfg.get("figure", {})
    fig_prefix = fig_cfg.get("prefix", "图")
    table_prefix = cfg.get("table", {}).get("prefix", "表")
    fig_caption_re = _build_caption_pattern((fig_prefix, "图", "Figure", "Fig"))
    table_caption_re = _build_caption_pattern((table_prefix, "表", "Table"))
    # 图与图题防分页：图题在下则图后紧跟，图题在上则图题后紧跟
    fig_keep = bool(fig_cfg.get("keepWithCaption", True))
    fig_caption_below = fig_cfg.get("captionPosition", "below") != "above"
    in_refs = False
    toc_title_index = _find_toc_title_index(doc)
    processed = 0

    for idx, para in enumerate(doc.paragraphs):
        if _is_toc_style(para):
            continue
        if toc_title_index is not None and idx == toc_title_index:
            continue
        sec_idx = para_to_section.get(idx)
        if sec_idx is not None and sec_idx in cover_sections:
            continue

        if _paragraph_has_drawing(para):
            processed += _handle_drawing_paragraph(
                para, cfg, fig_keep, fig_caption_below)
            continue

        text = _paragraph_text(para)
        if not text:
            set_zero_indent(para)
            set_para_spacing(para, line_spacing_mult=1.0, before_pt=0, after_pt=0)
            continue

        level = _detect_heading_level(para)
        if _is_reference_title(text, level):
            in_refs = True

        note = None
        kind = _classify_paragraph(text, level, fig_caption_re, table_caption_re, in_refs)
        before = _fingerprint(para._element) if on_change else None
        if kind == "heading":
            _apply_heading_format(para, cfg, level)
            processed += 1
            note = f"段落 {idx}：{level} 级标题「{_shorten(text)}」"
        elif kind == "figure_caption":
            _apply_caption_format(para, cfg, "figure")
            if fig_keep and not fig_caption_below:
                _set_keep_with_next(para)
            processed += 1
            note = f"段落 {idx}：图题注「{_shorten(text)}」"
        elif kind == "table_caption":
            _apply_caption_format(para, cfg, "table")
            processed += 1
            note = f"段落 {idx}：表题注「{_shorten(text)}」"
        elif kind == "reference":
            _apply_reference_format(para, cfg)
            processed += 1
            note = f"段落 {idx}：参考文献条目「{_shorten(text)}」"
        else:
            _apply_body_format(para, cfg)
            processed += 1
        # 幂等：只有真的改了 XML 才计入「变更」
        if note and on_change and _fingerprint(para._element) != before:
            on_change(note)
    return processed


def _table_style_element(tbl):
    """表格所引用的表格样式元素（用于级联修改与指纹比对），无则返回 None。"""
    try:
        style = tbl.style
    except (KeyError, ValueError):
        return None
    return style.element if style is not None else None


def _zero_style_indents(style_element):
    """样式 <w:ind>：删字符单位属性（firstLineChars 等），数值单位清零。

    段落级 set_zero_indent 只压制直接格式；样式里的 firstLineChars 仍会按字符数
    渲染（HANDOVER 坑 1），单元格引用到的样式必须一并清零。
    """
    pPr = style_element.find(qn("w:pPr"))
    if pPr is None:
        return
    ind = pPr.find(qn("w:ind"))
    if ind is None:
        return
    for attr in ("firstLineChars", "leftChars", "rightChars"):
        key = qn("w:" + attr)
        if key in ind.attrib:
            del ind.attrib[key]
    for attr in ("firstLine", "left", "right"):
        key = qn("w:" + attr)
        if key in ind.attrib:
            ind.set(key, "0")


def _style_by_id(doc, style_id):
    for style in doc.styles:
        if style.style_id == style_id:
            return style
    return None


def _clear_table_style_indents(doc, tbl):
    """单元格段落引用到的样式（含 basedOn 链）：清除缩进，样式级兜底。"""
    seen = set()
    for p in tbl._tbl.iter(qn("w:p")):
        pPr = p.find(qn("w:pPr"))
        if pPr is None:
            continue
        ps = pPr.find(qn("w:pStyle"))
        if ps is None:
            continue
        style_id = ps.get(qn("w:val"))
        if not style_id or style_id in seen:
            continue
        seen.add(style_id)
        style = _style_by_id(doc, style_id)
        visited = set()
        while style is not None and id(style.element) not in visited:
            visited.add(id(style.element))
            _zero_style_indents(style.element)
            try:
                style = style.base_style
            except (AttributeError, KeyError, ValueError):
                break


_BORDER_SIDE_TAGS = tuple(qn("w:" + side) for side in (
    "top", "left", "bottom", "right", "insideH", "insideV", "start", "end"))


def _strip_element_colors(el, border_color):
    """底纹清除 + 已有边框颜色恢复。作用于表格元素或其表格样式元素。"""
    # 1) w:shd：表/行/单元格/字符底纹全部移除（保留的显式 clear 由调用方补写）
    for shd in list(el.iter(qn("w:shd"))):
        parent = shd.getparent()
        if parent is None:
            continue
        if parent.tag in (qn("w:tblPr"), qn("w:trPr"), qn("w:tcPr"), qn("w:rPr")):
            parent.remove(shd)
    # 2) 已有边框（含段落/单元格/表级）颜色统一恢复为配置色
    for side in list(el.iter()):
        if side.tag in _BORDER_SIDE_TAGS:
            if side.get(qn("w:color")):
                side.set(qn("w:color"), border_color)


def _force_clear_cell_shading(tbl):
    """每个单元格显式写入无填充，压过表格样式自带的底色。"""
    for tc in tbl._tbl.iter(qn("w:tc")):
        tcPr = tc.find(qn("w:tcPr"))
        if tcPr is None:
            tcPr = OxmlElement("w:tcPr")
            tc.insert(0, tcPr)
        for shd in tcPr.findall(qn("w:shd")):
            tcPr.remove(shd)
        shd = OxmlElement("w:shd")
        shd.set(qn("w:val"), "clear")
        shd.set(qn("w:color"), "auto")
        shd.set(qn("w:fill"), "auto")
        anchor = None
        for tag in ("w:tcW", "w:gridSpan", "w:hMerge", "w:vMerge"):
            found = tcPr.find(qn(tag))
            if found is not None:
                anchor = found
        if anchor is not None:
            anchor.addnext(shd)
        else:
            tcPr.insert(0, shd)


def _format_table_cell(tc, align, valign, cell_font, cell_en, cell_size, line_spacing, bold):
    """格式化单个物理单元格（含合并延续格），完全独立于正文样式。"""
    tcPr = tc.get_or_add_tcPr()
    vAlign = tcPr.find(qn("w:vAlign"))
    if vAlign is None:
        vAlign = OxmlElement("w:vAlign")
        tcPr.append(vAlign)
    vAlign.set(qn("w:val"), valign)
    for p_el in tc.findall(qn("w:p")):
        para = Paragraph(p_el, None)
        para.alignment = align
        set_para_spacing(para, line_spacing_mult=line_spacing, before_pt=0, after_pt=0)
        set_zero_indent(para)
        _apply_run_defaults(para, cell_font, cell_en, cell_size, bold=bold)


def _normalize_tables(doc, cfg, table_to_section, cover_sections, on_warning=None, on_change=None):
    """表格规范化：独立排版 + 底纹/彩色边框清除 + 可选智能对齐。

    单元格字体默认五号宋体（不再回退到 body 的字体字号）；首行缩进在段落级
    与样式级双重清零；cellAlign="smart" 时按 normalizer.table_align 的判定
    逐格对齐，headerBold 只加在判定出的表头行上。
    """
    t = cfg.get("table", {})
    smart = str(t.get("cellAlign", "center")).lower() == "smart"
    threshold = t.get("smartAlignThresholdChars", 6)
    cell_font = t.get("cellFont", "宋体")
    cell_en = t.get("cellAsciiFont", "Times New Roman")
    cell_size = resolve_size(t.get("cellSize", "五号"))
    cell_align = resolve_align("center" if smart else t.get("cellAlign", "center"))
    valign = "center" if t.get("cellVAlign", "center") == "center" else "top"
    border_color = t.get("borderColor", "000000")
    line_spacing = t.get("cellLineSpacing", 1.0)
    header_bold = bool(t.get("headerBold", True))
    processed = 0

    for t_idx, tbl in enumerate(doc.tables):
        sec_idx = table_to_section.get(t_idx)
        if sec_idx is not None and sec_idx in cover_sections:
            continue
        processed += 1
        # 智能对齐检测必须先于底纹/边框清除（纵表头信号 ③ 要读底纹）
        plan = analyze_table(tbl, threshold) if smart else None
        style_el = _table_style_element(tbl)
        before = _fingerprint(tbl._tbl, style_el) if on_change else None

        # —— 表格独立于正文 ——
        _clear_table_style_indents(doc, tbl)
        _strip_element_colors(tbl._tbl, border_color)
        if style_el is not None:
            _strip_element_colors(style_el, border_color)
        _force_clear_cell_shading(tbl)
        tbl.alignment = resolve_table_align(t.get("align", "center"))

        bold_rows = set(plan.header_rows) if plan is not None else ({0} if header_bold else set())
        # smart 模式复用检测时的网格（plan.grid），避免重建 lxml 代理导致坐标错位
        grid_rows = plan.grid if plan is not None else table_grid(tbl)
        for row in grid_rows:
            for info in row:
                align = (resolve_align(plan.align_for(info.row, info.col))
                         if plan is not None else cell_align)
                _format_table_cell(info.tc, align, valign, cell_font, cell_en,
                                   cell_size, line_spacing,
                                   bold=header_bold and info.row in bold_rows)

        if on_change and _fingerprint(tbl._tbl, style_el) != before:
            if plan is not None:
                on_change(f"表格 {t_idx}：{plan.summary}")
            else:
                on_change(f"表格 {t_idx}：{len(tbl.rows)}×{len(tbl.columns)} 独立排版")
    return processed


def _set_keep_with_next(para, keep=True):
    """设置/清除 ``w:keepNext`` —— 让本段与下一段留在同一页（图与图题不被分页拆开）。"""
    ppr = para._element.get_or_add_pPr()
    if keep:
        if ppr.keepNext is None:
            ppr.get_or_add_keepNext()
    else:
        ppr._remove_keepNext()


def _cm_to_emu(value):
    """厘米 → EMU；非法值或 ≤0 返回 None（表示该维度不约束）。"""
    try:
        cm = float(value)
    except (TypeError, ValueError):
        return None
    return int(round(cm * EMU_PER_CM)) if cm > 0 else None


def _inside_any(element, local_names):
    """判断元素是否位于指定局部名祖先之内（不含自身）。"""
    parent = element.getparent()
    while parent is not None:
        if parent.tag.split("}", 1)[-1] in local_names:
            return True
        parent = parent.getparent()
    return False


def _iter_picture_frames(doc):
    """产出 body 中的图片框 ``(kind, frame_el)``，kind 为 "inline" 或 "anchor"。

    ``doc.inline_shapes`` 只覆盖嵌入型图片，周围环绕方式为"四周型/紧密型"等浮动图片
    完全不在其中；这类图片必须直接读 XML 才能改尺寸。
    只产出 ``graphicData`` 内含 ``pic:pic`` 的真图片，图表 / 形状 / SmartArt 不处理。
    """
    pic_tag = qn("pic:pic")
    for drawing in doc.element.body.iter(qn("w:drawing")):
        for frame in drawing:
            kind = frame.tag.split("}", 1)[-1]
            if kind not in ("inline", "anchor"):
                continue
            if frame.find(".//" + pic_tag) is None:
                continue
            yield kind, frame


def _frame_extent(frame_el):
    """读取图片框的 ``(cx, cy)``（EMU）；缺失或非法返回 ``(None, None)``。"""
    extent = frame_el.find(qn("wp:extent"))
    if extent is None:
        return None, None
    try:
        return int(extent.get("cx")), int(extent.get("cy"))
    except (TypeError, ValueError):
        return None, None


def _set_frame_extent(frame_el, cx, cy):
    """同时写 ``wp:extent`` 与 ``pic:spPr/a:xfrm/a:ext``。

    Word 以 ``wp:extent`` 为准，WPS / LibreOffice 更依赖 ``a:xfrm/a:ext``；
    只改其中之一会让不同软件显示出的图片尺寸不一致。
    注意只能精确定位 ``a:xfrm`` 下的 ``a:ext`` —— ``a:extLst`` 内也有同名 ``a:ext``，
    无差别遍历会写坏扩展节点。
    """
    extent = frame_el.find(qn("wp:extent"))
    if extent is None:
        return False
    extent.set("cx", str(int(cx)))
    extent.set("cy", str(int(cy)))
    for xfrm in frame_el.iter(qn("a:xfrm")):
        inner = xfrm.find(qn("a:ext"))
        if inner is not None:
            inner.set("cx", str(int(cx)))
            inner.set("cy", str(int(cy)))
    return True


def _available_text_width_emu(doc):
    """各节正文可用宽度的最小值（EMU）；无法计算时返回 None。"""
    widths = []
    for sec in doc.sections:
        try:
            usable = int(sec.page_width) - int(sec.left_margin) - int(sec.right_margin)
        except TypeError:
            continue
        if usable > 0:
            widths.append(usable)
    return min(widths) if widths else None


def _compute_target_size(cx, cy, mode, max_w, max_h, fixed_w, fixed_h, keep_ratio):
    """按尺寸策略算出目标 ``(cx, cy)``；无需改动时返回 None。"""
    if mode == "fixed":
        if not keep_ratio:
            return fixed_w or cx, fixed_h or cy
        scales = [limit / cur for limit, cur in ((fixed_w, cx), (fixed_h, cy)) if limit]
        if not scales:
            return None
        scale = min(scales)
        return int(round(cx * scale)), int(round(cy * scale))

    scales = [limit / cur for limit, cur in ((max_w, cx), (max_h, cy)) if limit]
    if not scales:
        return None
    scale = min(min(scales), 1.0)  # fit 模式只缩不放
    if scale >= 1.0:
        return None
    return int(round(cx * scale)), int(round(cy * scale))


def _count_header_footer_images(doc):
    """页眉/页脚中的图片数量（这些部件不在 body 内，本工具不处理）。"""
    total = 0
    for sec in doc.sections:
        for part in (sec.header, sec.footer):
            try:
                if part.is_linked_to_previous:
                    continue
                total += len(part._element.findall(".//" + qn("w:drawing")))
            except (AttributeError, ValueError):
                continue
    return total


def _normalize_inline_images(doc, cfg, on_warning=None, on_change=None):
    """统一/缩放文档中的图片，返回**实际被调整**的图片数量。

    ``figure.sizeMode``:

    - ``"fit"``（默认）：等比缩放到 ``maxWidthCm`` × ``maxHeightCm`` 上限内，只缩不放。
    - ``"fixed"``：统一到 ``fixedWidthCm`` × ``fixedHeightCm``。
      ``keepAspectRatio=false`` 为强制拉伸（等同 Word VBA 的
      ``LockAspectRatio = msoFalse``）；``true`` 则等比适配，只填宽度即"统一宽度"。

    ``clampToTextWidth`` 为安全阀：目标宽度不超过正文可用宽度，避免图片溢出页面
    （VBA 写死厘米数时最典型的翻车点）。
    """
    figure = cfg.get("figure", {})
    mode = str(figure.get("sizeMode") or "fit").lower()
    max_w = _cm_to_emu(figure.get("maxWidthCm"))
    max_h = _cm_to_emu(figure.get("maxHeightCm"))
    fixed_w = _cm_to_emu(figure.get("fixedWidthCm"))
    fixed_h = _cm_to_emu(figure.get("fixedHeightCm"))
    keep_ratio = bool(figure.get("keepAspectRatio", False))
    clamp = bool(figure.get("clampToTextWidth", True))
    skip_in_tables = bool(figure.get("skipInTables", True))
    limit = _available_text_width_emu(doc) if clamp else None

    if mode == "fixed":
        if not fixed_w and not fixed_h:
            if on_warning:
                on_warning("图片尺寸模式为「统一」但未设置目标宽/高，已跳过", location="图片")
            return 0
        if limit and fixed_w and fixed_w > limit:
            if on_warning:
                on_warning(
                    "目标宽度 {0:.2f} cm 超过正文可用宽度 {1:.2f} cm，已按可用宽度收窄".format(
                        fixed_w / EMU_PER_CM, limit / EMU_PER_CM),
                    location="图片",
                )
            fixed_w = limit
    elif not max_w and not max_h:
        return 0
    elif limit and max_w and max_w > limit:
        max_w = limit

    hf_count = _count_header_footer_images(doc)
    if hf_count and on_warning:
        on_warning(f"页眉/页脚中有 {hf_count} 张图片，本工具不处理，需手工调整", location="图片")

    adjusted = 0
    for idx, (kind, frame) in enumerate(_iter_picture_frames(doc), start=1):
        cx, cy = _frame_extent(frame)
        if not cx or not cy or cx <= 0 or cy <= 0:
            if on_warning:
                on_warning(f"图 {idx}：尺寸无效（{cx}×{cy}），保留原样", location="图片")
            continue
        if _inside_any(frame, {"txbxContent", "txbx"}):
            if on_change:
                on_change(f"图 {idx}：位于文本框内，跳过尺寸调整")
            continue
        if skip_in_tables and _inside_any(frame, {"tc"}):
            if on_change:
                on_change(f"图 {idx}：位于表格内，跳过尺寸调整")
            continue

        target = _compute_target_size(cx, cy, mode, max_w, max_h, fixed_w, fixed_h, keep_ratio)
        if target is None or target == (cx, cy):
            continue
        new_w, new_h = target
        if not _set_frame_extent(frame, new_w, new_h):
            if on_warning:
                on_warning(f"图 {idx}：写入尺寸失败，保留原样", location="图片")
            continue
        adjusted += 1
        if on_change:
            on_change(
                "图 {0}{1}：{2:.2f}×{3:.2f} cm → {4:.2f}×{5:.2f} cm".format(
                    idx, "（浮动）" if kind == "anchor" else "",
                    cx / EMU_PER_CM, cy / EMU_PER_CM,
                    new_w / EMU_PER_CM, new_h / EMU_PER_CM,
                )
            )
    return adjusted


def _mark_toc_fields_dirty(doc):
    """将 TOC 域标记为 dirty，使 Word 打开时提示更新目录以同步页码与条目。"""
    begin_stack = []
    for el in doc.element.body.iter():
        tag = el.tag.split("}", 1)[-1]
        if tag == "fldChar":
            ftype = el.get(qn("w:fldCharType"))
            if ftype == "begin":
                begin_stack.append(el)
            elif ftype == "end" and begin_stack:
                begin_stack.pop()
        elif tag == "instrText":
            if el.text and "TOC" in el.text and begin_stack:
                begin_stack[-1].set(qn("w:dirty"), "true")


# ----- 文档检查（inspect） -----
@dataclass(frozen=True)
class DocxInspection:
    paragraph_count: int
    table_count: int
    image_count: int
    heading_levels: tuple
    has_header_footer: bool
    has_complex_objects: bool


def _header_footer_has_content(part) -> bool:
    for para in part.paragraphs:
        if (para.text or "").strip():
            return True
        if _paragraph_has_drawing(para):
            return True
    if part.tables:
        return True
    for instr in part._element.iter(qn("w:instrText")):
        if instr.text and instr.text.strip():
            return True
    return False


def _has_header_footer(doc) -> bool:
    for section in doc.sections:
        for part in (section.header, section.footer):
            if part.is_linked_to_previous:
                continue
            if _header_footer_has_content(part):
                return True
    return False


def _has_complex_objects(doc) -> bool:
    inline_count = 0
    drawing_count = 0
    for el in doc.element.body.iter():
        tag = el.tag.split("}", 1)[-1]
        if tag in ("object", "pict", "sdt", "ins", "del"):
            return True
        if tag == "inline":
            inline_count += 1
        elif tag == "drawing":
            drawing_count += 1
    return drawing_count > inline_count


def _open_document(input_path):
    """打开 docx；损坏时抛 CorruptDocxError。"""
    try:
        return Document(input_path)
    except (zipfile.BadZipFile, lxml.etree.XMLSyntaxError) as e:
        raise CorruptDocxError(
            f"docx 文件已损坏: {input_path}",
            cause=e, hint="用 Word 打开重新保存一次",
        ) from e
    except KeyError as e:
        raise CorruptDocxError(
            f"docx 关键样式缺失: {input_path}",
            cause=e, hint="用 Word 打开重新保存一次",
        ) from e
    except Exception as e:  # noqa: BLE001
        from docx.opc.exceptions import PackageNotFoundError
        if isinstance(e, PackageNotFoundError):
            raise CorruptDocxError(
                f"docx 文件已损坏: {input_path}",
                cause=e, hint="用 Word 打开重新保存一次",
            ) from e
        raise


def inspect_docx(input_path):
    if not _is_docx_path(input_path):
        raise InvalidFileTypeError(
            f"不是有效的 docx 文件: {input_path}",
            hint="确认文件是 Word 导出的 .docx 格式（不是 .doc / .rtf / .wps）",
        )
    doc = _open_document(input_path)
    heading_levels = sorted({_detect_heading_level(p) for p in doc.paragraphs} - {0})
    return DocxInspection(
        paragraph_count=len(doc.paragraphs),
        table_count=len(doc.tables),
        image_count=sum(1 for _ in _iter_picture_frames(doc)),
        heading_levels=tuple(heading_levels),
        has_header_footer=_has_header_footer(doc),
        has_complex_objects=_has_complex_objects(doc),
    )
