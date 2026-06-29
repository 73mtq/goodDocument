"""Conservative normalization for existing Word documents."""

from dataclasses import dataclass
import os
import re

from docx import Document
from docx.shared import Cm, Pt
from docx.enum.text import WD_LINE_SPACING
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_ALIGN_VERTICAL

from formatter import (
    resolve_size,
    resolve_align,
    set_run_font,
    set_para_spacing,
    set_zero_indent,
    add_page_number_field,
    set_run_font_element,
)


EMU_PER_CM = 360000

HEADING_PATTERNS = {
    1: re.compile(r"^[一二三四五六七八九十百]+、"),
    2: re.compile(r"^（[一二三四五六七八九十百]+）|^\([一二三四五六七八九十百]+\)"),
    3: re.compile(r"^\d+[.、]"),
    4: re.compile(r"^（\d+）|^\(\d+\)"),
}


@dataclass(frozen=True)
class DocxInspection:
    paragraph_count: int
    table_count: int
    image_count: int
    heading_levels: tuple
    has_header_footer: bool
    has_complex_objects: bool


def inspect_docx(input_path):
    if not _is_docx_path(input_path):
        raise ValueError("输入文件必须是 .docx Word 文档")
    doc = Document(input_path)
    heading_levels = sorted({_detect_heading_level(p) for p in doc.paragraphs} - {0})
    return DocxInspection(
        paragraph_count=len(doc.paragraphs),
        table_count=len(doc.tables),
        image_count=len(doc.inline_shapes),
        heading_levels=tuple(heading_levels),
        has_header_footer=_has_header_footer(doc),
        has_complex_objects=_has_complex_objects(doc),
    )


def normalize_docx(config, input_path, output_path, assets_dir=None):
    """Conservatively apply the configured style system to an existing .docx."""
    del assets_dir  # Kept for a stable public interface.
    input_path = os.fspath(input_path)
    output_path = os.fspath(output_path)

    if not _is_docx_path(input_path):
        raise ValueError("输入文件必须是 .docx Word 文档")
    if not os.path.exists(input_path):
        raise FileNotFoundError("找不到输入文档: " + input_path)
    if not _is_docx_path(output_path):
        raise ValueError("输出文件必须是 .docx Word 文档")
    if _same_path(input_path, output_path):
        raise ValueError("输出文件不能覆盖输入文档，请选择不同的保存路径")

    doc = Document(input_path)
    _apply_page_settings(doc, config)
    _apply_document_styles(doc, config)
    _normalize_paragraphs(doc, config)
    _normalize_tables(doc, config)
    _normalize_inline_images(doc, config)

    out_dir = os.path.dirname(os.path.abspath(output_path))
    if out_dir:
        os.makedirs(out_dir, exist_ok=True)
    doc.save(output_path)
    return output_path


def _same_path(a, b):
    try:
        return os.path.abspath(os.fspath(a)).lower() == os.path.abspath(os.fspath(b)).lower()
    except TypeError:
        return False


def _is_docx_path(path):
    return str(path).lower().endswith(".docx")


def _paragraph_text(para):
    return (para.text or "").strip()


def _detect_heading_level(para):
    style_name = ""
    try:
        style_name = para.style.name or ""
    except Exception:
        style_name = ""
    match = re.search(r"Heading\s+([1-4])", style_name, re.I)
    if match:
        return int(match.group(1))

    text = _paragraph_text(para)
    for level, pattern in HEADING_PATTERNS.items():
        if pattern.search(text):
            return level
    return 0


def _has_header_footer(doc):
    for section in doc.sections:
        if any((p.text or "").strip() for p in section.header.paragraphs):
            return True
        if section.footer.paragraphs:
            return True
    return False


def _has_complex_objects(doc):
    xml = doc._element.xml
    markers = ("<w:drawing", "<w:object", "<w:pict", "<w:sdt", "<w:ins", "<w:del")
    inline_count = xml.count("<wp:inline")
    drawing_count = xml.count("<w:drawing")
    if drawing_count > inline_count:
        return True
    return any(marker in xml for marker in markers if marker != "<w:drawing")


def _is_caption_para(text, prefixes):
    clean = (text or "").strip()
    for prefix in prefixes:
        if prefix and clean.startswith(prefix):
            return True
    return False


def _apply_run_defaults(para, zh_font, en_font, size_pt, bold=None, color=None):
    runs = para.runs
    if not runs and para.text:
        runs = [para.add_run()]
    for run in runs:
        set_run_font(run, zh_font, en_font, size_pt=size_pt, bold=bold, color=color)


def _apply_page_settings(doc, cfg):
    page = cfg.get("page", {})
    margins = page.get("margins", {})
    pn = cfg.get("pageNumber", {})

    for sec in doc.sections:
        sec.page_width = Cm(21.0)
        sec.page_height = Cm(29.7)
        if "top" in margins:
            sec.top_margin = Cm(margins["top"])
        if "bottom" in margins:
            sec.bottom_margin = Cm(margins["bottom"])
        if "left" in margins:
            sec.left_margin = Cm(margins["left"])
        if "right" in margins:
            sec.right_margin = Cm(margins["right"])

        if pn.get("enabled", True):
            footer = sec.footer
            footer.is_linked_to_previous = False
            p = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
            p.clear()
            p.alignment = resolve_align(pn.get("align", "center"))
            set_para_spacing(p, line_spacing_mult=1.0, before_pt=0, after_pt=0)
            set_zero_indent(p)
            add_page_number_field(p, lambda r: set_run_font(
                r, pn.get("font", "宋体"), pn.get("asciiFont", "Times New Roman"),
                size_pt=resolve_size(pn.get("size", "小五"))))


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
    except Exception:
        pass
    para.alignment = resolve_align(hc.get("align", "left"))
    set_para_spacing(para, line_spacing_mult=hc.get("lineSpacing", 1.5),
                     before_pt=hc.get("spaceBeforePt", 0), after_pt=hc.get("spaceAfterPt", 0))
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
    para.paragraph_format.first_line_indent = (
        Pt(resolve_size(body.get("size", "小四")) * indent_chars) if indent_chars > 0 else Pt(0)
    )
    para.paragraph_format.left_indent = Pt(0)
    para.paragraph_format.right_indent = Pt(0)
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


def _normalize_paragraphs(doc, cfg):
    fig_prefix = cfg.get("figure", {}).get("prefix", "图")
    table_prefix = cfg.get("table", {}).get("prefix", "表")
    in_refs = False

    for para in doc.paragraphs:
        text = _paragraph_text(para)
        if not text:
            set_zero_indent(para)
            set_para_spacing(para, line_spacing_mult=1.0, before_pt=0, after_pt=0)
            continue

        if "参考文献" in text or "参考资料" in text:
            in_refs = True

        level = _detect_heading_level(para)
        if level:
            _apply_heading_format(para, cfg, level)
        elif _is_caption_para(text, (fig_prefix, "图", "Figure", "Fig")):
            _apply_caption_format(para, cfg, "figure")
        elif _is_caption_para(text, (table_prefix, "表", "Table")):
            _apply_caption_format(para, cfg, "table")
        elif in_refs and re.match(r"^\[\d+\]", text):
            _apply_reference_format(para, cfg)
        else:
            _apply_body_format(para, cfg)


def _normalize_tables(doc, cfg):
    t = cfg.get("table", {})
    cell_font = t.get("cellFont", cfg["body"].get("font", "宋体"))
    cell_en = t.get("cellAsciiFont", cfg["body"].get("asciiFont", "Times New Roman"))
    cell_size = resolve_size(t.get("cellSize", cfg["body"].get("size", "小四")))
    cell_align = resolve_align(t.get("cellAlign", "center"))
    cell_valign = WD_ALIGN_VERTICAL.CENTER if t.get("cellVAlign", "center") == "center" else WD_ALIGN_VERTICAL.TOP

    for tbl in doc.tables:
        tbl.alignment = WD_TABLE_ALIGNMENT.CENTER
        for r_idx, row in enumerate(tbl.rows):
            for cell in row.cells:
                cell.vertical_alignment = cell_valign
                for para in cell.paragraphs:
                    para.alignment = cell_align
                    set_para_spacing(para, line_spacing_mult=t.get("cellLineSpacing", 1.0), before_pt=0, after_pt=0)
                    set_zero_indent(para)
                    _apply_run_defaults(
                        para,
                        cell_font,
                        cell_en,
                        cell_size,
                        bold=t.get("headerBold", True) if r_idx == 0 else False,
                    )


def _normalize_inline_images(doc, cfg):
    figure = cfg.get("figure", {})
    max_width = figure.get("maxWidthCm")
    max_height = figure.get("maxHeightCm")
    if not max_width and not max_height:
        return

    max_w_emu = int(max_width * EMU_PER_CM) if max_width else None
    max_h_emu = int(max_height * EMU_PER_CM) if max_height else None
    for shape in getattr(doc, "inline_shapes", []):
        try:
            width = int(shape.width)
            height = int(shape.height)
        except Exception:
            continue
        if width <= 0 or height <= 0:
            continue
        scale = 1.0
        if max_w_emu and width > max_w_emu:
            scale = min(scale, max_w_emu / width)
        if max_h_emu and height > max_h_emu:
            scale = min(scale, max_h_emu / height)
        if scale < 1.0:
            shape.width = int(width * scale)
            shape.height = int(height * scale)
