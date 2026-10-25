"""测试 fixture 辅助：构造特殊 docx 用于边界场景测试。"""
import atexit
import os
import re
import sys
import tempfile
import zipfile
from pathlib import Path

from docx import Document
from docx.shared import Cm

ROOT = Path(__file__).resolve().parents[1]
FIGURE_PNG = ROOT / "assets" / "figure1.png"

# wp:anchor 的开头部分：simplePos / positionH / positionV 必须排在 wp:extent 之前
_ANCHOR_ATTRS = ('simplePos="0" relativeHeight="251658240" behindDoc="0" locked="0" '
                 'layoutInCell="1" allowOverlap="1"')
_ANCHOR_HEAD = ('<wp:simplePos x="0" y="0"/>'
                '<wp:positionH relativeFrom="column"><wp:posOffset>0</wp:posOffset></wp:positionH>'
                '<wp:positionV relativeFrom="paragraph"><wp:posOffset>0</wp:posOffset></wp:positionV>')


def _rewrite_zip_member(src, dst, member, transform):
    """把 zip 中某个成员的文本内容经 transform 处理后写回新 zip。"""
    with zipfile.ZipFile(src) as zin:
        entries = [(info, zin.read(info.filename)) for info in zin.infolist()]
    with zipfile.ZipFile(dst, "w", zipfile.ZIP_DEFLATED) as zout:
        for info, data in entries:
            if info.filename == member:
                data = transform(data.decode("utf-8")).encode("utf-8")
            zout.writestr(info, data)


def make_docx_with_floating_image(path, width_cm=18, height_cm=12):
    """构造含**浮动图片**（wp:anchor）的 docx。

    python-docx 只能生成嵌入型（wp:inline）图片，因此保存后把 document.xml 中的
    ``<wp:inline>`` 改写成 ``<wp:anchor>``，复现 Word 里"四周型/浮于文字上方"的图片。
    这类图片不在 ``doc.inline_shapes`` 中，是最容易漏掉的一类。
    """
    tmp = Path(tempfile.mkdtemp(prefix="gooddocument-float-")) / "inline.docx"
    doc = Document()
    doc.add_paragraph("正文段落")
    para = doc.add_paragraph()
    para.add_run().add_picture(str(FIGURE_PNG), width=Cm(width_cm), height=Cm(height_cm))
    doc.save(str(tmp))

    def to_anchor(xml):
        # 保留原 <wp:inline> 上的命名空间声明（xmlns:a / xmlns:pic / …），否则解析直接失败
        xml = re.sub(
            r"<wp:inline(?P<attrs>[^>]*)>",
            lambda m: "<wp:anchor " + _ANCHOR_ATTRS + m.group("attrs") + ">" + _ANCHOR_HEAD,
            xml,
        )
        # wrapNone 需位于 wp:extent 之后、wp:docPr 之前
        xml = re.sub(r"<wp:docPr", "<wp:wrapNone/><wp:docPr", xml)
        return xml.replace("</wp:inline>", "</wp:anchor>")

    _rewrite_zip_member(str(tmp), str(path), "word/document.xml", to_anchor)
    return path


def make_docx_with_images(path, specs, caption_position="below", table_image=False):
    """构造含多张图片的 docx。

    specs 为 ``[(宽cm, 高cm), ...]``；每张图后面跟一条图题段落。
    table_image=True 时额外在表格单元格内放一张 2×1.5 cm 的小图。
    """
    doc = Document()
    doc.add_paragraph("正文段落")
    for index, (width_cm, height_cm) in enumerate(specs, start=1):
        para = doc.add_paragraph()
        para.add_run().add_picture(str(FIGURE_PNG), width=Cm(width_cm), height=Cm(height_cm))
        caption = "图 %d  示例" % index
        if caption_position == "above":
            body = doc.paragraphs[-1]
            body._element.addprevious(doc.add_paragraph(caption)._element)
        else:
            doc.add_paragraph(caption)
    if table_image:
        table = doc.add_table(rows=1, cols=1)
        table.cell(0, 0).paragraphs[0].add_run().add_picture(
            str(FIGURE_PNG), width=Cm(2), height=Cm(1.5))
    doc.save(str(path))
    return path


def picture_sizes(path):
    """读取 docx 中每个图片框的 ``(kind, cx_emu, cy_emu, xfrm_cx_emu)``。"""
    from docx.oxml.ns import qn

    doc = Document(str(path))
    result = []
    for drawing in doc.element.body.iter(qn("w:drawing")):
        for frame in drawing:
            kind = frame.tag.split("}")[-1]
            if kind not in ("inline", "anchor"):
                continue
            if frame.find(".//" + qn("pic:pic")) is None:
                continue
            extent = frame.find(qn("wp:extent"))
            xfrm_cx = None
            for xfrm in frame.iter(qn("a:xfrm")):
                inner = xfrm.find(qn("a:ext"))
                if inner is not None:
                    xfrm_cx = int(inner.get("cx"))
            result.append((kind, int(extent.get("cx")), int(extent.get("cy")), xfrm_cx))
    return result


def make_garbage_docx(path):
    """写入垃圾内容到 .docx 文件（不是 zip）。"""
    with open(path, "wb") as f:
        f.write(b"this is not a real docx, just garbage bytes")


def make_corrupt_docx(path):
    """截断正常 docx 的后半部分（zip 中央目录损坏）。"""
    with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
        src = f.name
    try:
        doc = Document()
        doc.add_paragraph("test")
        doc.save(src)
        data = Path(src).read_bytes()
        Path(path).write_bytes(data[: max(100, len(data) // 3)])
    finally:
        os.remove(src)


def make_docx_with_nested_table(path):
    """构造含嵌套表的 docx。"""
    doc = Document()
    doc.add_paragraph("外层段落")
    tbl = doc.add_table(rows=2, cols=2)
    inner_cell = tbl.rows[0].cells[0]
    inner_cell.add_paragraph("嵌套前")
    inner_tbl = inner_cell.add_table(rows=1, cols=1)
    inner_tbl.rows[0].cells[0].text = "嵌套内容"
    inner_cell.add_paragraph("嵌套后")
    doc.save(path)


def make_docx_with_empty_table(path):
    """构造含空单元格的 docx（1 行 1 列但无文字）。"""
    doc = Document()
    doc.add_paragraph("前")
    doc.add_table(rows=1, cols=1)
    doc.add_paragraph("后")
    doc.save(path)


def make_docx_with_unrecognized_paragraph(path):
    """构造含无法识别段落的 docx。"""
    doc = Document()
    doc.add_paragraph("正常正文段落")
    doc.add_paragraph("12345")  # 纯数字，不匹配任何标题/正文模式
    doc.save(path)


_TK_ROOT = None


def tk_root():
    """进程级单例 Tk root（withdraw 不销毁）。

    Tcl 解释器在「全部 root 销毁后重建」时部分环境不可靠；且进程 DPI 上下文
    随首个窗口锁定，故首次创建前先开 DPI 感知。失败（无显示环境）返回 None。
    """
    global _TK_ROOT
    if _TK_ROOT is not None and bool(_TK_ROOT.winfo_exists()):
        return _TK_ROOT
        return _TK_ROOT
    try:
        import tkinter as tk

        try:
            sys_path = str(ROOT / "design" / "icon-metrics")
            if sys_path not in sys.path:
                sys.path.insert(0, sys_path)
            from inklab import _make_dpi_aware

            _make_dpi_aware()
        except ImportError:
            pass
        _TK_ROOT = tk.Tk()
        _TK_ROOT.withdraw()
        atexit.register(lambda: _TK_ROOT.winfo_exists() and _TK_ROOT.destroy())
        return _TK_ROOT
    except Exception as exc:  # noqa: BLE001
        print(f"tk_root 失败：{exc!r}", flush=True)
        return None
