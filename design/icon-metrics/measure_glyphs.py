"""图标字体原料测量（测量驱动对齐的第 1 步）。

产出两份东西到本目录：
- glyph_metrics.json   机器可读：两套字体的 OS/2·hhea 行盒、逐字形 hmtx/glyf 墨迹盒、
                       三档设计缩放（100/125/150%）×三档 token（16/20/40）的逐字形
                       Pillow 渲染墨迹数据，以及现行「token 当 points」做法的实际像素对照。
- metrics_report.md    人读摘要：行盒 vs 墨迹的差距、逐字形方差、现行做法的尺寸误差。

只读字体文件与 ui_icons.ICONS，不改任何运行时代码。字体缺失（非 Windows / 精简系统）
时对该字体记 missing，其余照常。
"""

import json
import sys
from pathlib import Path

from fontTools.pens.boundsPen import BoundsPen
from fontTools.ttLib import TTFont
from PIL import ImageFont

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

import ui_icons

OUT = Path(__file__).resolve().parent

# (字体族名, 字体文件)。SegoeIcons.ttf 是 Windows 11 上 Segoe Fluent Icons 的文件名。
FONTS = [
    ("Segoe Fluent Icons", r"C:\Windows\Fonts\SegoeIcons.ttf"),
    ("Segoe MDL2 Assets", r"C:\Windows\Fonts\segmdl2.ttf"),
]
SCALES = {"100%": 1.0, "125%": 1.25, "150%": 1.5}
TOKENS = {"icon_row": 16, "icon_group": 20, "icon_hero": 40}


def font_tables(path):
    """fontTools 层：行盒表 + 逐字形 advance/glyf 墨迹盒（字体单位）。"""
    tt = TTFont(str(path), fontNumber=0)
    head, hhea, os2 = tt["head"], tt["hhea"], tt["OS/2"]
    upm = head.unitsPerEm
    glyphset = tt.getGlyphSet()
    cmap = tt.getBestCmap()
    hmtx = tt["hmtx"]
    glyphs = {}
    for name, ch in ui_icons.ICONS.items():
        cp = ord(ch)
        gname = cmap.get(cp)
        if gname is None:
            glyphs[name] = {"cp": hex(cp), "present": False}
            continue
        pen = BoundsPen(glyphset)
        glyphset[gname].draw(pen)
        bounds = pen.bounds  # (xMin, yMin, xMax, yMax) 字体单位
        glyphs[name] = {
            "cp": hex(cp),
            "glyph_name": gname,
            "advance_units": hmtx[gname][0],
            "ink_units": list(bounds) if bounds else None,
        }
    return {
        "font_file": path.name,
        "units_per_em": upm,
        "hhea": {"ascent": hhea.ascent, "descent": hhea.descent, "lineGap": hhea.lineGap},
        "os2": {
            "usWinAscent": os2.usWinAscent,
            "usWinDescent": os2.usWinDescent,
            "sTypoAscender": os2.sTypoAscender,
            "sTypoDescender": os2.sTypoDescender,
            "sTypoLineGap": os2.sTypoLineGap,
            "fsSelection_USE_TYPO_METRICS": bool(os2.fsSelection & 1 << 7),
        },
        "line_box_px_at_16designpx": {
            scale_name: round(16 * scale * hhea.ascent / upm + 16 * scale * -hhea.descent / upm, 2)
            for scale_name, scale in SCALES.items()
        },
        "glyphs": glyphs,
    }


def render_metrics(path, ch, px):
    """Pillow 层：指定像素尺寸下逐字形墨迹盒（anchor 'la'：原点=行盒左上角）。"""
    font = ImageFont.truetype(str(path), px)
    x0, y0, x1, y1 = font.getbbox(ch)
    ascent, descent = font.getmetrics()
    return {
        "px": px,
        "pil_ascent": ascent,
        "pil_descent": descent,
        "ink_bbox_la": [x0, y0, x1, y1],
        "ink_w": x1 - x0,
        "ink_h": y1 - y0,
        "ink_center_from_origin": [(x0 + x1) / 2, (y0 + y1) / 2],
    }


def main():
    report = {"fonts": {}, "renders": {}}
    for family, font_path in FONTS:
        path = Path(font_path)
        if not path.exists():
            report["fonts"][family] = {"font_file": path.name, "missing": True}
            continue
        report["fonts"][family] = font_tables(path)
        for scale_name, scale in SCALES.items():
            for token, logical in TOKENS.items():
                key = f"{token}@{scale_name}"
                sizes = report["renders"].setdefault(key, {})
                px = round(logical * scale)
                per_font = sizes.setdefault(family, {"design_px": logical, "render_px": px})
                per_font[family + "_glyphs"] = {
                    name: render_metrics(path, ch, px) for name, ch in ui_icons.ICONS.items()
                }
                # 现行做法对照：token 被当作 tk points，Windows 上 scaling=dpi/72，
                # 设计基准应为 dpi/96，故恒放大 4/3。
                per_font["current_points_px"] = round(logical * scale * 4 / 3, 1)
    (OUT / "glyph_metrics.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")

    # ---- 人读摘要 ----
    lines = ["# 图标字体原料测量报告", ""]
    for family, data in report["fonts"].items():
        if data.get("missing"):
            lines += [f"## {family}", "", f"缺失：{data['font_file']} 不存在。", ""]
            continue
        upm = data["units_per_em"]
        hhea = data["hhea"]
        lines += [
            f"## {family}（{data['font_file']}）",
            "",
            f"- unitsPerEm={upm}，hhea ascent={hhea['ascent']} descent={hhea['descent']} "
            f"lineGap={hhea['lineGap']} → 行盒 {hhea['ascent'] - hhea['descent']} units "
            f"= {(hhea['ascent'] - hhea['descent']) / upm:.2f} em",
            f"- 16 设计px 图标的行盒高度：{data['line_box_px_at_16designpx']}",
            "",
            "| 图标 | 码点 | advance(u) | 墨迹盒(u) | 墨迹宽(em) | 墨迹高(em) |",
            "|---|---|---|---|---|---|",
        ]
        for name, g in data["glyphs"].items():
            if not g.get("present", True):
                lines.append(f"| {name} | {g['cp']} | — | 字形缺失 | — | — |")
                continue
            x0, y0, x1, y1 = g["ink_units"]
            lines.append(
                f"| {name} | {g['cp']} | {g['advance_units']} "
                f"| [{x0},{y0},{x1},{y1}] | {(x1 - x0) / upm:.3f} | {(y1 - y0) / upm:.3f} |"
            )
        lines.append("")
    lines += [
        "## 现行「token 当 points」的尺寸对照",
        "",
        "| 场景 | 设计像素 | 现行实际渲染（px） | 误差 |",
        "|---|---|---|---|",
    ]
    for key, fonts_data in report["renders"].items():
        for family, d in fonts_data.items():
            if isinstance(d, dict) and "current_points_px" in d:
                lines.append(
                    f"| {key}（{family.split()[-1]}） | {d['design_px']}→{d['render_px']} "
                    f"| {d['current_points_px']} | +33% |"
                )
    (OUT / "metrics_report.md").write_text("\n".join(lines) + "\n", encoding="utf-8")
    print(f"written: {OUT / 'glyph_metrics.json'}")
    print(f"written: {OUT / 'metrics_report.md'}")


if __name__ == "__main__":
    main()
