r"""生成 icon_slot_data.py：字体单位墨迹表（fontTools）+ grid-fit 校准表。

墨迹表：两套图标字体逐字形的 glyf 墨迹盒（字体单位），来自 C:\Windows\Fonts。
校准表：若 calibration_report.json 存在（design/icon-metrics/calibrate.py 产出），
把 |分量残差| ≥ CORR_THRESHOLD 的确定性 grid-fit 离群写入 CORRECTIONS；
否则 CORRECTIONS 置空（新环境先跑 calibrate.py）。

输出是仓库根的 icon_slot_data.py，自动生成勿手改；tests/test_icon_slot.py
含再生比对守卫。字体文件缺失（非 Windows/精简系统）时该字体记 None，
运行时由 ui_icons.select_icon_family 的检测结果自然走纯文字回退。
"""
import json
import sys
from pathlib import Path

from fontTools.pens.boundsPen import BoundsPen
from fontTools.ttLib import TTFont

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

import ui_icons

OUT = ROOT / "icon_slot_data.py"
HERE = Path(__file__).resolve().parent
CORR_THRESHOLD = 0.75  # 分量残差 ≥ 此值判为确定性 grid-fit 离群（见 DECISION.md）

FONTS = {
    "Segoe Fluent Icons": r"C:\Windows\Fonts\SegoeIcons.ttf",
    "Segoe MDL2 Assets": r"C:\Windows\Fonts\segmdl2.ttf",
}


def read_ink(font_path):
    tt = TTFont(str(font_path), fontNumber=0)
    glyphset = tt.getGlyphSet()
    cmap = tt.getBestCmap()
    glyphs = {}
    for name, ch in ui_icons.ICONS.items():
        gname = cmap.get(ord(ch))
        if gname is None:
            raise SystemExit(f"{font_path.name} 缺字形 {name}，码点表需与微软字形表核对")
        pen = BoundsPen(glyphset)
        glyphset[gname].draw(pen)
        if pen.bounds is None:
            raise SystemExit(f"{font_path.name} 的 {name} 墨迹为空（空白字形？）")
        glyphs[name] = [int(v) for v in pen.bounds]
    return {"upm": int(tt["head"].unitsPerEm), "glyphs": glyphs}


def load_corrections():
    report = HERE / "calibration_report.json"
    if not report.exists():
        return {}
    data = json.loads(report.read_text(encoding="utf-8"))
    corrections = {}
    for family, glyphs in data.get("residuals", {}).items():
        for name, sizes in glyphs.items():
            for ppem, (dx, dy) in sizes.items():
                if abs(dx) >= CORR_THRESHOLD or abs(dy) >= CORR_THRESHOLD:
                    corrections.setdefault(family, {}).setdefault(name, {})[str(ppem)] = [
                        round(dx, 2), round(dy, 2)]
    return corrections


def main():
    ink = {}
    for family, path in FONTS.items():
        if Path(path).exists():
            ink[family] = read_ink(path)
        else:
            print(f"警告：{path} 不存在，{family} 记 None（运行时纯文字回退）")
            ink[family] = None
    corrections = load_corrections()
    n_corr = sum(len(v) for f in corrections.values() for v in f.values())

    lines = [
        '"""自动生成（design/icon-metrics/build_icon_data.py），勿手改。',
        '',
        'INK：两套图标字体逐字形 glyf 墨迹盒（字体单位）；某字体缺失记 None。',
        'CORRECTIONS：GDI grid-fit 确定性离群校准（分量残差 ≥ '
        f'{CORR_THRESHOLD}px，来源 calibration_report.json；当前 {n_corr} 条）。',
        ' regeneration 见 design/icon-metrics/DECISION.md。',
        '"""',
        "",
        "INK = " + json.dumps(ink, ensure_ascii=False, indent=1).replace("null", "None"),
        "",
        "CORRECTIONS = " + json.dumps(corrections, ensure_ascii=False, indent=1),
        "",
        "",
        "def correction(family, name, size_px):",
        '    """该 (字体, 字形, ppem) 的 grid-fit 校准 [dx, dy]；无则 (0.0, 0.0)。"""',
        "    try:",
        "        dx, dy = CORRECTIONS[family][name][str(size_px)]",
        "        return dx, dy",
        "    except KeyError:",
        "        return 0.0, 0.0",
        "",
    ]
    OUT.write_text("\n".join(lines), encoding="utf-8")
    print(f"written: {OUT}（INK {len(ink)} 字体，CORRECTIONS {n_corr} 条）")


if __name__ == "__main__":
    main()
