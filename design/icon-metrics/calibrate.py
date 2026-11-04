"""全覆盖 grid-fit 校准：22 图标 × 8 档 ppem × 可用字体，量原始残差。

产出到本目录：
- calibration_report.json   逐格原始残差（未叠加校准），供 build_icon_data.py 生成
                            CORRECTIONS（|分量| ≥ 0.75px 判确定性离群）；
- calibration_screenshot.png 全覆盖渲染存档。

流程：build_icon_data.py（先出墨迹表、CORRECTIONS 空）→ 本脚本 → 再跑
build_icon_data.py（并入校准）→ tests/test_icon_alignment.py 验收。
注意：本脚本测「原始残差」，渲染前强制清空 CORRECTIONS，避免校准振荡。
"""

import json
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))
sys.path.insert(0, str(Path(__file__).resolve().parent))

import tkinter as tk

from build_icon_data import FONTS
from inklab import run_sheet

import icon_slot_data
import ui_icons

# (token, ppem) 对：3 个设计 token（16/20/40）× 3 档设计缩放（100/125/150%）
TOKEN_PPems = [(token, round(token * scale)) for token in (16, 20, 40) for scale in (1.0, 1.25, 1.5)]


def tol_for(token, ppem):
    """分量公差：0.5px@100% → 1.0px@150% 线性（DECISION.md）。"""
    return round(ppem / token - 0.5, 3)


def main():
    from inklab import _make_dpi_aware

    import icon_slot

    _make_dpi_aware()  # 必须在首个 Tk root 之前：进程 DPI 上下文随首个窗口锁定
    icon_slot_data.CORRECTIONS = {}  # 测原始残差：清空校准（防振荡），只影响本进程
    root = tk.Tk()
    root.withdraw()
    renderer = icon_slot.IconRenderer(root, None)
    root.destroy()

    rows = []
    tol_map = {}
    for family, font_path in FONTS.items():
        if not Path(font_path).exists():
            print(f"跳过缺失字体：{family}")
            continue
        renderer.family = family
        for token, ppem in TOKEN_PPems:
            row = []
            for name in ui_icons.ICONS:
                placement = renderer.text_placement(name, ppem)
                if placement is None:
                    raise SystemExit(f"{family} 的 {name}@{ppem}px 无摆位（墨迹表缺？）")
                x, y, glyph = placement
                row.append({"name": name, "family": family, "glyph": glyph, "size": ppem, "tx": x, "ty": y})
                tol_map[f"{family}/{name}/{ppem}"] = tol_for(token, ppem)
            rows.append(row)

    here = Path(__file__).resolve().parent
    results, _shot = run_sheet(
        rows,
        here / "calibration_screenshot.png",
        row_header=lambda c: f"{c['family']}  {c['size']}px",
        cell_label=False,
    )

    residuals = {}
    report_rows = []
    n_fail = 0
    for r in results:
        if "error" in r:
            print(f"PROBE-FAIL {r['cell']['family']} {r['cell']['name']}: {r['error']}")
            n_fail += 1
            continue
        c = r["cell"]
        key = f"{c['family']}/{c['name']}/{c['size']}"
        tol = tol_map[key]
        passed = abs(r["err_x"]) <= tol and abs(r["err_y"]) <= tol
        if not passed:
            n_fail += 1
        residuals.setdefault(c["family"], {}).setdefault(c["name"], {})[str(c["size"])] = [r["err_x"], r["err_y"]]
        report_rows.append(
            {
                **{k: v for k, v in r.items() if k != "cell"},
                "family": c["family"],
                "name": c["name"],
                "ppem": c["size"],
                "tol": tol,
                "pass": passed,
            }
        )

    report = {"token_ppems": [[t, p] for t, p in TOKEN_PPems], "residuals": residuals, "rows": report_rows}
    (here / "calibration_report.json").write_text(json.dumps(report, ensure_ascii=False, indent=1), encoding="utf-8")
    n_corr = sum(
        1 for f in residuals.values() for g in f.values() for dx, dy in g.values() if abs(dx) >= 0.75 or abs(dy) >= 0.75
    )
    worst = max(report_rows, key=lambda r: max(abs(r["err_x"]), abs(r["err_y"])))
    print(
        f"cells={len(report_rows)} fail={n_fail} corrections候选={n_corr} "
        f"worst={worst['family'][:12]}/{worst['name']}@{worst['ppem']} "
        f"err=({worst['err_x']:+.2f},{worst['err_y']:+.2f}) tol={worst['tol']}"
    )
    print("下一步：重跑 build_icon_data.py 并入 CORRECTIONS，再跑验收测试。")
    return 1 if n_fail else 0


if __name__ == "__main__":
    sys.exit(main())
