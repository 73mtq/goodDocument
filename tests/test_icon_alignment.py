"""图标对齐屏上验收量具（公差关）。

经真实 IconRenderer 渲染 22 图标 × 9 档 (token, ppem)（校准生效），
白框基准 + 截图墨迹检测，逐格断言分量公差：
    |err_x|, |err_y| ≤ tol，tol = ppem/token − 0.5（0.5px@100% → 1.0px@150%）。

产出 design/icon-metrics/：
- specimen.png            字样版（全图标全档位，人工关复核材料）
- acceptance_report.json  逐格偏差报告

无显示 / 无图标字体 / 缺 numpy·PIL（dev 依赖）时跳过。
"""
import json
import sys
import unittest
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
METRICS = ROOT / "design" / "icon-metrics"
sys.path.insert(0, str(METRICS))

try:
    import numpy as np  # noqa: F401  量具依赖（dev）
    from PIL import Image  # noqa: F401
    HAS_DEPS = True
except ImportError:
    HAS_DEPS = False


@unittest.skipUnless(HAS_DEPS, "缺 numpy/PIL（dev 依赖）")
class IconAlignmentAcceptance(unittest.TestCase):
    def test_all_icons_within_tolerance(self):
        import tkinter.font as tkfont

        import icon_slot
        import ui_icons
        from tests._helpers import tk_root  # noqa: E402
        try:
            from calibrate import TOKEN_PPems, tol_for
            from inklab import run_sheet
        except ImportError:
            self.skipTest("design/icon-metrics 工具不可用")

        root = tk_root()
        if root is None:
            self.skipTest("无显示环境")
        families = set(tkfont.families(root))
        family = ui_icons.select_icon_family(families)
        if family is None:
            self.skipTest("本机无 Segoe 图标字体")
        renderer = icon_slot.IconRenderer(root, family)

        rows = []
        tolerances = []
        for token, ppem in TOKEN_PPems:
            row = []
            for name in ui_icons.ICONS:
                placement = renderer.text_placement(name, ppem)
                self.assertIsNotNone(
                    placement, f"{family}/{name}@{ppem}px 无摆位（数据缺？）")
                x, y, glyph = placement
                row.append({"name": name, "family": family, "glyph": glyph,
                            "size": ppem, "tx": x, "ty": y,
                            "label": f"{family} · token{token} → {ppem}px"})
                tolerances.append(tol_for(token, ppem))
            rows.append(row)

        results, _shot = run_sheet(
            rows, METRICS / "specimen.png",
            row_header=lambda c: c["label"], cell_label=True)

        report_rows = []
        failures = []
        for tol, r in zip(tolerances, results, strict=True):
            c = r["cell"]
            if "error" in r:
                failures.append(f"{c['name']}@{c['size']}px: {r['error']}")
                continue
            passed = (abs(r["err_x"]) <= tol and abs(r["err_y"]) <= tol
                      and r["size_ok"])
            report_rows.append({
                "name": c["name"], "ppem": c["size"],
                "err_x": r["err_x"], "err_y": r["err_y"], "tol": tol,
                "size_ok": r["size_ok"], "pass": passed,
            })
            if not passed:
                failures.append(
                    f"{c['name']}@{c['size']}px err=({r['err_x']:+.2f},"
                    f"{r['err_y']:+.2f}) tol={tol}")

        (METRICS / "acceptance_report.json").write_text(
            json.dumps({"family": family, "rows": report_rows},
                       ensure_ascii=False, indent=1), encoding="utf-8")
        n_pass = sum(1 for r in report_rows if r["pass"])
        print(f"\n图标对齐验收：{n_pass}/{len(report_rows)} 达标 "
              f"（family={family}，报告与字样版在 design/icon-metrics/）")
        self.assertEqual(failures, [], "图标对齐超差：\n" + "\n".join(failures))


if __name__ == "__main__":
    unittest.main()
