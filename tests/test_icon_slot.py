"""图标槽数据与摆位测试（不依赖屏上渲染；屏上验收在 test_icon_alignment.py）。

锁三件事：
1. 数据形状与纪律：INK 覆盖全部字形、墨迹盒在 em 盒内、CORRECTIONS 只收确定性离群；
2. 摆位公式：预测墨迹中心 = 格心（校准叠加后仍成立）、墨迹永不越格（无裁切风险）；
3. 数据新鲜度：字体单位墨迹表与字体文件再读结果一致（字体缺失则跳过）。
"""

import unittest
from pathlib import Path

import icon_slot
import icon_slot_data
import ui_icons
from tests._helpers import tk_root  # noqa: E402

ROOT = Path(__file__).resolve().parents[1]

# 与 design/icon-metrics/calibrate.py 同源：3 token × 3 设计缩放
TOKEN_PPems = [(token, round(token * scale)) for token in (16, 20, 40) for scale in (1.0, 1.25, 1.5)]
FONT_FILES = {
    "Segoe Fluent Icons": Path(r"C:\Windows\Fonts\SegoeIcons.ttf"),
    "Segoe MDL2 Assets": Path(r"C:\Windows\Fonts\segmdl2.ttf"),
}


class IconDataTests(unittest.TestCase):
    def test_ink_covers_all_glyphs_within_em_box(self):
        for family, data in icon_slot_data.INK.items():
            if data is None:
                continue  # 字体缺失记 None（构建环境无该字体），运行时纯文字回退
            self.assertEqual(set(data["glyphs"]), set(ui_icons.ICONS), f"{family} 墨迹表与码点表不一致")
            upm = data["upm"]
            for name, (x0, y0, x1, y1) in data["glyphs"].items():
                self.assertTrue(
                    0 <= x0 <= x1 <= upm and 0 <= y0 <= y1 <= upm,
                    f"{family}/{name} 墨迹盒越出 em 盒：{x0},{y0},{x1},{y1}",
                )

    def test_corrections_are_deterministic_outliers_only(self):
        """校准只收 |分量|≥0.75px 的 grid-fit 离群（±0.5px 是量测粒度，修了会过校正）。"""
        for family, glyphs in icon_slot_data.CORRECTIONS.items():
            self.assertIn(family, icon_slot_data.INK)
            for name, sizes in glyphs.items():
                self.assertIn(name, ui_icons.ICONS)
                for ppem, (dx, dy) in sizes.items():
                    self.assertTrue(int(ppem) > 0)
                    self.assertTrue(
                        max(abs(dx), abs(dy)) >= 0.75, f"{family}/{name}@{ppem} 校准 {dx},{dy} 未达离群阈值"
                    )
                    self.assertTrue(max(abs(dx), abs(dy)) <= 2.0, f"{family}/{name}@{ppem} 校准 {dx},{dy} 异常大")

    def test_correction_lookup(self):
        self.assertEqual(icon_slot_data.correction("不存在", "page", 16), (0.0, 0.0))
        known = next(iter(icon_slot_data.CORRECTIONS), None)
        if known:
            name, sizes = next(iter(icon_slot_data.CORRECTIONS[known].items()))
            ppem, value = next(iter(sizes.items()))
            self.assertEqual(icon_slot_data.correction(known, name, int(ppem)), tuple(value))


class PlacementMathTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = tk_root()
        if cls.root is None:
            raise unittest.SkipTest("无显示环境，跳过 Tk 依赖测试")

    def test_placement_centers_ink_for_every_glyph_and_ppem(self):
        for family, data in icon_slot_data.INK.items():
            if data is None:
                continue
            renderer = icon_slot.IconRenderer(self.root, family)
            upm = data["upm"]
            for name in ui_icons.ICONS:
                x0, y0, x1, y1 = data["glyphs"][name]
                for ppem in sorted({ppem for _t, ppem in TOKEN_PPems}):
                    cdx, cdy = icon_slot_data.correction(family, name, ppem)
                    placement = renderer.text_placement(name, ppem)
                    self.assertIsNotNone(placement)
                    px, py, glyph = placement
                    self.assertEqual(glyph, ui_icons.glyph(name))
                    # 公式锁死：墨迹中心 = 格心（含校准）
                    ink_cx = (x0 + x1) / 2 * ppem / upm
                    ink_cy = ppem - (y0 + y1) / 2 * ppem / upm
                    self.assertAlmostEqual(px + ink_cx + cdx, ppem / 2, places=6)
                    self.assertAlmostEqual(py + ink_cy + cdy, ppem / 2, places=6)

    def test_rendered_ink_never_leaves_cell(self):
        """渲染墨迹（摆位+字形+校准）⊆ [0, s]——画布无裁切风险。
        校准正是按「渲染墨迹回到格心」标的，故含校准判定。"""
        for family, data in icon_slot_data.INK.items():
            if data is None:
                continue
            renderer = icon_slot.IconRenderer(self.root, family)
            upm = data["upm"]
            for name in ui_icons.ICONS:
                x0, y0, x1, y1 = data["glyphs"][name]
                for _, ppem in TOKEN_PPems:
                    cdx, cdy = icon_slot_data.correction(family, name, ppem)
                    px, py, _ = renderer.text_placement(name, ppem)
                    s = ppem
                    left = px + x0 * s / upm + cdx
                    right = px + x1 * s / upm + cdx
                    top = py + s - y1 * s / upm + cdy
                    bottom = py + s - y0 * s / upm + cdy
                    self.assertTrue(
                        -0.01 <= left and right <= s + 0.01 and -0.01 <= top and bottom <= s + 0.01,
                        f"{family}/{name}@{ppem}px 渲染墨迹越格",
                    )


class RendererTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.root = tk_root()
        if cls.root is None:
            raise unittest.SkipTest("无显示环境，跳过 Tk 依赖测试")

    def test_size_px_follows_dpi96_scale(self):
        family = next(f for f, d in icon_slot_data.INK.items() if d is not None)
        renderer = icon_slot.IconRenderer(self.root, family)
        scale = self.root.winfo_fpixels("1i") / 96.0
        for token in (16, 20, 40):
            self.assertEqual(renderer.size_px(token), max(1, round(token * scale)))

    def test_no_family_yields_nothing(self):
        renderer = icon_slot.IconRenderer(self.root, None)
        import tkinter as tk

        frame = tk.Frame(self.root)
        self.assertIsNone(renderer.cell(frame, "page", 16, "#fff", "#000"))
        self.assertIsNone(renderer.text_placement("page", 16))
        self.assertIsNone(renderer.text_pos(10, 10, "page", 16))

    def test_unknown_glyph_returns_none(self):
        family = next(f for f, d in icon_slot_data.INK.items() if d is not None)
        renderer = icon_slot.IconRenderer(self.root, family)
        self.assertIsNone(renderer.text_placement("不存在的图标", 16))
        import tkinter as tk

        frame = tk.Frame(self.root)
        self.assertIsNone(renderer.cell(frame, "不存在的图标", 16, "#fff", "#000"))

    def test_cell_is_fixed_pixel_canvas_and_mutable(self):
        family = next(f for f, d in icon_slot_data.INK.items() if d is not None)
        renderer = icon_slot.IconRenderer(self.root, family)
        import tkinter as tk

        frame = tk.Frame(self.root)
        cell = renderer.cell(frame, "settings", 16, "#111111", "#eeeeee")
        self.assertIsNotNone(cell)
        s = renderer.size_px(16)
        self.assertEqual(cell.canvas["width"], str(s))
        self.assertEqual(cell.canvas["height"], str(s))
        self.assertEqual(cell.canvas["bg"], "#111111")
        self.assertEqual(cell.canvas.itemcget(min(cell.canvas.find_all()), "text"), ui_icons.glyph("settings"))
        cell.set_icon("check")
        self.assertEqual(cell.canvas.itemcget(min(cell.canvas.find_all()), "text"), ui_icons.glyph("check"))
        cell.set_color("#ffffff")
        self.assertEqual(cell.canvas.itemcget(min(cell.canvas.find_all()), "fill"), "#ffffff")
        cell.set_bg("#222222")
        self.assertEqual(cell.canvas["bg"], "#222222")


class DataFreshnessTests(unittest.TestCase):
    def test_ink_matches_font_files(self):
        """守卫：icon_slot_data.INK 与字体文件再读结果一致（防手改/防过期）。"""
        try:
            from fontTools.pens.boundsPen import BoundsPen
            from fontTools.ttLib import TTFont
        except ImportError:
            self.skipTest("fontTools 未安装（dev 依赖）")
        for family, path in FONT_FILES.items():
            if not path.exists():
                self.skipTest(f"{path} 不存在")
            data = icon_slot_data.INK.get(family)
            self.assertIsNotNone(data, f"{family} 墨迹表缺失，请重跑 build_icon_data.py")
            tt = TTFont(str(path), fontNumber=0)
            glyphset = tt.getGlyphSet()
            cmap = tt.getBestCmap()
            for name, ch in ui_icons.ICONS.items():
                pen = BoundsPen(glyphset)
                glyphset[cmap[ord(ch)]].draw(pen)
                self.assertEqual(
                    data["glyphs"][name], [int(v) for v in pen.bounds], f"{family}/{name} 墨迹表与字体文件不符"
                )


class PuaDisciplineTests(unittest.TestCase):
    def test_no_pua_literals_outside_ui_icons(self):
        """码点唯一来源是 ui_icons.ICONS；槽与数据模块不得出现私用区字面量。"""
        for name in ("icon_slot.py", "icon_slot_data.py"):
            text = (ROOT / name).read_text(encoding="utf-8")
            hits = [hex(ord(c)) for c in text if ui_icons.PUA_RANGE[0] <= ord(c) <= ui_icons.PUA_RANGE[1]]
            self.assertEqual(hits, [], f"{name} 出现私用区字面量")


if __name__ == "__main__":
    unittest.main()
