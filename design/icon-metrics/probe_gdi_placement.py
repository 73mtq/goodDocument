"""GDI 落点探针 v2：Route A（字形补偿）的残差量具。

仪器设计（零活动部件）：一块纯黑大画布，白色 1px 矩形框 + 白色字形同画布同坐标系。
截图后：白色矩形框连通域 = 绝对基准框（闭合外环连通，墨迹劈不开它，尺寸自证），
框内白色连通域 = 墨迹。残差 = 墨迹盒中心 − 框中心，全程截图像素空间，
无 winfo→屏幕换算、无底色技巧、无 highlight 几何陷阱。

Route A 预测式（Segoe 图标字体三套度量一致：行盒=1em、基线在行盒底、advance=1em）：
  墨迹中心（相对行盒左上）= ((x0+x1)/2·s/upm, s − (y0+y1)/2·s/upm)
  create_text 摆到 (s/2−cx, s/2−cy, anchor=nw) 即「预测墨迹中心=格心」。
残差 ≤ 公差（0.5px@16 / 0.75px@24,30 / 1px@60）→ Route A 成立，否则 Route B。

只读不写运行时代码；退出码 0 = 全部样本在公差内。
"""

import ctypes
import json
import sys
import time
from pathlib import Path

import numpy as np
from fontTools.pens.boundsPen import BoundsPen
from fontTools.ttLib import TTFont
from PIL import ImageGrab

ROOT = Path(__file__).resolve().parents[2]
sys.path.insert(0, str(ROOT))

import tkinter as tk

import ui_icons

FONTS = [
    ("Segoe Fluent Icons", r"C:\Windows\Fonts\SegoeIcons.ttf"),
    ("Segoe MDL2 Assets", r"C:\Windows\Fonts\segmdl2.ttf"),
]
# 每档取墨迹形态差异最大的样本：矮（check）、宽扁（view/table）、满格（font）、异形（chevron）
SAMPLES = ("check", "view", "table", "font", "settings", "chevron_right", "warning", "play")
SIZES = (16, 24, 30, 60)  # 设计 16/20/40 @100% 与 @150%
TOL = {16: 0.5, 24: 0.75, 30: 0.75, 60: 1.0}  # 设计公差：0.5px@100%、1px@150%（中间线性）

GAP = 24  # 格距，保证矩形框与墨迹、邻格互不粘连
M = 20  # 画布边距
FG = (255, 255, 255)


def ink_units(font_path, ch):
    tt = TTFont(str(font_path), fontNumber=0)
    glyphset = tt.getGlyphSet()
    gname = tt.getBestCmap().get(ord(ch))
    pen = BoundsPen(glyphset)
    glyphset[gname].draw(pen)
    return pen.bounds, tt["head"].unitsPerEm


def components(mask):
    """4-连通域列表 [(ys, xs), ...]，带尺寸下限过滤由调用方做。"""
    visited = np.zeros(mask.shape, bool)
    out = []
    for py, px in np.argwhere(mask):
        if visited[py, px]:
            continue
        stack = [(py, px)]
        visited[py, px] = True
        ys, xs = [], []
        while stack:
            cy, cx = stack.pop()
            ys.append(cy)
            xs.append(cx)
            for dy, dx in ((1, 0), (-1, 0), (0, 1), (0, -1)):
                ny, nx = cy + dy, cx + dx
                if 0 <= ny < mask.shape[0] and 0 <= nx < mask.shape[1] and mask[ny, nx] and not visited[ny, nx]:
                    visited[ny, nx] = True
                    stack.append((ny, nx))
        out.append((np.array(ys), np.array(xs)))
    return out


def main():
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass

    # ---- 布局声明：一行一个 (family, size)；框=格+8（四周 4px 缓冲防墨迹粘框）----
    cursor_y = M
    row_h = max(SIZES) + 8  # 框尺寸
    cells_per_row = len(SAMPLES)
    cell_w = row_h + GAP
    W = M * 2 + cells_per_row * cell_w
    H = M * 2 + len(FONTS) * len(SIZES) * (row_h + GAP)
    declared = []
    for family, font_path in FONTS:
        if not Path(font_path).exists():
            continue
        for size in SIZES:
            cursor_x = M + (row_h - size - 8) // 2
            for name in SAMPLES:
                ch = ui_icons.glyph(name)
                (x0, y0, x1, y1), upm = ink_units(font_path, ch)
                cx = (x0 + x1) / 2 * size / upm  # 墨迹中心（相对行盒左上）
                cy = size - (y0 + y1) / 2 * size / upm  # 纵向：基线在行盒底
                # 内格（size×size）居中于框；文本摆到「预测墨迹中心=内格中心」
                ix0 = cursor_x + 4
                iy0 = cursor_y + 4
                declared.append(
                    {
                        "family": family,
                        "name": name,
                        "size": size,
                        "rect": [cursor_x, cursor_y, cursor_x + size + 7, cursor_y + size + 7],
                        "text_at": [ix0 + size / 2 - cx, iy0 + size / 2 - cy],
                        "glyph": ch,
                    }
                )
                cursor_x += cell_w
            cursor_y += row_h + GAP

    root = tk.Tk()
    root.attributes("-topmost", True)
    root.geometry("+40+40")
    canvas = tk.Canvas(root, width=W, height=H, bg="#000000", highlightthickness=0)
    canvas.pack()
    for d in declared:
        rx0, ry0, rx1, ry1 = d["rect"]
        canvas.create_rectangle(rx0, ry0, rx1, ry1, outline="#ffffff", width=1)
        canvas.create_text(
            d["text_at"][0],
            d["text_at"][1],
            text=d["glyph"],
            fill="#ffffff",
            font=(d["family"], -d["size"]),
            anchor="nw",
        )
    root.update_idletasks()
    root.update()
    time.sleep(0.4)

    x, y = root.winfo_rootx(), root.winfo_rooty()
    w, h = root.winfo_width(), root.winfo_height()
    shot = ImageGrab.grab(bbox=(x, y, x + w, y + h)).convert("RGB")
    arr = np.asarray(shot).astype(int)
    root.destroy()
    shot_path = Path(__file__).resolve().parent / "probe_screenshot.png"
    shot.save(shot_path)

    # ---- 检测：白色连通域；闭合矩形框（尺寸≈声明）为基准框，框内其余白域为墨迹 ----
    white = np.abs(arr - np.array(FG)).sum(axis=2) < 260
    comps = [(cys, cxs) for cys, cxs in components(white) if len(cys) >= 8]
    frames = []
    valid_sizes = {(s + 8, s + 8) for s in SIZES}
    for cys, cxs in comps:
        fx0, fx1 = cxs.min(), cxs.max()
        fy0, fy1 = cys.min(), cys.max()
        fw, fh = fx1 - fx0 + 1, fy1 - fy0 + 1
        # 矩形框特征：1px 细环（像素数≈周长），且外接尺寸=声明框尺寸
        if len(cys) <= 2 * (fw + fh) + 4 and (fw, fh) in valid_sizes:
            frames.append((fx0, fy0, fx1, fy1))
    frames.sort(key=lambda f: (f[1], f[0]))

    if len(frames) != len(declared):
        print(f"框检出 {len(frames)} != 声明 {len(declared)}，量具失效", file=sys.stderr)
        return 2

    out = []
    for d, (fx0, fy0, fx1, fy1) in zip(declared, frames, strict=True):
        fw, fh = fx1 - fx0 + 1, fy1 - fy0 + 1
        # 框内墨迹：内缩 3px（排除框线 1px + AA 1px；内格边缘在 4px 处，墨迹不丢）
        sub = white[fy0 + 3 : fy1 - 2, fx0 + 3 : fx1 - 2]
        iy, ix = np.nonzero(sub)
        if len(iy) == 0:
            out.append({"family": d["family"], "name": d["name"], "size": d["size"], "error": "no ink inside frame"})
            continue
        err_x = (ix.min() + ix.max()) / 2 + 3 - (fw - 1) / 2  # +3：子区映射回框坐标
        err_y = (iy.min() + iy.max()) / 2 + 3 - (fh - 1) / 2
        err = (err_x**2 + err_y**2) ** 0.5
        tol = TOL[d["size"]]
        out.append(
            {
                "family": d["family"],
                "name": d["name"],
                "size": d["size"],
                "frame_px": [int(fw), int(fh)],
                "declared_px": [d["size"] + 8, d["size"] + 8],
                "ink_bbox": [int(ix.min()) + 3, int(iy.min()) + 3, int(ix.max()) + 3, int(iy.max()) + 3],
                "err_x": round(err_x, 2),
                "err_y": round(err_y, 2),
                "err": round(err, 2),
                "tol": tol,
                "size_ok": bool(fw == d["size"] + 8 and fh == d["size"] + 8),
                "pass": bool(err <= tol + 0.05),
            }
        )

    out_path = Path(__file__).resolve().parent / "probe_results.json"
    out_path.write_text(json.dumps(out, ensure_ascii=False, indent=1), encoding="utf-8")
    measured = [r for r in out if "err" in r]
    failed_probes = [r for r in out if "error" in r]
    for r in failed_probes:
        print(f"PROBE-FAIL {r['family']} {r['name']}@{r['size']}: {r['error']}")
    bad_size = [r for r in measured if not r["size_ok"]]
    if bad_size:
        print(f"警告：{len(bad_size)} 格实测框尺寸 ≠ 声明（DPI 虚拟化？结果不可信）")
    for r in measured:
        flag = "PASS" if r["pass"] else "FAIL"
        print(
            f"{flag} {r['family'][:24]:24} {r['name']:14} {r['size']:3}px "
            f"err=({r['err_x']:+.2f},{r['err_y']:+.2f}) |{r['err']:.2f}| tol={r['tol']}"
        )
    worst = max(measured, key=lambda r: r["err"])
    n_pass = sum(1 for r in measured if r["pass"])
    print(
        f"worst: {worst['family'][:12]}/{worst['name']}@{worst['size']} "
        f"err={worst['err']:.2f}px tol={worst['tol']}  ->  {n_pass}/{len(measured)} pass"
    )
    ok = not failed_probes and not bad_size and all(r["pass"] for r in measured)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())
