"""inklab：图标对齐量具的共用仪器（画布白框基准 + 截图墨迹检测）。

仪器原理（DECISION.md 第四节的教训全部内置）：
- 一块纯黑大画布；每格一个白色 1px 矩形框（=格+8，四周 4px 缓冲防墨迹粘框）
  + 白色字形。框与字形同画布同坐标系，框即绝对基准——无 winfo→屏幕换算、
  无底色技巧、无 highlight 几何陷阱。
- 检测：白色连通域中「细环且尺寸=声明」者为框（闭合外环连通，环形字形劈不开）；
  框内缩 3px 后的白域（可多连通域，如 BulletedList）取并集包围盒为墨迹。
- 残差 = 墨迹盒中心 − 框中心（=内格中心），全在截图像素空间，分量公差判定。

依赖 numpy/PIL/tk + 真屏（dev 依赖；运行时零依赖）。仅供 calibrate.py 与
tests/test_icon_alignment.py 使用。
"""

import sys
import time
import tkinter as tk
from pathlib import Path

import numpy as np
from PIL import ImageGrab

FRAME_PAD = 4  # 框比内格大出的缓冲（每侧）
INSET = 3  # 框内检测墨迹时排除框线 1px + AA 2px
GAP = 16
MARGIN = 12
LABEL_LEFT = 265  # 行头（字体名+ppem）文字区
FG = (255, 255, 255)


def _make_dpi_aware():
    """像素级量测必须 DPI 感知，否则 DWM 拉伸既糊又改几何。"""
    import ctypes

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass


def components(mask):
    """4-连通域列表 [(ys, xs), ...]。"""
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


def run_sheet(cells, shot_path, row_header=None, cell_label=True):
    """渲染 + 截图 + 量测。

    cells: 逐行给的渲染项列表 [[cell, ...], ...]，cell = dict(name, family, glyph,
    size(内格px), tx, ty) —— tx/ty 为字形 anchor="nw" 摆位相对内格左上的偏移，
    内格 = 框内缩 FRAME_PAD。行内从左到右，行间从上到下。
    返回 (results, screenshot_image)。results 逐格 dict：
    err_x/err_y（墨迹盒中心−格心）、size_ok、ink_bbox。
    row_header: 可选 callable(family, ppem)->str 行头文字；cell_label 是否给格下标注名。
    """
    max_frame = max(c["size"] for row in cells for c in row) + 2 * FRAME_PAD
    cell_w = max_frame + GAP
    W = MARGIN + LABEL_LEFT + len(cells[0]) * cell_w + MARGIN
    H = MARGIN * 2 + len(cells) * (max_frame + GAP + (18 if cell_label else 0))

    _make_dpi_aware()
    root = tk.Tk()
    root.attributes("-topmost", True)
    root.geometry("+40+20")
    if H > root.winfo_screenheight() - 60 or W > root.winfo_screenwidth() - 80:
        root.destroy()
        raise RuntimeError(f"量具画布 {W}x{H} 超出屏幕（含边框余量），落屏外的框测不到")
    canvas = tk.Canvas(root, width=W, height=H, bg="#000000", highlightthickness=0)
    canvas.pack()
    declared = []
    # 行头/标注用暗灰：白色检测阈值（sum<260）看不到它，不产生干扰连通域
    y = MARGIN
    for row in cells:
        x = MARGIN + LABEL_LEFT
        if row_header:
            canvas.create_text(
                MARGIN,
                y + max_frame / 2,
                text=row_header(row[0]),
                fill="#404040",
                font=("Microsoft YaHei UI", 10),
                anchor="w",
            )
        for c in row:
            fx0, fy0 = x, y
            fx1, fy1 = x + c["size"] + 2 * FRAME_PAD - 1, y + c["size"] + 2 * FRAME_PAD - 1
            canvas.create_rectangle(fx0, fy0, fx1, fy1, outline="#ffffff", width=1)
            ix0, iy0 = fx0 + FRAME_PAD, fy0 + FRAME_PAD
            canvas.create_text(
                ix0 + c["tx"],
                iy0 + c["ty"],
                text=c["glyph"],
                fill="#ffffff",
                font=(c["family"], -c["size"]),
                anchor="nw",
            )
            if cell_label:
                canvas.create_text(
                    (fx0 + fx1) / 2, fy1 + 8, text=c["name"], fill="#404040", font=("Microsoft YaHei UI", 8)
                )
            declared.append({"cell": c, "frame": (fx0, fy0, fx1, fy1)})
            x += cell_w
        y += max_frame + GAP + (18 if cell_label else 0)

    root.update_idletasks()
    root.update()
    time.sleep(0.4)
    wx, wy = root.winfo_rootx(), root.winfo_rooty()
    ww, wh = root.winfo_width(), root.winfo_height()
    shot = ImageGrab.grab(bbox=(wx, wy, wx + ww, wy + wh)).convert("RGB")
    root.destroy()
    if shot_path:
        Path(shot_path).parent.mkdir(parents=True, exist_ok=True)
        shot.save(shot_path)

    arr = np.asarray(shot).astype(int)
    white = np.abs(arr - np.array(FG)).sum(axis=2) < 260
    comps = [(cys, cxs) for cys, cxs in components(white) if len(cys) >= 8]

    frames = []
    valid = {(c["size"] + 2 * FRAME_PAD, c["size"] + 2 * FRAME_PAD) for row in cells for c in row}
    for cys, cxs in comps:
        fx0, fy0 = cxs.min(), cys.min()
        fw = int(cxs.max()) - fx0 + 1
        fh = int(cys.max()) - fy0 + 1
        # 环判定：1px 细环（像素数≈周长）且尺寸恰为某个声明框
        if len(cys) <= 2 * (fw + fh) + 4 and (fw, fh) in valid:
            frames.append((int(fx0), int(fy0), fx0 + fw - 1, fy0 + fh - 1))
    frames.sort(key=lambda f: (f[1], f[0]))

    if len(frames) != len(declared):
        got = {(f[0], f[1]) for f in frames}
        missing = [d["cell"] for d in declared if (d["frame"][0], d["frame"][1]) not in got]
        for c in missing[:8]:
            print(
                f"  缺框: {c['family'][:20]} {c['name']}@{c['size']}px 声明位 {c.get('_frame_hint', '?')}",
                file=sys.stderr,
            )
        raise RuntimeError(f"框检出 {len(frames)} != 声明 {len(declared)}，量具失效")

    results = []
    for d, (fx0, fy0, fx1, fy1) in zip(declared, frames, strict=True):
        c = d["cell"]
        fw, fh = fx1 - fx0 + 1, fy1 - fy0 + 1
        sub = white[fy0 + INSET : fy1 - INSET + 1, fx0 + INSET : fx1 - INSET + 1]
        iy, ix = np.nonzero(sub)
        if len(iy) == 0:
            results.append({"cell": c, "error": "no ink inside frame"})
            continue
        err_x = (ix.min() + ix.max()) / 2 + INSET - (fw - 1) / 2
        err_y = (iy.min() + iy.max()) / 2 + INSET - (fh - 1) / 2
        results.append(
            {
                "cell": c,
                "err_x": round(float(err_x), 2),
                "err_y": round(float(err_y), 2),
                "ink_bbox": [
                    int(ix.min()) + INSET,
                    int(iy.min()) + INSET,
                    int(ix.max()) + INSET,
                    int(iy.max()) + INSET,
                ],
                "size_ok": bool(fw == c["size"] + 2 * FRAME_PAD and fh == c["size"] + 2 * FRAME_PAD),
            }
        )
    return results, shot
