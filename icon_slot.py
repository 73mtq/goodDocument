"""图标槽：唯一合法的图标排版入口（测量驱动，v0.8.1）。

排版模型（design/icon-metrics/DECISION.md，字体表 + GDI 落点探针实测确立）：
- Segoe Fluent Icons / Segoe MDL2 三套度量一致：行盒=1em、基线在行盒底、advance=1em 方格；
- tk 文本项按行盒排布，墨迹中心（行盒左上原点、像素尺寸 s）=
  ((x0+x1)/2·s/upm, s − (y0+y1)/2·s/upm)；
- 文本以 anchor="nw" 摆到「格心 − 预测墨迹中心 − 校准值」，墨迹即光学居中；
- 图标像素尺寸 = round(设计 token × dpi/96)，恒为整数 ppem；token 按 points 传入字体
  的旧做法在 Windows 上恒放大 4/3，禁止恢复。

调用方只许摆放槽本身（pack/place 这个 Canvas），不许对槽内做任何 padding/微调；
换色 set_color、换字形 set_icon、改底 set_bg；字体或字形数据缺失时工厂返回 None
（纯文字回退契约不变，调用方必须判 None 跳过布局）。
墨迹表与校准表由 design/icon-metrics/build_icon_data.py 生成到 icon_slot_data.py，
禁止手改；字形码点仍以 ui_icons.ICONS 为唯一来源（PUA 纪律）。
"""

import tkinter as tk

import icon_slot_data
import ui_icons


class IconCell:
    """一个图标槽实例：固定像素格 Canvas，墨迹光学居中。"""

    def __init__(self, canvas, renderer, name, size_px, fg):
        self._canvas = canvas
        self._renderer = renderer
        self._size = size_px
        self._name = name
        self._fg = fg
        self._item = None
        self._draw()

    @property
    def canvas(self):
        return self._canvas

    def _draw(self):
        canvas = self._canvas
        if self._item is not None:
            canvas.delete(self._item)
            self._item = None
        placement = self._renderer.text_placement(self._name, self._size)
        if placement is None:
            return
        x, y, glyph = placement
        self._item = canvas.create_text(
            x, y, text=glyph, fill=self._fg, font=(self._renderer.family, -self._size), anchor="nw"
        )

    def set_icon(self, name):
        self._name = name
        self._draw()

    def set_color(self, fg):
        self._fg = fg
        if self._item is not None:
            self._canvas.itemconfigure(self._item, fill=fg)

    def set_bg(self, bg):
        self._canvas.configure(bg=bg)


class IconRenderer:
    """按字体族持有的图标槽工厂；family=None（纯文字回退）时一切产出 None。"""

    def __init__(self, root, family):
        self.family = family
        self._dpi_scale = root.winfo_fpixels("1i") / 96.0

    def size_px(self, token):
        """设计 token（px@100%）→ 整数像素 ppem。"""
        return max(1, round(token * self._dpi_scale))

    def text_placement(self, name, size_px):
        """(x, y, glyph)：使墨迹盒中心落在 (size/2, size/2) 的 anchor="nw" 摆位。
        字形或墨迹数据缺失返回 None。"""
        glyph = ui_icons.glyph(name)
        if not glyph or self.family is None:
            return None
        font_data = icon_slot_data.INK.get(self.family)
        if not font_data:
            return None
        entry = font_data["glyphs"].get(name)
        if entry is None:
            return None
        x0, y0, x1, y1 = entry
        upm = font_data["upm"]
        s = size_px
        cx = (x0 + x1) / 2 * s / upm  # 墨迹中心（相对行盒左上）
        cy = s - (y0 + y1) / 2 * s / upm  # 纵向：基线在行盒底，y 向下
        corr_x, corr_y = icon_slot_data.correction(self.family, name, s)
        return s / 2 - cx - corr_x, s / 2 - cy - corr_y, glyph

    def cell(self, parent, name, token, bg, fg):
        """创建图标槽；family/glyph/数据缺失返回 None，调用方跳过布局。"""
        if self.family is None:
            return None
        if ui_icons.glyph(name) is None:
            return None
        size_px = self.size_px(token)
        if self.text_placement(name, size_px) is None:
            return None
        canvas = tk.Canvas(parent, width=size_px, height=size_px, bg=bg, highlightthickness=0, borderwidth=0)
        return IconCell(canvas, self, name, size_px, fg)

    def text_pos(self, center_x, center_y, name, token):
        """既有画布内嵌图标（如拖放区主图）的摆位：墨迹盒中心落在 (center_x, center_y)。
        返回 (x, y) 或 None；需搭配 anchor="nw" 与负号像素字号使用。"""
        if self.family is None:
            return None
        s = self.size_px(token)
        placement = self.text_placement(name, s)
        if placement is None:
            return None
        return center_x - s / 2 + placement[0], center_y - s / 2 + placement[1]
