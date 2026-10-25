"""Apple 系统色主题：设计令牌、ttk 样式注册、取值中英映射。

令牌数值全部来自一手测量，来源与计算见 design/apple-measurements.md（颜色）
与 design/measurements.md（结构尺寸：行高、列宽、分隔节奏，标本为 macOS
系统设置）。视觉规则：近白分组底、白卡片承载内容、近黑文字；全界面只有
一个强调色，且只出现在「开始整理」与聚焦态；每组文字/背景对比度不低于
4.5:1。无阴影、1px 分隔线、间距与行高全部来自令牌表且落在 4 的倍数；
不用浅字重、不仿 macOS 窗口部件。
"""
import tkinter.font as tkfont
from tkinter import ttk

import ui_icons

COLORS = {
    "bg": "#F7F7F7",
    "surface": "#FFFFFF",
    "surface_alt": "#EDEDED",
    "surface_hover": "#F0F0F0",
    "surface_pressed": "#E4E4E4",
    "border": "#D8D8D8",
    "border_strong": "#BEBEBE",
    "ink": "#1F1F1F",
    "ink_muted": "#636368",
    "ink_disabled": "#BEBEBE",
    "accent": "#2F6EED",
    "accent_hover": "#2860D8",
    "accent_pressed": "#2455C4",
    "accent_soft": "#E5EEFD",
    "accent_disabled_bg": "#C9D3EA",
    "accent_disabled_fg": "#FBFCFE",
    "success": "#187734",
    "warning": "#7E560F",
    "danger": "#C03527",
}

SPACING = {"xs": 4, "sm": 8, "md": 12, "lg": 16, "xl": 24, "xxl": 32}

# 结构令牌（px）：量测依据与换算见 design/measurements.md。
MEASUREMENTS = {
    "row_height": 44,           # 设置行高（系统设置内容行距 34pt）
    "group_header_height": 32,  # 组头高（侧栏行 25pt）
    "control_height": 28,       # 尾部控件高（弹窗按钮 bezel 19-20pt）
    "panel_width": 480,         # 高级设置面板宽
    "name_column": 128,         # 行/主流程名称列宽
    "group_gap": 24,            # 组卡间距
    "separator_inset": 44,      # 行分隔线左缩进
    "indent_step": 36,          # 主流程步骤名/设置行名文字列起点（12+16+8）
    "indent_group": 40,         # 组头标题列起点（12+20+8）
    "drop_height": 96,          # 拖放区高
    "wrap_main": 760,           # 主卡说明折行宽
    "wrap_panel": 440,          # 面板说明折行宽
    "icon_group": 20,           # 组头图标（微软字体文档推荐档位）
    "icon_row": 16,             # 行内图标
    "icon_hero": 40,            # 拖放区主图
}

FAMILY_CANDIDATES = ("Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", "Segoe UI", "Helvetica")
MONO_CANDIDATES = ("Cascadia Mono", "Consolas", "Menlo", "Courier New")

VALUE_LABELS = {
    "left": "左",
    "center": "居中",
    "right": "右",
    "justify": "两端",
    "above": "上方",
    "below": "下方",
    "top": "靠上",
    "fit": "等比缩放",
    "fixed": "统一尺寸",
    "smart": "智能识别",
}
LABEL_VALUES = {label: value for value, label in VALUE_LABELS.items()}


def _luminance(color):
    r, g, b = (int(color[i:i + 2], 16) / 255 for i in (1, 3, 5))

    def channel(value):
        return value / 12.92 if value <= 0.04045 else ((value + 0.055) / 1.055) ** 2.4

    r, g, b = channel(r), channel(g), channel(b)
    return 0.2126 * r + 0.7152 * g + 0.0722 * b


def contrast_ratio(foreground, background):
    """两色对比度，用于测试断言（WCAG 计算方式）。"""
    first, second = _luminance(foreground), _luminance(background)
    lighter, darker = max(first, second), min(first, second)
    return (lighter + 0.05) / (darker + 0.05)


def display_label(value):
    return VALUE_LABELS.get(value, value)


def option_values(values):
    """把配置取值列表转成界面显示用的中文标签列表。"""
    return [display_label(v) for v in values]


def build_fonts(root):
    available = set(tkfont.families(root))
    family = next((f for f in FAMILY_CANDIDATES if f in available), "TkDefaultFont")
    mono = next((f for f in MONO_CANDIDATES if f in available), "TkFixedFont")
    return {
        "display": (family, 17),
        "section": (family, 13, "bold"),
        "body": (family, 13),
        "body_bold": (family, 13, "bold"),
        "caption": (family, 11),
        "mono": (mono, 11),
        # 图标字体族；检测不到时为 None，界面回退纯文字（见 ui_icons.py）。
        "icon": ui_icons.select_icon_family(available),
    }


def apply_theme(root):
    """在 root 上启用 clam 主题并注册全部样式，返回字体令牌。"""
    fonts = build_fonts(root)
    style = ttk.Style(root)
    style.theme_use("clam")
    root.configure(background=COLORS["bg"])
    _configure_base(style, fonts)
    _configure_buttons(style, fonts)
    _configure_fields(style, fonts)
    _configure_misc(style)
    _configure_popdown(root)
    return fonts


def _configure_base(style, fonts):
    style.configure(
        ".",
        background=COLORS["bg"],
        foreground=COLORS["ink"],
        font=fonts["body"],
        bordercolor=COLORS["border"],
        focuscolor=COLORS["accent"],
        troughcolor=COLORS["surface_alt"],
        selectbackground=COLORS["accent_soft"],
        selectforeground=COLORS["ink"],
        fieldbackground=COLORS["surface"],
    )
    style.configure("TFrame", background=COLORS["bg"])
    style.configure("Card.TFrame", background=COLORS["surface"])
    style.configure("Panel.TFrame", background=COLORS["bg"])
    style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["ink"])
    style.configure("Card.TLabel", background=COLORS["surface"], foreground=COLORS["ink"])
    style.configure("Panel.TLabel", background=COLORS["bg"], foreground=COLORS["ink"])
    style.configure("Display.TLabel", background=COLORS["bg"], foreground=COLORS["ink"],
                    font=fonts["display"])
    style.configure("Section.TLabel", background=COLORS["bg"], foreground=COLORS["ink"],
                    font=fonts["section"])
    style.configure("CardSection.TLabel", background=COLORS["surface"], foreground=COLORS["ink"],
                    font=fonts["section"])
    style.configure("PanelSection.TLabel", background=COLORS["bg"], foreground=COLORS["ink"],
                    font=fonts["section"])
    style.configure("Muted.TLabel", background=COLORS["bg"], foreground=COLORS["ink_muted"])
    style.configure("CardMuted.TLabel", background=COLORS["surface"], foreground=COLORS["ink_muted"])
    style.configure("PanelMuted.TLabel", background=COLORS["bg"], foreground=COLORS["ink_muted"])
    style.configure("Caption.TLabel", background=COLORS["bg"], foreground=COLORS["ink_muted"],
                    font=fonts["caption"])
    style.configure("CardCaption.TLabel", background=COLORS["surface"], foreground=COLORS["ink_muted"],
                    font=fonts["caption"])
    style.configure("PanelCaption.TLabel", background=COLORS["bg"], foreground=COLORS["ink_muted"],
                    font=fonts["caption"])
    style.configure("Danger.TLabel", background=COLORS["bg"], foreground=COLORS["danger"])
    style.configure("Success.TLabel", background=COLORS["bg"], foreground=COLORS["success"])
    style.configure("TSeparator", background=COLORS["border"])


def _configure_buttons(style, fonts):
    style.configure(
        "TButton",
        background=COLORS["surface"],
        foreground=COLORS["ink"],
        bordercolor=COLORS["border"],
        lightcolor=COLORS["surface"],
        darkcolor=COLORS["border"],
        padding=(16, 8),
        relief="flat",
        font=fonts["body"],
    )
    style.map(
        "TButton",
        background=[("disabled", COLORS["surface_alt"]), ("pressed", COLORS["surface_pressed"]),
                    ("active", COLORS["surface_hover"])],
        foreground=[("disabled", COLORS["ink_disabled"])],
        bordercolor=[("focus", COLORS["accent"])],
        lightcolor=[("pressed", COLORS["surface_pressed"]), ("active", COLORS["surface_hover"])],
        darkcolor=[("pressed", COLORS["surface_pressed"]), ("active", COLORS["surface_hover"])],
    )
    style.configure(
        "Accent.TButton",
        background=COLORS["accent"],
        foreground=COLORS["surface"],
        bordercolor=COLORS["accent"],
        lightcolor=COLORS["accent"],
        darkcolor=COLORS["accent"],
        padding=(24, 8),
        relief="flat",
        font=fonts["body_bold"],
    )
    style.map(
        "Accent.TButton",
        background=[("disabled", COLORS["accent_disabled_bg"]), ("pressed", COLORS["accent_pressed"]),
                    ("active", COLORS["accent_hover"])],
        foreground=[("disabled", COLORS["accent_disabled_fg"])],
        bordercolor=[("focus", COLORS["ink"]), ("disabled", COLORS["accent_disabled_bg"])],
        lightcolor=[("pressed", COLORS["accent_pressed"]), ("active", COLORS["accent_hover"])],
        darkcolor=[("pressed", COLORS["accent_pressed"]), ("active", COLORS["accent_hover"])],
    )
    style.configure(
        "AccentLead.TButton",
        background=COLORS["accent"],
        foreground=COLORS["surface"],
        bordercolor=COLORS["accent"],
        lightcolor=COLORS["accent"],
        darkcolor=COLORS["accent"],
        padding=(8, 24),
        relief="flat",
        font=fonts["body_bold"],
    )
    style.map(
        "AccentLead.TButton",
        background=[("disabled", COLORS["accent_disabled_bg"]), ("pressed", COLORS["accent_pressed"]),
                    ("active", COLORS["accent_hover"])],
        foreground=[("disabled", COLORS["accent_disabled_fg"])],
        bordercolor=[("focus", COLORS["ink"]), ("disabled", COLORS["accent_disabled_bg"])],
        lightcolor=[("pressed", COLORS["accent_pressed"]), ("active", COLORS["accent_hover"])],
        darkcolor=[("pressed", COLORS["accent_pressed"]), ("active", COLORS["accent_hover"])],
    )
    style.configure(
        "Quiet.TButton",
        background=COLORS["bg"],
        foreground=COLORS["ink_muted"],
        bordercolor=COLORS["bg"],
        lightcolor=COLORS["bg"],
        darkcolor=COLORS["bg"],
        padding=(12, 8),
        relief="flat",
    )
    style.map(
        "Quiet.TButton",
        background=[("disabled", COLORS["bg"]), ("pressed", COLORS["surface_pressed"]),
                    ("active", COLORS["surface"])],
        foreground=[("disabled", COLORS["ink_disabled"]), ("active", COLORS["ink"])],
    )
    style.configure(
        "PanelQuiet.TButton",
        background=COLORS["bg"],
        foreground=COLORS["ink_muted"],
        bordercolor=COLORS["bg"],
        lightcolor=COLORS["bg"],
        darkcolor=COLORS["bg"],
        padding=(12, 8),
        relief="flat",
    )
    style.map(
        "PanelQuiet.TButton",
        background=[("disabled", COLORS["bg"]), ("pressed", COLORS["surface_pressed"]),
                    ("active", COLORS["surface"])],
        foreground=[("disabled", COLORS["ink_disabled"]), ("active", COLORS["ink"])],
    )
    style.configure(
        "CardQuiet.TButton",
        background=COLORS["surface"],
        foreground=COLORS["ink_muted"],
        bordercolor=COLORS["surface"],
        lightcolor=COLORS["surface"],
        darkcolor=COLORS["surface"],
        padding=(8, 4),
        relief="flat",
        font=fonts["caption"],
    )
    style.map(
        "CardQuiet.TButton",
        background=[("disabled", COLORS["surface"]), ("pressed", COLORS["surface_pressed"]),
                    ("active", COLORS["surface_hover"])],
        foreground=[("disabled", COLORS["ink_disabled"]), ("active", COLORS["ink"])],
    )
    style.configure(
        "Seg.TButton",
        background=COLORS["surface"],
        foreground=COLORS["ink_muted"],
        bordercolor=COLORS["border"],
        lightcolor=COLORS["surface"],
        darkcolor=COLORS["border"],
        padding=(16, 8),
        relief="flat",
    )
    style.map(
        "Seg.TButton",
        background=[("pressed", COLORS["surface_pressed"]), ("active", COLORS["surface_hover"])],
        foreground=[("active", COLORS["ink"])],
    )
    # 选中态用中性灰 + 加粗，强调蓝只留给「开始整理」与聚焦态。
    style.configure(
        "SegActive.TButton",
        background=COLORS["surface_pressed"],
        foreground=COLORS["ink"],
        bordercolor=COLORS["surface_pressed"],
        lightcolor=COLORS["surface_pressed"],
        darkcolor=COLORS["surface_pressed"],
        padding=(16, 8),
        relief="flat",
        font=fonts["body_bold"],
    )


def _configure_fields(style, fonts):
    style.configure(
        "TEntry",
        fieldbackground=COLORS["surface"],
        foreground=COLORS["ink"],
        bordercolor=COLORS["border"],
        lightcolor=COLORS["border"],
        darkcolor=COLORS["border"],
        padding=8,
    )
    style.map(
        "TEntry",
        bordercolor=[("focus", COLORS["accent"])],
        lightcolor=[("focus", COLORS["accent"])],
        darkcolor=[("focus", COLORS["accent"])],
        foreground=[("disabled", COLORS["ink_disabled"])],
        fieldbackground=[("disabled", COLORS["surface_alt"])],
    )
    style.configure(
        "TCombobox",
        fieldbackground=COLORS["surface"],
        background=COLORS["surface"],
        foreground=COLORS["ink"],
        bordercolor=COLORS["border"],
        lightcolor=COLORS["border"],
        darkcolor=COLORS["border"],
        padding=4,
        arrowcolor=COLORS["ink_muted"],
    )
    style.map(
        "TCombobox",
        bordercolor=[("focus", COLORS["accent"])],
        lightcolor=[("focus", COLORS["accent"])],
        darkcolor=[("focus", COLORS["accent"])],
        arrowcolor=[("active", COLORS["accent"]), ("disabled", COLORS["ink_disabled"])],
        fieldbackground=[("readonly", COLORS["surface"]), ("disabled", COLORS["surface_alt"])],
        foreground=[("disabled", COLORS["ink_disabled"])],
    )
    style.configure(
        "TSpinbox",
        fieldbackground=COLORS["surface"],
        background=COLORS["surface"],
        foreground=COLORS["ink"],
        bordercolor=COLORS["border"],
        lightcolor=COLORS["border"],
        darkcolor=COLORS["border"],
        arrowcolor=COLORS["ink_muted"],
        padding=4,
    )
    style.map(
        "TSpinbox",
        bordercolor=[("focus", COLORS["accent"])],
        lightcolor=[("focus", COLORS["accent"])],
        darkcolor=[("focus", COLORS["accent"])],
        arrowcolor=[("active", COLORS["accent"]), ("disabled", COLORS["ink_disabled"])],
        fieldbackground=[("disabled", COLORS["surface_alt"])],
        foreground=[("disabled", COLORS["ink_disabled"])],
    )


def _configure_misc(style):
    style.configure(
        "TProgressbar",
        troughcolor=COLORS["surface_alt"],
        background=COLORS["accent"],
        bordercolor=COLORS["surface_alt"],
        lightcolor=COLORS["accent"],
        darkcolor=COLORS["accent"],
        thickness=6,
    )
    for orient in ("Vertical", "Horizontal"):
        style.configure(
            f"{orient}.TScrollbar",
            background=COLORS["border_strong"],
            troughcolor=COLORS["bg"],
            bordercolor=COLORS["bg"],
            arrowcolor=COLORS["ink_muted"],
            relief="flat",
        )
        style.map(
            f"{orient}.TScrollbar",
            background=[("active", COLORS["accent"])],
            arrowcolor=[("active", COLORS["accent"])],
        )
    style.configure(
        "TCheckbutton",
        background=COLORS["bg"],
        foreground=COLORS["ink"],
        focuscolor=COLORS["accent"],
        padding=0,
    )
    style.map(
        "TCheckbutton",
        indicatorcolor=[("selected", COLORS["accent"]), ("!selected", COLORS["surface"]),
                        ("disabled", COLORS["surface_alt"])],
        foreground=[("disabled", COLORS["ink_disabled"])],
    )
    style.configure(
        "Panel.TCheckbutton",
        background=COLORS["surface"],
        foreground=COLORS["ink"],
        focuscolor=COLORS["accent"],
        padding=0,
    )
    style.map(
        "Panel.TCheckbutton",
        indicatorcolor=[("selected", COLORS["accent"]), ("!selected", COLORS["surface"]),
                        ("disabled", COLORS["surface_alt"])],
        foreground=[("disabled", COLORS["ink_disabled"])],
    )


def _configure_popdown(root):
    root.option_add("*TCombobox*Listbox.background", COLORS["surface"])
    root.option_add("*TCombobox*Listbox.foreground", COLORS["ink"])
    root.option_add("*TCombobox*Listbox.selectBackground", COLORS["surface_hover"])
    root.option_add("*TCombobox*Listbox.selectForeground", COLORS["ink"])
    root.option_add("*TCombobox*Listbox.borderWidth", 0)
    root.option_add("*TCombobox*Listbox.highlightThickness", 1)
    root.option_add("*TCombobox*Listbox.highlightColor", COLORS["border"])
