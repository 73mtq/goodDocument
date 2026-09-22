"""朱批纸面主题：颜色与字体令牌、ttk 样式注册、取值中英映射。

视觉规则：纸色打底、墨色承载信息、朱砂只出现在动作上；无阴影、1px 分隔线、
以字号与字重建立层级。
"""
import tkinter.font as tkfont
from tkinter import ttk


COLORS = {
    "bg": "#F5F1E8",
    "surface": "#FBF9F4",
    "surface_alt": "#EEE8DB",
    "surface_hover": "#EFE9DC",
    "surface_pressed": "#E6DECD",
    "border": "#DED7C8",
    "border_strong": "#C9C0AC",
    "ink": "#211F1B",
    "ink_muted": "#6B665C",
    "ink_disabled": "#A9A296",
    "accent": "#A8402F",
    "accent_hover": "#93372A",
    "accent_pressed": "#7E2F24",
    "accent_soft": "#EBD8D2",
    "accent_disabled_bg": "#D8C3BD",
    "accent_disabled_fg": "#F2EAE4",
    "success": "#2F6B4F",
    "warning": "#9A6B1F",
    "danger": "#8C2F24",
}

SPACING = {"xs": 4, "sm": 8, "md": 12, "lg": 16, "xl": 24, "xxl": 32}

FAMILY_CANDIDATES = ("Microsoft YaHei UI", "Microsoft YaHei", "PingFang SC", "Segoe UI", "Helvetica")
LIGHT_CANDIDATES = ("Microsoft YaHei UI Light", "Microsoft YaHei Light", "PingFang SC Light", "Segoe UI Light")
MONO_CANDIDATES = ("Cascadia Mono", "Consolas", "Menlo", "Courier New")

VALUE_LABELS = {
    "left": "左",
    "center": "居中",
    "right": "右",
    "justify": "两端",
    "above": "上方",
    "below": "下方",
    "top": "顶部",
    "fit": "等比缩放",
    "fixed": "统一尺寸",
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
    light = next((f for f in LIGHT_CANDIDATES if f in available), family)
    mono = next((f for f in MONO_CANDIDATES if f in available), "TkFixedFont")
    return {
        "display": (light, 21),
        "step": (light, 15),
        "section": (family, 12, "bold"),
        "body": (family, 10),
        "body_bold": (family, 10, "bold"),
        "caption": (family, 9),
        "mono": (mono, 9),
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
    style.configure("Panel.TFrame", background=COLORS["surface"])
    style.configure("Card.TFrame", background=COLORS["surface"])
    style.configure("TLabel", background=COLORS["bg"], foreground=COLORS["ink"])
    style.configure("Panel.TLabel", background=COLORS["surface"], foreground=COLORS["ink"])
    style.configure("Display.TLabel", background=COLORS["bg"], foreground=COLORS["ink"],
                    font=fonts["display"])
    style.configure("Step.TLabel", background=COLORS["bg"], foreground=COLORS["accent"],
                    font=fonts["step"])
    style.configure("Section.TLabel", background=COLORS["bg"], foreground=COLORS["ink"],
                    font=fonts["section"])
    style.configure("PanelSection.TLabel", background=COLORS["surface"], foreground=COLORS["ink"],
                    font=fonts["section"])
    style.configure("Muted.TLabel", background=COLORS["bg"], foreground=COLORS["ink_muted"])
    style.configure("PanelMuted.TLabel", background=COLORS["surface"], foreground=COLORS["ink_muted"])
    style.configure("Caption.TLabel", background=COLORS["bg"], foreground=COLORS["ink_muted"],
                    font=fonts["caption"])
    style.configure("PanelCaption.TLabel", background=COLORS["surface"], foreground=COLORS["ink_muted"],
                    font=fonts["caption"])
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
        padding=(14, 7),
        relief="flat",
        font=fonts["body_bold"],
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
        padding=(20, 8),
        relief="flat",
        font=fonts["body_bold"],
    )
    style.map(
        "Accent.TButton",
        background=[("disabled", COLORS["accent_disabled_bg"]), ("pressed", COLORS["accent_pressed"]),
                    ("active", COLORS["accent_hover"])],
        foreground=[("disabled", COLORS["accent_disabled_fg"])],
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
        padding=(10, 6),
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
        background=COLORS["surface"],
        foreground=COLORS["ink_muted"],
        bordercolor=COLORS["surface"],
        lightcolor=COLORS["surface"],
        darkcolor=COLORS["surface"],
        padding=(10, 6),
        relief="flat",
    )
    style.map(
        "PanelQuiet.TButton",
        background=[("disabled", COLORS["surface"]), ("pressed", COLORS["surface_pressed"]),
                    ("active", COLORS["surface_alt"])],
        foreground=[("disabled", COLORS["ink_disabled"]), ("active", COLORS["ink"])],
    )
    style.configure(
        "Group.TButton",
        background=COLORS["surface"],
        foreground=COLORS["ink"],
        bordercolor=COLORS["surface"],
        lightcolor=COLORS["surface"],
        darkcolor=COLORS["surface"],
        padding=(10, 8),
        anchor="w",
        relief="flat",
        font=fonts["section"],
    )
    style.map(
        "Group.TButton",
        background=[("pressed", COLORS["surface_pressed"]), ("active", COLORS["surface_alt"])],
        foreground=[("active", COLORS["accent"])],
    )
    style.configure(
        "Seg.TButton",
        background=COLORS["surface_alt"],
        foreground=COLORS["ink_muted"],
        bordercolor=COLORS["surface_alt"],
        lightcolor=COLORS["surface_alt"],
        darkcolor=COLORS["surface_alt"],
        padding=(16, 7),
        relief="flat",
    )
    style.map(
        "Seg.TButton",
        background=[("pressed", COLORS["surface_pressed"]), ("active", COLORS["surface_hover"])],
        foreground=[("active", COLORS["ink"])],
    )
    style.configure(
        "SegActive.TButton",
        background=COLORS["surface"],
        foreground=COLORS["accent"],
        bordercolor=COLORS["surface"],
        lightcolor=COLORS["surface"],
        darkcolor=COLORS["surface"],
        padding=(16, 7),
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
        padding=6,
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
        padding=2,
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
        padding=2,
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
    root.option_add("*TCombobox*Listbox.selectBackground", COLORS["accent_soft"])
    root.option_add("*TCombobox*Listbox.selectForeground", COLORS["ink"])
    root.option_add("*TCombobox*Listbox.borderWidth", 0)
    root.option_add("*TCombobox*Listbox.highlightThickness", 1)
    root.option_add("*TCombobox*Listbox.highlightColor", COLORS["border"])
