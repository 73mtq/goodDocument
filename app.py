"""论文格式整理工具（桌面应用 / tkinter GUI）

结构取 macOS 系统设置（量结构不量颜色，依据见 design/measurements.md）：
浅灰底、白色分组卡片、近黑文字；唯一的强调蓝只出现在「开始整理」与
聚焦/拖入态。主流程三步：选论文、选学校模板、保存位置；细节调整在
「高级设置」面板：分组设置行（图标 + 名称 + 当前生效值 + 尾部控件）、
顶部过滤框、每组恢复默认、与学校模板逐行对照（按学校模板 / 已修改）。

图标全部来自 Segoe Fluent Icons（Windows 11，Win10 回退 Segoe MDL2
Assets），运行时检测字体，缺失即纯文字回退；不捆绑字体文件。
文案面向不懂排版术语的人：说结果不说机制，每条帮助一句话，
错误信息写清下一步动作。

用法：
    python app.py
"""
import copy
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
import traceback
from tkinter import filedialog, messagebox, ttk

from tkinterdnd2 import DND_FILES, TkinterDnD

import icon_slot
import ui_icons
import ui_theme
from config_model import SIZE_NAME_TO_PT, list_presets, load_config_file, save_config_file
from paths import app_dir, assets_dir, bundled_dir, config_path, default_config_path, make_normalized_output_path

# 导入核心库
sys.path.insert(0, bundled_dir())
from formatter import build_sample_content, generate  # noqa: E402
from normalizer import NormalizeError, inspect_docx, normalize_docx  # noqa: E402
from version import __version__  # noqa: E402

SIZE_OPTIONS = [
    "初号", "小初", "一号", "小一", "二号", "小二",
    "三号", "小三", "四号", "小四", "五号", "小五",
    "六号", "小六", "七号", "八号",
]
ALIGN_OPTIONS = ["left", "center", "right", "justify"]
CELL_ALIGN_OPTIONS = ALIGN_OPTIONS + ["smart"]
FONT_ZH_OPTIONS = ["宋体", "黑体", "楷体", "仿宋", "微软雅黑"]
FONT_EN_OPTIONS = ["Times New Roman", "Arial", "Calibri", "Cambria"]
POSITION_OPTIONS = ["above", "below"]
VALIGN_OPTIONS = ["center", "top"]
SIZE_MODE_OPTIONS = ["fit", "fixed"]
FIGURE_PREFIX_OPTIONS = ["图", "Figure", "Fig"]
TABLE_PREFIX_OPTIONS = ["表", "Table"]

CUSTOM_PRESET = "自定义"

EMPTY_RESULT_TEXT = "选择论文后，这里会列出将要改动的格式。"
DROP_TITLE = "把论文拖到这里"
DROP_SUBTITLE = "或点击选择 Word 文件"
DROP_SELECTED_SUBTITLE = "点击可重新选择"
STATUS_READY = "先选择一份论文。"
STATUS_BUSY_NORMALIZE = "正在整理，请稍候。"
STATUS_BUSY_PREVIEW = "正在对比改动，请稍候。"
STATUS_BUSY_INSPECT = "正在读取文档，请稍候。"
STATUS_BUSY_GENERATE = "正在生成示例文档，请稍候。"
FILTER_PLACEHOLDER = "输入要找的设置"
FILTER_EMPTY_TEXT = "没有匹配的设置"
TAG_FROM_TEMPLATE = "按学校模板"
TAG_MODIFIED = "已修改"
RESET_GROUP_TEXT = "恢复默认"


def _heading_rows(level):
    prefix = ("headings", "h" + str(level))
    return [
        {"name": "中文字体", "key": (*prefix, "font"), "kind": "enum", "options": FONT_ZH_OPTIONS},
        {"name": "英文字体", "key": (*prefix, "asciiFont"), "kind": "enum", "options": FONT_EN_OPTIONS},
        {"name": "字号", "key": (*prefix, "size"), "kind": "size", "options": SIZE_OPTIONS},
        {"name": "行距", "key": (*prefix, "lineSpacing"), "kind": "number",
         "unit": "倍", "lo": 1.0, "hi": 3.0, "step": 0.5},
        {"name": "加粗", "key": (*prefix, "bold"), "kind": "bool"},
        {"name": "段前间距", "key": (*prefix, "spaceBeforePt"), "kind": "number",
         "unit": "磅", "lo": 0, "hi": 40, "step": 1},
        {"name": "段后间距", "key": (*prefix, "spaceAfterPt"), "kind": "number",
         "unit": "磅", "lo": 0, "hi": 40, "step": 1},
        {"name": "对齐", "key": (*prefix, "align"), "kind": "enum", "options": ALIGN_OPTIONS},
    ]


# 高级设置信息架构：按用户目的分 11 组（卡片分类过程见 design/ia-regroup.md）。
# kind: number(数字+单位) / bool(开关) / enum(下拉) / size(下拉+磅值对照)。
SETTING_GROUPS = [
    {"id": "page", "title": "页面", "icon": "page", "about": "纸张四边的空白宽度。",
     "rows": [
         {"name": "上边距", "key": ("page", "margins", "top"), "kind": "number",
          "unit": "厘米", "lo": 0, "hi": 15, "step": 0.5},
         {"name": "下边距", "key": ("page", "margins", "bottom"), "kind": "number",
          "unit": "厘米", "lo": 0, "hi": 15, "step": 0.5},
         {"name": "左边距", "key": ("page", "margins", "left"), "kind": "number",
          "unit": "厘米", "lo": 0, "hi": 15, "step": 0.5},
         {"name": "右边距", "key": ("page", "margins", "right"), "kind": "number",
          "unit": "厘米", "lo": 0, "hi": 15, "step": 0.5},
     ]},
    {"id": "pageNumber", "title": "页码", "icon": "dialpad", "about": "页码放哪、多大、用什么字体。",
     "rows": [
         {"name": "显示页码", "key": ("pageNumber", "enabled"), "kind": "bool"},
         {"name": "只更新已有页码", "key": ("pageNumber", "onlyUpdateExisting"), "kind": "bool"},
         {"name": "页码位置", "key": ("pageNumber", "align"), "kind": "enum", "options": ALIGN_OPTIONS},
         {"name": "页码大小", "key": ("pageNumber", "size"), "kind": "size", "options": SIZE_OPTIONS},
         {"name": "中文字体", "key": ("pageNumber", "font"), "kind": "enum", "options": FONT_ZH_OPTIONS},
         {"name": "英文字体", "key": ("pageNumber", "asciiFont"), "kind": "enum", "options": FONT_EN_OPTIONS},
     ]},
    {"id": "body", "title": "正文", "icon": "font", "about": "论文主体文字的样子。",
     "rows": [
         {"name": "中文字体", "key": ("body", "font"), "kind": "enum", "options": FONT_ZH_OPTIONS},
         {"name": "英文字体", "key": ("body", "asciiFont"), "kind": "enum", "options": FONT_EN_OPTIONS},
         {"name": "字号", "key": ("body", "size"), "kind": "size", "options": SIZE_OPTIONS},
         {"name": "行距", "key": ("body", "lineSpacing"), "kind": "number",
          "unit": "倍", "lo": 1.0, "hi": 3.0, "step": 0.5},
         {"name": "段首缩进", "key": ("body", "firstLineIndentChars"), "kind": "number",
          "unit": "字", "lo": 0, "hi": 10, "step": 1},
         {"name": "对齐", "key": ("body", "align"), "kind": "enum", "options": ALIGN_OPTIONS},
         {"name": "段后间距", "key": ("body", "spaceAfterPt"), "kind": "number",
          "unit": "磅", "lo": 0, "hi": 30, "step": 1},
     ]},
    {"id": "h1", "title": "章标题（第1章）", "icon": "bold", "about": "一级标题（第1章）的样式。",
     "rows": _heading_rows(1)},
    {"id": "h2", "title": "节标题（1.1）", "icon": "bold", "about": "二级标题（1.1）的样式。",
     "rows": _heading_rows(2)},
    {"id": "h3", "title": "条标题（1.1.1）", "icon": "bold", "about": "三级标题（1.1.1）的样式。",
     "rows": _heading_rows(3)},
    {"id": "h4", "title": "四级标题（1.1.1.1）", "icon": "bold", "about": "四级标题（1.1.1.1）的样式。",
     "rows": _heading_rows(4)},
    {"id": "figure", "title": "图片", "icon": "photo", "about": "图片统一成多大。",
     "rows": [
         {"name": "图片大小方式", "key": ("figure", "sizeMode"), "kind": "enum",
          "options": SIZE_MODE_OPTIONS},
         {"name": "等比最大宽度", "key": ("figure", "maxWidthCm"), "kind": "number",
          "unit": "厘米", "lo": 5, "hi": 25, "step": 0.5},
         {"name": "等比最大高度", "key": ("figure", "maxHeightCm"), "kind": "number",
          "unit": "厘米", "lo": 5, "hi": 30, "step": 0.5},
         {"name": "统一宽度", "key": ("figure", "fixedWidthCm"), "kind": "number",
          "unit": "厘米", "lo": 1, "hi": 25, "step": 0.5},
         {"name": "统一高度", "key": ("figure", "fixedHeightCm"), "kind": "number",
          "unit": "厘米", "lo": 1, "hi": 30, "step": 0.5},
         {"name": "统一时保持比例", "key": ("figure", "keepAspectRatio"), "kind": "bool"},
         {"name": "宽度不超正文", "key": ("figure", "clampToTextWidth"), "kind": "bool"},
         {"name": "跳过表格内图片", "key": ("figure", "skipInTables"), "kind": "bool"},
         {"name": "图和图题不分开", "key": ("figure", "keepWithCaption"), "kind": "bool"},
     ]},
    {"id": "figureCaption", "title": "图题与说明", "icon": "caption",
     "about": "图下面「图1 …」和说明文字的样式。",
     "rows": [
         {"name": "图题位置", "key": ("figure", "captionPosition"), "kind": "enum",
          "options": POSITION_OPTIONS},
         {"name": "图题字号", "key": ("figure", "captionSize"), "kind": "size", "options": SIZE_OPTIONS},
         {"name": "图题加粗", "key": ("figure", "captionBold"), "kind": "bool"},
         {"name": "图题对齐", "key": ("figure", "captionAlign"), "kind": "enum", "options": ALIGN_OPTIONS},
         {"name": "图题行距", "key": ("figure", "captionLineSpacing"), "kind": "number",
          "unit": "倍", "lo": 1.0, "hi": 3.0, "step": 0.5},
         {"name": "注释字号", "key": ("figure", "noteSize"), "kind": "size", "options": SIZE_OPTIONS},
         {"name": "注释对齐", "key": ("figure", "noteAlign"), "kind": "enum", "options": ALIGN_OPTIONS},
         {"name": "编号开头", "key": ("figure", "prefix"), "kind": "enum",
          "options": FIGURE_PREFIX_OPTIONS},
     ]},
    {"id": "table", "title": "表格", "icon": "table", "about": "表格、表题与表内文字的样式。",
     "rows": [
         {"name": "表题位置", "key": ("table", "captionPosition"), "kind": "enum",
          "options": POSITION_OPTIONS},
         {"name": "表题字号", "key": ("table", "captionSize"), "kind": "size", "options": SIZE_OPTIONS},
         {"name": "表题加粗", "key": ("table", "captionBold"), "kind": "bool"},
         {"name": "表题对齐", "key": ("table", "captionAlign"), "kind": "enum", "options": ALIGN_OPTIONS},
         {"name": "表题行距", "key": ("table", "captionLineSpacing"), "kind": "number",
          "unit": "倍", "lo": 1.0, "hi": 3.0, "step": 0.5},
         {"name": "注释字号", "key": ("table", "noteSize"), "kind": "size", "options": SIZE_OPTIONS},
         {"name": "注释对齐", "key": ("table", "noteAlign"), "kind": "enum", "options": ALIGN_OPTIONS},
         {"name": "内容字号", "key": ("table", "cellSize"), "kind": "size", "options": SIZE_OPTIONS},
         {"name": "内容对齐", "key": ("table", "cellAlign"), "kind": "enum",
          "options": CELL_ALIGN_OPTIONS},
         {"name": "上下位置", "key": ("table", "cellVAlign"), "kind": "enum", "options": VALIGN_OPTIONS},
         {"name": "内容行距", "key": ("table", "cellLineSpacing"), "kind": "number",
          "unit": "倍", "lo": 1.0, "hi": 3.0, "step": 0.5},
         {"name": "表头加粗", "key": ("table", "headerBold"), "kind": "bool"},
         {"name": "编号开头", "key": ("table", "prefix"), "kind": "enum",
          "options": TABLE_PREFIX_OPTIONS},
     ]},
    {"id": "references", "title": "参考文献", "icon": "library",
     "about": "参考文献列表的样式；条目缩进指第二条起相对第一条缩进的距离。",
     "rows": [
         {"name": "中文字体", "key": ("references", "font"), "kind": "enum", "options": FONT_ZH_OPTIONS},
         {"name": "英文字体", "key": ("references", "asciiFont"), "kind": "enum",
          "options": FONT_EN_OPTIONS},
         {"name": "字号", "key": ("references", "size"), "kind": "size", "options": SIZE_OPTIONS},
         {"name": "行距", "key": ("references", "lineSpacing"), "kind": "number",
          "unit": "倍", "lo": 1.0, "hi": 3.0, "step": 0.5},
         {"name": "对齐", "key": ("references", "align"), "kind": "enum", "options": ALIGN_OPTIONS},
         {"name": "条目缩进", "key": ("references", "hangingIndentPt"), "kind": "number",
          "unit": "磅", "lo": 0, "hi": 60, "step": 1},
     ]},
]


def spec_default(spec):
    if spec["kind"] == "bool":
        return False
    if spec["kind"] == "number":
        return spec.get("lo", 0)
    if "options" in spec:
        return spec["options"][0]
    return ""


def fmt_number(value):
    try:
        number = float(value)
    except (TypeError, ValueError):
        return str(value)
    if number == int(number):
        return str(int(number))
    return str(number)


def humanize_value(spec, raw):
    """当前生效值的人话展示：附单位说明（磅/厘米/倍/字），字号补磅值对照。"""
    kind = spec["kind"]
    if kind == "bool":
        return "开" if raw else "关"
    if kind == "number":
        return fmt_number(raw) + " " + spec.get("unit", "")
    if kind == "size":
        text = str(raw)
        pt = SIZE_NAME_TO_PT.get(text)
        if pt is not None:
            return text + " · " + fmt_number(pt) + " 磅"
        return text
    return ui_theme.display_label(str(raw))


def raw_equal(spec, first, second):
    """两个配置取值是否等价（数字按值比较，枚举按字符串比较）。"""
    if first is None or second is None:
        return first is None and second is None
    if spec["kind"] == "number":
        try:
            return abs(float(first) - float(second)) < 1e-9
        except (TypeError, ValueError):
            return str(first) == str(second)
    if spec["kind"] == "bool":
        return bool(first) == bool(second)
    return str(first) == str(second)


def cfg_get(cfg, chain, default=None):
    current = cfg
    for key in chain:
        if not isinstance(current, dict) or key not in current:
            return default
        current = current[key]
    return current


class App:
    def __init__(self, root):
        self.root = root
        self.cfg = {}
        self.vars = {}
        self._fonts = {}
        self._icon_family = None
        self._advanced_visible = False
        self._busy = False
        self._worker_queue = queue.Queue()
        self._open_target = None
        self._results = {"改动": [], "提醒": [], "记录": []}
        self._active_segment = "改动"
        self._segment_buttons = {}
        self._segment_icons = {}
        self._action_widgets = []
        self._panel_controls = {}
        self._rows = []
        self._groups = []
        self._baseline_cfg = {}
        self._preset_active = False
        self._menubar = None
        self._filter_focused = False

        self._input_path = tk.StringVar()
        self._output_path = tk.StringVar()
        self._file_name = tk.StringVar(value="")
        self._file_meta = tk.StringVar(value="")
        self._status = tk.StringVar(value=STATUS_READY)
        self._preset_var = tk.StringVar(value=CUSTOM_PRESET)
        self._presets = []
        self._filter_var = tk.StringVar(value="")
        self._output_ghost = None
        self._output_entry = None

        self._load_config()
        self._fonts = ui_theme.apply_theme(self.root)
        self._icon_family = self._fonts.get("icon")
        self._icons = icon_slot.IconRenderer(self.root, self._icon_family)
        self._configure_root()
        self._build_menu()
        self._build_ui()
        self._bind_keys()
        self._select_segment("改动")

    # ---- 配置读写 ----
    def _load_config(self):
        path = config_path()
        if not os.path.exists(path):
            path = default_config_path()
        if not os.path.exists(path):
            self.cfg = {}
        else:
            self.cfg = load_config_file(path)
        self._preset_active = False
        self._baseline_cfg = self._builtin_default_cfg()

    @staticmethod
    def _builtin_default_cfg():
        path = default_config_path()
        if not os.path.exists(path):
            return {}
        try:
            return load_config_file(path)
        except (OSError, ValueError):
            return {}

    def _save_config(self):
        save_config_file(self.cfg, config_path())

    # ---- 窗口与菜单 ----
    def _configure_root(self):
        self.root.title(f"文档整理 v{__version__}")
        self.root.minsize(980, 720)
        width, height = 1120, 860
        self.root.update_idletasks()
        x = max(0, (self.root.winfo_screenwidth() - width) // 2)
        y = max(0, (self.root.winfo_screenheight() - height) // 2 - 24)
        self.root.geometry(f"{width}x{height}+{x}+{y}")
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)

    def _build_menu(self):
        colors = ui_theme.COLORS
        self._menubar = tk.Menu(self.root, tearoff=0)
        more = tk.Menu(self._menubar, tearoff=0)
        for menu in (self._menubar, more):
            menu.configure(bg=colors["surface"], fg=colors["ink"],
                           activebackground=colors["surface_hover"], activeforeground=colors["ink"],
                           bd=0, relief="flat", font=self._fonts["body"])
        more.add_command(label="生成示例文档", command=self._on_generate)
        more.add_command(label="保存设置", command=self._on_save_cfg)
        more.add_separator()
        more.add_command(label="恢复默认设置", command=self._on_reset)
        more.add_command(label="关于文档整理", command=self._show_about)
        more.add_separator()
        more.add_command(label="退出", command=self.root.destroy)
        self._menubar.add_cascade(label="更多", menu=more)
        self.root.configure(menu=self._menubar)

    # ---- 界面构建 ----
    def _build_ui(self):
        shell = ttk.Frame(self.root, padding=(24, 24, 24, 16))
        shell.grid(row=0, column=0, sticky="nsew")

        self._build_header(shell)
        self._build_main_card(shell)
        self._build_action_row(shell)
        self._build_status_bar(shell)
        self._build_result_area(shell)
        self._build_advanced_panel()

    def _build_header(self, parent):
        colors = ui_theme.COLORS
        head = ttk.Frame(parent)
        head.pack(fill="x")
        title_box = ttk.Frame(head)
        title_box.pack(side="left")
        ttk.Label(title_box, text="文档整理", style="Display.TLabel").pack(anchor="w")
        ttk.Label(title_box, text="把论文排成学校要求的格式", style="Caption.TLabel").pack(anchor="w")
        self._advanced_button_frame = ttk.Frame(head)
        self._advanced_button_frame.pack(side="right", pady=(8, 0))
        self._advanced_button_icon = self._make_icon_cell(
            self._advanced_button_frame, "settings", ui_theme.MEASUREMENTS["icon_row"],
            bg=colors["bg"])
        if self._advanced_button_icon is not None:
            self._advanced_button_icon.canvas.pack(side="left", padx=(0, 4))
        self._advanced_button = ttk.Button(self._advanced_button_frame, text="高级设置",
                                           style="Quiet.TButton", command=self._toggle_advanced)
        self._advanced_button.pack(side="left")
        self._action_widgets.append(self._advanced_button)

    def _build_main_card(self, parent):
        """主流程白卡：拖入论文、选学校模板、确认保存位置。"""
        colors = ui_theme.COLORS
        m = ui_theme.MEASUREMENTS
        card = ttk.Frame(parent, style="Card.TFrame", padding=(16, 16))
        card.pack(fill="x", pady=(16, 0))

        self._drop = tk.Canvas(card, height=m["drop_height"], highlightthickness=0,
                               cursor="hand2", bg=colors["surface"], takefocus=1)
        self._drop.pack(fill="x")
        self._drop.bind("<Configure>", lambda event: self._draw_drop())
        self._drop.bind("<Button-1>", lambda event: self._browse_input())
        self._drop.bind("<Return>", self._drop_activate)
        self._drop.bind("<space>", self._drop_activate)
        self._drop.bind("<FocusIn>", lambda event: self._draw_drop())
        self._drop.bind("<FocusOut>", lambda event: self._draw_drop())
        self._drop.bind("<Enter>", lambda event: self._draw_drop(hover=True))
        self._drop.bind("<Leave>", lambda event: self._draw_drop())
        self._drop.drop_target_register(DND_FILES)
        self._drop.dnd_bind("<<Drop>>", self._on_drop)
        self._drop.dnd_bind("<<DropEnter>>", lambda event: self._draw_drop(drag=True))
        self._drop.dnd_bind("<<DropLeave>>", lambda event: self._draw_drop())

        info = ttk.Frame(card, style="Card.TFrame")
        info.pack(fill="x", pady=(8, 0))
        doc_icon = self._make_icon_cell(info, "document", m["icon_row"])
        if doc_icon is not None:
            doc_icon.canvas.pack(side="left", padx=(0, 8))
        ttk.Label(info, textvariable=self._file_name, style="CardSection.TLabel").pack(side="left")
        self._select_button = ttk.Button(info, text="选择文档…", command=self._browse_input)
        self._select_button.pack(side="right")
        self._action_widgets.append(self._select_button)
        ttk.Label(card, textvariable=self._file_meta, style="CardCaption.TLabel",
                  wraplength=m["wrap_main"]).pack(anchor="w", pady=(4, 0))

        ttk.Separator(card).pack(fill="x", pady=(8, 8))

        # 学校模板
        self._presets = list_presets()
        self._step_row(card, "library", "学校模板", self._build_preset_control,
                       "选好学校，格式就按学校的规定来；没有的学校选「自定义」。")
        # 保存位置
        self._step_row(card, "save", "保存位置", self._build_output_control,
                       "整理结果保存为新文件，原来的论文不会被改动。")

    def _step_row(self, parent, icon_name, name, control_builder, caption):
        """主流程步骤行：图标 + 步骤名（固定列宽）+ 尾部控件，说明一句话换行放行下。"""
        colors = ui_theme.COLORS
        m = ui_theme.MEASUREMENTS
        frame = tk.Frame(parent, bg=colors["surface"])
        frame.pack(fill="x")
        icon = self._make_icon_cell(frame, icon_name, m["icon_row"])
        if icon is not None:
            icon.canvas.place(x=ui_theme.SPACING["sm"] + ui_theme.SPACING["xs"],
                              rely=0.5, anchor="w")
        label = tk.Label(frame, text=name, font=self._fonts["body_bold"],
                         fg=colors["ink"], bg=colors["surface"], anchor="w")
        label.place(x=m["indent_step"], rely=0.5, anchor="w", width=m["name_column"])
        control_builder(frame)
        tip = tk.Label(parent, text=caption, font=self._fonts["caption"],
                       fg=colors["ink_muted"], bg=colors["surface"], anchor="w",
                       justify="left", wraplength=m["wrap_main"])
        tip.pack(anchor="w", padx=(m["indent_step"], 0), pady=(ui_theme.SPACING["xs"], 0))

    def _build_preset_control(self, frame):
        names = [CUSTOM_PRESET] + [preset["name"] for preset in self._presets]
        combo = ttk.Combobox(frame, textvariable=self._preset_var, values=names,
                             state="readonly", width=28, font=self._fonts["body"])
        combo.pack(side="right", padx=(0, ui_theme.SPACING["md"]))
        combo.bind("<<ComboboxSelected>>", lambda event: self._on_preset_selected())
        self._register_panel_control(combo, "readonly")

    def _build_output_control(self, frame):
        colors = ui_theme.COLORS
        m = ui_theme.MEASUREMENTS
        spacer = tk.Frame(frame, bg=colors["surface"], width=m["indent_step"] + m["name_column"])
        spacer.pack(side="left")
        entry = ttk.Entry(frame, textvariable=self._output_path, font=self._fonts["body"])
        entry.pack(side="left", fill="x", expand=True, padx=(0, ui_theme.SPACING["sm"]))
        self._output_ghost = tk.Label(frame, text="选好论文后自动填写",
                                      font=self._fonts["caption"], fg=colors["ink_disabled"],
                                      bg=colors["surface"])
        self._output_ghost.place(in_=entry, x=ui_theme.SPACING["sm"], rely=0.5, anchor="w")
        self._output_entry = entry
        self._sync_output_ghost()
        save_as = ttk.Button(frame, text="更改…", command=self._browse_normalized_out)
        save_as.pack(side="left", padx=(0, ui_theme.SPACING["md"]))
        self._action_widgets.append(save_as)

    def _sync_output_ghost(self):
        if self._output_ghost is None or not self._output_ghost.winfo_exists():
            return
        if self._output_path.get().strip():
            self._output_ghost.place_forget()
        else:
            self._output_ghost.place(in_=self._output_entry, x=ui_theme.SPACING["sm"],
                                     rely=0.5, anchor="w")

    def _build_action_row(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(16, 0))
        ttk.Label(row, text="预览只列出改动，不会改动文件。", style="Caption.TLabel").pack(side="left")
        self._preview_button = ttk.Button(row, text="预览改动", command=self._on_preview)
        self._preview_button.pack(side="right")
        self._build_normalize_button(row).pack(side="right", padx=(0, ui_theme.SPACING["sm"]))
        self._action_widgets.append(self._preview_button)

    def _build_normalize_button(self, parent):
        """主按钮：图标字体 Play + 文字合成一个整体，蓝底只此一处。"""
        colors = ui_theme.COLORS
        m = ui_theme.MEASUREMENTS
        shell = tk.Frame(parent, bg=colors["accent"])
        self._normalize_shell = shell
        self._normalize_icon = self._make_icon_cell(shell, "play", m["icon_row"],
                                                    fg=colors["surface"])
        if self._normalize_icon is not None:
            self._normalize_icon.canvas.pack(
                side="left", padx=(ui_theme.SPACING["lg"], 0),
                pady=(ui_theme.SPACING["sm"], ui_theme.SPACING["sm"]))
        self._normalize_button = ttk.Button(shell, text="开始整理", style="AccentLead.TButton",
                                            command=self._on_normalize)
        self._normalize_button.pack(side="left", padx=(0, ui_theme.SPACING["lg"]))
        self._action_widgets.append(self._normalize_button)
        self._sync_accent_shell()
        for target in (shell, self._normalize_button):
            target.bind("<Enter>", lambda event: self._sync_accent_shell("active"))
            target.bind("<Leave>", lambda event: self._sync_accent_shell())
        return shell

    def _sync_accent_shell(self, visual=None):
        """主按钮三件套（底、图标、按钮）同色；指针离开整体才回落常态。"""
        colors = ui_theme.COLORS
        if visual is None:
            pointer = self._normalize_shell.winfo_containing(
                self._normalize_shell.winfo_pointerx(), self._normalize_shell.winfo_pointery())
            if pointer is not None and self._widget_inside(pointer, self._normalize_shell):
                visual = "active"
            else:
                visual = "normal"
        if self._normalize_button.instate(["disabled"]):
            bg, fg = colors["accent_disabled_bg"], colors["accent_disabled_fg"]
        elif visual == "active":
            bg, fg = colors["accent_hover"], colors["surface"]
        else:
            bg, fg = colors["accent"], colors["surface"]
        self._normalize_shell.configure(bg=bg)
        if self._normalize_icon is not None:
            self._normalize_icon.set_bg(bg)
            self._normalize_icon.set_color(fg)

    def _build_result_area(self, parent):
        colors = ui_theme.COLORS
        area = ttk.Frame(parent, style="Card.TFrame", padding=(16, 12))
        area.pack(fill="both", expand=True, pady=(16, 0))
        seg_row = ttk.Frame(area, style="Card.TFrame")
        seg_row.pack(fill="x")
        for name, icon_name in (("改动", "edit"), ("提醒", "warning"), ("记录", "list")):
            pair = ttk.Frame(seg_row, style="Card.TFrame")
            pair.pack(side="left", padx=(0, ui_theme.SPACING["xs"]))
            icon = self._make_icon_cell(pair, icon_name, ui_theme.MEASUREMENTS["icon_row"])
            if icon is not None:
                icon.canvas.pack(side="left", padx=(ui_theme.SPACING["xs"], 0))
                icon.canvas.bind("<Button-1>",
                                 lambda event, item=name: self._select_segment(item))
                self._segment_icons[name] = icon
            button = ttk.Button(pair, text=name, style="Seg.TButton",
                                command=lambda item=name: self._select_segment(item))
            button.pack(side="left")
            self._segment_buttons[name] = button
        body = ttk.Frame(area, style="Card.TFrame")
        body.pack(fill="both", expand=True, pady=(8, 0))
        self._text = tk.Text(body, wrap="word", relief="flat", highlightthickness=1,
                             highlightbackground=colors["border"], highlightcolor=colors["accent"],
                             bg=colors["surface"], fg=colors["ink"],
                             insertbackground=colors["ink"],
                             selectbackground=colors["surface_pressed"], selectforeground=colors["ink"],
                             font=self._fonts["body"], padx=12, pady=12,
                             spacing1=4, spacing3=4, state="disabled")
        scrollbar = ttk.Scrollbar(body, orient="vertical", command=self._text.yview)
        self._text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self._text.pack(side="left", fill="both", expand=True)
        self._text.tag_configure("change", foreground=colors["ink"])
        self._text.tag_configure("warning", foreground=colors["warning"])
        self._text.tag_configure("error", foreground=colors["danger"])
        self._text.tag_configure("meta", foreground=colors["ink_muted"])
        self._text.tag_configure("heading", foreground=colors["ink"], font=self._fonts["body_bold"])

    def _build_status_bar(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(side="bottom", fill="x", pady=(12, 0))
        ttk.Separator(bar).pack(fill="x", pady=(0, 8))
        inner = ttk.Frame(bar)
        inner.pack(fill="x")
        self._status_label = ttk.Label(inner, textvariable=self._status, style="Muted.TLabel")
        self._status_label.pack(side="left")
        open_frame = ttk.Frame(inner)
        open_frame.pack(side="right")
        self._open_icon = self._make_icon_cell(open_frame, "folder",
                                               ui_theme.MEASUREMENTS["icon_row"],
                                               bg=ui_theme.COLORS["bg"])
        if self._open_icon is not None:
            self._open_icon.canvas.pack(side="left", padx=(0, ui_theme.SPACING["xs"]))
            self._open_icon.canvas.bind("<Button-1>",
                                        lambda event: self._open_folder_clicked())
        self._open_button = ttk.Button(open_frame, text="打开所在文件夹", style="Quiet.TButton",
                                       command=self._open_folder_clicked)
        self._progress = ttk.Progressbar(inner, mode="indeterminate", length=140)

    # ---- 图标辅助 ----
    def _make_icon_cell(self, parent, name, size, bg=None, fg=None):
        """图标槽（唯一排版入口，见 icon_slot.py）；字体/字形缺失返回 None，
        调用方跳过布局。返回 IconCell：摆放用 .canvas，换色/换形用 set_* 方法。"""
        colors = ui_theme.COLORS
        return self._icons.cell(parent, name, size,
                                colors["surface"] if bg is None else bg,
                                colors["ink_muted"] if fg is None else fg)

    # ---- 高级设置面板 ----
    def _build_advanced_panel(self):
        colors = ui_theme.COLORS
        m = ui_theme.MEASUREMENTS
        self._advanced_panel = ttk.Frame(self.root, style="Panel.TFrame")
        ttk.Separator(self._advanced_panel, orient="vertical").pack(side="left", fill="y")
        content = ttk.Frame(self._advanced_panel, style="Panel.TFrame", padding=(16, 16, 12, 12))
        content.pack(side="left", fill="both", expand=True)

        head = ttk.Frame(content, style="Panel.TFrame")
        head.pack(fill="x")
        ttk.Label(head, text="高级设置", style="PanelSection.TLabel").pack(side="left")
        ttk.Button(head, text="收起", style="PanelQuiet.TButton",
                   command=self._toggle_advanced).pack(side="right")

        filter_row = ttk.Frame(content, style="Panel.TFrame")
        filter_row.pack(fill="x", pady=(8, 0))
        search_icon = self._make_icon_cell(filter_row, "search", m["icon_row"],
                                           bg=colors["bg"])
        if search_icon is not None:
            search_icon.canvas.pack(side="left", padx=(0, ui_theme.SPACING["xs"]))
        self._filter_entry = ttk.Entry(filter_row, textvariable=self._filter_var,
                                       font=self._fonts["body"])
        self._filter_entry.pack(side="left", fill="x", expand=True)
        self._filter_ghost = tk.Label(filter_row, text=FILTER_PLACEHOLDER,
                                      font=self._fonts["caption"], fg=colors["ink_disabled"],
                                      bg=colors["surface"])
        self._filter_ghost.place(in_=self._filter_entry, x=ui_theme.SPACING["sm"],
                                 rely=0.5, anchor="w")
        self._filter_entry.bind("<FocusIn>", lambda event: self._set_filter_focus(True))
        self._filter_entry.bind("<FocusOut>", lambda event: self._set_filter_focus(False))
        self._filter_entry.bind("<Escape>", lambda event: self._filter_var.set(""))
        self._filter_var.trace_add("write", self._on_filter_change)

        ttk.Label(content, text="一般不用改，改动会影响之后整理的文件。",
                  style="PanelCaption.TLabel", wraplength=m["wrap_panel"]).pack(
            anchor="w", pady=(4, 8))

        holder = ttk.Frame(content, style="Panel.TFrame")
        holder.pack(fill="both", expand=True)
        self._panel_canvas = tk.Canvas(holder, bg=colors["bg"], highlightthickness=0,
                                       width=m["panel_width"])
        scrollbar = ttk.Scrollbar(holder, orient="vertical", command=self._panel_canvas.yview)
        self._panel_canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self._panel_canvas.pack(side="left", fill="both", expand=True)
        inner = ttk.Frame(self._panel_canvas, style="Panel.TFrame")
        window = self._panel_canvas.create_window((0, 0), window=inner, anchor="nw")
        inner.bind("<Configure>", lambda event: self._panel_canvas.configure(
            scrollregion=self._panel_canvas.bbox("all")))
        self._panel_canvas.bind("<Configure>", lambda event: self._panel_canvas.itemconfigure(
            window, width=event.width))

        self._filter_empty = ttk.Label(inner, text=FILTER_EMPTY_TEXT, style="PanelCaption.TLabel")
        for group in SETTING_GROUPS:
            self._build_setting_group(inner, group)
        self._bind_wheel_local(self._advanced_panel)
        self._sync_filter_ghost()

    def _bind_wheel_local(self, widget):
        """滚轮只在面板内生效：逐控件局部绑定，不用 bind_all 全局接管。"""
        widget.bind("<MouseWheel>", self._on_panel_wheel)
        for child in widget.winfo_children():
            self._bind_wheel_local(child)

    def _on_panel_wheel(self, event):
        self._panel_canvas.yview_scroll(int(-event.delta / 120), "units")

    def _toggle_advanced(self):
        if self._advanced_visible:
            self._advanced_panel.grid_forget()
            self._advanced_button.configure(text="高级设置")
            self._set_header_icon("settings")
            self._advanced_visible = False
        else:
            self._advanced_panel.grid(row=0, column=1, sticky="ns")
            self._advanced_button.configure(text="收起")
            self._set_header_icon("chevron_right")
            self._advanced_visible = True

    def _set_header_icon(self, name):
        if self._advanced_button_icon is None:
            return
        if ui_icons.glyph(name):
            self._advanced_button_icon.set_icon(name)

    def _build_setting_group(self, parent, group):
        colors = ui_theme.COLORS
        m = ui_theme.MEASUREMENTS
        card = tk.Frame(parent, bg=colors["surface"])
        card.pack(fill="x", pady=(0, m["group_gap"]))
        header = tk.Frame(card, bg=colors["surface"], height=m["group_header_height"])
        header.pack(fill="x")
        header.pack_propagate(False)
        icon = self._make_icon_cell(header, group["icon"], m["icon_group"], fg=colors["ink"])
        if icon is not None:
            icon.canvas.place(x=ui_theme.SPACING["md"], rely=0.5, anchor="w")
        title = tk.Label(header, text=group["title"], font=self._fonts["body_bold"],
                         fg=colors["ink"], bg=colors["surface"], anchor="w")
        title.place(x=m["indent_group"], rely=0.5, anchor="w")
        reset = ttk.Button(header, text=RESET_GROUP_TEXT, style="CardQuiet.TButton",
                           command=lambda item=group: self._reset_group(item))
        reset.pack(side="right", padx=(0, ui_theme.SPACING["md"]))
        if group.get("about"):
            about = tk.Label(card, text=group["about"], font=self._fonts["caption"],
                             fg=colors["ink_muted"], bg=colors["surface"], anchor="w",
                             justify="left", wraplength=m["wrap_panel"])
            about.pack(fill="x", padx=(m["indent_group"], ui_theme.SPACING["md"]))
        body = tk.Frame(card, bg=colors["surface"])
        body.pack(fill="x")

        entries = []
        frames = []
        rows = group["rows"]
        for index, spec in enumerate(rows):
            entry, frame = self._build_setting_row(body, spec, group, with_separator=index + 1 < len(rows))
            entries.append(entry)
            frames.append(frame)
        self._groups.append({"spec": group, "card": card, "entries": entries, "frames": frames})

    def _build_setting_row(self, parent, spec, group, with_separator):
        colors = ui_theme.COLORS
        m = ui_theme.MEASUREMENTS
        chain = list(spec["key"])
        var_key = ".".join(chain)
        raw = self._get(chain, spec_default(spec))
        if spec["kind"] == "bool":
            var = tk.BooleanVar(value=bool(raw))
        else:
            var = tk.StringVar(value=self._display_of(spec, raw))
        self.vars[var_key] = var

        row = tk.Frame(parent, bg=colors["surface"], height=m["row_height"])
        row.pack(fill="x")
        row.pack_propagate(False)
        hover_targets = [row]

        icon = self._make_icon_cell(row, group["icon"], m["icon_row"])
        if icon is not None:
            icon.canvas.place(x=ui_theme.SPACING["md"], rely=0.5, anchor="w")
            hover_targets.append(icon.canvas)
        name = tk.Label(row, text=spec["name"], font=self._fonts["body"], fg=colors["ink"],
                        bg=colors["surface"], anchor="w")
        name.place(x=m["indent_step"], rely=0.5, anchor="w", width=m["name_column"])
        hover_targets.append(name)
        tag = tk.Label(row, font=self._fonts["caption"], fg=colors["ink_muted"],
                       bg=colors["surface"], anchor="w")
        tag.place(x=m["indent_step"] + m["name_column"] + ui_theme.SPACING["sm"], rely=0.5,
                  anchor="w")
        hover_targets.append(tag)
        value = tk.Label(row, font=self._fonts["body"], fg=colors["ink_muted"],
                         bg=colors["surface"], anchor="e")
        hover_targets.append(value)

        control = None
        if spec["kind"] == "bool":
            control = ttk.Checkbutton(row, variable=var, style="Panel.TCheckbutton")
            control.pack(side="right", padx=(0, ui_theme.SPACING["md"]))
            value.pack_forget()
            hover_targets.remove(value)
        elif spec["kind"] == "number":
            control = ttk.Spinbox(row, from_=spec["lo"], to=spec["hi"],
                                  increment=spec.get("step", 1), textvariable=var, width=8)
            control.pack(side="right", padx=(ui_theme.SPACING["sm"], ui_theme.SPACING["md"]))
            value.pack(side="right")
            validate = (row.register(self._number_ok), "%P")
            control.configure(validate="key", validatecommand=validate)
            control.bind("<FocusOut>", lambda event, item=spec: self._clamp_number(item, var))
        else:
            options_text = " ".join(spec.get("options", []))
            width = 16 if "New Roman" in options_text else 12
            control = ttk.Combobox(row, textvariable=var,
                                   values=ui_theme.option_values(spec.get("options", [])),
                                   state="readonly", width=width)
            control.pack(side="right", padx=(ui_theme.SPACING["sm"], ui_theme.SPACING["md"]))
            value.pack(side="right")
            self._register_panel_control(control, "readonly")
        if control is not None and spec["kind"] != "enum":
            self._register_panel_control(control, "normal")

        if with_separator:
            separator = tk.Frame(row, bg=colors["border"], height=1)
            separator.pack(side="bottom", fill="x", padx=(m["separator_inset"], 0))

        self._bind_row_hover(hover_targets, row)
        entry = {"spec": spec, "group": group, "var": var, "frame": row, "tag": tag,
                 "value": value, "control": control}
        var.trace_add("write", lambda *args, item=entry: self._refresh_row_state(item))
        self._rows.append(entry)
        self._refresh_row_state(entry)
        return entry, row

    @staticmethod
    def _number_ok(proposed):
        if proposed == "":
            return True
        if proposed.count(".") > 1:
            return False
        return all(char in "0123456789." for char in proposed)

    def _clamp_number(self, spec, var):
        try:
            number = float(var.get())
        except (TypeError, ValueError):
            var.set(fmt_number(self._get(list(spec["key"]), spec_default(spec))))
            return
        clamped = min(max(number, spec["lo"]), spec["hi"])
        var.set(fmt_number(clamped))

    def _display_of(self, spec, raw):
        if spec["kind"] == "number":
            return fmt_number(raw)
        if spec["kind"] == "bool":
            return ""
        return ui_theme.display_label(str(raw))

    def _bind_row_hover(self, targets, row):
        colors = ui_theme.COLORS

        def enter(_event):
            for widget in targets:
                widget.configure(bg=colors["surface_hover"])

        def leave(_event):
            pointer = row.winfo_containing(row.winfo_pointerx(), row.winfo_pointery())
            if pointer is not None and self._widget_inside(pointer, row):
                return
            for widget in targets:
                widget.configure(bg=colors["surface"])

        for widget in targets:
            widget.bind("<Enter>", enter)
            widget.bind("<Leave>", leave)

    @staticmethod
    def _widget_inside(widget, ancestor):
        current = widget
        while current is not None:
            if current is ancestor:
                return True
            current = getattr(current, "master", None)
        return False

    def _register_panel_control(self, control, normal_state):
        self._panel_controls[control] = normal_state

    def _refresh_row_state(self, entry):
        spec = entry["spec"]
        current = self._var_raw(spec, entry["var"])
        baseline = cfg_get(self._baseline_cfg, spec["key"], spec_default(spec))
        colors = ui_theme.COLORS
        if not raw_equal(spec, current, baseline):
            entry["tag"].configure(text=TAG_MODIFIED, fg=colors["accent"])
        elif self._preset_active:
            entry["tag"].configure(text=TAG_FROM_TEMPLATE, fg=colors["ink_muted"])
        else:
            entry["tag"].configure(text="", fg=colors["ink_muted"])
        if spec["kind"] != "bool":
            entry["value"].configure(text=humanize_value(spec, current))

    def _var_raw(self, spec, var):
        if spec["kind"] == "bool":
            return bool(var.get())
        text = str(var.get())
        if spec["kind"] in ("enum", "size"):
            return ui_theme.LABEL_VALUES.get(text, text)
        try:
            number = float(text)
            return int(number) if number == int(number) else number
        except ValueError:
            return text

    def _reset_group(self, group):
        for entry in self._rows:
            if entry["group"] is not group:
                continue
            spec = entry["spec"]
            baseline = cfg_get(self._baseline_cfg, spec["key"], spec_default(spec))
            if spec["kind"] == "bool":
                entry["var"].set(bool(baseline))
            elif spec["kind"] == "number":
                entry["var"].set(fmt_number(baseline))
            else:
                entry["var"].set(ui_theme.display_label(str(baseline)))
        self._status.set("已恢复「" + group["title"] + "」的默认设置。")

    # ---- 过滤框 ----
    def _set_filter_focus(self, focused):
        self._filter_focused = focused
        self._sync_filter_ghost()

    def _on_filter_change(self, *args):
        self._sync_filter_ghost()
        self._apply_filter()

    def _sync_filter_ghost(self):
        text = self._filter_var.get()
        if text or self._filter_focused:
            self._filter_ghost.place_forget()
        else:
            self._filter_ghost.place(in_=self._filter_entry, x=ui_theme.SPACING["sm"],
                                     rely=0.5, anchor="w")

    def _apply_filter(self):
        query = self._filter_var.get().strip().lower()
        group_frames = [group["card"] for group in self._groups]
        any_visible = False
        for group in self._groups:
            group_visible = False
            for entry in group["entries"]:
                spec = entry["spec"]
                haystack = " ".join([
                    group["spec"]["title"], spec["name"],
                    humanize_value(spec, self._var_raw(spec, entry["var"])),
                ]).lower()
                visible = not query or query in haystack
                if visible:
                    self._pack_in_order(entry["frame"], group["frames"])
                    group_visible = True
                else:
                    entry["frame"].pack_forget()
            if group_visible:
                self._pack_in_order(group["card"], group_frames)
                any_visible = True
            else:
                group["card"].pack_forget()
        if any_visible or not query:
            self._filter_empty.pack_forget()
        else:
            self._filter_empty.pack(anchor="w", pady=(0, ui_theme.SPACING["sm"]))

    @staticmethod
    def _pack_in_order(widget, order):
        """pack_forget 后按原始顺序回到原位（pack() 默认会排到队尾）。"""
        index = order.index(widget)
        for later in order[index + 1:]:
            if later.winfo_manager() == "pack":
                widget.pack(before=later, fill="x")
                return
        for earlier in reversed(order[:index]):
            if earlier.winfo_manager() == "pack":
                widget.pack(after=earlier, fill="x")
                return
        widget.pack(fill="x")

    # ---- 拖拽与画布 ----
    def _drop_activate(self, event):
        self._browse_input()
        return "break"

    def _draw_drop(self, hover=False, drag=False):
        canvas = self._drop
        canvas.delete("all")
        colors = ui_theme.COLORS
        m = ui_theme.MEASUREMENTS
        width = max(canvas.winfo_width(), 200)
        height = int(canvas["height"])
        focus = self.root.focus_get() is canvas
        if drag:
            bg, outline, line_width, icon_color = (
                colors["accent_soft"], colors["accent"], 2, colors["accent"])
        elif hover:
            bg, outline, line_width, icon_color = (
                colors["surface_hover"], colors["border_strong"], 1, colors["ink"])
        elif focus:
            bg, outline, line_width, icon_color = (
                colors["surface"], colors["accent"], 2, colors["ink_muted"])
        else:
            bg, outline, line_width, icon_color = (
                colors["surface"], colors["border_strong"], 1, colors["ink_muted"])
        canvas.configure(bg=bg)
        canvas.create_rectangle(1, 1, width - 2, height - 2, outline=outline, width=line_width)
        path = self._input_path.get().strip()
        title = self._shorten_path(os.path.basename(path)) if path else DROP_TITLE
        subtitle = DROP_SELECTED_SUBTITLE if path else DROP_SUBTITLE
        if self._icon_family:
            pos = self._icons.text_pos(width / 2, height / 2 - 20, "document",
                                       m["icon_hero"])
            if pos:
                canvas.create_text(pos[0], pos[1], text=ui_icons.glyph("document"),
                                   fill=icon_color,
                                   font=(self._icon_family,
                                         -self._icons.size_px(m["icon_hero"])),
                                   anchor="nw")
            canvas.create_text(width / 2, height / 2 + 12, text=title, fill=colors["ink"],
                               font=self._fonts["body_bold"])
            canvas.create_text(width / 2, height / 2 + 32, text=subtitle,
                               fill=colors["ink_muted"], font=self._fonts["caption"])
        else:
            canvas.create_text(width / 2, height / 2 - 8, text=title, fill=colors["ink"],
                               font=self._fonts["body_bold"])
            canvas.create_text(width / 2, height / 2 + 12, text=subtitle,
                               fill=colors["ink_muted"], font=self._fonts["caption"])

    @staticmethod
    def _shorten_path(text, limit=64):
        if len(text) <= limit:
            return text
        return "…" + text[-(limit - 1):]

    def _on_drop(self, event):
        self._draw_drop()
        if self._busy:
            return
        paths = self.root.tk.splitlist(event.data)
        if not paths:
            return
        path = paths[0]
        if not path.lower().endswith(".docx"):
            messagebox.showwarning(
                "文件类型不支持",
                "只支持 Word 文档（.docx）。\n请在 Word 里把这份文件另存为 .docx，再拖进来。",
            )
            return
        self._set_input(path)

    def _set_input(self, path):
        self._input_path.set(path)
        if not self._output_path.get().strip():
            self._output_path.set(make_normalized_output_path(path))
        self._sync_output_ghost()
        self._file_name.set(os.path.basename(path))
        self._file_meta.set(STATUS_BUSY_INSPECT)
        self._draw_drop()
        self._clear_open_target()
        self._set_status("已选择：" + os.path.basename(path), "info")
        self._run_worker(lambda: inspect_docx(path), self._show_inspection,
                         lambda exc: self._set_status(
                             "读取失败，请确认文件是 Word 文档（.docx）。", "error"))

    def _show_inspection(self, info):
        levels = "、".join(str(level) for level in info.heading_levels) + " 级" if info.heading_levels else "无标题"
        parts = [
            f"段落 {info.paragraph_count}",
            f"表格 {info.table_count}",
            f"图片 {info.image_count}",
            f"标题 {levels}",
            "含页眉页脚" if info.has_header_footer else "无页眉页脚",
        ]
        self._file_meta.set(" · ".join(parts))

    # ---- 文件对话框 ----
    def _browse_input(self):
        if self._busy:
            return
        path = filedialog.askopenfilename(
            title="选择要整理的论文",
            filetypes=[("Word 文档", "*.docx")],
            initialdir=app_dir(),
        )
        if path:
            self._set_input(path)

    def _browse_normalized_out(self):
        if self._busy:
            return
        initial = self._output_path.get().strip()
        if not initial:
            input_path = self._input_path.get().strip()
            initial = make_normalized_output_path(input_path) if input_path else ""
        if not initial:
            initial = os.path.join(app_dir(), "整理文档.docx")
        path = filedialog.asksaveasfilename(
            title="选择保存位置",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
            initialdir=os.path.dirname(initial),
            initialfile=os.path.basename(initial),
        )
        if path:
            self._output_path.set(path)
            self._sync_output_ghost()

    # ---- 后台任务 ----
    def _run_worker(self, work, on_success, on_error):
        if self._busy:
            return
        self._set_busy(True)

        def target():
            try:
                self._worker_queue.put(("ok", work()))
            except Exception as exc:  # noqa: BLE001
                self._worker_queue.put(("error", exc))

        threading.Thread(target=target, daemon=True).start()
        self.root.after(60, lambda: self._poll_worker(on_success, on_error))

    def _poll_worker(self, on_success, on_error):
        try:
            status, payload = self._worker_queue.get_nowait()
        except queue.Empty:
            self.root.after(60, lambda: self._poll_worker(on_success, on_error))
            return
        self._set_busy(False)
        if status == "ok":
            on_success(payload)
        else:
            on_error(payload)

    def _set_busy(self, busy):
        self._busy = busy
        state = "disabled" if busy else "normal"
        for widget in self._action_widgets:
            widget.configure(state=state)
        for control, normal_state in self._panel_controls.items():
            control.configure(state="disabled" if busy else normal_state)
        if busy:
            self._drop.configure(cursor="watch")
            self._progress.pack(side="right")
            self._progress.start(12)
        else:
            self._drop.configure(cursor="hand2")
            self._progress.stop()
            self._progress.pack_forget()
        self._sync_accent_shell("normal")

    def _set_status(self, text, tone):
        self._status.set(text)
        style = {"info": "Muted.TLabel", "success": "Success.TLabel",
                 "error": "Danger.TLabel"}.get(tone, "Muted.TLabel")
        self._status_label.configure(style=style)

    # ---- 主操作 ----
    def _on_normalize(self):
        self._collect()
        self._save_config()
        input_path = self._input_path.get().strip()
        if not input_path:
            messagebox.showwarning("未选择文档", "请先把论文拖进窗口，或点击选择文件。")
            return
        output_path = self._output_path.get().strip()
        if not output_path:
            output_path = make_normalized_output_path(input_path)
            self._output_path.set(output_path)
        self._set_status(STATUS_BUSY_NORMALIZE, "info")

        def work():
            normalize_docx(self.cfg, input_path, output_path, assets_dir=assets_dir())
            return output_path

        def done(path):
            self._set_status("整理完成，已保存为「" + os.path.basename(path) + "」。", "success")
            self._add_result("记录", "已保存到：" + path, "change")
            self._set_open_target(os.path.dirname(path), os.path.basename(path))
            self._render_results()

        def failed(exc):
            self._set_status("整理没有完成", "error")
            self._add_result("提醒", "整理没有完成：" + str(exc), "error")
            self._select_segment("提醒")
            self._show_error("整理没有完成", exc)

        self._run_worker(work, done, failed)

    def _on_preview(self):
        self._collect()
        self._save_config()
        input_path = self._input_path.get().strip()
        if not input_path:
            messagebox.showwarning("未选择文档", "请先把论文拖进窗口，或点击选择文件。")
            return
        output_guess = self._output_path.get().strip() or make_normalized_output_path(input_path)
        self._set_status(STATUS_BUSY_PREVIEW, "info")

        def work():
            return normalize_docx(self.cfg, input_path, output_guess,
                                  dry_run=True, return_result=True)

        def done(result):
            self._results["改动"] = [(item, "change") for item in result.changes]
            self._results["提醒"] = [(item, "warning") for item in result.warnings]
            summary = (
                f"{os.path.basename(result.input_path)}：将处理 {result.paragraphs_processed} 段、"
                f"{result.tables_processed} 个表格、{result.images_processed} 张图片（未写入文件）"
            )
            self._results["记录"].append((summary, "meta"))
            self._set_status("预览完成，文件没有被改动。", "info")
            if result.changes:
                self._select_segment("改动")
            elif result.warnings:
                self._select_segment("提醒")
            else:
                self._select_segment("记录")
            self._render_results()

        def failed(exc):
            self._set_status("预览没有完成", "error")
            self._add_result("提醒", "预览没有完成：" + str(exc), "error")
            self._select_segment("提醒")
            self._show_error("预览没有完成", exc)

        self._run_worker(work, done, failed)

    def _on_generate(self):
        self._collect()
        self._save_config()
        default_name = self._get(["output", "filename"], "示例论文.docx")
        out_path = os.path.join(app_dir(), default_name)
        self._set_status(STATUS_BUSY_GENERATE, "info")

        def work():
            generate(self.cfg, out_path, assets_dir=assets_dir(), content_fn=build_sample_content)
            return out_path

        def done(path):
            self._set_status("示例文档已生成：" + path, "success")
            self._add_result("记录", "已生成：" + path, "change")
            self._set_open_target(os.path.dirname(path), os.path.basename(path))
            self._render_results()

        def failed(exc):
            self._set_status("示例文档没有生成", "error")
            self._add_result("提醒", "示例文档没有生成：" + str(exc), "error")
            self._select_segment("提醒")
            self._show_error("示例文档没有生成", exc)

        self._run_worker(work, done, failed)

    # ---- 模板与设置 ----
    def _on_preset_selected(self):
        name = self._preset_var.get()
        if name == CUSTOM_PRESET:
            self._load_config()
            self._rebuild_ui()
            self._set_status("已切换到自定义设置。", "info")
            return
        preset = next((item for item in self._presets if item["name"] == name), None)
        if preset is None:
            return
        self.cfg = preset["config"]
        self._preset_active = True
        self._baseline_cfg = copy.deepcopy(self.cfg)
        self._rebuild_ui()
        self._set_status("已套用「" + name + "」，点「开始整理」即可。", "info")

    def _on_save_cfg(self):
        self._collect()
        self._save_config()
        self._set_status("设置已保存。", "info")
        self._add_result("记录", "已保存到：" + config_path(), "meta")
        self._render_results()

    def _on_reset(self):
        path = default_config_path()
        if not os.path.exists(path):
            messagebox.showwarning("恢复默认设置", "找不到默认设置文件。重新安装本软件可以恢复。")
            return
        self._load_config()
        self.cfg = load_config_file(path)
        self._preset_var.set(CUSTOM_PRESET)
        self._rebuild_ui()
        self._set_status("已恢复默认设置。", "info")

    def _show_about(self):
        messagebox.showinfo(
            "关于文档整理",
            f"文档整理 v{__version__}\n\n"
            "把论文排成学校要求的格式。\n"
            "整理结果另存为新文件，原来的论文不会被改动。\n\n"
            "常用快捷键：\n"
            "Ctrl+O 选择文档\n"
            "Ctrl+P 预览改动\n"
            "Ctrl+Enter 开始整理\n"
            "Ctrl+S 保存设置\n"
            "Ctrl+G 生成示例文档",
        )

    # ---- 结果区 ----
    def _select_segment(self, name):
        self._active_segment = name
        colors = ui_theme.COLORS
        for key, button in self._segment_buttons.items():
            button.configure(style="SegActive.TButton" if key == name else "Seg.TButton")
        for key, icon in self._segment_icons.items():
            icon.set_color(colors["ink"] if key == name else colors["ink_muted"])
        self._render_results()

    def _render_results(self):
        text = self._text
        text.configure(state="normal")
        text.delete("1.0", "end")
        items = self._results.get(self._active_segment, [])
        if not items:
            text.insert("end", EMPTY_RESULT_TEXT, "meta")
        else:
            for content, tag in items:
                text.insert("end", "· ", "meta")
                text.insert("end", content + "\n", tag or "change")
        text.configure(state="disabled")
        text.see("end")

    def _add_result(self, segment, content, tag=None):
        self._results[segment].append((content, tag))

    def _show_error(self, title, exc):
        if isinstance(exc, NormalizeError):
            messagebox.showerror(
                title, str(exc) + "\n\n可以换一份文档再试，或到「记录」里查看详细信息。")
            return
        detail = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        self._add_result("记录", f"{title}：{exc}", "error")
        self._add_result("记录", detail, "error")
        self._select_segment("记录")
        messagebox.showerror(
            title, f"{exc}\n\n详细信息已放入「记录」，可以复制下来发给帮忙维护的人。")

    # ---- 打开输出位置 ----
    def _set_open_target(self, folder, select=None):
        self._open_target = (folder, select)
        self._open_button.pack(side="right")
        if self._open_icon is not None:
            self._open_icon.pack(side="left", padx=(0, ui_theme.SPACING["xs"]))

    def _clear_open_target(self):
        self._open_target = None
        self._open_button.pack_forget()
        if self._open_icon is not None:
            self._open_icon.pack_forget()

    def _open_folder_clicked(self):
        if self._open_target:
            open_folder(*self._open_target)

    # ---- 配置收集 ----
    def _collect(self):
        for chain_str, var in self.vars.items():
            chain = chain_str.split(".")
            if isinstance(var, tk.BooleanVar):
                value = bool(var.get())
            else:
                raw = var.get()
                if raw in ui_theme.LABEL_VALUES:
                    value = ui_theme.LABEL_VALUES[raw]
                else:
                    try:
                        number = float(raw)
                        value = int(number) if number == int(number) else number
                    except (ValueError, TypeError):
                        value = raw
            self._set(chain, value)

    def _set(self, chain, value):
        current = self.cfg
        for key in chain[:-1]:
            if key not in current or not isinstance(current.get(key), dict):
                current[key] = {}
            current = current[key]
        current[chain[-1]] = value

    def _get(self, chain, default):
        current = self.cfg
        for key in chain:
            if isinstance(current, dict):
                current = current.get(key)
            else:
                return default
            if current is None:
                return default
        return current if current is not None else default

    # ---- 快捷键与重建 ----
    def _bind_keys(self):
        self.root.bind("<Control-o>", self._shortcut(self._browse_input))
        self.root.bind("<Control-p>", self._shortcut(self._on_preview))
        self.root.bind("<Control-Return>", self._shortcut(self._on_normalize))
        self.root.bind("<Control-s>", self._shortcut(self._on_save_cfg))
        self.root.bind("<Control-g>", self._shortcut(self._on_generate))

    def _shortcut(self, handler):
        def wrapped(event):
            handler()
            return "break"
        return wrapped

    def _rebuild_ui(self):
        for child in self.root.winfo_children():
            if child is self._menubar:
                continue
            child.destroy()
        self.vars = {}
        self._segment_buttons = {}
        self._segment_icons = {}
        self._action_widgets = []
        self._panel_controls = {}
        self._rows = []
        self._groups = []
        self._advanced_visible = False
        self._worker_queue = queue.Queue()
        self._results = {"改动": [], "提醒": [], "记录": []}
        self._clear_open_target()
        self._file_name.set("")
        self._file_meta.set("")
        self._build_ui()
        self._bind_keys()
        self._select_segment("改动")
        self._draw_drop()


def open_folder(folder, select=None):
    """跨平台打开文件夹并可选选中文件。"""
    folder = os.path.abspath(folder)
    target = os.path.join(folder, select) if select else folder
    try:
        if sys.platform.startswith("win"):
            if select:
                subprocess.run(["explorer", "/select,", target], check=False)
            else:
                os.startfile(folder)  # noqa: S606
        elif sys.platform == "darwin":
            if select:
                subprocess.run(["open", "-R", target], check=False)
            else:
                subprocess.run(["open", folder], check=False)
        else:
            subprocess.run(["xdg-open", folder], check=False)
    except OSError:
        return


def enable_dpi_awareness():
    """Windows 高 DPI 感知，避免缩放不为 100% 时界面模糊。"""
    if sys.platform != "win32":
        return
    import ctypes
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except (AttributeError, OSError):
        pass


def main():
    enable_dpi_awareness()
    root = TkinterDnD.Tk()
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
