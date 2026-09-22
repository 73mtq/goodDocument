"""app.py —— 规范文档排版工具（桌面应用 / tkinter GUI）

朱批纸面：纸色打底、墨色承载信息、朱砂只出现在动作上。
把 .docx 拖入窗口或点击选择，一键统一页面、标题、正文、图表、参考文献与页码；
参数在右侧「高级设置」面板中调整，全部来自 config.json。

用法：
    python app.py
"""
import os
import queue
import sys
import threading
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from tkinterdnd2 import DND_FILES, TkinterDnD

import ui_theme
from config_model import load_config_file, save_config_file
from paths import app_dir, assets_dir, bundled_dir, config_path, default_config_path, make_normalized_output_path


# 导入核心库
sys.path.insert(0, bundled_dir())
from formatter import generate, build_sample_content  # noqa: E402
from normalizer import inspect_docx, normalize_docx, NormalizeError  # noqa: E402
from version import __version__


SIZE_OPTIONS = [
    "初号", "小初", "一号", "小一", "二号", "小二",
    "三号", "小三", "四号", "小四", "五号", "小五",
    "六号", "小六", "七号", "八号",
]
ALIGN_OPTIONS = ["left", "center", "right", "justify"]
FONT_ZH_OPTIONS = ["宋体", "黑体", "楷体", "仿宋", "微软雅黑"]
FONT_EN_OPTIONS = ["Times New Roman", "Arial", "Calibri", "Cambria"]
POSITION_OPTIONS = ["above", "below"]
VALIGN_OPTIONS = ["center", "top"]
SIZE_MODE_OPTIONS = ["fit", "fixed"]

EMPTY_RESULT_TEXT = "尚未选择文档 — 选择一份 .docx，可先「预览变更」查看会改什么，不会写入文件。"


class App:
    def __init__(self, root):
        self.root = root
        self.cfg = {}
        self.vars = {}
        self._fonts = {}
        self._advanced_visible = False
        self._busy = False
        self._worker_queue = queue.Queue()
        self._open_target = None
        self._results = {"变更": [], "警告": [], "日志": []}
        self._active_segment = "变更"
        self._segment_buttons = {}
        self._action_widgets = []
        self._menubar = None

        self._input_path = tk.StringVar()
        self._output_path = tk.StringVar()
        self._file_name = tk.StringVar(value="尚未选择文档")
        self._file_meta = tk.StringVar(value="")
        self._status = tk.StringVar(value="就绪：拖入或选择一份 .docx 文档。")
        default_name = self._get(["output", "filename"], "规范文档示例.docx")
        self._gen_path = tk.StringVar(value=os.path.join(app_dir(), default_name))

        self._load_config()
        self._fonts = ui_theme.apply_theme(self.root)
        self._configure_root()
        self._build_menu()
        self._build_ui()
        self._bind_keys()
        self._select_segment("变更")

    # ---- 配置读写 ----
    def _load_config(self):
        path = config_path()
        if not os.path.exists(path):
            path = default_config_path()
        if not os.path.exists(path):
            self.cfg = {}
            return
        self.cfg = load_config_file(path)

    def _save_config(self):
        save_config_file(self.cfg, config_path())

    # ---- 窗口与菜单 ----
    def _configure_root(self):
        self.root.title(f"Word 文档规范化 v{__version__}")
        self.root.minsize(900, 640)
        width, height = 1000, 730
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
                           activebackground=colors["accent_soft"], activeforeground=colors["ink"],
                           bd=0, relief="flat", font=self._fonts["body"])
        more.add_command(label="生成示例文档", command=self._on_generate)
        more.add_command(label="保存当前配置", command=self._on_save_cfg)
        more.add_separator()
        more.add_command(label="恢复默认配置", command=self._on_reset)
        more.add_command(label="关于与快捷键", command=self._show_about)
        more.add_separator()
        more.add_command(label="退出", command=self.root.destroy)
        self._menubar.add_cascade(label="更多", menu=more)
        self.root.configure(menu=self._menubar)

    # ---- 界面构建 ----
    def _build_ui(self):
        shell = ttk.Frame(self.root, padding=(24, 20, 24, 12))
        shell.grid(row=0, column=0, sticky="nsew")

        self._build_header(shell)
        self._build_select_step(shell)
        self._build_output_step(shell)
        self._build_action_step(shell)
        self._build_status_bar(shell)
        self._build_result_area(shell)
        self._build_advanced_panel()

    def _build_header(self, parent):
        head = ttk.Frame(parent)
        head.pack(fill="x")
        ttk.Label(head, text="Word 文档规范化", style="Display.TLabel").pack(side="left")
        ttk.Label(head, text=f"v{__version__} · 参照 GB/T 7714 与公文格式", style="Caption.TLabel").pack(
            side="left", padx=(12, 0), pady=(9, 0))
        self._advanced_button = ttk.Button(head, text="高级设置", style="Quiet.TButton",
                                           command=self._toggle_advanced)
        self._advanced_button.pack(side="right")
        self._action_widgets.append(self._advanced_button)

    def _section_header(self, parent, number, title):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(18, 8))
        tk.Frame(row, bg=ui_theme.COLORS["accent"], width=2, height=14).pack(side="left", padx=(0, 8))
        ttk.Label(row, text=number, style="Step.TLabel").pack(side="left", padx=(0, 6))
        ttk.Label(row, text=title, style="Section.TLabel").pack(side="left")

    def _build_select_step(self, parent):
        self._section_header(parent, "1", "选择文档")
        self._drop = tk.Canvas(parent, height=96, highlightthickness=0, cursor="hand2",
                               bg=ui_theme.COLORS["surface"])
        self._drop.pack(fill="x")
        self._drop.bind("<Configure>", lambda event: self._draw_drop())
        self._drop.bind("<Button-1>", lambda event: self._browse_input())
        self._drop.bind("<Enter>", lambda event: self._draw_drop(hover=True))
        self._drop.bind("<Leave>", lambda event: self._draw_drop())
        self._drop.drop_target_register(DND_FILES)
        self._drop.dnd_bind("<<Drop>>", self._on_drop)
        self._drop.dnd_bind("<<DropEnter>>", lambda event: self._draw_drop(drag=True))
        self._drop.dnd_bind("<<DropLeave>>", lambda event: self._draw_drop())

        info = ttk.Frame(parent)
        info.pack(fill="x", pady=(8, 0))
        ttk.Label(info, textvariable=self._file_name, font=self._fonts["body_bold"]).pack(side="left")
        self._select_button = ttk.Button(info, text="选择文档…", command=self._browse_input)
        self._select_button.pack(side="right")
        self._action_widgets.append(self._select_button)
        ttk.Label(parent, textvariable=self._file_meta, style="Caption.TLabel",
                  wraplength=760).pack(anchor="w", pady=(2, 0))

    def _build_output_step(self, parent):
        self._section_header(parent, "2", "规范化后另存为")
        row = ttk.Frame(parent)
        row.pack(fill="x")
        entry = ttk.Entry(row, textvariable=self._output_path)
        entry.pack(side="left", fill="x", expand=True)
        save_as = ttk.Button(row, text="另存为…", command=self._browse_normalized_out)
        save_as.pack(side="left", padx=(8, 0))
        self._action_widgets.append(save_as)
        ttk.Label(parent, text="留空则保存为原文件同目录的「原文件名_规范化.docx」",
                  style="Caption.TLabel").pack(anchor="w", pady=(4, 0))

    def _build_action_step(self, parent):
        row = ttk.Frame(parent)
        row.pack(fill="x", pady=(18, 0))
        self._preview_button = ttk.Button(row, text="预览变更", command=self._on_preview)
        self._preview_button.pack(side="left")
        self._normalize_button = ttk.Button(row, text="开始规范化", style="Accent.TButton",
                                            command=self._on_normalize)
        self._normalize_button.pack(side="right")
        self._action_widgets.extend([self._preview_button, self._normalize_button])
        ttk.Label(parent, text="预览不会写入文件；规范化另存为新文件，不改动原文档。",
                  style="Caption.TLabel").pack(anchor="w", pady=(6, 0))

    def _build_status_bar(self, parent):
        bar = ttk.Frame(parent)
        bar.pack(side="bottom", fill="x", pady=(10, 0))
        ttk.Separator(bar).pack(fill="x", pady=(0, 8))
        inner = ttk.Frame(bar)
        inner.pack(fill="x")
        ttk.Label(inner, textvariable=self._status, style="Muted.TLabel").pack(side="left")
        self._open_button = ttk.Button(inner, text="打开文件夹", style="Quiet.TButton",
                                       command=self._open_folder_clicked)
        self._progress = ttk.Progressbar(inner, mode="indeterminate", length=140)

    def _build_result_area(self, parent):
        area = ttk.Frame(parent)
        area.pack(fill="both", expand=True, pady=(18, 0))
        seg_row = ttk.Frame(area)
        seg_row.pack(fill="x")
        for name in ("变更", "警告", "日志"):
            button = ttk.Button(seg_row, text=name, style="Seg.TButton",
                                command=lambda item=name: self._select_segment(item))
            button.pack(side="left")
            self._segment_buttons[name] = button
        ttk.Label(seg_row, text="预览与规范化的结果会显示在这里", style="Caption.TLabel").pack(
            side="right", padx=(0, 4))
        body = ttk.Frame(area)
        body.pack(fill="both", expand=True, pady=(6, 0))
        colors = ui_theme.COLORS
        self._text = tk.Text(body, wrap="word", relief="flat", highlightthickness=1,
                             highlightbackground=colors["border"], highlightcolor=colors["accent"],
                             bg=colors["surface"], fg=colors["ink"],
                             insertbackground=colors["ink"],
                             selectbackground=colors["accent_soft"], selectforeground=colors["ink"],
                             font=self._fonts["body"], padx=12, pady=10,
                             spacing1=2, spacing3=4, state="disabled")
        scrollbar = ttk.Scrollbar(body, orient="vertical", command=self._text.yview)
        self._text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self._text.pack(side="left", fill="both", expand=True)
        self._text.tag_configure("change", foreground=colors["ink"])
        self._text.tag_configure("warning", foreground=colors["warning"])
        self._text.tag_configure("error", foreground=colors["danger"])
        self._text.tag_configure("meta", foreground=colors["ink_muted"])
        self._text.tag_configure("heading", foreground=colors["ink"], font=self._fonts["body_bold"])

    # ---- 高级设置面板 ----
    def _build_advanced_panel(self):
        colors = ui_theme.COLORS
        self._advanced_panel = ttk.Frame(self.root, style="Panel.TFrame")
        ttk.Separator(self._advanced_panel, orient="vertical").pack(side="left", fill="y")
        content = ttk.Frame(self._advanced_panel, style="Panel.TFrame", padding=(16, 16, 12, 12))
        content.pack(side="left", fill="both", expand=True)

        head = ttk.Frame(content, style="Panel.TFrame")
        head.pack(fill="x")
        ttk.Label(head, text="高级设置", style="PanelSection.TLabel").pack(side="left")
        ttk.Button(head, text="收起", style="PanelQuiet.TButton",
                   command=self._toggle_advanced).pack(side="right")
        ttk.Label(content, text="常用时无需调整；改动后点「开始规范化」即生效。",
                  style="PanelCaption.TLabel", wraplength=360).pack(anchor="w", pady=(4, 8))

        holder = ttk.Frame(content, style="Panel.TFrame")
        holder.pack(fill="both", expand=True)
        self._panel_canvas = tk.Canvas(holder, bg=colors["surface"], highlightthickness=0, width=380)
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
        self._panel_canvas.bind("<Enter>", self._bind_panel_wheel)
        self._panel_canvas.bind("<Leave>", self._unbind_panel_wheel)
        self._build_advanced_groups(inner)

    def _bind_panel_wheel(self, event=None):
        self.root.bind_all("<MouseWheel>", self._on_panel_wheel)

    def _unbind_panel_wheel(self, event=None):
        self.root.unbind_all("<MouseWheel>")

    def _on_panel_wheel(self, event):
        self._panel_canvas.yview_scroll(int(-event.delta / 120), "units")

    def _toggle_advanced(self):
        if self._advanced_visible:
            self._advanced_panel.grid_forget()
            self._advanced_button.configure(text="高级设置")
            self._advanced_visible = False
        else:
            self._advanced_panel.grid(row=0, column=1, sticky="ns")
            self._advanced_button.configure(text="收起设置")
            self._advanced_visible = True

    def _build_advanced_groups(self, parent):
        self._group(parent, "页面", self._build_page_group)
        self._group(parent, "正文", self._build_body_group)
        self._group(parent, "标题", self._build_headings_group)
        self._group(parent, "图", self._build_figure_group)
        self._group(parent, "表", self._build_table_group)
        self._group(parent, "参考文献", self._build_refs_group)
        self._group(parent, "示例文档", self._build_sample_group)

    def _group(self, parent, title, builder):
        frame = ttk.Frame(parent, style="Panel.TFrame")
        frame.pack(fill="x", pady=(0, 2))
        body = ttk.Frame(frame, style="Panel.TFrame")
        state = {"open": False}

        def toggle():
            state["open"] = not state["open"]
            if state["open"]:
                body.pack(fill="x", padx=(14, 4), pady=(0, 8))
                header.configure(text="▾ " + title)
            else:
                body.pack_forget()
                header.configure(text="▸ " + title)

        header = ttk.Button(frame, text="▸ " + title, style="Group.TButton", command=toggle)
        header.pack(fill="x")
        builder(body)

    # ---- 面板字段构建器 ----
    def _field_row(self, parent, label):
        row = ttk.Frame(parent, style="Panel.TFrame")
        row.pack(fill="x", pady=3)
        ttk.Label(row, text=label, width=13, anchor="e", style="Panel.TLabel").pack(side="left")
        return row

    def _entry(self, parent, label, key_chain, width=20):
        value = tk.StringVar(value=str(self._get(key_chain, "")))
        self.vars[".".join(key_chain)] = value
        row = self._field_row(parent, label)
        entry = ttk.Entry(row, textvariable=value, width=width)
        entry.pack(side="left", padx=(8, 0), fill="x", expand=True)
        return value

    def _combo(self, parent, label, key_chain, raw_options, width=16):
        raw = self._get(key_chain, "")
        value = tk.StringVar(value=ui_theme.display_label(str(raw)))
        self.vars[".".join(key_chain)] = value
        row = self._field_row(parent, label)
        combo = ttk.Combobox(row, textvariable=value, values=ui_theme.option_values(raw_options),
                             width=width, state="normal")
        combo.pack(side="left", padx=(8, 0))
        return value

    def _check(self, parent, label, key_chain):
        value = tk.BooleanVar(value=bool(self._get(key_chain, False)))
        self.vars[".".join(key_chain)] = value
        row = self._field_row(parent, label)
        ttk.Checkbutton(row, variable=value, style="Panel.TCheckbutton").pack(side="left", padx=(8, 0))
        return value

    def _spin(self, parent, label, key_chain, lo=0, hi=999, step=1):
        raw = self._get(key_chain, 0)
        try:
            number = float(raw)
            if number == int(number):
                number = int(number)
        except (ValueError, TypeError):
            number = 0
        value = tk.StringVar(value=str(number))
        self.vars[".".join(key_chain)] = value
        row = self._field_row(parent, label)
        spin = ttk.Spinbox(row, from_=lo, to=hi, increment=step, textvariable=value, width=8)
        spin.pack(side="left", padx=(8, 0))
        return value

    # ---- 各设置分组 ----
    def _build_page_group(self, parent):
        self._spin(parent, "上边距(cm)", ["page", "margins", "top"], 0, 10, 0.5)
        self._spin(parent, "下边距(cm)", ["page", "margins", "bottom"], 0, 10, 0.5)
        self._spin(parent, "左边距(cm)", ["page", "margins", "left"], 0, 10, 0.5)
        self._spin(parent, "右边距(cm)", ["page", "margins", "right"], 0, 10, 0.5)
        self._check(parent, "页码启用", ["pageNumber", "enabled"])
        self._check(parent, "仅更新已有页码", ["pageNumber", "onlyUpdateExisting"])
        self._combo(parent, "页码排列", ["pageNumber", "align"], ALIGN_OPTIONS)
        self._combo(parent, "页码字号", ["pageNumber", "size"], SIZE_OPTIONS)
        self._combo(parent, "页码中文字体", ["pageNumber", "font"], FONT_ZH_OPTIONS)
        self._combo(parent, "页码英数字体", ["pageNumber", "asciiFont"], FONT_EN_OPTIONS)

    def _build_body_group(self, parent):
        self._combo(parent, "中文字体", ["body", "font"], FONT_ZH_OPTIONS)
        self._combo(parent, "英数字体", ["body", "asciiFont"], FONT_EN_OPTIONS)
        self._combo(parent, "字号", ["body", "size"], SIZE_OPTIONS)
        self._spin(parent, "行距(倍)", ["body", "lineSpacing"], 1.0, 3.0, 0.5)
        self._spin(parent, "首行缩进(字)", ["body", "firstLineIndentChars"], 0, 10, 1)
        self._combo(parent, "排列方式", ["body", "align"], ALIGN_OPTIONS)
        self._spin(parent, "段后(pt)", ["body", "spaceAfterPt"], 0, 30, 1)

    def _build_headings_group(self, parent):
        for level, name in [(1, "H1"), (2, "H2"), (3, "H3"), (4, "H4")]:
            ttk.Label(parent, text=name, style="PanelSection.TLabel").pack(
                anchor="w", padx=(6, 0), pady=(10, 2))
            self._combo(parent, "中文字体", ["headings", "h" + str(level), "font"], FONT_ZH_OPTIONS)
            self._combo(parent, "英数字体", ["headings", "h" + str(level), "asciiFont"], FONT_EN_OPTIONS)
            self._combo(parent, "字号", ["headings", "h" + str(level), "size"], SIZE_OPTIONS)
            self._spin(parent, "行距(倍)", ["headings", "h" + str(level), "lineSpacing"], 1.0, 3.0, 0.5)
            self._check(parent, "加粗", ["headings", "h" + str(level), "bold"])
            self._spin(parent, "段前(pt)", ["headings", "h" + str(level), "spaceBeforePt"], 0, 40, 1)
            self._spin(parent, "段后(pt)", ["headings", "h" + str(level), "spaceAfterPt"], 0, 40, 1)
            self._combo(parent, "排列方式", ["headings", "h" + str(level), "align"], ALIGN_OPTIONS)

    def _build_figure_group(self, parent):
        self._combo(parent, "尺寸模式", ["figure", "sizeMode"], SIZE_MODE_OPTIONS)
        self._spin(parent, "等比最大宽(cm)", ["figure", "maxWidthCm"], 5, 25, 0.5)
        self._spin(parent, "等比最大高(cm)", ["figure", "maxHeightCm"], 5, 30, 0.5)
        self._spin(parent, "统一宽度(cm)", ["figure", "fixedWidthCm"], 1, 25, 0.5)
        self._spin(parent, "统一高度(cm)", ["figure", "fixedHeightCm"], 1, 30, 0.5)
        self._check(parent, "统一时保持比例", ["figure", "keepAspectRatio"])
        self._check(parent, "宽度不超出正文", ["figure", "clampToTextWidth"])
        self._check(parent, "跳过表格内图片", ["figure", "skipInTables"])
        self._check(parent, "图题跟随图片", ["figure", "keepWithCaption"])
        self._combo(parent, "图题位置", ["figure", "captionPosition"], POSITION_OPTIONS)
        self._combo(parent, "图题字号", ["figure", "captionSize"], SIZE_OPTIONS)
        self._check(parent, "图题加粗", ["figure", "captionBold"])
        self._combo(parent, "图题排列", ["figure", "captionAlign"], ALIGN_OPTIONS)
        self._spin(parent, "图题行距", ["figure", "captionLineSpacing"], 1.0, 3.0, 0.5)
        self._combo(parent, "注释字号", ["figure", "noteSize"], SIZE_OPTIONS)
        self._combo(parent, "注释排列", ["figure", "noteAlign"], ALIGN_OPTIONS)
        self._combo(parent, "前缀", ["figure", "prefix"], ["图", "Figure", "Fig"])
        ttk.Label(parent, text="「统一尺寸」下：只填宽度即为统一宽度；"
                               "不勾「保持比例」则强制拉伸（等同 VBA 的 LockAspectRatio=False）。",
                  style="PanelCaption.TLabel", wraplength=340).pack(anchor="w", pady=(6, 0))

    def _build_table_group(self, parent):
        self._combo(parent, "表题位置", ["table", "captionPosition"], POSITION_OPTIONS)
        self._combo(parent, "表题字号", ["table", "captionSize"], SIZE_OPTIONS)
        self._check(parent, "表题加粗", ["table", "captionBold"])
        self._combo(parent, "表题排列", ["table", "captionAlign"], ALIGN_OPTIONS)
        self._spin(parent, "表题行距", ["table", "captionLineSpacing"], 1.0, 3.0, 0.5)
        self._combo(parent, "注释字号", ["table", "noteSize"], SIZE_OPTIONS)
        self._combo(parent, "注释排列", ["table", "noteAlign"], ALIGN_OPTIONS)
        self._combo(parent, "单元格字号", ["table", "cellSize"], SIZE_OPTIONS)
        self._combo(parent, "单元格排列", ["table", "cellAlign"], ALIGN_OPTIONS)
        self._combo(parent, "垂直位置", ["table", "cellVAlign"], VALIGN_OPTIONS)
        self._spin(parent, "单元格行距", ["table", "cellLineSpacing"], 1.0, 3.0, 0.5)
        self._check(parent, "表头加粗", ["table", "headerBold"])
        self._combo(parent, "前缀", ["table", "prefix"], ["表", "Table"])

    def _build_refs_group(self, parent):
        self._combo(parent, "中文字体", ["references", "font"], FONT_ZH_OPTIONS)
        self._combo(parent, "英数字体", ["references", "asciiFont"], FONT_EN_OPTIONS)
        self._combo(parent, "字号", ["references", "size"], SIZE_OPTIONS)
        self._spin(parent, "行距(倍)", ["references", "lineSpacing"], 1.0, 3.0, 0.5)
        self._combo(parent, "排列方式", ["references", "align"], ALIGN_OPTIONS)
        self._spin(parent, "悬挂缩进(pt)", ["references", "hangingIndentPt"], 0, 60, 1)

    def _build_sample_group(self, parent):
        self._entry(parent, "文件名", ["output", "filename"], width=24)
        row = self._field_row(parent, "保存路径")
        entry = ttk.Entry(row, textvariable=self._gen_path)
        entry.pack(side="left", padx=(8, 0), fill="x", expand=True)
        ttk.Button(row, text="浏览…", style="PanelQuiet.TButton",
                   command=self._browse_out).pack(side="left", padx=(6, 0))
        ttk.Button(parent, text="生成示例文档", command=self._on_generate).pack(anchor="w", pady=(10, 2))
        ttk.Label(parent, text="示例包含示例正文、一张示例图、一个示例表与 GB/T 7714 参考文献。",
                  style="PanelCaption.TLabel", wraplength=340).pack(anchor="w", pady=(4, 0))

    # ---- 拖拽与画布 ----
    def _draw_drop(self, hover=False, drag=False):
        canvas = self._drop
        canvas.delete("all")
        width = max(canvas.winfo_width(), 200)
        height = int(canvas["height"])
        active = hover or drag
        colors = ui_theme.COLORS
        outline = colors["accent"] if active else colors["border_strong"]
        fill = colors["accent_soft"] if active else colors["surface"]
        canvas.configure(bg=fill)
        canvas.create_rectangle(3, 3, width - 3, height - 3, dash=(6, 4), outline=outline, width=2)
        path = self._input_path.get().strip()
        title = self._shorten_path(os.path.basename(path)) if path else "把 .docx 文件拖到这里"
        subtitle = "或点击此处选择文档" if not path else self._shorten_path(path)
        canvas.create_text(width / 2, height / 2 - 13, text=title, fill=colors["ink"],
                           font=self._fonts["body_bold"])
        canvas.create_text(width / 2, height / 2 + 11, text=subtitle, fill=colors["ink_muted"],
                           font=self._fonts["caption"])

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
            messagebox.showwarning("文件类型", "只接受 .docx 文件：\n" + os.path.basename(path))
            return
        self._set_input(path)

    def _set_input(self, path):
        self._input_path.set(path)
        if not self._output_path.get().strip():
            self._output_path.set(make_normalized_output_path(path))
        self._file_name.set(os.path.basename(path))
        self._file_meta.set("正在读取文档信息…")
        self._draw_drop()
        self._clear_open_target()
        self._status.set("已选择：" + path)
        self._run_worker(lambda: inspect_docx(path), self._show_inspection,
                         lambda exc: self._file_meta.set("文档信息读取失败：" + str(exc)))

    def _show_inspection(self, info):
        levels = "、".join(f"H{level}" for level in info.heading_levels) if info.heading_levels else "无"
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
            title="选择要规范化的 Word 文档",
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
            initial = os.path.join(app_dir(), "规范化文档.docx")
        path = filedialog.asksaveasfilename(
            title="选择规范化输出位置",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
            initialdir=os.path.dirname(initial),
            initialfile=os.path.basename(initial),
        )
        if path:
            self._output_path.set(path)

    def _browse_out(self):
        if self._busy:
            return
        current = self._gen_path.get()
        path = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
            initialdir=os.path.dirname(current) if current else app_dir(),
            initialfile=os.path.basename(current) if current else "规范文档示例.docx",
        )
        if path:
            self._gen_path.set(path)

    # ---- 后台任务 ----
    def _run_worker(self, work, on_success, on_error):
        if self._busy:
            return
        self._set_busy(True)

        def target():
            try:
                self._worker_queue.put(("ok", work()))
            except BaseException as exc:  # noqa: BLE001
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
        if busy:
            self._progress.pack(side="right")
            self._progress.start(12)
        else:
            self._progress.stop()
            self._progress.pack_forget()

    # ---- 主操作 ----
    def _on_normalize(self):
        self._collect()
        self._save_config()
        input_path = self._input_path.get().strip()
        if not input_path:
            messagebox.showwarning("缺少输入文档", "请先拖入或选择一份 .docx 文档。")
            return
        output_path = self._output_path.get().strip()
        if not output_path:
            output_path = make_normalized_output_path(input_path)
            self._output_path.set(output_path)

        def work():
            normalize_docx(self.cfg, input_path, output_path, assets_dir=assets_dir())
            return output_path

        def done(path):
            self._status.set("已规范化：" + path)
            self._add_result("日志", "已规范化：" + path, "change")
            self._set_open_target(os.path.dirname(path), os.path.basename(path))
            self._render_results()

        def failed(exc):
            self._status.set("规范化失败")
            self._add_result("警告", "规范化失败：" + str(exc), "error")
            self._select_segment("警告")
            self._show_error("规范化失败", exc)

        self._run_worker(work, done, failed)

    def _on_preview(self):
        self._collect()
        self._save_config()
        input_path = self._input_path.get().strip()
        if not input_path:
            messagebox.showwarning("缺少输入文档", "请先拖入或选择一份 .docx 文档。")
            return
        output_guess = self._output_path.get().strip() or make_normalized_output_path(input_path)

        def work():
            return normalize_docx(self.cfg, input_path, output_guess,
                                  dry_run=True, return_result=True)

        def done(result):
            self._results["变更"] = [(item, "change") for item in result.changes]
            self._results["警告"] = [(item, "warning") for item in result.warnings]
            summary = (
                f"{os.path.basename(result.input_path)}：将处理 {result.paragraphs_processed} 段、"
                f"{result.tables_processed} 个表、调整 {result.images_processed} 张图"
                f"（{result.duration_ms}ms，未写入文件）"
            )
            self._results["日志"].append((summary, "meta"))
            self._status.set("预览完成，未修改文件")
            if result.changes:
                self._select_segment("变更")
            elif result.warnings:
                self._select_segment("警告")
            else:
                self._select_segment("日志")
            self._render_results()

        def failed(exc):
            self._status.set("预览失败")
            self._add_result("警告", "预览失败：" + str(exc), "error")
            self._select_segment("警告")
            self._show_error("预览失败", exc)

        self._run_worker(work, done, failed)

    def _on_generate(self):
        self._collect()
        self._save_config()
        out_path = self._gen_path.get().strip()
        if not out_path:
            out_path = os.path.join(app_dir(), self._get(["output", "filename"], "规范文档示例.docx"))
            self._gen_path.set(out_path)

        def work():
            generate(self.cfg, out_path, assets_dir=assets_dir(), content_fn=build_sample_content)
            return out_path

        def done(path):
            self._status.set("已生成示例文档：" + path)
            self._add_result("日志", "已生成示例文档：" + path, "change")
            self._set_open_target(os.path.dirname(path), os.path.basename(path))
            self._render_results()

        def failed(exc):
            self._status.set("生成失败")
            self._add_result("警告", "生成失败：" + str(exc), "error")
            self._select_segment("警告")
            self._show_error("生成失败", exc)

        self._run_worker(work, done, failed)

    # ---- 配置与关于 ----
    def _on_save_cfg(self):
        self._collect()
        self._save_config()
        self._status.set("配置已保存 → " + config_path())
        self._add_result("日志", "配置已保存：" + config_path(), "meta")
        self._render_results()

    def _on_reset(self):
        path = default_config_path()
        if not os.path.exists(path):
            messagebox.showwarning("提示", "未找到默认配置文件。")
            return
        self.cfg = load_config_file(path)
        self._rebuild_ui()
        self._status.set("已恢复默认配置（尚未写回 config.json）")

    def _show_about(self):
        messagebox.showinfo(
            "关于与快捷键",
            f"Word 文档规范化工具 v{__version__}\n\n"
            "把 .docx 拖入窗口或点击选择，按 config.json 统一页面、标题、正文、"
            "图表、参考文献与页码。\n\n"
            "快捷键：\n"
            "Ctrl+O  选择文档\n"
            "Ctrl+P  预览变更\n"
            "Ctrl+Enter  开始规范化\n"
            "Ctrl+S  保存当前配置\n"
            "Ctrl+G  生成示例文档\n"
            "F1  本窗口",
        )

    # ---- 结果区 ----
    def _select_segment(self, name):
        self._active_segment = name
        for key, button in self._segment_buttons.items():
            button.configure(style="SegActive.TButton" if key == name else "Seg.TButton")
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
            messagebox.showerror(title, str(exc))
            return
        detail = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__))
        self._add_result("日志", f"{title}：{exc}", "error")
        self._add_result("日志", detail, "error")
        self._select_segment("日志")
        messagebox.showerror(title, f"{exc}\n\n完整详情已写入「日志」。")

    # ---- 打开输出位置 ----
    def _set_open_target(self, folder, select=None):
        self._open_target = (folder, select)
        self._open_button.pack(side="right")

    def _clear_open_target(self):
        self._open_target = None
        self._open_button.pack_forget()

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
        self.root.bind("<F1>", self._shortcut(self._show_about))

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
        self._action_widgets = []
        self._advanced_visible = False
        self._worker_queue = queue.Queue()
        self._results = {"变更": [], "警告": [], "日志": []}
        self._clear_open_target()
        self._file_name.set("尚未选择文档")
        self._file_meta.set("")
        self._build_ui()
        self._bind_keys()
        self._select_segment("变更")
        self._draw_drop()


def open_folder(folder, select=None):
    """跨平台打开文件夹并可选选中文件。"""
    folder = os.path.abspath(folder)
    if sys.platform.startswith("win"):
        if select:
            os.system('explorer /select,"%s"' % os.path.join(folder, select))
        else:
            os.startfile(folder)
    elif sys.platform == "darwin":
        if select:
            os.system('open -R "%s"' % os.path.join(folder, select))
        else:
            os.system('open "%s"' % folder)
    else:
        os.system('xdg-open "%s"' % folder)


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
