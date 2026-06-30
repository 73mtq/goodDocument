"""
app.py —— 规范文档排版生成器（桌面应用 / tkinter GUI）
=====================================================
仅"调参与生成"：在界面里修改各排版参数，点击"生成文档"即可输出 .docx。
可打包为 .exe（PyInstaller）。

用法：
    python app.py              # 启动 GUI
    pyinstaller --onefile --windowed app.py   # 打包
"""

import os
import sys
import traceback
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext

from config_model import load_config_file, save_config_file
from paths import app_dir, assets_dir, bundled_dir, config_path, default_config_path, make_normalized_output_path


# 导入核心库
sys.path.insert(0, bundled_dir())
from formatter import generate, build_sample_content  # noqa: E402
from normalizer import inspect_docx, normalize_docx, NormalizeError  # noqa: E402


# ------------------------------------------------------------------
# 字号选项
# ------------------------------------------------------------------

SIZE_OPTIONS = [
    "初号", "小初", "一号", "小一", "二号", "小二",
    "三号", "小三", "四号", "小四", "五号", "小五",
    "六号", "小六", "七号", "八号",
]
ALIGN_OPTIONS = ["left", "center", "right", "justify"]
FONT_ZH_OPTIONS = ["宋体", "黑体", "楷体", "仿宋", "微软雅黑"]
FONT_EN_OPTIONS = ["Times New Roman", "Arial", "Calibri", "Cambria"]


# ------------------------------------------------------------------
# GUI
# ------------------------------------------------------------------

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Word 文档规范化工具")
        self.root.geometry("820x760")
        self.root.minsize(760, 620)
        self.cfg = {}
        self.vars = {}  # 控件变量缓存
        self._advanced_visible = False

        self._load_config()
        self._build_ui()

    # ---- 配置读写 ----
    def _load_config(self):
        p = config_path()
        if not os.path.exists(p):
            p = default_config_path()
        if not os.path.exists(p):
            self.cfg = {}
            return
        self.cfg = load_config_file(p)

    def _save_config(self):
        p = config_path()
        save_config_file(self.cfg, p)

    # ---- UI 构建 ----
    def _build_ui(self):
        self._advanced_visible = False
        self._configure_style()

        shell = ttk.Frame(self.root, padding=18)
        shell.pack(fill="both", expand=True)

        ttk.Label(shell, text="Word 文档规范化", style="Title.TLabel").pack(anchor="w")
        ttk.Label(
            shell,
            text="选择已有 .docx，按当前规范统一页面、标题、正文、图表、参考文献和页码。",
            style="Muted.TLabel",
        ).pack(anchor="w", pady=(4, 16))

        self._build_workflow(shell)
        self._build_status(shell)
        self._build_advanced(shell)

    def _configure_style(self):
        style = ttk.Style()
        style.configure("Title.TLabel", font=("Microsoft YaHei UI", 18, "bold"))
        style.configure("Section.TLabel", font=("Microsoft YaHei UI", 11, "bold"))
        style.configure("Muted.TLabel", foreground="#666666")
        style.configure("Primary.TButton", font=("Microsoft YaHei UI", 10, "bold"))

    def _build_workflow(self, parent):
        box = ttk.LabelFrame(parent, text="一键规范化", padding=14)
        box.pack(fill="x")

        self._input_path = tk.StringVar()
        self._output_path = tk.StringVar()
        self._gen_path = tk.StringVar(value=os.path.join(app_dir(), self._get(["output", "filename"], "规范文档示例.docx")))

        row1 = ttk.Frame(box)
        row1.pack(fill="x", pady=(0, 8))
        ttk.Label(row1, text="输入文档", width=10, anchor="e").pack(side="left")
        ttk.Entry(row1, textvariable=self._input_path).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row1, text="选择…", command=self._browse_input).pack(side="left")

        row2 = ttk.Frame(box)
        row2.pack(fill="x", pady=(0, 12))
        ttk.Label(row2, text="输出位置", width=10, anchor="e").pack(side="left")
        ttk.Entry(row2, textvariable=self._output_path).pack(side="left", fill="x", expand=True, padx=8)
        ttk.Button(row2, text="另存为…", command=self._browse_normalized_out).pack(side="left")

        actions = ttk.Frame(box)
        actions.pack(fill="x")
        ttk.Button(actions, text="开始规范化", command=self._on_normalize, style="Primary.TButton").pack(side="right")
        ttk.Button(actions, text="预览变更", command=self._on_preview).pack(side="right", padx=8)
        ttk.Button(actions, text="生成示例文档", command=self._on_generate).pack(side="right", padx=8)
        ttk.Button(actions, text="保存配置", command=self._on_save_cfg).pack(side="left")
        ttk.Button(actions, text="重置默认", command=self._on_reset).pack(side="left", padx=8)

    def _build_status(self, parent):
        status = ttk.LabelFrame(parent, text="状态", padding=10)
        status.pack(fill="both", expand=True, pady=12)
        self._status = tk.StringVar(value="就绪：请选择一个 .docx 文档。")
        ttk.Label(status, textvariable=self._status, wraplength=720).pack(anchor="w")
        self._status_text = scrolledtext.ScrolledText(
            status, height=10, wrap="word", font=("Microsoft YaHei UI", 9)
        )
        self._status_text.pack(fill="both", expand=True, pady=(8, 0))
        self._status_text.insert("end", "提示：选好输入文档后，点'预览变更'可先看规范化会改什么（不修改文件）。\n")

    def _build_advanced(self, parent):
        header = ttk.Frame(parent)
        header.pack(fill="x", pady=(2, 8))
        self._advanced_label = tk.StringVar(value="展开高级设置")
        ttk.Button(header, textvariable=self._advanced_label, command=self._toggle_advanced).pack(side="left")
        ttk.Label(header, text="常用时无需调整；复杂文档再展开修改格式参数。", style="Muted.TLabel").pack(side="left", padx=10)

        self._advanced_frame = ttk.Frame(parent)
        self._build_advanced_content(self._advanced_frame)

    def _build_advanced_content(self, parent):
        nb = ttk.Notebook(parent)
        nb.pack(fill="both", expand=True)

        self._tab_page(nb)
        self._tab_body(nb)
        self._tab_headings(nb)
        self._tab_figure(nb)
        self._tab_table(nb)
        self._tab_refs(nb)
        self._tab_output(nb)

    def _toggle_advanced(self):
        if self._advanced_visible:
            self._advanced_frame.pack_forget()
            self._advanced_label.set("展开高级设置")
            self._advanced_visible = False
        else:
            self._advanced_frame.pack(fill="both", expand=True)
            self._advanced_label.set("收起高级设置")
            self._advanced_visible = True

    # ---- 通用字段构建器 ----
    def _row(self, parent, label, widget):
        f = ttk.Frame(parent)
        f.pack(fill="x", padx=6, pady=3)
        ttk.Label(f, text=label, width=16, anchor="e").pack(side="left")
        widget.pack(side="left", fill="x", expand=True, padx=4)

    def _entry(self, parent, label, key_chain, width=20):
        v = tk.StringVar(value=str(self._get(key_chain, "")))
        self.vars[".".join(key_chain)] = v
        e = ttk.Entry(f := ttk.Frame(parent), textvariable=v, width=width)
        ttk.Label(f, text=label, width=16, anchor="e").pack(side="left")
        e.pack(side="left", padx=4)
        f.pack(fill="x", padx=6, pady=3)
        return v

    def _combo(self, parent, label, key_chain, options, width=18):
        v = tk.StringVar(value=str(self._get(key_chain, "")))
        self.vars[".".join(key_chain)] = v
        f = ttk.Frame(parent)
        ttk.Label(f, text=label, width=16, anchor="e").pack(side="left")
        c = ttk.Combobox(f, textvariable=v, values=options, width=width, state="normal")
        c.pack(side="left", padx=4)
        f.pack(fill="x", padx=6, pady=3)
        return v

    def _check(self, parent, label, key_chain):
        v = tk.BooleanVar(value=bool(self._get(key_chain, False)))
        self.vars[".".join(key_chain)] = v
        f = ttk.Frame(parent)
        ttk.Label(f, text=label, width=16, anchor="e").pack(side="left")
        ttk.Checkbutton(f, variable=v).pack(side="left", padx=4)
        f.pack(fill="x", padx=6, pady=3)
        return v

    def _spin(self, parent, label, key_chain, lo=0, hi=999, step=1, fmt=None):
        raw = self._get(key_chain, 0)
        try:
            val = float(raw)
            if val == int(val):
                val = int(val)
        except (ValueError, TypeError):
            val = 0
        v = tk.StringVar(value=str(val))
        self.vars[".".join(key_chain)] = v
        f = ttk.Frame(parent)
        ttk.Label(f, text=label, width=16, anchor="e").pack(side="left")
        sp = ttk.Spinbox(f, from_=lo, to=hi, increment=step, textvariable=v, width=10)
        sp.pack(side="left", padx=4)
        f.pack(fill="x", padx=6, pady=3)
        return v

    # ---- 配置取值 ----
    def _get(self, chain, default):
        cur = self.cfg
        for k in chain:
            if isinstance(cur, dict):
                cur = cur.get(k)
            else:
                return default
            if cur is None:
                return default
        return cur if cur is not None else default

    # ---- 各 Tab ----
    def _tab_page(self, nb):
        t = ttk.Frame(nb)
        nb.add(t, text="页面")
        m = self.cfg.get("page", {}).get("margins", {})
        self._spin(t, "上边距(cm)", ["page", "margins", "top"], 0, 10, 0.5)
        self._spin(t, "下边距(cm)", ["page", "margins", "bottom"], 0, 10, 0.5)
        self._spin(t, "左边距(cm)", ["page", "margins", "left"], 0, 10, 0.5)
        self._spin(t, "右边距(cm)", ["page", "margins", "right"], 0, 10, 0.5)
        ttk.Separator(t).pack(fill="x", padx=6, pady=6)
        self._check(t, "页码启用", ["pageNumber", "enabled"])
        self._combo(t, "页码对齐", ["pageNumber", "align"], ALIGN_OPTIONS)
        self._combo(t, "页码字号", ["pageNumber", "size"], SIZE_OPTIONS)
        self._combo(t, "页码中文字体", ["pageNumber", "font"], FONT_ZH_OPTIONS)
        self._combo(t, "页码英数字体", ["pageNumber", "asciiFont"], FONT_EN_OPTIONS)

    def _tab_body(self, nb):
        t = ttk.Frame(nb)
        nb.add(t, text="正文")
        self._combo(t, "中文字体", ["body", "font"], FONT_ZH_OPTIONS)
        self._combo(t, "英数字体", ["body", "asciiFont"], FONT_EN_OPTIONS)
        self._combo(t, "字号", ["body", "size"], SIZE_OPTIONS)
        self._spin(t, "行距(倍)", ["body", "lineSpacing"], 1.0, 3.0, 0.5)
        self._spin(t, "首行缩进(字)", ["body", "firstLineIndentChars"], 0, 10, 1)
        self._combo(t, "对齐", ["body", "align"], ALIGN_OPTIONS)
        self._spin(t, "段后(pt)", ["body", "spaceAfterPt"], 0, 30, 1)

    def _tab_headings(self, nb):
        t = ttk.Frame(nb)
        nb.add(t, text="标题")
        for lv, name in [(1, "一级标题 H1"), (2, "二级标题 H2"), (3, "三级标题 H3"), (4, "四级标题 H4")]:
            ttk.Label(t, text=name, font=("", 10, "bold")).pack(anchor="w", padx=6, pady=(8, 2))
            self._combo(t, "  中文字体", ["headings", "h" + str(lv), "font"], FONT_ZH_OPTIONS)
            self._combo(t, "  英数字体", ["headings", "h" + str(lv), "asciiFont"], FONT_EN_OPTIONS)
            self._combo(t, "  字号", ["headings", "h" + str(lv), "size"], SIZE_OPTIONS)
            self._spin(t, "  行距(倍)", ["headings", "h" + str(lv), "lineSpacing"], 1.0, 3.0, 0.5)
            self._check(t, "  加粗", ["headings", "h" + str(lv), "bold"])
            self._spin(t, "  段前(pt)", ["headings", "h" + str(lv), "spaceBeforePt"], 0, 40, 1)
            self._spin(t, "  段后(pt)", ["headings", "h" + str(lv), "spaceAfterPt"], 0, 40, 1)
            self._combo(t, "  对齐", ["headings", "h" + str(lv), "align"], ALIGN_OPTIONS)
            ttk.Separator(t).pack(fill="x", padx=6, pady=4)

    def _tab_figure(self, nb):
        t = ttk.Frame(nb)
        nb.add(t, text="图")
        self._spin(t, "最大宽(cm)", ["figure", "maxWidthCm"], 5, 25, 0.5)
        self._spin(t, "最大高(cm)", ["figure", "maxHeightCm"], 5, 30, 0.5)
        self._combo(t, "图题位置", ["figure", "captionPosition"], ["above", "below"])
        self._combo(t, "图题字号", ["figure", "captionSize"], SIZE_OPTIONS)
        self._check(t, "图题加粗", ["figure", "captionBold"])
        self._combo(t, "图题对齐", ["figure", "captionAlign"], ALIGN_OPTIONS)
        self._spin(t, "图题行距", ["figure", "captionLineSpacing"], 1.0, 3.0, 0.5)
        self._combo(t, "注释字号", ["figure", "noteSize"], SIZE_OPTIONS)
        self._combo(t, "注释对齐", ["figure", "noteAlign"], ALIGN_OPTIONS)
        self._combo(t, "前缀", ["figure", "prefix"], ["图", "Figure", "Fig"])

    def _tab_table(self, nb):
        t = ttk.Frame(nb)
        nb.add(t, text="表")
        self._combo(t, "表题位置", ["table", "captionPosition"], ["above", "below"])
        self._combo(t, "表题字号", ["table", "captionSize"], SIZE_OPTIONS)
        self._check(t, "表题加粗", ["table", "captionBold"])
        self._combo(t, "表题对齐", ["table", "captionAlign"], ALIGN_OPTIONS)
        self._spin(t, "表题行距", ["table", "captionLineSpacing"], 1.0, 3.0, 0.5)
        self._combo(t, "注释字号", ["table", "noteSize"], SIZE_OPTIONS)
        self._combo(t, "注释对齐", ["table", "noteAlign"], ALIGN_OPTIONS)
        self._combo(t, "单元格字号", ["table", "cellSize"], SIZE_OPTIONS)
        self._combo(t, "单元格对齐", ["table", "cellAlign"], ALIGN_OPTIONS)
        self._combo(t, "垂直对齐", ["table", "cellVAlign"], ["center", "top"])
        self._spin(t, "单元格行距", ["table", "cellLineSpacing"], 1.0, 3.0, 0.5)
        self._check(t, "表头加粗", ["table", "headerBold"])
        self._combo(t, "前缀", ["table", "prefix"], ["表", "Table"])

    def _tab_refs(self, nb):
        t = ttk.Frame(nb)
        nb.add(t, text="参考文献")
        self._combo(t, "中文字体", ["references", "font"], FONT_ZH_OPTIONS)
        self._combo(t, "英数字体", ["references", "asciiFont"], FONT_EN_OPTIONS)
        self._combo(t, "字号", ["references", "size"], SIZE_OPTIONS)
        self._spin(t, "行距(倍)", ["references", "lineSpacing"], 1.0, 3.0, 0.5)
        self._combo(t, "对齐", ["references", "align"], ALIGN_OPTIONS)
        self._spin(t, "悬挂缩进(pt)", ["references", "hangingIndentPt"], 0, 60, 1)

    def _tab_output(self, nb):
        t = ttk.Frame(nb)
        nb.add(t, text="输出")
        self._entry(t, "文件名", ["output", "filename"], width=30)
        ttk.Label(t, text="（生成时可选保存位置，默认保存在应用同目录）",
                  foreground="gray").pack(anchor="w", padx=22, pady=4)
        self._gen_path = tk.StringVar(value=os.path.join(app_dir(), self._get(["output", "filename"], "规范文档示例.docx")))
        f = ttk.Frame(t)
        ttk.Label(f, text="保存路径", width=16, anchor="e").pack(side="left")
        ttk.Entry(f, textvariable=self._gen_path, width=40).pack(side="left", padx=4, fill="x", expand=True)
        ttk.Button(f, text="浏览…", command=self._browse_out).pack(side="left")
        f.pack(fill="x", padx=6, pady=6)
        ttk.Label(t, text="提示：示例内容包含占位正文、一张示例图、一个示例表与 GB/T 7714 参考文献示例。",
                  foreground="gray", wraplength=560).pack(anchor="w", padx=22, pady=8)

    def _browse_out(self):
        fn = self._gen_path.get()
        p = filedialog.asksaveasfilename(
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
            initialdir=os.path.dirname(fn) if fn else app_dir(),
            initialfile=os.path.basename(fn) if fn else "规范文档示例.docx",
        )
        if p:
            self._gen_path.set(p)

    def _browse_input(self):
        p = filedialog.askopenfilename(
            title="选择要规范化的 Word 文档",
            filetypes=[("Word 文档", "*.docx")],
            initialdir=app_dir(),
        )
        if p:
            self._input_path.set(p)
            if not self._output_path.get().strip():
                self._output_path.set(make_normalized_output_path(p))
            self._status.set("已选择文档：" + p)

    def _browse_normalized_out(self):
        input_path = self._input_path.get().strip()
        initial = self._output_path.get().strip()
        if not initial and input_path:
            initial = make_normalized_output_path(input_path)
        if not initial:
            initial = os.path.join(app_dir(), "规范化文档.docx")
        p = filedialog.asksaveasfilename(
            title="选择规范化输出位置",
            defaultextension=".docx",
            filetypes=[("Word 文档", "*.docx")],
            initialdir=os.path.dirname(initial) if initial else app_dir(),
            initialfile=os.path.basename(initial) if initial else "规范化文档.docx",
        )
        if p:
            self._output_path.set(p)

    # ---- 从控件收集配置 ----
    def _collect(self):
        for chain_str, var in self.vars.items():
            chain = chain_str.split(".")
            val = var.get()
            # 类型推断
            if isinstance(var, tk.BooleanVar):
                val = bool(var.get())
            else:
                s = var.get()
                # 尝试数字
                try:
                    f = float(s)
                    val = int(f) if f == int(f) else f
                except (ValueError, TypeError):
                    val = s
            self._set(chain, val)

    def _set(self, chain, val):
        cur = self.cfg
        for k in chain[:-1]:
            if k not in cur or not isinstance(cur.get(k), dict):
                cur[k] = {}
            cur = cur[k]
        cur[chain[-1]] = val

    # ---- 事件 ----
    def _on_save_cfg(self):
        self._collect()
        self._save_config()
        self._status.set("配置已保存 → " + config_path())

    def _on_reset(self):
        p = default_config_path()
        if os.path.exists(p):
            self.cfg = load_config_file(p)
            self._rebuild_ui()
            self._status.set("已重置为默认配置")
        else:
            messagebox.showwarning("提示", "未找到默认配置文件。")

    def _rebuild_ui(self):
        for w in self.root.winfo_children():
            w.destroy()
        self.vars = {}
        self._build_ui()

    def _on_generate(self):
        self._collect()
        self._save_config()
        out_path = self._gen_path.get().strip()
        if not out_path:
            out_path = os.path.join(app_dir(), self._get(["output", "filename"], "规范文档示例.docx"))
        try:
            self._status.set("生成中…")
            self.root.update()
            generate(self.cfg, out_path, assets_dir=assets_dir(), content_fn=build_sample_content)
            self._status.set("已生成: " + out_path)
            if messagebox.askyesno("完成", "文档已生成：\n" + out_path + "\n\n是否打开所在文件夹？"):
                open_folder(os.path.dirname(out_path) or ".", os.path.basename(out_path))
        except Exception as e:
            self._status.set("生成失败")
            messagebox.showerror("生成失败", str(e) + "\n\n" + traceback.format_exc())

    def _on_normalize(self):
        self._collect()
        self._save_config()
        input_path = self._input_path.get().strip()
        output_path = self._output_path.get().strip()

        if not input_path:
            messagebox.showwarning("缺少输入文档", "请先选择一个要规范化的 .docx 文档。")
            self._status.set("等待输入文档")
            return
        if not output_path:
            output_path = make_normalized_output_path(input_path)
            self._output_path.set(output_path)

        try:
            self._status.set("正在规范化…")
            self.root.update()
            normalize_docx(self.cfg, input_path, output_path, assets_dir=assets_dir())
            self._status.set("已规范化: " + output_path)
            if messagebox.askyesno("完成", "文档已规范化：\n" + output_path + "\n\n是否打开所在文件夹？"):
                open_folder(os.path.dirname(output_path) or ".", os.path.basename(output_path))
        except NormalizeError as e:
            self._status.set("规范化失败")
            messagebox.showerror("规范化失败", str(e))
        except Exception as e:
            self._status.set("规范化失败")
            messagebox.showerror("规范化失败", str(e) + "\n\n" + traceback.format_exc())

    def _on_preview(self):
        """预览规范化会改什么（不写文件）。"""
        self._collect()
        self._save_config()
        input_path = self._input_path.get().strip()

        if not input_path:
            messagebox.showwarning("缺少输入文档", "请先选择一个要规范化的 .docx 文档。")
            self._status.set("等待输入文档")
            return

        # 预览不依赖 output 路径，自动用 _validate_paths 推断一个临时路径
        output_path = self._output_path.get().strip() or make_normalized_output_path(input_path)

        try:
            self._status.set("正在预览…")
            self.root.update()
            result = normalize_docx(
                self.cfg, input_path, output_path,
                dry_run=True, return_result=True,
            )
            self._render_preview(result)
            self._status.set("预览完成（未修改文件）")
        except NormalizeError as e:
            self._status.set("预览失败")
            messagebox.showerror("预览失败", str(e))
            self._append_status(f"[ERROR] {e.message}\n{e.hint or ''}\n")
        except Exception as e:
            self._status.set("预览失败")
            messagebox.showerror("未预期错误", str(e) + "\n\n" + traceback.format_exc())

    def _render_preview(self, result):
        """把 NormalizeResult 渲染到状态文本框。"""
        self._status_text.delete("1.0", "end")
        self._status_text.insert("end", f"预览：{result.input_path}\n")
        self._status_text.insert(
            "end",
            f"将处理：{result.paragraphs_processed} 段、{result.tables_processed} 个表、{result.images_processed} 张图\n",
        )
        if result.warnings:
            self._status_text.insert("end", "\n警告：\n")
            for w in result.warnings:
                self._status_text.insert("end", f"  · {w}\n")
        if result.changes:
            self._status_text.insert("end", "\n变更：\n")
            for c in result.changes:
                self._status_text.insert("end", f"  · {c}\n")
        self._status_text.insert("end", f"\n完成（dry-run，未修改文件）。耗时 {result.duration_ms}ms。\n")

    def _append_status(self, line):
        self._status_text.insert("end", line)
        self._status_text.see("end")


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


def main():
    root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
