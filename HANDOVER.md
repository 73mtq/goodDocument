# 项目交接文档

> 接手人第一份必读。讲清楚这是什么、做到哪了、坑在哪、别走哪些路、下一步做什么。
>
> 最后更新：v0.5.0（2026-09-22）

---

## 1. 项目目标

`goodDocument` 是一个**配置驱动的 Word 文档规范化排版工具**，参照 GB/T 7714 参考文献著录规则与党政机关公文格式。

**核心定位**：

- **不**替代 Word（不是所见即所得编辑器）
- **不**做"AI 自动排版"（不做内容理解）
- **是**一套可复用的排版规范系统 + 工具链

**用户场景**：

1. 选一份已有的 `.docx`，一键统一页面/标题/正文/图表/参考文献/页码格式
2. 从零生成符合规范的示例文档
3. 所有排版参数集中在 `config.json`，改配置不改代码

**双语言实现**（共享同一份 `config.json` 语义）：

- **Python 版**（主力）：tkinter GUI + python-docx，提供 .exe 桌面应用
- **Node.js 版**：CLI + docx 库，提供命令行/脚本场景

---

## 2. 当前进度（v0.5.0）

### 2.1 版本时间线

| 版本 | commit | 主要变化 |
|---|---|---|
| 0.1.0 | `1b44f0e` | 初始版本（双实现 + 14 测试）|
| 0.1.1 | `a861ab7` | 保留封面/目录原版式（不再一刀切规范化）|
| 0.1.2 | `e2fb833` | 修表格首行缩进 bug（OOXML 字符单位）|
| 0.2.0 | `62c6c34` | 健壮性与 UX 全面升级（5 类异常 / dry_run / 备份 / warnings）|
| 0.2.1 | *(历史)* | **`backup` 默认 `True` → `False`（修 .bak 无限累计）** |
| 0.3.0 | *(历史)* | **全项目审查修复：识别误判、页脚页码、变更收集、跨实现统一、测试与 CI** |
| 0.4.0 | *(历史)* | **桌面界面重做（朱批纸面）+ 拖拽 + 后台执行 + 高 DPI + 快捷键** |
| **0.5.0** | *(本次)* | **图片能力补全：统一尺寸模式 + 浮动图片 + 安全阀 + 图题防分页** |

### 2.2 0.2.0 交付清单

| 模块 | 状态 |
|---|---|
| 5 种 `NormalizeError` 子类（中文错误信息）| ✅ |
| `NormalizeResult` 数据类 | ✅ |
| `WarningCollector` 上下文管理器 | ✅ |
| 自动备份（`.bak` / `.bak.1` / ...）| ✅ |
| `dry_run` 模式（GUI 按钮 + CLI flag）| ✅ |
| GUI 错误弹窗 + ScrolledText 状态区 | ✅ |
| CLI `--dry-run` / `--examples` flag | ✅ |
| `normalizer.py` 拆为 `normalizer/` 包（5 文件）| ✅ |
| README 全面重写（4 场景 + 10 FAQ）| ✅ |
| CHANGELOG.md | ✅ |
| GitHub Actions 自动构建 workflow | ✅ |
| v0.2.0 tag 推送（已触发 CI）| ✅ |
| 42 个测试全过 | ✅ |

### 2.3 4 个子项目（B/C/D 全部完成 + A 已完）

| 子项目 | 范围 | commits | 状态 |
|---|---|---|---|
| **A 健壮性与 UX** | 错误处理 / 备份 / dry-run / warnings | `730c7cf` … `3dc1e2e` | ✅ |
| **B 文档与示例** | README / FAQ / CLI --examples | `efc741e` `5d79e02` | ✅ |
| **C 代码结构** | normalizer.py 拆为 normalizer/ 包 | `51e782b` | ✅ |
| **D 打包与发布** | v0.2.0 / GitHub Actions | `9c6cf20` `46fc204` | ✅ |

### 2.5 0.2.1 调整（`backup` 默认关闭）

**问题**：0.2.0 默认 `backup=True`，每次规范化产生 `.bak` / `.bak.1` / `.bak.2` ... 无限累计（实测工作目录累积 27 个 = 38MB 垃圾）。

**修复**（1 个 fix 提交）：

| 文件 | 变更 |
|---|---|
| `normalizer/core.py` | `backup=True` → `backup=False`（默认） |
| `tests/test_normalize_docx.py` | `BackupAndDryRunTests` 重写：新增 `test_default_no_backup_does_not_create_bak`；`test_backup_creates_bak_file` 改名 `test_backup_true_creates_bak_file` 并显式 `backup=True`；`test_backup_increments_when_exists` 同理 |
| `CHANGELOG.md` | [Unreleased] 加 Changed + Fixed 条目 |
| `README.md` | 健壮性保证 + FAQ Q7 改描述 |
| `HANDOVER.md` | 2.4 签名 + 备份策略 + 新增 7.7 决策 |
| `version.py` / `package.json` | `0.2.0` → `0.2.1` |
| 工作目录 | 删 27 个 `规范文档示例.docx.bak*` |

**测试**：43 passed（原 42 + 新增 1）。

**教训**：自动备份类功能默认**关**（opt-in），opt-out 容易污染用户工作目录。GUI/CLI 暂不加 backup 开关（用户呼声再加）。

### 2.6 0.3.0 变更摘要（全项目审查修复）

**识别的缺陷**（均已修复，测试覆盖）：

- Node.js 版正文首行缩进单位错误（半磅值当 twips，2 字小四实际输出 48 而非 480）
- Node.js 版 `detectHeadingLevel` 自定义规则因 `lv` 未定义而失效
- 题注识别范围过宽：正文段"图片…""表题…"被当作题注
- 标题识别范围过宽："1.5 倍行距…"被当作三级标题
- `inspect_docx().has_header_footer` 恒为 True
- 页脚页码更新固定清空首段，页码位于其他段落时产生两个 PAGE 域
- `NormalizeResult.changes` 与处理过程中的 warnings 从未产生（设计文档承诺未实现）
- Python 版忽略 `page.size` / `table.align`；`config.json` 有 5 个无效键
- CI 无法成功（缺 pytest、`cache: pip` 缺少依赖文件）

**交付**：

| 模块 | 状态 |
|---|---|
| 题注 / 标题 / 参考文献识别收紧 | ✅ |
| 页脚 PAGE 段落定位改写 | ✅ |
| 变更收集与准确计数（跳过封面/目录）| ✅ |
| 警告回调与 `NormalizeResult.warnings` 打通 | ✅ |
| Python 版 `page.size`（A4/Letter）与 `table.align` | ✅ |
| Node.js 首行缩进 twips 修复、自定义标题规则修复、单元格段距统一 | ✅ |
| 根目录 `normalizer.py` 兼容层删除；`normalizer/__init__.py` 只保留公开 API | ✅ |
| 测试样例改为运行时生成；删除提交的二进制样本 | ✅ |
| `tests/node/` Node.js 测试 + `npm test` | ✅ |
| `tests/test_formatter.py`、`tests/test_version_consistency.py` | ✅ |
| `requirements.txt` / `requirements-dev.txt` / `pyproject.toml`（ruff） | ✅ |
| CI 拆分为 test / build 两个任务 | ✅ |
| 版本 0.3.0；CHANGELOG / README / HANDOVER 同步 | ✅ |

**当时测试**：Python 64 个 + Node.js 11 个，全部通过。

### 2.7 0.4.0 变更摘要（桌面界面重做）

**视觉方向**：朱批纸面 —— 纸色打底、墨色承载信息、朱砂只出现在主操作上；无阴影、1px 分隔线、以字号与字重建立层级。

| 模块 | 状态 |
|---|---|
| 新增 `ui_theme.py`（颜色/字体令牌 + clam 主题全部 ttk 样式 + 取值映射）| ✅ |
| 主界面改为三步单列流程：选择文档 → 规范化后另存为 → 预览/开始规范化 | ✅ |
| 文件拖拽（`tkinterdnd2`，含空格与中文路径）| ✅ |
| 选中文档后显示摘要（段落/表格/图片/标题层级/页眉页脚）| ✅ |
| 结果区「变更 / 警告 / 日志」分段展示 | ✅ |
| 后台线程执行 + 进度条 + 忙碌时禁用操作 | ✅ |
| 完成后状态栏内联「打开文件夹」，去掉模态询问 | ✅ |
| 高级设置改为右侧可滚动折叠面板（7 组）| ✅ |
| 下拉选项值中文化（center → 居中、above → 上方、top → 顶部）| ✅ |
| 快捷键 Ctrl+O / Ctrl+P / Ctrl+Enter / Ctrl+S / Ctrl+G / F1 | ✅ |
| Windows 高 DPI 感知 | ✅ |
| 主操作唯一突出，其余收进「更多」菜单 | ✅ |
| `tests/test_ui_theme.py`（令牌、对比度、映射）| ✅ |
| exe 构建验证：tkdnd 已打入归档，启动后拖拽可用 | ✅ |

**当时测试**：Python 70 个 + Node.js 11 个，全部通过。

### 2.8 0.5.0 变更摘要（图片能力补全）

**起因**：用户给出一个 Word VBA 宏（`LockAspectRatio = msoFalse` 后把所有图片写死成 12×9 cm），要求把这个「照片处理」能力并入本工具，并顺带审查还有哪些可优化处。

**识别出的既有缺陷**（均已修复）：

- **浮动图片完全未被处理** —— `doc.inline_shapes` 只覆盖嵌入型（`wp:inline`），环绕方式为四周型 / 浮于文字上方的图片（`wp:anchor`）一张都不动。VBA 的 `InlineShapes` 同样有这个盲区
- **`a:ext` 未同步** —— 图片尺寸有 `wp:extent` 与 `pic:spPr/a:xfrm/a:ext` 两处，只改前者时 WPS / LibreOffice 可能仍按旧尺寸渲染
- **图片会溢出页面** —— 早期只做「只缩不放」，没有把目标宽度与正文可用宽度对齐的安全阀
- **`images_processed` 名不副实** —— 等于文档中图片总数，而非实际调整数，预览摘要「将处理 N 张图」会误导
- **`inspect_docx().image_count`** 漏掉浮动图片
- **图与图题会被分页拆开** —— 未设置 `w:keepNext`

**交付**：

| 模块 | 状态 |
|---|---|
| `figure.sizeMode`（`fit` 等比 / `fixed` 统一尺寸）| ✅ |
| `fixedWidthCm` / `fixedHeightCm` / `keepAspectRatio`（强制拉伸 vs 等比适配）| ✅ |
| 浮动图片（`wp:anchor`）与嵌入图片统一处理 | ✅ |
| `clampToTextWidth` 宽度安全阀（默认开，带警告）| ✅ |
| `skipInTables` 表格内图片默认跳过（带变更说明）| ✅ |
| 文本框内图片跳过；页眉/页脚图片给出提示 | ✅ |
| `keepWithCaption` 图题防分页（`w:keepNext`）| ✅ |
| `wp:extent` 与 `a:xfrm/a:ext` 同步写入 | ✅ |
| GUI「图」分组新增 6 个控件 + 中文取值标签（等比缩放 / 统一尺寸）| ✅ |
| Node.js `computeImageTransform` 与 Python 端语义对齐 | ✅ |
| 生成端 `Formatter._figure_size` 支持两种模式 | ✅ |
| 测试：Python 70 → 89，Node.js 11 → 16 | ✅ |

**当前测试**：Python 89 个 + Node.js 16 个，全部通过。

### 2.4 已确认的方案

- **架构**：Python（GUI + 库）+ Node.js（CLI + 库）双实现，共享 config.json 语义
- **核心库入口**：`from normalizer import normalize_docx, inspect_docx, NormalizeError, ...`
- **签名（新）**：`normalize_docx(cfg, input, output, *, backup=False, dry_run=False, return_result=False, on_warning=None)`
- **签名（向后兼容）**：不传新参数时返回 `str(output_path)`
- **错误体系**：`NormalizeError` 异常类，5 子类带中文 message + hint
- **备份策略（0.2.1 调整）**：默认 `backup=False`，需要时显式传 `backup=True`。递增策略 `.bak` → `.bak.1` → ... 保留。原因：默认开启会无限累计污染工作目录（实测累积 27 个 .bak）。
- **优雅降级**：非致命 warning 继续执行，致命 error 抛 NormalizeError
- **图片尺寸（0.5.0）**：两种模式共存，`fit` 为默认（向后兼容，行为不变）；`fixed` 才启用统一尺寸。安全阀 `clampToTextWidth` 与 `skipInTables` 默认**开** —— 宁可少改也不能把版面弄坏，跳过时在「变更」列表写明原因
- **CI**：GitHub Actions + 推送 `v*` tag 自动构建 exe + zip + Release

---

## 3. 目录结构

```
goodDocument/
├── app.py                       # tkinter GUI 入口（拖拽、后台执行、快捷键）
├── settings_model.py            # 高级设置分组/行定义与配置值转换（纯数据层）
├── ui_theme.py                  # ⭐ 界面主题（令牌 + ttk 样式 + 取值映射）
├── formatter.py                 # 核心排版引擎（生成文档）
├── font_sizes.py                # 中文字号名与磅值映射（GUI / 配置校验 / 排版共用）
├── normalizer/                  # ⭐ 子项目 C 拆分后的包
│   ├── __init__.py              # 公共 API（仅公开入口）
│   ├── errors.py                # NormalizeError + 5 子类 + NormalizeResult + WarningCollector
│   ├── paths.py                 # 路径校验 + 备份（_validate_paths / _backup_source）
│   ├── normalize.py             # 样式应用 + 段落/表格/图片 + 标题检测 + TOC 处理
│   └── core.py                  # normalize_docx 主入口（薄包装层）
├── version.py                   # ⭐ 单一来源 __version__ = "0.5.0"
├── config_model.py              # 配置加载/保存/校验
├── paths.py                     # 路径工具（PyInstaller 兼容）
├── config.json                  # 排版参数配置
├── DocFormatter.spec            # PyInstaller 打包配置
├── requirements.txt             # ⭐ Python 运行依赖（含 tkinterdnd2）
├── requirements-dev.txt         # ⭐ 开发依赖（pytest / pyinstaller / ruff）
├── pyproject.toml               # ⭐ ruff 与 pytest 配置
│
├── package.json                 # Node.js 项目配置（version: 0.5.0）
├── generate.js                  # Node.js CLI 入口（含 require.main 守卫）
├── formatter.js                 # Node.js 排版核心库
│
├── .github/
│   └── workflows/
│       └── build.yml            # ⭐ CI：test（push/PR）+ build（v* tag）
│
├── assets/
│   └── figure1.png              # 示例图片
│
├── tests/                       # ⭐ Python 89 个测试全过
│   ├── __init__.py
│   ├── test_app_paths.py        # 路径工具测试
│   ├── test_config_model.py     # 配置测试
│   ├── test_formatter.py        # ⭐ 生成器与工具函数测试
│   ├── test_normalize_docx.py   # 核心规范化 + 回归测试（样例运行时生成）
│   ├── test_ui_theme.py         # ⭐ 主题令牌、对比度、取值映射
│   ├── test_version_consistency.py  # ⭐ version.py 与 package.json 一致性
│   ├── _helpers.py              # 边界与图片 docx 构造器（垃圾/截断/嵌套表/空表/浮动图片）
│   └── node/                    # ⭐ Node.js 测试（node --test）
│
├── docs/
│   └── superpowers/
│       ├── specs/               # 设计文档
│       │   └── 2026-06-30-normalizer-robustness-design.md
│       └── plans/               # 实施计划
│           └── 2026-06-30-normalizer-robustness-plan.md
│
├── .gitattributes               # ⭐ 统一行尾（LF / 自动检测）
├── .gitignore                   # 已 gitignore：.omc/ .omo/ *.docx.bak __pycache__/ build/ dist/ node_modules/
│
├── dist/                        # .gitignore，打包产物
│   ├── DocFormatter.exe
│   └── config.json
│
├── README.md                    # ⭐ 使用说明
├── CHANGELOG.md                 # ⭐ 版本变更记录
└── HANDOVER.md                  # ⭐ 本文件
```

---

## 4. 关键代码

### 4.1 入口（要改时从这里开始）

| 需求 | 看这个 |
|---|---|
| 改规范化逻辑 | `normalizer/normalize.py` 的 `_normalize_*` 函数 |
| 改样式应用 | `normalizer/normalize.py` 的 `_apply_*` 函数 |
| 改错误信息 | `normalizer/errors.py` 各子类 + `normalizer/normalize.py` 的 `_open_document` |
| 改备份策略 | `normalizer/paths.py` 的 `_backup_source` |
| 改 GUI | `app.py` |
| 改高级设置项/分组 | `settings_model.py` 的 `SETTING_GROUPS` |
| 改 CLI | `generate.js` 的 `main()` |
| 改字号/字体工具 | `font_sizes.py` 的字号表 + `formatter.py` 的 `set_run_font` / `set_zero_indent` |

### 4.2 必须知道的"魔法函数"

```python
# formatter.py 中的关键工具
resolve_size(val)              # "小四" 或 12 → 12.0 (pt)
resolve_align(a)               # "left" → WD_ALIGN_PARAGRAPH.LEFT
set_run_font(run, zh, en, ...) # 设置 run 中英文字体（关键：写 w:rFonts + eastAsia）
set_para_spacing(...)          # 段落行距/段前段后
set_zero_indent(para)          # 显式清零所有缩进（含字符单位 firstLineChars）
clear_style_indent_chars(s)    # 清样式 <w:ind> 里的字符单位属性（Normal 样式必调）
apply_document_styles(doc, cfg)  # Normal / Heading 基础样式（生成与规范化共用）
add_page_number_field(para, fn)  # 插 PAGE 域（页码）

# normalizer 包中的关键函数
normalize_docx(cfg, in, out, *, backup, dry_run, return_result, on_warning)  # 主入口
_validate_paths(in, out)       # 抛 NormalizeError 子类
_backup_source(in)              # 返回 .bak 路径或 None
_open_document(path)            # 打开 docx，损坏时抛 CorruptDocxError
_normalize_paragraphs(doc, ...) # 段落遍历：标题/正文/题注/参考文献（返回处理计数，支持 on_change）
_normalize_tables(doc, ...)     # 表格遍历：单元格段落全部独立
_normalize_inline_images(...)   # 图片尺寸策略（fit/fixed）；返回实际调整数
_iter_picture_frames(doc)       # 遍历 body 中所有图片框 (kind, frame)：inline + anchor
_set_frame_extent(...)          # 写尺寸，同时落 wp:extent 与 a:xfrm/a:ext
_available_text_width_emu(doc)  # 正文可用宽度（页宽 − 左右边距），宽度安全阀用
_set_keep_with_next(para)       # w:keepNext，图与图题防分页
_mark_toc_fields_dirty(doc)     # 标记 TOC 域 dirty（让 Word 重新生成目录）
inspect_docx(path)              # 返回 DocxInspection 数据类
```

### 4.3 异常体系（用户友好化的关键）

```python
NormalizeError(message, *, cause=None, hint=None)
├── InputNotFoundError      # 找不到文件
├── InvalidFileTypeError    # 不是 .docx
├── CorruptDocxError        # zip 或 XML 损坏
├── OutputNotWritableError  # 权限/磁盘
└── SameInputOutputError    # 输出 = 输入

# 使用模式（normalizer/normalize.py 的 _open_document）
try:
    doc = Document(input_path)
except (zipfile.BadZipFile, lxml.etree.XMLSyntaxError) as e:
    raise CorruptDocxError(f"docx 文件已损坏: {input_path}",
                           cause=e, hint="用 Word 打开重新保存一次") from e
```

---

## 5. 环境配置

### 5.1 开发环境

| 工具 | 版本 |
|---|---|
| Python | 3.12+（conda / system 都可）|
| Node.js | 20+（`node` + `npm`）|
| PyInstaller | 6.x（需 `pyinstaller-hooks-contrib`，自带 tkinterdnd2 hook）|
| python-docx | 1.2.x（`pip install -r requirements.txt`）|
| tkinterdnd2 | 0.6.x（GUI 拖拽）|
| docx (npm) | 9.7.1 |
| pytest | 8.x |
| ruff | 0.15+ |

### 5.2 一次性 setup

```bash
# Python
pip install -r requirements-dev.txt
python -m pytest tests/ -v  # 89 passed
ruff check .

# Node.js
npm ci
npm test                    # 11 passed
node generate.js --examples
```

### 5.3 打包 exe

```bash
pyinstaller DocFormatter.spec --noconfirm --clean
# 产物：dist/DocFormatter.exe（~20 MB）
```

### 5.4 发新版本（CI 自动）

```bash
# 1. 改 version.py 和 package.json 的 version
# 2. 改 CHANGELOG.md 加新条目
# 3. 提交 + 打 tag + 推送
git add . && git commit -m "..."
git tag -a v0.X.0 -m "v0.X.0 — ..."
git push origin main v0.X.0
# GitHub Actions 自动构建 + 发 Release
```

---

## 6. 已经踩过的坑（重要！）

### 坑 1：OOXML 字符单位（`firstLineChars`）继承 Normal 样式

**症状**：表格里"明明 XML 里 firstLine=0"，Word 段落对话框还是显示"首行缩进 2 字符"。

**根因**（两层叠加）：
1. OOXML `<w:ind>` 同时支持数值单位（`firstLine`）和字符单位（`firstLineChars`），**字符单位优先**
2. 段落会从样式继承，**Normal 样式的 `firstLineChars` 也会被继承**
3. `python-docx` **完全不读写** `firstLineChars` 属性（源码 0 处提及）

**修复**（`formatter.py`）：
- `set_zero_indent(para)`：清段落级 `firstLineChars` 等
- `clear_style_indent_chars(style)`：清样式级 `firstLineChars` 等
- 两层都要清！

**教训**：OOXML 解析不要只看 python-docx 的 API，必须直接看 XML 输出。诊断时用 `lxml.etree` 手动跟踪属性链。

### 坑 2：python-docx `PackageNotFoundError`

**症状**：传入非 zip 的垃圾文件时，python-docx 抛 `docx.opc.exceptions.PackageNotFoundError`，不是 `BadZipFile`。

**现状**（`normalizer/normalize.py`）：显式捕获 `zipfile.BadZipFile`、`lxml.etree.XMLSyntaxError` 与 `docx.opc.exceptions.PackageNotFoundError`，统一转为 `CorruptDocxError`；其他异常继续上抛。

**教训**：python-docx 对损坏包抛出的异常类型不只一种，逐一列出并在测试中覆盖。

### 坑 3：GitHub OAuth 缺 `workflow` scope

**症状**：`git push` 拒推 `.github/workflows/*.yml`，提示 `refusing to allow an OAuth App to create or update workflow`。

**解决**：
- **方案 A**（最简）：用 PAT 时勾选 `workflow` scope
- **方案 B**：本机用 `gh auth refresh -h github.com -s workflow` 后重推
- **方案 C**（当前做法）：先 commit 不带 workflow 的代码，push 完再单独 commit workflow

**教训**：涉及 GitHub Actions 推送要么用 PAT（带 workflow scope），要么把 workflow 文件单独管理。

### 坑 4：中文引号在 Python 字符串里的转义

**症状**：`app.py` 里写 `"提示：点"预览变更"..."` 报 SyntaxError。

**修复**：用单引号包字符串 `'提示：点"预览变更"...'`，或者用三引号 `"""..."""`。

**教训**：中文 + Python 字符串里出现英文引号必转义或换引号。

### 坑 5：pytest `assertRaises(ValueError)` 在新代码里失败

**症状**：旧测试用 `with self.assertRaises(ValueError):` 期望 `normalize_docx` 抛 ValueError，0.2.0 后改成抛 `NormalizeError`（不继承 ValueError），测试失败。

**修复**：统一改 `assertRaises(NormalizeError)`（也是基类，会捕获 5 个子类）。

**教训**：自定义异常层级设计好之后，**所有调用方要更新 catch 类型**。

### 坑 6：PowerShell 中文 stdout 乱码

**症状**：跑 `node generate.js --examples` 在 PowerShell 里输出 `?` 而不是中文。

**修复**：先 `chcp 65001` + `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8`，或者输出到文件再读。

**教训**：跨平台 CLI 工具的测试要避免 PowerShell 默认 GBK 编码干扰。

### 坑 7：`a:extLst` 里也有 `a:ext`，无差别遍历会写坏文件

**症状**：想同步图片的 `a:xfrm/a:ext`，写了 `frame_el.iter(qn("a:ext"))` 一起改 —— 结果把 `<a:extLst><a:ext uri="...">` 这类**扩展节点**也写了 `cx`/`cy`，文件出现莫名的渲染问题。

**根因**：DrawingML 里有两个不同层级的同名元素：
- `pic:spPr/a:xfrm/a:ext` —— 图片容器的真实尺寸，需要同步
- `a:extLst/a:ext` —— 扩展列表项，**只带 `uri` 属性**，与尺寸无关

**修复**（`normalizer/normalize.py` 的 `_set_frame_extent`）：只精确定位 `a:xfrm` 下的 `a:ext`：

```python
for xfrm in frame_el.iter(qn("a:xfrm")):
    inner = xfrm.find(qn("a:ext"))
    if inner is not None:
        inner.set("cx", ...)
```

**教训**：OOXML 的元素名在不同命名空间/层级下会重复，用 `findall(".//tag")` 前先想清楚会不会误伤。

### 坑 8：改 `wp:inline` 的 XML 时丢掉命名空间声明，文件直接解析失败

**症状**：测试夹具里把 `document.xml` 的 `<wp:inline ...>` 用正则替换成 `<wp:anchor ...>`，保存后用 python-docx 打开报 `XMLSyntaxError: Namespace prefix a on graphicFrameLocks is not defined`。

**根因**：python-docx 生成的 `<wp:inline>` **把 `xmlns:a` / `xmlns:pic` / `xmlns:wp` 等声明挂在自己的开始标签上**。用 `<wp:inline[^>]*>` 整个替换掉，声明就一起没了，后续所有 `a:` 前缀都成了未定义。

**修复**（`tests/_helpers.py` 的 `make_docx_with_floating_image`）：捕获原标签的属性串并原样带过去：

```python
re.sub(r"<wp:inline(?P<attrs>[^>]*)>",
       lambda m: "<wp:anchor " + ATTRS + m.group("attrs") + ">" + HEAD, xml)
```

顺带注意 `wp:anchor` 的子元素顺序（schema 要求）：`simplePos` → `positionH` → `positionV` → `extent` → `effectExtent?` → `wrapNone/wrapSquare/...` → `docPr` → `cNvGraphicFramePr` → `graphic`。把 `wrapNone` 塞在 `extent` 前面同样是错的。

**教训**：改写 OOXML 片段时，命名空间声明和子元素顺序两件事都不能想当然。

### 坑 9：页面尺寸经 twips 往返换算，厘米值不精确

**症状**：测试断言 `wp:extent` 钳制后的宽度等于 `15.5 × 360000 = 5580000` EMU，实际得到 `5580380`。

**根因**：OOXML 的 `w:pgSz/@w:w` 与 `w:pgMar/@w:*` 单位是 **twips（1 twip = 635 EMU）**，python-docx 写入时按 twips 取整。`21cm → 11906 twips → 7560310 EMU`，边距同理，累加后与「厘米直接换 EMU」差几百 EMU。

**修复**：测试改为按厘米比较并给容差：

```python
def assertCm(self, emu, expected_cm):
    self.assertAlmostEqual(emu / 360000, expected_cm, delta=0.01)
```

**教训**：凡是与页面几何相关的数值断言，**不要拿 EMU 精确相等**，按厘米给容差。

---

## 7. 不能再走的错误路线（明确放弃）

### 7.1 拆分 formatter.py

**讨论过**但**决定不做**。原因：
- formatter.py 501 行虽然不小，但**核心是 `Formatter` 类 + 工具函数**，结构清晰
- 现在 A/B/C/D 任务紧，动了风险高
- 收益小于风险
- **后续时机**：如果未来加更多 caption 类型（公式、代码块等），再拆

### 7.2 国际化（i18n）错误信息

**讨论过**但**不做**。原因：
- 用户群体是中文
- 错误信息中文 + hint 是核心卖点之一
- 加 i18n 增加复杂度但收益小
- **未来时机**：如有英文用户需求再做

### 7.3 GUI 进度条（0.4.0 已落地，方案见此）

**曾经的判断**"规范化通常很快（<5 秒），进度条意义不大" —— 0.4.0 已实现：后台线程执行 + `ttk.Progressbar` 不确定态动画 + 忙碌时禁用操作，见 `app.py` 的 `_set_busy`。

**仍然没做**：`_normalize_*` 里的**百分比进度回调**（当前只是转圈，不显示 37%）。如果将来出现「几千张图的大文档」场景，再往里加 `on_progress` 回调。

### 7.4 Python / Node.js 公共 API 抽象

**讨论过**但**不做**。原因：
- 两个实现的 API 形状不完全相同（Python 有 `normalize_docx`，Node.js 是 `createFormatter`）
- 强行抽象会让两边都用得不顺手
- 当前用户基本只用 Python，Node.js 是辅助
- **未来时机**：如果有 30% 公共逻辑，可考虑用 protobuf / JSON Schema 抽象

### 7.5 损坏 docx 修复工具

**讨论过**但**不做**。原因：
- 工具只规范化，不修复
- 修复 docx 是另一个工具的事
- **未来时机**：单独的 docx 修复工具项目

### 7.6 用户配置 `.bak` 命名策略

**讨论过**但**不做配置暴露**。原因：
- 增加 GUI/CLI 复杂度
- 固定 `.bak` → `.bak.N` 满足 99% 场景
- **未来时机**：如有用户反馈再考虑

### 7.7 `backup` 默认值（0.2.1 决策）

**0.2.0** 默认 `backup=True` → 无限累计 `.bak` / `.bak.1` / ... 污染工作目录（实测 27 个）。

**0.2.1** 改成默认 `backup=False`。原因：
- 默认开 → 用户每规范化一次就多一个 .bak，27 次后 38MB
- 99% 场景下输出文件就是用户想要的，不需要保留原文件
- 真要恢复 → 用户的源文件本来就有（normalize 是规范化，不是覆盖原文件）
- 想要备份 → 显式 `backup=True`

**教训**：自动备份类功能默认**关**，opt-in 比 opt-out 安全得多。GUI/CLI 暂不加 backup 开关（如果有用户呼声再加）。

---

## 8. 下一步要做什么

### 8.1 立即（接手后第一周）

- [ ] **看 `dist/DocFormatter.exe` 在真实机器上跑起来**
  - Windows 10/11 双击启动
  - 拖入一份真实 docx（含空格与中文路径），确认拖拽可用
  - 验证错误信息是否友好
  - 验证「预览变更」「打开文件夹」与快捷键
- [ ] **拿一份真实的「图片尺寸杂乱」文档验证 0.5.0 的图片能力**
  - 「统一尺寸」模式下确认所有目标图片宽度一致
  - 确认表格内图标被跳过且「变更」列表写明了原因
  - 用 WPS 打开输出的 docx，确认图片显示尺寸与 Word 一致（验证 `a:ext` 同步）
- [ ] **打 v0.5.0 tag 并确认 CI 通过**
  - test 任务：ruff + pytest + npm test
  - build 任务：构建 exe 并发布 Release
  - https://github.com/73mtq/goodDocument/actions
- [ ] **编辑 GitHub Release v0.5.0 的 release notes**
  - 建议概括统一尺寸模式、浮动图片、安全阀与图题防分页

### 8.2 短期（1-2 周）

- [ ] **加 i18n 错误信息基础架构**（如果需要支持英文用户）
  - 把 `NormalizeError` message 改成模板 `not_found: "{path} 不存在"`
  - 加 `locale` 参数
- [ ] **GUI 里把两种图片模式做成预设按钮**（"统一宽度 / 等比缩放"一键切换）
  - 现在要点开「图」分组逐项改，场景化预设更顺手

### 8.3 中期（1-2 月）

- [ ] **批量规范化**
  - GUI 加"批量选择文件"按钮
  - CLI 加 `--batch <dir>` 遍历目录
  - 进度报告（HTML/PDF）
- [ ] **损坏 docx 修复工具**（独立项目）
  - 用 lxml 解析 + 重建
  - 对应文档：哪些是 fatal error / 哪些可恢复
- [ ] **配置文件模板**
  - 加 `config/templates/` 目录
  - 提供 `学术论文.json` / `行政公文.json` / `技术报告.json` 等预设
  - GUI 加"加载模板"按钮

### 8.4 长期

- [ ] **VSCode 插件**（如果需要）
  - 在 VSCode 里直接预览规范化效果
  - 实时调整 config.json
- [ ] **Web 版**（如果有需求）
  - Python 后端 + 简单前端
  - 服务化部署
- [ ] **GB/T 7714 参考文献自动格式化**
  - 解析 DOI / arxiv 链接
  - 自动生成 GB/T 7714 规范的参考文献条目
  - 当前只负责悬挂缩进排版，**内容**还是用户写

---

## 9. 联系与继承

**项目地址**：https://github.com/73mtq/goodDocument

**关键 commit**（按时间顺序）：

- `e2fb833` — 表格首行缩进 bug 修复（OOXML 字符单位）
- `a861ab7` — 封面/目录原版式保留
- `3247513` — A 子项目设计文档
- `49d21b3` — A 子项目实施计划
- `730c7cf` … `3dc1e2e` — A 子项目实现（10 个 commit）
- `efc741e` — CLI --examples
- `5d79e02` — README + CHANGELOG
- `51e782b` — normalizer 包拆分
- `9c6cf20` — v0.2.0 版本号
- `46fc204` — GitHub Actions workflow
- `62c6c34` — .gitattributes

**设计/计划文档**：
- `docs/superpowers/specs/2026-06-30-normalizer-robustness-design.md`（A 子项目设计）
- `docs/superpowers/plans/2026-06-30-normalizer-robustness-plan.md`（A 子项目实施计划）

**核心用户**：
- 中文学术/公文写作场景
- 关注 GB/T 7714 参考文献规范
- 想要"一键统一"排版而不想手动改

**核心约束**：
- 不替换 Word
- 不做 AI 内容理解
- 配置驱动（不改代码改 JSON）
- 公开 API 必须向后兼容
- 中文错误信息是核心卖点

---

## 10. 一句话总结

**这是从 bug 修复出发 → 重构 → 拆分 → 文档化 → CI/CD 全套的 0.1.0 → 0.2.0 升级。下一步看真实用户反馈决定方向。**
