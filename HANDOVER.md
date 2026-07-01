# 项目交接文档

> 接手人第一份必读。讲清楚这是什么、做到哪了、坑在哪、别走哪些路、下一步做什么。
>
> 最后更新：v0.2.0（2026-06-30）

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

## 2. 当前进度（v0.2.0）

### 2.1 版本时间线

| 版本 | commit | 主要变化 |
|---|---|---|
| 0.1.0 | `1b44f0e` | 初始版本（双实现 + 14 测试）|
| 0.1.1 | `a861ab7` | 保留封面/目录原版式（不再一刀切规范化）|
| 0.1.2 | `e2fb833` | 修表格首行缩进 bug（OOXML 字符单位）|
| **0.2.0** | `62c6c34` | **健壮性与 UX 全面升级（本次）** |

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

### 2.4 已确认的方案

- **架构**：Python（GUI + 库）+ Node.js（CLI + 库）双实现，共享 config.json 语义
- **核心库入口**：`from normalizer import normalize_docx, inspect_docx, NormalizeError, ...`
- **签名（新）**：`normalize_docx(cfg, input, output, *, backup=True, dry_run=False, return_result=False, on_warning=None)`
- **签名（向后兼容）**：不传新参数时返回 `str(output_path)`
- **错误体系**：`NormalizeError` 异常类，5 子类带中文 message + hint
- **备份策略**：固定 `.bak` → `.bak.1` → `.bak.2` ...（不暴露配置）
- **优雅降级**：非致命 warning 继续执行，致命 error 抛 NormalizeError
- **CI**：GitHub Actions + 推送 `v*` tag 自动构建 exe + zip + Release

---

## 3. 目录结构

```
goodDocument/
├── app.py                       # tkinter GUI 入口（486 行）
├── formatter.py                 # 核心排版引擎（生成文档，501 行）
├── normalizer.py                # ⚠️ shim 文件（re-export 自 normalizer/ 包）
├── normalizer/                  # ⭐ 子项目 C 拆分后的包
│   ├── __init__.py              # 公共 API + 内部辅助（re-export）
│   ├── errors.py                # NormalizeError + 5 子类 + NormalizeResult + WarningCollector
│   ├── paths.py                 # 路径校验 + 备份（_validate_paths / _backup_source）
│   ├── normalize.py             # 样式应用 + 段落/表格/图片 + 标题检测 + TOC 处理
│   └── core.py                  # normalize_docx 主入口（薄包装层）
├── version.py                   # ⭐ 单一来源 __version__ = "0.2.0"
├── config_model.py              # 配置加载/保存/校验
├── paths.py                     # 路径工具（PyInstaller 兼容）
├── config.json                  # 排版参数配置
├── DocFormatter.spec            # PyInstaller 打包配置
│
├── package.json                 # Node.js 项目配置（version: 0.2.0）
├── generate.js                  # Node.js CLI 入口
├── formatter.js                 # Node.js 排版核心库
│
├── .github/
│   └── workflows/
│       └── build.yml            # ⭐ GitHub Actions：推送 v* tag 自动构建
│
├── assets/
│   └── figure1.png              # 示例图片
│
├── tests/                       # ⭐ 42 个测试全过
│   ├── __init__.py
│   ├── test_app_paths.py        # 4 个路径工具测试
│   ├── test_config_model.py     # 2 个配置测试
│   ├── test_normalize_docx.py   # 36 个核心测试（14 旧 + 22 新）
│   ├── _helpers.py              # 边界 docx 构造器（垃圾/截断/嵌套表/空表/无法识别）
│   ├── _corrupt_source.docx     # _helpers 用的 fixture（自动生成）
│   └── __pycache__/             # .gitignore
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
│   ├── DocFormatter.exe         # 19.98 MB
│   └── config.json
│
├── README.md                    # ⭐ 全面重写（200 行，4 场景 + 10 FAQ + 健壮性章节）
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
| 改错误信息 | `normalizer/errors.py` 各子类 + `normalizer/core.py` 的 try/except |
| 改备份策略 | `normalizer/paths.py` 的 `_backup_source` |
| 改 GUI | `app.py`（486 行，结构清晰） |
| 改 CLI | `generate.js` 的 `main()` |
| 改字号/字体工具 | `formatter.py` 顶部常量 + `set_run_font` / `set_zero_indent` |

### 4.2 必须知道的"魔法函数"

```python
# formatter.py 中的关键工具
resolve_size(val)              # "小四" 或 12 → 12.0 (pt)
resolve_align(a)               # "left" → WD_ALIGN_PARAGRAPH.LEFT
set_run_font(run, zh, en, ...) # 设置 run 中英文字体（关键：写 w:rFonts + eastAsia）
set_para_spacing(...)          # 段落行距/段前段后
set_zero_indent(para)          # 显式清零所有缩进（含字符单位 firstLineChars）
clear_style_indent_chars(s)    # 清样式 <w:ind> 里的字符单位属性（Normal 样式必调）
add_page_number_field(para, fn)  # 插 PAGE 域（页码）

# normalizer 包中的关键函数
normalize_docx(cfg, in, out, *, backup, dry_run, return_result, on_warning)  # 主入口
_validate_paths(in, out)       # 抛 NormalizeError 子类
_backup_source(in)              # 返回 .bak 路径或 None
_normalize_paragraphs(doc, ...) # 段落遍历：标题/正文/题注/参考文献
_normalize_tables(doc, ...)     # 表格遍历：单元格段落全部独立
_mark_toc_fields_dirty(doc)     # 标记 TOC 域 dirty（让 Word 重新生成目录）
inspect_docx(path)               # 返回 DocxInspection 数据类
```

### 4.3 异常体系（用户友好化的关键）

```python
NormalizeError(message, *, cause=None, hint=None)
├── InputNotFoundError      # 找不到文件
├── InvalidFileTypeError    # 不是 .docx
├── CorruptDocxError        # zip 或 XML 损坏
├── OutputNotWritableError  # 权限/磁盘
└── SameInputOutputError    # 输出 = 输入

# 使用模式（normalizer/core.py 入口）
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
| PyInstaller | 6.x |
| python-docx | 最新（`pip install python-docx`）|
| docx (npm) | 9.7.1 |
| pytest | 8.x |

### 5.2 一次性 setup

```bash
# Python
pip install python-docx pyinstaller pytest
python -m pytest tests/ -v  # 42 passed

# Node.js
npm install
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

**修复**（`normalizer/core.py`）：catch `Exception` 兜底，按类名（`Package` / `Zip` / `Found`）判断是否归为 `CorruptDocxError`。

**教训**：python-docx 抛的异常类型**不只一种**包损坏情况，写 try/except 要宽。

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

### 7.3 GUI 进度条

**讨论过**但**延后**。原因：
- 规范化通常很快（<5 秒），进度条意义不大
- 大文档场景下需要，但当前未优化
- **未来时机**：性能优化时一起做

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

---

## 8. 下一步要做什么

### 8.1 立即（接手后第一周）

- [ ] **看 `dist/DocFormatter.exe` v0.2.0 在真实机器上跑起来**
  - Windows 10/11 双击启动
  - 选个真实的 docx 规范化
  - 验证错误信息是否友好
  - 验证 dry-run 是否正常
- [ ] **等 GitHub Actions v0.2.0 构建完成**
  - https://github.com/73mtq/goodDocument/actions
  - 下载自动构建的 exe 与本地对比
- [ ] **编辑 GitHub Release v0.2.0 的 release notes**
  - CI 自动生成的可能太简单
  - 建议加：4 子项目摘要 + 14 commit 列表 + 截图

### 8.2 短期（1-2 周）

- [ ] **加 i18n 错误信息基础架构**（如果需要支持英文用户）
  - 把 `NormalizeError` message 改成模板 `not_found: "{path} 不存在"`
  - 加 `locale` 参数
- [ ] **formatter.py 拆分**（如果感觉 500 行+ 太大）
  - 拆成 `font.py` + `spacing.py` + `Formatter` 类
- [ ] **加 GUI 进度条**（如果大文档场景下规范化慢）
  - 用 `ttk.Progressbar`
  - 在 `_normalize_*` 函数里加进度回调

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
