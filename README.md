# GoodDocument

> **配置驱动的 Word 文档规范化排版工具**，参照 GB/T 7714 参考文献著录规则与公文格式标准。

选一份已有的 `.docx`，**一键统一**页面、标题、正文、图表、参考文献和页码格式；也可以**从零生成**符合规范的示例文档。所有排版参数（字体、字号、行距、边距、对齐）均由 `config.json` 控制，无需改代码即可调整。

---

## 快速导航

- [功能特性](#功能特性)
- [快速开始](#快速开始)
- [使用场景](#使用场景)
- [配置说明](#配置说明)
- [中文字号对照表](#中文字号对照表)
- [CLI 速查](#cli-速查)
- [故障排查 FAQ](#故障排查-faq)
- [健壮性保证](#健壮性保证)
- [从源码构建](#从源码构建)
- [项目结构](#项目结构)
- [运行测试](#运行测试)
- [License](#license)

---

## 功能特性

- **一键规范化** — 打开任意 `.docx`，自动识别标题层级（一、/（一）/1./（1））、图表题注、参考文献段落，统一套用配置的格式系统
- **从零生成** — 生成包含正文、图片、表格、GB/T 7714 参考文献的完整示例文档
- **配置驱动** — 字体、字号、行距、边距、对齐、段前段后等全部参数集中在 `config.json` 中
- **表格独立排版** — 单元格使用独立的字体/字号/行距，显式置零首行缩进（OOXML 字符单位也清零）
- **图片智能缩放** — 默认宽度等于正文宽度，按宽高比等比缩放
- **页码自动插入** — 底部居中页码，PAGE 域自动编号
- **健壮性优先** — 中文错误信息、自动备份、dry-run 预览、优雅降级（详见 [健壮性保证](#健壮性保证)）
- **双语言实现** — Python（tkinter GUI + python-docx）和 Node.js（CLI + docx 库），共享同一份配置

---

## 快速开始

### 方式一：直接运行可执行文件（Windows，推荐）

从 [Releases](../../releases) 下载 `DocFormatter.exe`，双击即可启动 GUI。将 `config.json` 放在 exe 同目录即可自定义排版参数。

### 方式二：Python 运行

```bash
pip install python-docx
python app.py
```

### 方式三：Node.js CLI 运行

```bash
npm install
npm start                                       # 用默认 config.json 生成
node generate.js --examples                     # 看 6 个典型场景
node generate.js --set body.size=14 --set body.lineSpacing=2.0   # 临时改参数
node generate.js --dry-run                      # 预览，不写文件
```

---

## 使用场景

### 场景 1：学生论文（标准学术格式）

**目标**：宋体小四 1.5 倍行距首行缩进 2 字，标题黑体分级，图表题五号加粗，参考文献 GB/T 7714。

**操作**：
1. 启动 GUI（`DocFormatter.exe` 或 `python app.py`）
2. 选输入 `.docx`，点"开始规范化"
3. 默认配置就是学术格式

**等价 CLI**：
```bash
node generate.js --out 我的论文.docx
```

### 场景 2：工程师技术报告（紧凑型）

**目标**：表题在上方、行距紧凑、无首行缩进。

**操作**：GUI 里打开"高级设置" → "表" → 把"表题位置"改成 `above`；"正文" → 行距改 1.0、首行缩进改 0。

**等价 CLI**：
```bash
node generate.js --set table.captionPosition=above
node generate.js --set body.lineSpacing=1.0 --set body.firstLineIndentChars=0
```

### 场景 3：行政公文（方正仿宋）

**目标**：仿宋_GB2312 三号、行距 28-30 磅、上下边距 3.5cm 左右。

**等价 CLI**：
```bash
node generate.js --set body.font=仿宋_GB2312 --set body.size=16 --set body.lineSpacing=1.75
node generate.js --set page.margins.top=3.7 --set page.margins.bottom=3.5
```

### 场景 4：批量规范化多份文档

```python
from pathlib import Path
from normalizer import normalize_docx
import json

cfg = json.loads(Path("config.json").read_text(encoding="utf-8"))
for src in Path("papers").glob("*.docx"):
    out = src.with_name(src.stem + "_规范化.docx")
    normalize_docx(cfg, src, out)
    print(f"OK: {out}")
```

### 更多场景

`node generate.js --examples` 查看全部 6 个场景。

---

## 配置说明

所有排版参数集中在 `config.json` 中，结构如下：

| 区段 | 说明 | 关键参数 |
|------|------|----------|
| `page` | 页面设置 | `margins`（上下左右边距，cm）、`forceA4` |
| `pageNumber` | 页码 | `enabled`、`align`、`size`、`font` |
| `body` | 正文 | `font`（中文字体）、`asciiFont`（英数字体）、`size`、`lineSpacing`、`firstLineIndentChars`（首行缩进字数）、`align` |
| `headings` | 标题（H1-H4） | 每级独立设置 `font`、`size`、`bold`、`color`、`spaceBeforePt`、`spaceAfterPt`、`align` |
| `figure` | 图片 | `maxWidthCm`、`maxHeightCm`、`captionPosition`（above/below）、`captionSize`、`prefix` |
| `table` | 表格 | `captionPosition`、`cellSize`、`cellAlign`、`cellVAlign`、`headerBold`、`prefix` |
| `references` | 参考文献 | `style`（GB/T 7714）、`hangingIndentPt`（悬挂缩进）、`size`、`lineSpacing` |
| `output` | 输出 | `filename` |

默认配置参照公文格式：正文宋体小四 1.5 倍行距首行缩进 2 字，一级标题黑体小二，图表题注五号加粗居中，参考文献小五悬挂缩进。

**字号两种写法都支持**：
```json
"size": "小四"   // 中文名
"size": 12       // 磅值数字
```

---

## 中文字号对照表

| 字号名 | 磅值 (pt) | 字号名 | 磅值 (pt) |
|--------|-----------|--------|-----------|
| 初号 | 42 | 小四 | 12 |
| 小初 | 36 | 五号 | 10.5 |
| 一号 | 26 | 小五 | 9 |
| 小一 | 24 | 六号 | 7.5 |
| 二号 | 22 | 小六 | 6.5 |
| 小二 | 18 | 七号 | 5.5 |
| 三号 | 16 | 八号 | 5 |
| 小三 | 15 | | |
| 四号 | 14 | | |

CLI 查表：`node generate.js --list-sizes`

---

## CLI 速查

```bash
node generate.js                            # 用默认 config.json 生成
node generate.js --config 公文.json         # 指定配置
node generate.js --out 文件.docx            # 指定输出
node generate.js --set a.b=c                # 临时改配置，可多次
node generate.js --list-sizes               # 中文字号表
node generate.js --show-config              # 看生效配置
node generate.js --examples                 # 6 个场景示例
node generate.js --dry-run                  # 预览，不写文件
node generate.js -h                          # 详细 help
```

常用 `--set` 示例：
```bash
# 改正文格式
--set body.size=14
--set body.font=黑体
--set body.lineSpacing=2.0
--set body.firstLineIndentChars=0

# 改页边距
--set page.margins.left=3.0

# 改表格
--set table.captionPosition=below    # 或 above
--set table.cellAlign=center

# 改页码
--set pageNumber.enabled=true
--set pageNumber.align=center
```

---

## 故障排查 FAQ

### Q1: 双击 exe 没反应 / 一闪而过
A: 用命令行启动看错误：
```bash
cd /d "C:\Program Files\GoodDocument"
DocFormatter.exe
```
看具体异常信息。常见原因：缺 `config.json`、缺 `assets` 目录、杀毒拦截。

### Q2: 规范化后表格里的字还有首行缩进 2 字符
A: 这是 Word 段落对话框里看"特殊格式 → 首行缩进 2 字符"的问题。**0.1.0+ 已修复**（清掉 OOXML 字符单位属性 + Normal 样式 firstLineChars）。如果还出现：
1. 检查 exe 是不是新版本（看打包时间）
2. 在 Word 里把光标放在那个段落 → 右键 → 段落 → "特殊格式"应该是"(无)"、"首行"是"0 字符"
3. 如果还是 2 字符，截图发我（可能是个新版 Word 的渲染问题）

### Q3: 中文显示乱码
A: 两种情况：
- **Word 打开 docx 乱码** → 检查 `config.json` 的 `body.font` 是不是常用中文字体（宋体/黑体/楷体）
- **CLI 输出乱码** → 在 PowerShell 里先 `chcp 65001` 切到 UTF-8

### Q4: 报告"找不到文件"但文件明明在
A: 检查路径：
- Windows 路径用 `\` 还是 `/`？两种都支持
- 路径含中文/空格？应该没问题（程序用 `os.fspath` 标准化）
- 是不是拖拽文件到 PowerShell 时漏了引号？用 `& "path\with space.docx"` 形式

### Q5: 报告"docx 文件已损坏"
A: docx 内部 XML 损坏，工具无法继续。解决：
1. 用 Word 打开原文件 → 另存为 → 关闭
2. 或者用 `unzip -l 文件.docx` 看 zip 列表，排查哪个 part 损坏

### Q6: 报告"输出与输入相同"
A: 工具拒绝覆盖原文件（防误操作）。选一个不同的输出文件名。

### Q7: 想恢复规范化前的版本
A: 同目录有 `<原文件名>.docx.bak`，把它改回原名即可。

### Q8: GUI 启动后界面错位
A: 显示器缩放不是 100% 时会错位。右键 exe → 属性 → 兼容性 → 改高 DPI 设置 → "替代高 DPI 缩放行为"。

### Q9: 想让表格在文档中居中
A: 表格默认居中（`<w:jc w:val="center"/>`）。如果不对，可能是 Normal 样式被改动 — 用"重置默认"按钮恢复。

### Q10: 参考文献格式不对（不是 GB/T 7714）
A: GB/T 7714 是一组规范，工具只负责悬挂缩进。文献条目本身的标点格式需要按规范排。常见类型标识：`[M]` 专著 / `[J]` 期刊 / `[D]` 学位论文 / `[J/OL]` 网络期刊 / `[EB/OL]` 网络资源。

---

## 健壮性保证

工具对异常场景做了友好化处理，所有错误信息均为中文：

| 异常 | 触发场景 | 错误信息示例 |
|---|---|---|
| `InputNotFoundError` | 输入文件不存在 | `找不到文件: X` + 建议"检查路径" |
| `InvalidFileTypeError` | 不是 .docx | `不是有效的 docx 文件: X` |
| `CorruptDocxError` | docx 内部损坏 | `docx 文件已损坏: X` + 建议"用 Word 重新保存" |
| `OutputNotWritableError` | 输出路径不可写 | `无法写入输出: X` + 建议"检查权限/磁盘" |
| `SameInputOutputError` | 输出=输入 | `输出与输入相同: X` |

**自动备份**：规范化前自动在同目录创建 `<原文件名>.docx.bak`（已有 .bak 时用 .bak.1, .bak.2, ...）。失败不阻塞规范化。

**dry-run 模式**：规范化前先预览，**不修改任何文件**：
- Python 库：`normalize_docx(..., dry_run=True, return_result=True)` → 返回 `NormalizeResult`
- GUI：点"预览变更"按钮
- CLI：`node generate.js --dry-run`

**优雅降级**：遇到无法识别的段落、损坏的表格单元格、读取失败的图片等**非致命问题**时，记录为警告后继续执行，**不中断**整体规范化。

---

## 从源码构建

```bash
# Python 版打包（生成 dist\DocFormatter.exe）
pip install pyinstaller
pyinstaller DocFormatter.spec --noconfirm --clean

# Node.js 依赖
npm install
```

或用命令行参数：

```bash
pyinstaller --onefile --windowed --name DocFormatter ^
    --add-data "formatter.py;." ^
    --add-data "normalizer.py;." ^
    --add-data "config_model.py;." ^
    --add-data "paths.py;." ^
    --add-data "config.json;." ^
    --add-data "assets;assets" ^
    app.py
```

生成的 `dist/DocFormatter.exe` 可独立运行，不需要 Python 环境。

---

## 项目结构

```
goodDocument/
├── app.py              # tkinter GUI 入口
├── formatter.py        # 核心排版引擎（生成文档）
├── normalizer.py       # 已有文档规范化处理（含健壮性异常层）
├── config_model.py     # 配置加载/保存/校验
├── paths.py            # 路径工具（支持 PyInstaller 打包模式）
├── config.json         # 排版参数配置
├── DocFormatter.spec   # PyInstaller 打包配置
├── generate.js         # Node.js CLI 入口
├── formatter.js        # Node.js 排版核心库
├── package.json        # Node.js 项目配置
├── assets/             # 示例图片资源
│   └── figure1.png
├── tests/              # 单元测试
│   ├── test_app_paths.py
│   ├── test_config_model.py
│   ├── test_normalize_docx.py
│   ├── _helpers.py
│   └── _corrupt_source.docx
└── docs/
    └── superpowers/
        ├── specs/      # 设计文档
        └── plans/      # 实施计划
```

---

## 运行测试

```bash
python -m pytest tests/ -v
```

**42 个测试**全部覆盖：核心规范化、配置加载、异常抛出、备份、dry-run、边界场景（损坏 docx、嵌套表、空表、无法识别段落等）。

---

## 技术栈

- **Python 版** — Python 3.12 / python-docx / tkinter / PyInstaller
- **Node.js 版** — Node.js / docx / image-size
- **排版规范参照** — GB/T 7714 参考文献著录规则、党政机关公文格式

---

## License

ISC
