# GoodDocument

> 配置驱动的 Word 文档规范化排版工具，参照 GB/T 7714 参考文献著录规则与公文格式标准。

选一份已有的 `.docx`，一键统一页面、标题、正文、图表、参考文献和页码格式；也可以从零生成符合规范的示例文档。所有排版参数（字体、字号、行距、边距、对齐）均由 `config.json` 控制，无需改代码即可调整。

---

## 功能特性

**文档规范化** — 打开任意 `.docx`，自动识别标题层级（一、/（一）/1./（1））、图表题注、参考文献段落，统一套用配置的格式系统。

**文档生成** — 从零生成一份包含正文、图片、表格、GB/T 7714 参考文献的完整示例文档，适合作为模板或排版验证。

**配置驱动** — 字体、字号、行距、边距、对齐、段前段后等全部参数集中在 `config.json` 中，修改后立即生效。字号支持中文名（小二/小三/小四/五号）和磅值数字（18/15/12/10.5）两种写法。

**表格独立排版** — 表格单元格使用独立的字体、字号、行距，显式置零首行缩进，杜绝"表格文字继承正文格式"的常见问题。

**图片智能缩放** — 图片默认宽度等于正文宽度，按宽高比等比缩放，超高时以高度上限反向缩放。

**页码自动插入** — 底部居中页码，PAGE 域自动编号。

**双语言实现** — Python 版（tkinter GUI + python-docx）和 Node.js 版（CLI + docx 库），共享同一份配置文件。

---

## 快速开始

### 方式一：直接运行可执行文件（Windows）

从 [Releases](../../releases) 下载 `DocFormatter.exe`，双击即可启动 GUI。将 `config.json` 放在 exe 同目录即可自定义排版参数。

### 方式二：Python 运行

```bash
# 安装依赖
pip install python-docx

# 启动 GUI
python app.py
```

### 方式三：Node.js CLI 运行

```bash
# 安装依赖
npm install

# 用默认配置生成文档
npm start

# 临时修改参数并生成
node generate.js --set body.size=14 --set body.lineSpacing=2.0

# 查看中文字号对照表
node generate.js --list-sizes
```

---

## GUI 使用说明

启动后界面分为三个区域：

**一键规范化区** — 选择输入 `.docx` 文档，指定输出位置，点击"开始规范化"即可。点击"生成示例文档"则从零生成一份完整的排版示例。

**状态区** — 实时显示操作进度和结果。

**高级设置区**（可折叠） — 分为页面、正文、标题、图、表、参考文献、输出七个标签页，每个参数都有对应的输入框或下拉选择。修改后点击"保存配置"即可持久化，下次启动自动加载。

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
| 三小 | 15 | | |
| 四号 | 14 | | |

---

## 从源码构建可执行文件

```bash
# 安装 PyInstaller
pip install pyinstaller

# 用 spec 文件打包（推荐）
pyinstaller DocFormatter.spec --noconfirm

# 或用命令行参数打包
pyinstaller --onefile --windowed --name DocFormatter ^
    --add-data "formatter.py;." ^
    --add-data "normalizer.py;." ^
    --add-data "config_model.py;." ^
    --add-data "paths.py;." ^
    --add-data "config.json;." ^
    --add-data "assets;assets" ^
    app.py
```

生成的 `dist/DocFormatter.exe` 可独立运行，无需安装 Python 环境。

---

## 项目结构

```
goodDocument/
├── app.py              # tkinter GUI 入口
├── formatter.py        # 核心排版引擎（生成文档）
├── normalizer.py       # 已有文档规范化处理
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
│   └── test_normalize_docx.py
└── .gitignore
```

---

## 运行测试

```bash
python -m pytest tests/ -v
```

---

## 技术栈

**Python 版** — Python 3.12 / python-docx / tkinter / PyInstaller

**Node.js 版** — Node.js / docx / image-size

**排版规范参照** — GB/T 7714 参考文献著录规则、党政机关公文格式

---

## License

ISC
