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

- **排版预设一键切换** — GUI 主卡片「学校模板」下拉一键载入整套学校/场景规格（首个内置「成都理工大学学士论文」），载入后仍可在高级设置微调；模板放 `config/templates/`，复用 config.json schema
- **一键规范化** — 打开任意 `.docx`，自动识别标题层级（一、/（一）/1./（1）与 第X章/X.1/X.1.1/X.1.1.1，以及结论/致谢/参考文献/附录/摘要等未编号章）、图表题注、参考文献段落，统一套用配置的格式系统
- **表格智能对齐** — `cellAlign: smart` 自动判定表头形态（横/纵/横+纵/无表头）并逐格对齐：表头居中、行名列左、数字列整列右对齐（上下对齐成竖线）、日期/状态居中、短文本居中、长文本左对齐；每个表格报告判定与决策，dry-run 同样可见
- **从零生成** — 生成包含正文、图片、表格、GB/T 7714 参考文献的完整示例文档
- **配置驱动** — 字体、字号、行距、边距、段落排列、段前段后等全部参数集中在 `config.json` 中
- **表格独立排版** — 单元格默认五号宋体（不回退到正文字号），段落级与样式级首行缩进双重清零（OOXML 字符单位也清零）；自动清除彩色底纹、彩色边框恢复黑色；合并单元格同样规范化
- **图片智能缩放** — 两种模式：`fit` 等比缩放到上限内（默认），`fixed` 把全部图片统一到指定尺寸；可只填宽度实现「统一宽度」，也可强制拉伸
- **浮动图片同样处理** — `wp:anchor`（四周型 / 浮于文字上方）也在处理范围内，不只是嵌入型
- **图片安全阀** — 目标宽度自动收窄到正文可用宽度、表格内图标默认跳过，避免图片溢出页面或被拉爆
- **图与图题不分页** — 自动设置 `keepNext`，图题不会被分页甩到下一页
- **页码样式统一** — 已有页码统一为底部居中并套用配置的字体字号；无页码的封面/前置页保持原样（可用 `pageNumber.onlyUpdateExisting` 改为给所有分节新增页码）；可选 `header` 区段把已有页眉统一为宋体小五居中（从正文开始，不创建不改文字）
- **幂等可重入** — 同一文件连续规范化两次，第二次「变更」计数为 0
- **桌面界面** — 支持把 `.docx` 拖入窗口；选中文档后显示段落/表格/图片摘要；结果按「改动 / 提醒 / 记录」分段展示；耗时操作在后台执行并显示进度。「高级设置」按使用目的分组（图标 + 名称 + 当前生效值 + 尾部控件），顶部过滤框输入即筛，每组可单独恢复默认，与学校模板一致的行标「按学校模板」、改过的标「已修改」
- **健壮性优先** — 中文错误信息、自动备份、dry-run 预览、优雅降级（详见 [健壮性保证](#健壮性保证)）
- **双语言实现** — Python（tkinter GUI + python-docx）和 Node.js（CLI + docx 库），共享同一份配置；表格智能对齐仅 Python 规范化端实现

---

## 快速开始

### 方式一：直接运行可执行文件（Windows，推荐）

从 [Releases](../../releases) 下载 `DocFormatter.exe`，双击即可启动 GUI。将 `config.json` 放在 exe 同目录即可自定义排版参数。把 `.docx` 拖入窗口或点击选择，点「开始整理」即可；「预览改动」不会写入文件。

GUI 快捷键：`Ctrl+O` 选择文档、`Ctrl+P` 预览改动、`Ctrl+Enter` 开始整理、`Ctrl+S` 保存设置、`Ctrl+G` 生成示例文档。

### 方式二：Python 运行

```bash
pip install -r requirements.txt
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
2. 「排版预设」下拉选择学校规格（如「成都理工大学学士论文」，表格自动启用智能对齐），或保持「自定义」用当前 config.json
3. 把 `.docx` 拖入窗口（或点「选择文档…」）
4. 可先点「预览改动」确认会改什么（含每个表格的对齐决策），再点「开始整理」

**等价 CLI**：
```bash
node generate.js --out 我的论文.docx
```

### 场景 2：工程师技术报告（紧凑型）

**目标**：表题在上方、行距紧凑、无首行缩进。

**操作**：打开右侧「高级设置」→「表格」组的「表题位置」选「上方」；「正文」组「行距」改 1、「段首缩进」改 0；想撤销就点对应组的「恢复默认」。

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

### 场景 5：把文档里所有照片统一成同样大小

**目标**：一批尺寸不一的照片（截图、手机拍照）统一为宽 12cm，高度各自等比；如果要求每张都严格 12×9 cm（允许多少变形）改成强制拉伸。

**操作**（GUI）：右侧「高级设置」→「图片」组 →
- 「图片大小方式」选「统一尺寸」
- 「统一宽度」填 12（高度不动即只约束宽度）
- 「统一时保持比例」打开

**等价配置**：

```json
"figure": {
  "sizeMode": "fixed",
  "fixedWidthCm": 12,
  "fixedHeightCm": null,
  "keepAspectRatio": true
}
```

强制拉伸成 12×9（等同 Word VBA 的 `LockAspectRatio = msoFalse`）：

```json
"figure": { "sizeMode": "fixed", "fixedWidthCm": 12, "fixedHeightCm": 9, "keepAspectRatio": false }
```

### 更多场景

`node generate.js --examples` 查看全部 6 个场景。

---

## 配置说明

所有排版参数集中在 `config.json` 中，结构如下：

| 区段 | 说明 | 关键参数 |
|------|------|----------|
| `page` | 页面设置 | `margins`（上下左右边距，cm；可选 `header` / `footer` 页眉页脚距边）、`size`（A4 / Letter） |
| `pageNumber` | 页码 | `enabled`、`align`、`size`、`font`、`onlyUpdateExisting`（只更新已有页码，默认 true） |
| `header` | 页眉（可选） | `enabled`、`font`、`size`、`align`；只格式化已有页眉文字，不创建不修改内容，封面/前置节跳过 |
| `body` | 正文 | `font`（中文字体）、`asciiFont`（英数字体）、`size`、`lineSpacing`、`firstLineIndentChars`（首行缩进字数）、`align` |
| `headings` | 标题（H1-H4） | 每级独立设置 `font`、`size`、`bold`、`color`、`spaceBeforePt`、`spaceAfterPt`、`spaceBeforeLines` / `spaceAfterLines`（段前/段后「行」单位，0.5 → 半行）、`align` |
| `figure` | 图片 | `sizeMode`（`fit` 等比 / `fixed` 统一）、`maxWidthCm`、`maxHeightCm`、`fixedWidthCm`、`fixedHeightCm`、`keepAspectRatio`、`clampToTextWidth`、`skipInTables`、`keepWithCaption`、`captionPosition`（above/below）、`captionSize`、`prefix` |
| `table` | 表格 | `captionPosition`、`align`、`cellSize`、`cellAlign`（`left` / `center` / `right` / `justify` / **`smart` 智能对齐**）、`smartAlignThresholdChars`（smart 模式短文本阈值，默认 6，中文计 1 半角计 0.5）、`cellVAlign`、`headerBold`、`prefix`、`borderColor`（清除彩色边框后的恢复色） |
| `references` | 参考文献 | `style`（GB/T 7714）、`hangingIndentPt`（悬挂缩进）、`size`、`lineSpacing` |
| `output` | 输出 | `filename` |

默认配置参照公文格式：正文宋体小四 1.5 倍行距首行缩进 2 字，一级标题黑体小二，图表题注五号加粗居中，参考文献小五悬挂缩进。

### 排版预设（config/templates/）

预设是「填好参数的 config.json」，可加 `_preset.name` / `_preset.description` 供 GUI 展示；放入 `config/templates/*.json` 即被自动发现（`config_model.list_presets()`，校验失败的模板自动跳过）。GUI 顶部下拉选择即载入，「更多 → 保存设置」可把当前参数（含预设微调结果）写回 `config.json`。内置：

- **成都理工大学学士论文** — 依据学校《学生学士学位论文（设计）编写格式》：A4、边距上下左右 2.54cm、页眉 1.5cm、页脚 1.75cm；正文宋体小四 1.5 倍行距首行缩进 2 字；页码小五底端居中；章（第X章/结论/致谢/参考文献）小二黑体居中段前 0.5 行、节 X.1 小三黑体、条 X.1.1 小四黑体段前后 0.5 行；表题上方五号黑体、表内容五号宋体智能对齐；图题下方五号黑体；参考文献五号宋体 GB/T 7714 悬挂缩进

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
A: 0.2.1+ **默认不生成 .bak 备份**（之前版本会生成）。如果规范化时显式传了 `backup=True`，同目录会有 `<原文件名>.docx.bak`，把它改回原名即可。

### Q8: GUI 启动后界面模糊或错位
A: 0.4.0+ 已在启动时声明 Windows DPI 感知，正常显示。如果仍异常，右键 exe → 属性 → 兼容性 → 改高 DPI 设置 → "替代高 DPI 缩放行为"。

### Q9: 想让表格在文档中居中
A: 表格默认居中。如果不对，可能是 Normal 样式被改动 — 用「更多 → 恢复默认配置」恢复。

### Q10: 参考文献格式不对（不是 GB/T 7714）
A: GB/T 7714 是一组规范，工具只负责悬挂缩进。文献条目本身的标点格式需要按规范排。常见类型标识：`[M]` 专著 / `[J]` 期刊 / `[D]` 学位论文 / `[J/OL]` 网络期刊 / `[EB/OL]` 网络资源。

### Q11: 设了「统一尺寸」，但有的图片没变
A: 按以下顺序排查：
1. **在表格里** — `figure.skipInTables` 默认为 `true`，表格单元格内的图标会跳过（单元格远窄于正文，强制 12cm 会被拉爆）。「变更」列表里会写明「图 N：位于表格内，跳过尺寸统一」。确实要一起改就把 `skipInTables` 设为 `false`
2. **在文本框里** — 文本框内的小图标一律跳过（会写明原因）
3. **在页眉/页脚里** — 这些部件不属于正文（body），本工具不处理，会给出警告提示
4. **本来就是目标尺寸** — 已经等于目标大小，无需改动，也不计入「调整 N 张图」

### Q12: 统一宽度设了 20cm，结果变成了 15.5cm
A: 这是 `clampToTextWidth`（默认开）在起作用 —— 目标宽度超过正文可用宽度时会自动收窄，否则图片会溢出页面。正文可用宽度 = 页宽 − 左右边距，A4 加 3 / 2.5cm 边距时为 15.5cm。确实想要更宽就调小页边距，或关闭 `clampToTextWidth`。

### Q13: 想让图片严格 12×9 cm，变形不变形无所谓
A: 用 `sizeMode: "fixed"` + `keepAspectRatio: false`，这就是 Word VBA 里解开长宽比锁定后写死宽高的效果。若希望不变形，把 `keepAspectRatio` 设为 `true`（按框等比适配）。

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

**自动备份（默认关闭）**：早期版本默认开启，会无限累计 `.bak` / `.bak.1` / ... 污染工作目录。**0.2.1+ 默认不备份**。需要时显式开启：
- Python 库：`normalize_docx(..., backup=True)`
- GUI / CLI：暂无开关，需直接传参（如有需求可加）

备份开启时策略：同目录创建 `<原文件名>.docx.bak`（已有时用 .bak.1, .bak.2, ...）。备份失败不阻塞规范化。

**dry-run 模式**：规范化前先预览，**不修改任何文件**：
- Python 库：`normalize_docx(..., dry_run=True, return_result=True)` → 返回 `NormalizeResult`
- GUI：点"预览改动"按钮
- CLI：`node generate.js --dry-run`

**优雅降级**：图片尺寸读取失败等**非致命问题**记录为警告（进入 `NormalizeResult.warnings`，同时回调 `on_warning`）后继续执行，**不中断**整体规范化。

---

## 从源码构建

```bash
# Python 依赖与打包工具
pip install -r requirements-dev.txt

# 打包（生成 dist\DocFormatter.exe）
pyinstaller DocFormatter.spec --noconfirm --clean

# Node.js 依赖
npm ci
```

生成的 `dist/DocFormatter.exe` 可独立运行，不需要 Python 环境。

---

## 项目结构

```
goodDocument/
├── app.py               # tkinter GUI 入口（拖拽、设置面板、后台执行、快捷键）
├── settings_model.py    # 高级设置目录（分组/行定义）与配置值转换
├── ui_theme.py          # 界面主题（颜色/字体/结构令牌、ttk 样式、取值映射）
├── ui_icons.py          # Segoe 图标字体系统（码点表、运行时检测、纯文字回退）
├── formatter.py         # 核心排版引擎（生成文档）
├── font_sizes.py        # 中文字号名与磅值映射（GUI / 配置校验 / 排版共用）
├── normalizer/          # 已有文档规范化（core / normalize / errors / paths）
├── config_model.py      # 配置加载/保存/校验
├── paths.py             # 路径工具（支持 PyInstaller 打包模式）
├── config.json          # 排版参数配置
├── DocFormatter.spec    # PyInstaller 打包配置
├── generate.js          # Node.js CLI 入口
├── formatter.js         # Node.js 排版核心库
├── package.json         # Node.js 项目配置
├── requirements.txt     # Python 运行依赖
├── requirements-dev.txt # Python 开发依赖（pytest / pyinstaller / ruff）
├── assets/              # 示例图片资源
│   └── figure1.png
├── tests/               # 测试
│   ├── test_app_paths.py
│   ├── test_config_model.py
│   ├── test_formatter.py
│   ├── test_normalize_docx.py
│   ├── test_ui_theme.py
│   ├── test_version_consistency.py
│   ├── _helpers.py
│   └── node/            # Node.js 测试（node --test）
└── docs/
    └── superpowers/
        ├── specs/       # 设计文档
        └── plans/       # 实施计划
```

---

## 运行测试

```bash
pip install -r requirements-dev.txt
python -m pytest tests/ -v   # Python 测试
npm test                     # Node.js 测试
ruff check .                 # 静态检查
```

覆盖核心规范化、配置加载、生成器、异常抛出、备份、dry-run、边界场景（损坏 docx、嵌套表、空表、无法识别段落等）、变更收集、界面主题令牌与版本一致性。

---

## 使用的技术

- **Python 版** — Python 3.12 / python-docx / tkinter + tkinterdnd2 / PyInstaller
- **Node.js 版** — Node.js / docx / image-size
- **排版规范参照** — GB/T 7714 参考文献著录规则、党政机关公文格式

---

## License

ISC
