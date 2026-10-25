# Changelog

所有重要变更记录在此文件。格式基于 [Keep a Changelog](https://keepachangelog.com/)。

## [0.8.1] - 2026-10-25

图标偏移修复：对齐从「目测微调」升级为「测量驱动的系统属性」。取证与决策：`design/icon-metrics/DECISION.md`（根因三层 + 路线定案 + 量具工程记录）。

### Fixed
- **图标偏移根因修复**（三层）：① 居中模型错误——tk 按字体行盒居中，而 Segoe 图标字体行盒=1em、基线压在行盒底、字形墨迹全部向上长，矮字形（check 0.53em）沉底、满格字形占满，偏移量逐字形不同；② 尺寸单位错误——图标 token（设计像素）被当作 points 传入字体，Windows 上恒放大 4/3（150% 缩放下 16 设计px 渲染成 32px 行盒）；③ 两套回退字体逐字形墨迹差异大（如 font：Fluent 1.0em vs MDL2 0.75em 高），单一常数补偿在数学上不可能
- 全部 9 处图标调用点（含拖放区主图）收编到唯一图标槽 `icon_slot.py`：固定像素格 + 按字体单位墨迹盒光学居中 + grid-fit 稀疏校准表（5 条，全为 warning 字形），调用方只许摆放格子、禁止槽内微调

### Added
- **测量链**（design/icon-metrics/，dev 依赖，运行时零依赖）：字体表与逐字形墨迹盒（fontTools）、GDI 落点探针、全覆盖校准器（22 图标 × 9 档 ppem × 2 字体 = 396 格）、字样版 specimen.png 与偏差报告
- `icon_slot_data.py`：构建期生成的字体单位墨迹表 + 校准表（自动生成勿手改）
- `tests/test_icon_slot.py`：数据形状/摆位公式/渲染墨迹不越格/数据新鲜度（与字体文件比对）/PUA 纪律
- `tests/test_icon_alignment.py`：屏上验收关——经真实渲染管线逐格断言分量公差（≤0.5px@100% → ≤1px@150% 线性），产出验收报告；本机实测 198/198 达标、136 格残差恰为 0

## [0.8.0] - 2026-10-25

v0.7.0 被否决后的结构层重构：换配色与文案不够，这次动信息架构、组件结构与图标系统。
取证与决策记录：`design/ui-audit.md`（现状审计）、`design/ia-regroup.md`（卡片分类 + 状态矩阵）、`design/measurements.md`（macOS 系统设置结构量测）、`design/acceptance-v0.8/`（四道关卡材料）。

### Changed
- **高级设置面板重做**（标本 macOS 系统设置，量结构不量颜色）：旧分组（页面/正文/标题/图片/表格/参考文献）是 config.json 节名照搬，按用户目的卡片分类重排为 11 组（页面/页码/正文/章标题（第1章）/节标题（1.1）/条标题（1.1.1）/四级标题（1.1.1.1）/图片/图题与说明/表格/参考文献）；每行 = 图标 + 设置名 + 当前生效值 + 尾部控件，行高固定 44px、列缘对齐、1px 分隔线规律节奏（量测换算见 `design/measurements.md`，落 `ui_theme.MEASUREMENTS` 令牌）
- **默认值透明**：选了学校模板后，与模板一致的行标「按学校模板」，改过的标「已修改」；每组组头提供「恢复默认」；「示例文档」不是设置，移出面板（菜单保留入口）
- **面板顶部过滤框**：输入即筛字段（匹配组名/行名/当前值），空结果显示「没有匹配的设置」
- **滚轮改逐控件局部绑定**，删除 `bind_all` 全局劫持；下拉全部 readonly（不再接受自由输入的非法字体/字号），数字行输入校验（只放行数字与一个小数点）+ 失焦钳制范围
- **排版术语清洗**：行距(倍)→行距（值列「1.5 倍」）、首行缩进(字)→段首缩进（值列「2 字」）、段后(pt)→段后间距（值列「0 磅」）、悬挂缩进(pt)→条目缩进（组说明一句话解释）、字号值列补磅值对照（「小四 · 12 磅」）；cm/pt 等单位缩写不再出现在行名
- **主流程层级强化**：三步（选文档/学校模板/保存位置）改为步骤行（图标 + 加粗步骤名固定列宽 + 控件右对齐 + 一句话说明换行行下）；拖放区弃虚线矩形字符画，改为图标字体大文档图形 + 两行文案，悬停/拖入/聚焦各有明确反馈（拖入变蓝底）；拖放区可键盘聚焦（Tab + Enter/空格打开文件）
- **强调蓝收敛**：只出现在「开始整理」与聚焦/拖入态；分段按钮选中态改中性灰加粗，菜单高亮与文本选区弃用蓝色系
- 状态栏分语气：成功绿、失败红（新 Danger.TLabel）、日常灰；主按钮禁用态（忙碌）图标+文字+底色整体变灰蓝

### Added
- **Segoe 图标系统**（`ui_icons.py`）：全部图标来自微软官方字形表（Segoe Fluent Icons 与 Segoe MDL2 Assets 共有的 E7xx-E9xx 码点，Win11/Win10 一套通吃）；运行时检测字体可用性，缺失即纯文字回退；不捆绑字体文件（系统字体授权跟随 Windows）。22 个图标覆盖 11 个设置组、主流程步骤、结果区分段、过滤框与状态动作
- `tests/test_ui_icons.py`：码点私用区合法性、去重、字体检测优先级与回退
- `tests/test_settings_panel.py`：11 组 85 行覆盖旧面板全部字段（含标题四级）、组名/行名零英文术语、单位全是人话、人话取值格式化、默认值对照比较
- `tests/test_ui_lint.py` 扩展：字符图形禁令（几何/箭头/emoji 区段全部禁用）、私用区字形只准出现在 `ui_icons.py`、`.bind_all` 与 `dash=` 源码禁令、`height`/`wraplength` 关键字实参必须来自令牌表、结构令牌同样落 4px 网格
- `tests/test_ui_theme.py` 对比度断言扩到 18 组（+悬停/按下底、+accent on white、+danger on bg）

### Removed
- 面板分组折叠（▸/▾ 文本三角）与全部字符图形；「示例文档」设置组（移菜单）

## [0.7.0] - 2026-10-25

### Changed
- **界面重做为 Apple 系统色**（在 Windows 上实现，不仿 macOS 窗口部件）：浅灰分组底 `#F7F7F7`、白色卡片 `#FFFFFF`、近黑文字 `#1F1F1F`，全界面单一强调蓝 `#2F6EED`（macOS「系统设置」官方截图实测色；HIG 文档蓝 `#007AFF` 因与白字对比度仅 4.02 不达标而弃用，量色记录见 `design/apple-measurements.md`）
- **主流程压缩为一步**：选论文、选学校模板、开始整理收进同一张白卡，删除分节编号；「高级设置」保持默认折叠，面板宽度与字号阶梯适配
- **全部界面文案重写**（先过 humanizer-zh 去 AI 痕迹，再按展签规范）：陈述句、说结果不说机制、每条帮助一句话、按钮动词开头、错误信息写下一步动作；主流程用「整理」代替「规范化」，藏起 config.json、VBA、H1-H4 等词汇；结果页签改名「改动 / 提醒 / 记录」
- 字号阶梯改用 macOS 官方文字样式映射（窗口标题 Title2 17pt、正文 Body 13pt、说明 Subheadline 11pt），弃用 Light 字重；默认窗口 1120×860（最小 980×720）
- 移除 F1 快捷键（与用户截图软件冲突），「关于文档整理」入口保留在「更多」菜单
- README 同步新界面文案

### Added
- `tests/test_ui_lint.py`：界面机械 lint（零 emoji / 零感叹号 / 零长破折号 / 零箭头；间距全在令牌表且全落 4px 网格）
- `tests/test_ui_theme.py` 对比度断言扩到界面实际用到的全部文字/背景组合（14 组，全部 ≥ 4.5:1）
- `design/` 取证目录：改造前全部界面状态截图（`audit/`，已 gitignore）、Apple 一手量色与规格记录（`apple-measurements.md`）、改造前文案清单（`copy-inventory.md`）

## [0.6.0] - 2026-10-25

### Added
- **排版预设**：GUI 三步流程顶部新增「排版预设」下拉，一键载入整套排版参数；模板放 `config/templates/*.json`（复用 config.json schema，经 `config_model` 校验），首个内置预设「成都理工大学学士论文」（A4、边距 2.54/页眉 1.5/页脚 1.75cm、正文宋体小四 1.5 倍行距、章小二黑体居中段前 0.5 行、表题上方五号黑体、参考文献五号宋体 GB/T 7714）
- **表格智能对齐**：`table.cellAlign` 新增 `smart`（默认值不变，向后兼容）——自动检测表头形态（横 / 纵 / 横+纵 / 无表头，四形态组合判定），按内容类型逐格对齐：表头居中 → 行名列左 → 数字列按列投票整列右对齐（千分位/百分比/带单位金额，「—」占位跟随列）→ 日期与状态居中 → 短文本（≤ `table.smartAlignThresholdChars`，默认 6，中文计 1 半角计 0.5）居中 → 长文本左对齐；合并单元格按自身文本归类；每个表格在「变更」与 dry-run 中报告判定形态与对齐决策（如「表格 2：横+纵表头；首列左对齐、数字 2 列右对齐」）；GUI「表」分组「单元格排列」下拉新增「智能识别」
- 标题识别扩展：新增「第X章 / X.1 / X.1.1 / X.1.1.1」与未编号章级标题（结论 / 致谢 / 参考文献 / 附录 / 摘要，允许「摘  要」内部空格）两套体系，原有「一、/（一）/1./（1）」保持不变；正文小数（如「1.1 是小数」「1.5 倍行距是最常见的选择。」）不误判
- `page.margins.header` / `page.margins.footer`：页眉/页脚距边（cm）；可选 `header` 配置区段：格式化已有页眉（宋体小五居中，从正文开始，不创建不改文字）
- `headings.hN.spaceBeforeLines` / `spaceAfterLines`：段前/段后「行」单位（0.5 → 半行），写入 `w:beforeLines`/`w:afterLines` 并以磅值兜底
- 规范化幂等：同一文件连续规范化两次，第二次「变更」计数为 0（变更改为按 XML 实际差异上报；对齐不改变文本，天然可重入）

### Changed
- **表格完全独立于正文**：单元格字体默认五号宋体（不再回退到 body 的字体字号）；首行缩进在段落级与样式级双重清零（单元格引用到的样式及其 basedOn 链一并清零，根治 HANDOVER 坑 1）；主动清除表格底纹与彩色边框（单元格显式无填充压过表格样式底色，已有边框颜色统一恢复为 `table.borderColor`）；`headerBold` 在智能模式下只加在判定出的表头行上
- `set_zero_indent` / 正文缩进改直接写 `w:ind` 属性：python-docx 的 first_line_indent setter 会删重建属性导致两次规范化间属性顺序往返翻转，破坏幂等
- 合并单元格（跨列与纵向延续格）现在也会被规范化——此前经 `row.cells` 迭代会跳过延续格、留下原格式
- `borderColor` 由「仅 Node.js 生效」改为规范化时同样生效（清除彩色边框的恢复色）

## [0.5.0] - 2026-09-22

### Added
- 图片「统一尺寸」模式：`figure.sizeMode` 支持 `fit`（等比缩放到上限内，默认）与 `fixed`（全部统一到指定尺寸），对应 Word VBA 里 `LockAspectRatio = msoFalse` 后写死 `Width` / `Height` 的做法
- `figure.fixedWidthCm` / `figure.fixedHeightCm` / `figure.keepAspectRatio`：`fixed` 模式下可强制拉伸，也可只填宽度实现「统一宽度、高度各自等比」
- 支持**浮动图片**（`wp:anchor`，四周型 / 浮于文字上方）：此前 `doc.inline_shapes` 只覆盖嵌入型，浮动图片完全未被处理
- `figure.clampToTextWidth`（默认开）：目标宽度自动收窄到正文可用宽度，避免写死厘米数导致图片溢出页面
- `figure.skipInTables`（默认开）：表格单元格内的图标默认跳过统一尺寸，避免被拉爆
- `figure.keepWithCaption`（默认开）：图与图题设置 `w:keepNext`，不再被分页拆到两页
- 页眉/页脚中存在图片时给出提示（这些部件不在 body 内，本工具不处理）
- GUI「图」分组新增尺寸模式、统一宽/高、保持比例、宽度不超出正文、跳过表格内图片、图题跟随图片等控件
- `tests/_helpers.py`：`make_docx_with_floating_image`（把 `wp:inline` 改写成 `wp:anchor`）、`make_docx_with_images`、`picture_sizes`
- Python 新增 19 个、Node.js 新增 5 个图片尺寸相关测试
- `config_model` 校验 `figure.sizeMode` 取值，非法值给出中文错误

### Changed
- `NormalizeResult.images_processed` 语义修正为**实际被调整**的图片数量（此前等于文档中图片总数）
- `inspect_docx().image_count` 同时统计浮动图片（此前只统计嵌入型）
- 图片尺寸写入同时覆盖 `wp:extent` 与 `pic:spPr/a:xfrm/a:ext`，保证 Word / WPS / LibreOffice 显示一致
- Node.js 版 `computeImageTransform` 与 Python 版对齐：支持 `sizeMode`、比例控制与宽度钳制

## [0.4.0] - 2026-09-19

### Added
- 界面重做（朱批纸面方向）：纸色打底、墨色文字、朱砂只出现在主操作上；节标题带朱砂竖条
- 新增 `ui_theme.py`：颜色与字体令牌、`clam` 主题下的完整 `ttk` 样式（按钮三档、输入、下拉、数字框、复选框、分段、滚动条、进度条）
- 文件拖拽：把 `.docx` 拖入窗口即可选择（`tkinterdnd2`）
- 选中文档后显示摘要（段落 / 表格 / 图片 / 标题层级 / 页眉页脚）
- 结果区改为「变更 / 警告 / 日志」分段展示；完成后状态栏内联「打开文件夹」，不再弹窗询问
- 耗时操作在后台线程执行，界面保持响应并显示进度条
- 键盘快捷键：Ctrl+O 选择文档、Ctrl+P 预览变更、Ctrl+Enter 开始规范化、Ctrl+S 保存配置、Ctrl+G 生成示例文档、F1 关于
- Windows 高 DPI 感知声明，缩放不为 100% 时不再模糊
- 高级设置改为右侧可滚动面板，7 个分组可折叠
- 下拉选项值中文化（center → 居中、above → 上方、top → 顶部）
- `tests/test_ui_theme.py`：令牌完整性、文字对比度、取值映射

### Changed
- 主操作唯一突出，「生成示例文档 / 保存配置 / 恢复默认 / 关于」收进「更多」菜单
- 术语区分：「规范化后另存为」与「示例文档保存位置」
- 错误提示不再直接抛 traceback，完整详情写入「日志」分段
- `requirements.txt` 增加 `tkinterdnd2`

## [0.3.0] - 2026-09-19

### Fixed
- Node.js 版正文首行缩进按 twips 计算（2 字小四输出 `w:firstLine="480"`，此前为 48）
- Node.js 版 `detectHeadingLevel` 自定义规则失效（`lv` 未定义导致异常被吞）
- Node.js 版表格单元格段前段后与 Python 版统一为 0
- 题注识别要求"前缀 + 编号"，正文段"图片…""表题…"不再被当作题注
- 标题识别不再把"1.5 倍行距…"这类段首小数判为三级标题
- 正文提及"参考文献"不再切换参考文献模式，改为标题段或整段为"参考文献/参考资料"时开启
- `inspect_docx` 的 `has_header_footer` 按真实内容判断（此前对任意文档恒为 True）
- 页脚已有页码时改写所在段落，不再清空首段或产生两个 PAGE 域
- `normalizer/paths.py` 先校验文件存在，再校验扩展名
- `normalize_docx(dry_run=True)` 未请求结果时返回 None

### Added
- `NormalizeResult.changes` 变更列表（标题、题注、参考文献、表格、图片缩放）
- 图片尺寸读取失败的降级警告；`on_warning` 回调与 `NormalizeResult.warnings` 同时生效
- 处理计数只统计实际处理的段落/表格/图片，跳过封面与目录
- Python 版支持 `page.size`（A4 / Letter）
- `tests/node/` Node.js 测试与 `npm test`
- `tests/test_formatter.py`、`tests/test_version_consistency.py`
- `requirements.txt` / `requirements-dev.txt` / `pyproject.toml`（ruff 与 pytest 配置）
- GUI 预览变更按钮现在显示变更列表

### Changed
- Python 版表格整体位置读取 `table.align`（此前固定居中）
- CI 拆分为 test（push / PR / 手工触发）与 build（v* 标签）两个任务，新增 ruff 与 Node 测试
- `config.json` 删除无效配置键：`page.forceA4`、`pageNumber.position`、`figure.borderless`、`table.fullWidth`、`page.marginUnit`
- 删除根目录 `normalizer.py` 兼容层（同名包始终优先，该文件不会生效）
- 测试样例文档改为运行时生成，不再提交二进制样本

## [0.2.1] - 2026-07-02

### Fixed
- `backup` 默认值从 `True` 改为 `False`，避免 `.bak` 文件无限累计

## [0.2.0] - 2026-06-30

### Added
- 5 种 `NormalizeError` 子类（中文错误信息 + 修复建议）
- `NormalizeResult` 数据类、`WarningCollector` 上下文管理器
- `normalize_docx` 新签名：`backup` / `dry_run` / `return_result` / `on_warning`
- GUI "预览变更" 按钮 + ScrolledText 状态区
- CLI `--dry-run` 与 `--examples`
- README 大改版

### Changed
- 旧 `ValueError` / `FileNotFoundError` 替换为 `NormalizeError` 子类
- `normalizer.py` 整理为 `normalizer/` 包

### Fixed
- 表格文字继承正文首行缩进 2 字符（OOXML 字符单位 `firstLineChars` 残留）

## [0.1.0] - 2026-06-29

### Added
- 初始版本
- Python GUI（tkinter）+ Node.js CLI
- 配置驱动排版（config.json）
- 表格独立排版、图片智能缩放、页码插入
- 14 个 Python 单元测试

## 版本约定

- 主版本号（X.0.0）：不兼容的 API 变更
- 次版本号（0.X.0）：向后兼容的功能新增
- 修订号（0.0.X）：向后兼容的问题修复
