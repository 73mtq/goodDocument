# Changelog

所有重要变更记录在此文件。格式基于 [Keep a Changelog](https://keepachangelog.com/)。

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
