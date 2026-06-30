# Changelog

所有重要变更记录在此文件。格式基于 [Keep a Changelog](https://keepachangelog.com/)。

## [Unreleased]

### Added
- 子项目 A 健壮性与 UX：5 种 `NormalizeError` 子类（中文错误信息 + 修复建议）
- `NormalizeResult` 数据类（输入/输出/备份路径/计数/changes/warnings）
- `WarningCollector` 上下文管理器
- `_validate_paths` 路径校验辅助
- `_backup_source` 自动备份（`.bak` / `.bak.1` / `.bak.2` 递增）
- `normalize_docx` 新签名：`backup` / `dry_run` / `return_result` / `on_warning`（旧签名 100% 向后兼容）
- GUI "预览变更" 按钮 + ScrolledText 状态区
- GUI 友好错误弹窗（`NormalizeError.message` + hint）
- CLI `--dry-run` flag
- CLI `--examples` flag（6 个典型使用场景示例）
- README 大改版：4 个场景 + FAQ + 故障排查 + 健壮性保证章节

### Changed
- 旧 `ValueError` / `FileNotFoundError` 替换为 `NormalizeError` 子类
- `_apply_run_defaults` 等接受可选 `on_warning` 回调
- `formatter.js` 的 `cellPara` 加 `indent: { left: 0, right: 0, firstLine: 0 }`（Node.js 端也清零缩进）

### Fixed
- 表格里文字会继承正文首行缩进 2 字符（OOXML 字符单位 `firstLineChars` 残留问题）
- 表格里文字会从 Normal 样式继承 `firstLineChars=200`（Normal 样式清理 bug）
- `formatter.js` Node.js 端表格单元格段落无显式 indent

## [0.1.0] - 2026-06-29

### Added
- 初始版本
- Python GUI（tkinter）+ Node.js CLI 双实现
- 配置驱动排版（config.json 9 大区段）
- 表格独立排版（独立字体/字号/行距）
- 图片智能缩放
- 页码自动插入
- 14 个 Python 单元测试

## 版本约定

- 主版本号（X.0.0）：不兼容的 API 变更
- 次版本号（0.X.0）：向后兼容的功能新增
- 修订号（0.0.X）：向后兼容的问题修复
