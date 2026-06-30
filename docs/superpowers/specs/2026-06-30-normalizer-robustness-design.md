# 子项目 A：健壮性与 UX — 设计文档

**状态**：待用户 review
**作者**：Sisyphus
**日期**：2026-06-30
**范围**：子项目 A（健壮性与 UX），4 段设计

---

## 1. 目标

在不破坏现有 `normalizer.py` 核心规范化逻辑的前提下，通过**薄包装层**提升工具的健壮性和用户体验，让工具：

- **更易用**：错误信息具体可操作，GUI 有预览，CLI 有 dry-run
- **更易理解**：规范化过程透明（变更摘要、警告列表）
- **更健壮**：边界情况优雅降级，不因个别段落 / 表格损坏而整体失败

明确**不在范围**内（延后到其他子项目）：

- 拆分 formatter.py / normalizer.py（→ 子项目 C）
- 国际化错误信息（i18n）— 只用中文
- GUI 进度条 — 延后到 A1.5
- 损坏 docx 修复（→ 单独工具）
- 用户配置 `.bak` 命名策略 — 固定 `.bak` + `.bak.N`

---

## 2. 架构

```
        ┌──────────────────────────────────────────┐
        │ 入口层（薄包装，加在 normalizer.py 里）    │
        │  • 路径校验 + 错误信息友好化                │
        │  • 自动备份（input.docx → input.docx.bak） │
        │  • dry_run 模式（不改文件只返回摘要）       │
        │  • warnings 收集（不致命问题）             │
        └──────────┬───────────────────────────────┘
                   ↓ 调
        ┌──────────────────────────────────────────┐
        │ 现有 normalizer（行为不变）                │
        │  加载 docx → 样式 / 段落 / 表格 / 图片     │
        └──────────────────────────────────────────┘
```

**核心原则**：
- `normalizer.normalize_docx` 行为向后兼容（不传新参数 = 旧行为）
- 新增可选参数：`backup=True, dry_run=False, return_result=False, on_warning=None`
  - `on_warning`：可选回调 `def warn(msg, location=None) -> None`；传 None 时用内部 `WarningCollector`
- 默认返回 `str`（旧行为，`return_result=False`）；新调用方用 `return_result=True` 拿 `NormalizeResult`
- GUI 和 CLI 调用同一个入口，行为一致
- 现有 14 个测试全部继续通过

---

## 3. 数据结构

### 3.1 异常层级

```python
class NormalizeError(Exception):
    """规范化失败的统一异常，message 直接给最终用户看。"""
    def __init__(self, message, *, cause=None, hint=None):
        super().__init__(message)
        self.message = message
        self.cause = cause        # 原始异常（用于日志）
        self.hint = hint          # 修复建议

class InputNotFoundError(NormalizeError):       # 文件不存在
class InvalidFileTypeError(NormalizeError):     # 不是 .docx
class CorruptDocxError(NormalizeError):         # zip / XML 损坏
class OutputNotWritableError(NormalizeError):   # 权限/磁盘
class SameInputOutputError(NormalizeError):     # 输出=输入
```

每个异常的 `__str__` 格式：
```
<message>

原因：<cause 一句话摘要>
建议：<hint>
```

### 3.2 NormalizeResult

```python
@dataclass
class NormalizeResult:
    input_path: str
    output_path: Optional[str]      # dry_run 时为 None
    backup_path: Optional[str]      # 备份失败时为 None
    dry_run: bool
    paragraphs_processed: int
    tables_processed: int
    images_processed: int
    changes: list[str]              # 人类可读变更描述
    warnings: list[str]             # 非致命问题
    errors: list[str]               # 致命问题（dry_run 时为 []）
    duration_ms: int
```

向后兼容：默认 `return_result=False` 时返回 `str`（即 `str(output_path)`，与旧行为完全一致），
新调用方用 `return_result=True` 拿完整 `NormalizeResult`。

### 3.3 WarningCollector 上下文管理器

```python
class WarningCollector:
    """规范化过程中收集非致命警告。"""
    def __init__(self): self.warnings = []
    def warn(self, msg, *, location=None):
        prefix = f"[{location}] " if location else ""
        self.warnings.append(f"{prefix}{msg}")
    def __enter__(self): return self
    def __exit__(self, *args): pass
```

---

## 4. 错误信息模板

| 异常类 | message 模板 | hint |
|---|---|---|
| `InputNotFoundError` | `找不到文件: {path}` | `检查路径是否正确，文件名是否包含特殊字符` |
| `InvalidFileTypeError` | `不是有效的 docx 文件: {path}` | `确认文件是 Word 导出的 .docx 格式（不是 .doc / .rtf / .wps）` |
| `CorruptDocxError` | `docx 文件已损坏: {path}` | `用 Word 打开重新保存一次`（cause 字段带具体原因）|
| `OutputNotWritableError` | `无法写入输出: {path}` | `检查目录权限和磁盘空间`（cause 带 OSError 信息）|
| `SameInputOutputError` | `输出与输入相同: {path}` | `选择不同的输出文件` |

所有 message 中文，hint 在 `__str__` 时附加。

---

## 5. 自动备份

**行为**：
- 规范化前，复制输入文件到同目录，命名 `原文件名.docx.bak`
- 已有 `.bak` 时不覆盖，命名 `.bak.1`、`.bak.2`、…
- 备份**失败不阻塞**规范化，归为 warning：`备份失败: {原因}（不影响规范化继续）`
- 备份成功路径通过 `NormalizeResult.backup_path` 返回

**示例**：
```
输入:  D:\论文.docx
输出:  D:\论文_规范化.docx
备份:  D:\论文.docx.bak
```

**实现**：`normalizer.py` 内私有函数 `_backup_source(input_path) -> Optional[str]`

---

## 6. dry-run 模式

**接口**：
```python
result = normalize_docx(
    cfg, input_path, output_path,
    dry_run=True,
    return_result=True,
)
```

**行为**：
- 不写输出文件（`output_path` 字段为 None）
- 不创建备份（不必要）
- 仍走完整规范化 pipeline 收集 `changes` 和 `warnings`
- `changes` 是人类可读字符串列表

**changes 格式示例**：
```
· 段落 5：检测到一级标题"一、绪论"，套用 h1 样式
· 段落 8：检测到二级标题"（一）研究背景"
· 表格 1：3×4，应用独立排版
· 段落 42：参考文献段，套用 GB/T 7714 样式
· 图 1：3.5cm × 2.3cm，宽度等于正文宽度
```

**index 语义**（明确）：
- `段落 N` 中的 N 是输入 docx 里的**段落索引**（从 0 开始）
- `表格 N` 中的 N 是输入 docx 里的**表格索引**（从 0 开始），`M×K` 表示 M 行 × K 列
- `图 N` 中的 N 是输入 docx 里的**图片序号**（从 1 开始，与图题编号一致）

**跳过规则**：封面/前置页 section 里的段落/表格（被 `cover_sections` 排除）**不计入** `changes` 也不计入计数。

---

## 7. GUI 预览按钮

**位置**：`app.py` 主窗口"一键规范化区"，在"开始规范化"按钮左边加一个 **预览变更** 按钮。

**视觉**（ASCII 模拟）：
```
┌─ 一键规范化区 ──────────────────────────────┐
│ 输入: [D:\论文.docx          ] [浏览]         │
│ 输出: [D:\论文_规范化.docx    ] [浏览]         │
│                                                │
│        [ 预览变更 ]    [ 开始规范化 ]          │
│                                                │
│ ─ 状态 ─────────────────────────────────────│
│ 预览：D:\论文.docx                            │
│ 共处理 42 段、1 个表、3 张图                      │
│ · 段落 5：检测到一级标题"一、绪论"...           │
│ · 段落 8：检测到二级标题...                     │
│ · 表格 1：3×4，应用独立排版                     │
└──────────────────────────────────────────────┘
```

**实现细节**：
- 预览不需要选输出
- 状态区用 `scrolledtext.ScrolledText` 显示 `changes` 和 `warnings`
- 错误用 `messagebox.showerror("预览失败", str(e))` 弹窗
- 高级设置区折叠逻辑保持不变

**GUI 端错误处理**：
```python
try:
    result = normalize_docx(cfg, in_path, out_path, dry_run=True, return_result=True)
    render_preview(text_widget, result)
except NormalizeError as e:
    messagebox.showerror("预览失败", str(e))
    text_widget.insert("end", f"[ERROR] {e.message}\n{e.hint or ''}\n")
```

---

## 8. CLI `--dry-run` flag

**位置**：`generate.js`（Node.js CLI），与 `--out` / `--set` 平级。

**示例**：
```bash
node generate.js --dry-run                            # 干跑默认 config
node generate.js --dry-run --config 公文.json        # 指定配置
node generate.js --dry-run --set body.size=14        # 临时覆盖
```

**输出格式**（纯文本，便于脚本解析）：
```
[DRY-RUN] 不会修改任何文件

输入:   D:\论文.docx
处理摘要:
  · 段落: 42
  · 表格: 1 (3×4)
  · 图片: 3
变更:
  · 段落 5: 一级标题 "一、绪论"
  · 段落 8: 二级标题 "（一）研究背景"
  · 表格 1: 3×4 → 独立排版
警告:
  · 封面段落 2 跳过（封面保留规则）

完成。耗时 234ms。
```

**Node.js 端的 --dry-run 实现**：因为 Node.js CLI 用于**生成新文档**而非规范化现有文档，dry-run 行为略有不同：
- 不写输出 .docx
- 把 `createFormatter(config)` 跑出的所有元素 dump 到 stdout（每个段落、表格、图一个条目）
- 退出码 0（成功生成预览）或 1（生成失败）
- 与 Python 的 `dry-run`（规范化现有文档）**语义不同**：Node.js 是"先生成再丢弃"，Python 是"读 → 模拟规范化 → 不写"

---

## 9. 优雅降级：哪些算非致命

| 场景 | 行为 |
|---|---|
| 段落无法识别为标题/正文/参考文献 | warning：`段落 {idx}：未识别类型，跳过格式套用` |
| 表格的 tc 元素解析失败 | warning：`表 {t_idx} 单元格 {c_idx}：tc 解析失败，保留原样` |
| 图片尺寸读取失败 | warning：`图 {fig_idx}：尺寸读取失败，使用默认宽度` |
| 配置字段缺失（如 `cellSize`）| 用 fallback 默认值（不 warning）|
| 段落 `rFonts` 解析失败 | warning：保留原字体 |
| Normal 样式不存在 | 跳过 `_apply_document_styles`，warning |

**哪些算致命**（抛 NormalizeError）：
- 输入文件不存在
- 不是 `.docx` 后缀
- `Document(input_path)` 打开失败（zip 损坏）
- 内部 XML 解析失败（`lxml.etree.XMLSyntaxError` / `KeyError` 关键样式）
- 输出路径不可写（`PermissionError` / `OSError`）

---

## 10. 改动文件清单

| 文件 | 改动 |
|---|---|
| `normalizer.py` | 加 NormalizeError 层级、NormalizeResult、WarningCollector、`_backup_source`、`_validate_paths`；扩展 `normalize_docx` 签名加 `backup / dry_run / return_result` 参数 |
| `formatter.py` | 改一行：`_apply_run_defaults` 等接受可选 `on_warning` 回调（向后兼容）|
| `app.py` | 加"预览变更"按钮、错误弹窗、状态区 ScrolledText |
| `generate.js` | 加 `--dry-run` flag 与解析 |
| `tests/test_normalize_docx.py` | 加 14 个边界测试 |
| `tests/_helpers.py`（新）| `_make_garbage_docx` / `_make_corrupt_docx` / `_make_docx_with_nested_table` |
| `README.md` | 增"健壮性保证"小节：异常类、错误信息、备份、dry-run 用法 |

---

## 11. 测试计划

**新增 14 个测试**（覆盖 §4 §5 §6 §9）：

| # | 测试名 | 覆盖 |
|---|---|---|
| 1 | `test_normalize_docx_raises_input_not_found` | InputNotFoundError |
| 2 | `test_normalize_docx_raises_invalid_file_type` | InvalidFileTypeError |
| 3 | `test_normalize_docx_raises_same_input_output` | SameInputOutputError |
| 4 | `test_normalize_docx_raises_corrupt_docx_for_garbage_zip` | CorruptDocxError |
| 5 | `test_normalize_docx_raises_corrupt_docx_for_truncated_zip` | CorruptDocxError |
| 6 | `test_normalize_docx_raises_output_not_writable` | OutputNotWritableError |
| 7 | `test_normalize_docx_backup_creates_bak_file` | §5 |
| 8 | `test_normalize_docx_backup_uses_incremented_name` | §5 |
| 9 | `test_normalize_docx_dry_run_does_not_write_output` | §6 |
| 10 | `test_normalize_docx_dry_run_returns_normalize_result` | §6 |
| 11 | `test_normalize_docx_continues_with_warnings_on_unrecognized_paragraph` | §9 |
| 12 | `test_normalize_docx_handles_nested_table` | §9 |
| 13 | `test_normalize_docx_handles_empty_table` | §9 |
| 14 | `test_normalize_docx_message_chinese_human_readable` | §4 |

**目标覆盖率**：14 → 28 测试，全部通过。

---

## 12. 实施顺序

按风险/收益排序：

1. **阶段 1**（半天）：异常类 + 错误信息 + 路径校验 → 立即改善崩溃体验
2. **阶段 2**（半天）：WarningCollector + 优雅降级 → 改善边界体验
3. **阶段 3**（半天）：自动备份 + dry_run + NormalizeResult → 改善安全感
4. **阶段 4**（半天）：GUI 预览按钮 + 错误弹窗
5. **阶段 5**（半天）：CLI --dry-run
6. **阶段 6**（半天）：14 个测试 + README 更新

每阶段独立可发版，回归测试不破。

---

## 13. 验收标准

- [ ] 所有 28 个测试通过（14 旧 + 14 新）
- [ ] 用以下 6 个故障 docx 测试，错误信息中文且可操作：
  - 不存在的路径
  - `.txt` 改名的假 docx
  - 截断的 zip
  - 只读目录的输出
  - 嵌套表的复杂 docx
  - 0 字节表的 docx
- [ ] GUI 预览按钮可用，状态区显示变更列表
- [ ] CLI `--dry-run` 输出可被 grep 解析
- [ ] exe 重新打包，新功能在 GUI 端可见
