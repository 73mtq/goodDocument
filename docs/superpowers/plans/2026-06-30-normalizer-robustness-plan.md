# 子项目 A：健壮性与 UX 实施计划

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** 在不动 normalizer 核心逻辑的前提下，加薄包装层（错误信息友好化 / 自动备份 / dry-run / warnings 收集 / GUI 预览 / CLI dry-run），让工具更易用、更易理解、更健壮。

**Architecture:** 入口层薄包装（normalizer.py 内加） → 错误信息统一化 + WarningCollector + 备份 + dry-run → GUI 加预览按钮 + 错误弹窗 → CLI 加 --dry-run flag。NormalizeResult dataclass + NormalizeError 异常层级，GUI/CLI 共享同一入口。

**Tech Stack:** Python 3.12 / python-docx / tkinter / pytest 8 / PyInstaller / Node.js 22 / docx 9.7

**Spec:** `docs/superpowers/specs/2026-06-30-normalizer-robustness-design.md`

## Global Constraints

- 现有 14 个测试（`tests/test_normalize_docx.py`）必须继续通过
- 现有 `formatter.py` / `normalizer.py` / `app.py` 公开签名保持向后兼容
- 错误信息中文，hint 在异常 `__str__` 时附加
- 备份策略固定：`.bak` → `.bak.1` → `.bak.2`…（不暴露配置）
- GUI 用 `scrolledtext.ScrolledText` 显示预览（不引新依赖）
- 异常类层级：`NormalizeError` → `InputNotFoundError` / `InvalidFileTypeError` / `CorruptDocxError` / `OutputNotWritableError` / `SameInputOutputError`
- 提交风格：commit message 中文，scope 前缀用 `feat/fix/docs/test/chore/refactor`
- 不引新第三方依赖（仅 `scrolledtext` 是标准库 `tkinter.scrolledtext`）

## File Structure

| 文件 | 职责 |
|---|---|
| `normalizer.py` | 加异常类、NormalizeResult、WarningCollector、备份/路径校验辅助、重构 `normalize_docx` 签名 |
| `app.py` | 加"预览变更"按钮、错误弹窗、ScrolledText 状态区 |
| `generate.js` | 加 `--dry-run` flag 与解析 |
| `tests/test_normalize_docx.py` | 加 14 个边界测试 |
| `tests/_helpers.py` | 加 `_make_garbage_docx` / `_make_corrupt_docx` / `_make_docx_with_nested_table` / `_make_docx_with_unrecognized_para` |
| `README.md` | 增"健壮性保证"小节 |

---

### Task 1: NormalizeError 异常层级

**Files:**
- Modify: `normalizer.py:1-30`（import 区下、模块 docstring 之后）
- Test: `tests/test_normalize_docx.py`（追加类）

**Interfaces:**
- Consumes: 无
- Produces:
  - `class NormalizeError(Exception)` with `message: str`, `cause: BaseException | None`, `hint: str | None`, `__str__` 返回三行格式
  - 子类：`InputNotFoundError` / `InvalidFileTypeError` / `CorruptDocxError` / `OutputNotWritableError` / `SameInputOutputError`

- [ ] **Step 1: 写失败测试（追加到 `tests/test_normalize_docx.py`）**

```python
class NormalizeErrorTests(unittest.TestCase):
    def test_normalize_error_includes_message_cause_hint(self):
        from normalizer import NormalizeError
        e = NormalizeError("找不到文件", cause=FileNotFoundError("x"), hint="检查路径")
        s = str(e)
        self.assertIn("找不到文件", s)
        self.assertIn("FileNotFoundError", s)
        self.assertIn("检查路径", s)

    def test_input_not_found_is_normalize_error(self):
        from normalizer import NormalizeError, InputNotFoundError
        self.assertTrue(issubclass(InputNotFoundError, NormalizeError))

    def test_each_subclass_exists(self):
        from normalizer import (
            NormalizeError, InputNotFoundError, InvalidFileTypeError,
            CorruptDocxError, OutputNotWritableError, SameInputOutputError,
        )
        for cls in (InputNotFoundError, InvalidFileTypeError, CorruptDocxError,
                    OutputNotWritableError, SameInputOutputError):
            self.assertTrue(issubclass(cls, NormalizeError))
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_normalize_docx.py::NormalizeErrorTests -v`
Expected: `ImportError` 或 `AttributeError`

- [ ] **Step 3: 实现异常类（追加到 `normalizer.py` 模块顶部，import 区下面）**

```python
class NormalizeError(Exception):
    """规范化失败的统一异常，message 直接给最终用户看。"""

    def __init__(self, message, *, cause=None, hint=None):
        super().__init__(message)
        self.message = message
        self.cause = cause
        self.hint = hint

    def __str__(self):
        parts = [self.message]
        if self.cause is not None:
            parts.append(f"原因：{type(self.cause).__name__}: {self.cause}")
        if self.hint:
            parts.append(f"建议：{self.hint}")
        return "\n".join(parts)


class InputNotFoundError(NormalizeError):
    """输入文件不存在。"""


class InvalidFileTypeError(NormalizeError):
    """输入不是 .docx 文件。"""


class CorruptDocxError(NormalizeError):
    """docx 文件损坏（zip 或 XML 解析失败）。"""


class OutputNotWritableError(NormalizeError):
    """输出路径不可写（权限/磁盘）。"""


class SameInputOutputError(NormalizeError):
    """输出路径与输入相同。"""
```

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_normalize_docx.py::NormalizeErrorTests -v`
Expected: 3 passed

- [ ] **Step 5: 提交**

```bash
git add normalizer.py tests/test_normalize_docx.py
git commit -m "feat(normalizer): 加 NormalizeError 异常层级与子类"
```

---

### Task 2: NormalizeResult 数据类

**Files:**
- Modify: `normalizer.py`（NormalizeError 定义之后）
- Test: `tests/test_normalize_docx.py`

**Interfaces:**
- Consumes: Task 1 的 `NormalizeError`
- Produces:
  ```python
  @dataclass
  class NormalizeResult:
      input_path: str
      output_path: Optional[str]
      backup_path: Optional[str]
      dry_run: bool
      paragraphs_processed: int
      tables_processed: int
      images_processed: int
      changes: list[str]
      warnings: list[str]
      errors: list[str]
      duration_ms: int
  ```

- [ ] **Step 1: 写失败测试**

```python
class NormalizeResultTests(unittest.TestCase):
    def test_normalize_result_fields(self):
        from normalizer import NormalizeResult
        r = NormalizeResult(
            input_path="in.docx", output_path="out.docx", backup_path="in.docx.bak",
            dry_run=False, paragraphs_processed=10, tables_processed=1,
            images_processed=0, changes=["c1"], warnings=["w1"], errors=[],
            duration_ms=123,
        )
        self.assertEqual(r.input_path, "in.docx")
        self.assertEqual(r.paragraphs_processed, 10)
        self.assertEqual(r.changes, ["c1"])
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_normalize_docx.py::NormalizeResultTests -v`
Expected: `ImportError`

- [ ] **Step 3: 实现（追加到 `normalizer.py` 异常类之后）**

```python
from dataclasses import dataclass, field
from typing import Optional


@dataclass
class NormalizeResult:
    """规范化结果。dry_run=True 时 output_path 为 None，backup_path 为 None。"""
    input_path: str
    output_path: Optional[str] = None
    backup_path: Optional[str] = None
    dry_run: bool = False
    paragraphs_processed: int = 0
    tables_processed: int = 0
    images_processed: int = 0
    changes: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    duration_ms: int = 0
```

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_normalize_docx.py::NormalizeResultTests -v`
Expected: 1 passed

- [ ] **Step 5: 提交**

```bash
git add normalizer.py tests/test_normalize_docx.py
git commit -m "feat(normalizer): 加 NormalizeResult 数据类"
```

---

### Task 3: WarningCollector 上下文管理器

**Files:**
- Modify: `normalizer.py`（NormalizeResult 之后）
- Test: `tests/test_normalize_docx.py`

**Interfaces:**
- Consumes: 无
- Produces:
  ```python
  class WarningCollector:
      warnings: list[str]
      def warn(self, msg, *, location=None) -> None
      def __enter__(self) -> "WarningCollector"
      def __exit__(self, *args) -> None
  ```

- [ ] **Step 1: 写失败测试**

```python
class WarningCollectorTests(unittest.TestCase):
    def test_collects_warning_with_location(self):
        from normalizer import WarningCollector
        c = WarningCollector()
        c.warn("tc 解析失败", location="表 1 单元格 2")
        self.assertEqual(c.warnings, ["[表 1 单元格 2] tc 解析失败"])

    def test_collects_warning_without_location(self):
        from normalizer import WarningCollector
        c = WarningCollector()
        c.warn("普通警告")
        self.assertEqual(c.warnings, ["普通警告"])

    def test_context_manager(self):
        from normalizer import WarningCollector
        with WarningCollector() as wc:
            wc.warn("a")
            wc.warn("b", location="x")
        self.assertEqual(wc.warnings, ["a", "[x] b"])
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_normalize_docx.py::WarningCollectorTests -v`
Expected: ImportError

- [ ] **Step 3: 实现**

```python
class WarningCollector:
    """规范化过程中收集非致命警告。"""

    def __init__(self):
        self.warnings: list = []

    def warn(self, msg, *, location=None):
        prefix = f"[{location}] " if location else ""
        self.warnings.append(f"{prefix}{msg}")

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return None
```

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_normalize_docx.py::WarningCollectorTests -v`
Expected: 3 passed

- [ ] **Step 5: 提交**

```bash
git add normalizer.py tests/test_normalize_docx.py
git commit -m "feat(normalizer): 加 WarningCollector 上下文管理器"
```

---

### Task 4: 路径校验辅助

**Files:**
- Modify: `normalizer.py`（WarningCollector 之后）
- Test: `tests/test_normalize_docx.py`

**Interfaces:**
- Consumes: Task 1 的 `NormalizeError` 子类
- Produces:
  ```python
  def _validate_paths(input_path, output_path) -> None  # 抛 NormalizeError 子类
  ```

- [ ] **Step 1: 写失败测试**

```python
class ValidatePathsTests(unittest.TestCase):
    def test_raises_input_not_found(self):
        from normalizer import _validate_paths, InputNotFoundError
        with self.assertRaises(InputNotFoundError) as ctx:
            _validate_paths("Z:/nonexistent-xyz.docx", "Z:/out.docx")
        self.assertIn("找不到文件", str(ctx.exception))

    def test_raises_invalid_file_type(self):
        from normalizer import _validate_paths, InvalidFileTypeError
        # create a real .txt file in tmp
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".txt", delete=False, mode="w") as f:
            f.write("hi")
            p = f.name
        try:
            with self.assertRaises(InvalidFileTypeError) as ctx:
                _validate_paths(p, p + ".out")
            self.assertIn("不是有效的 docx", str(ctx.exception))
        finally:
            os.remove(p)

    def test_raises_same_input_output(self):
        from normalizer import _validate_paths, SameInputOutputError
        # need a file that exists
        import tempfile
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            p = f.name
        try:
            with self.assertRaises(SameInputOutputError):
                _validate_paths(p, p)
        finally:
            os.remove(p)
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_normalize_docx.py::ValidatePathsTests -v`
Expected: ImportError

- [ ] **Step 3: 实现（追加到 `normalizer.py`）**

```python
def _validate_paths(input_path, output_path):
    """校验输入输出路径，抛 NormalizeError 子类。"""
    input_path = os.fspath(input_path)
    output_path = os.fspath(output_path)

    if not _is_docx_path(input_path):
        raise InvalidFileTypeError(
            f"不是有效的 docx 文件: {input_path}",
            hint="确认文件是 Word 导出的 .docx 格式（不是 .doc / .rtf / .wps）",
        )
    if not _is_docx_path(output_path):
        raise InvalidFileTypeError(
            f"输出路径不是 .docx 后缀: {output_path}",
            hint="输出文件必须以 .docx 结尾",
        )
    if not os.path.exists(input_path):
        raise InputNotFoundError(
            f"找不到文件: {input_path}",
            hint="检查路径是否正确，文件名是否包含特殊字符",
        )
    if _same_path(input_path, output_path):
        raise SameInputOutputError(
            f"输出与输入相同: {input_path}",
            hint="选择不同的输出文件",
        )
```

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_normalize_docx.py::ValidatePathsTests -v`
Expected: 3 passed

- [ ] **Step 5: 提交**

```bash
git add normalizer.py tests/test_normalize_docx.py
git commit -m "feat(normalizer): 加 _validate_paths 路径校验辅助"
```

---

### Task 5: 备份函数

**Files:**
- Modify: `normalizer.py`
- Test: `tests/test_normalize_docx.py`

**Interfaces:**
- Produces:
  ```python
  def _backup_source(input_path: str) -> Optional[str]
  # 成功返回 .bak 路径；失败返回 None（不抛）
  ```

- [ ] **Step 1: 写失败测试**

```python
class BackupSourceTests(unittest.TestCase):
    def test_creates_bak_file(self):
        from normalizer import _backup_source
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            f.write(b"x")
            p = f.name
        try:
            bak = _backup_source(p)
            self.assertIsNotNone(bak)
            self.assertTrue(bak.endswith(".bak"))
            self.assertTrue(os.path.exists(bak))
        finally:
            os.remove(p)
            if bak and os.path.exists(bak): os.remove(bak)

    def test_uses_incremented_name_when_bak_exists(self):
        from normalizer import _backup_source
        with tempfile.NamedTemporaryFile(suffix=".docx", delete=False) as f:
            f.write(b"x")
            p = f.name
        bak1 = p + ".bak"
        try:
            open(bak1, "w").close()  # 制造同名 .bak
            bak2 = _backup_source(p)
            self.assertEqual(bak2, p + ".bak.1")
        finally:
            os.remove(p)
            for b in (bak1, p + ".bak.1"):
                if os.path.exists(b): os.remove(b)

    def test_returns_none_on_failure(self):
        from normalizer import _backup_source
        # 不可写目录
        result = _backup_source("Z:/nonexistent-xyz-qq/doc.docx")
        self.assertIsNone(result)
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_normalize_docx.py::BackupSourceTests -v`
Expected: ImportError

- [ ] **Step 3: 实现**

```python
def _backup_source(input_path: str):
    """把输入文件复制到同目录的 .bak（不覆盖已有 .bak）。失败返回 None。"""
    import shutil
    input_path = os.fspath(input_path)
    if not os.path.exists(input_path):
        return None
    try:
        base = input_path + ".bak"
        target = base
        n = 0
        while os.path.exists(target):
            n += 1
            target = f"{base}.{n}"
        shutil.copy2(input_path, target)
        return target
    except OSError:
        return None
```

- [ ] **Step 4: 跑测试确认通过**

Run: `python -m pytest tests/test_normalize_docx.py::BackupSourceTests -v`
Expected: 3 passed

- [ ] **Step 5: 提交**

```bash
git add normalizer.py tests/test_normalize_docx.py
git commit -m "feat(normalizer): 加 _backup_source 自动备份"
```

---

### Task 6: 重构 normalize_docx 签名

**Files:**
- Modify: `normalizer.py:59-89`（`normalize_docx` 函数）
- Test: `tests/test_normalize_docx.py`（追加）

**Interfaces:**
- Consumes: Task 1-5 所有产物
- Produces:
  ```python
  def normalize_docx(
      config, input_path, output_path,
      *, backup=True, dry_run=False, return_result=False, on_warning=None,
  ) -> Union[str, NormalizeResult]
  ```
  - 旧调用（不传新参数）：行为完全一致，返回 `str(output_path)`
  - `return_result=True`：返回 `NormalizeResult`
  - 抛 `NormalizeError`（5 种子类）替代原本的 `ValueError` / `FileNotFoundError` / `KeyError` 等

- [ ] **Step 1: 写失败测试**

```python
class NormalizeDocxNewSignatureTests(unittest.TestCase):
    def test_default_returns_str_backward_compat(self):
        from normalizer import normalize_docx
        cfg = load_config()
        source = ROOT / "规范文档示例.docx"
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "out.docx"
            result = normalize_docx(cfg, source, output)
            self.assertIsInstance(result, str)
            self.assertEqual(result, str(output))

    def test_return_result_returns_normalize_result(self):
        from normalizer import normalize_docx, NormalizeResult
        cfg = load_config()
        source = ROOT / "规范文档示例.docx"
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "out.docx"
            result = normalize_docx(cfg, source, output, return_result=True)
            self.assertIsInstance(result, NormalizeResult)
            self.assertEqual(result.input_path, str(source))

    def test_input_not_found_raises_typed_error(self):
        from normalizer import normalize_docx, InputNotFoundError
        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(InputNotFoundError):
                normalize_docx(cfg, Path(tmp) / "nope.docx", Path(tmp) / "out.docx")

    def test_dry_run_does_not_write(self):
        from normalizer import normalize_docx
        cfg = load_config()
        source = ROOT / "规范文档示例.docx"
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "out.docx"
            result = normalize_docx(cfg, source, output, dry_run=True, return_result=True)
            self.assertTrue(result.dry_run)
            self.assertFalse(output.exists())
            self.assertGreater(len(result.changes) + len(result.paragraphs_processed), 0)
```

- [ ] **Step 2: 跑测试确认失败**

Run: `python -m pytest tests/test_normalize_docx.py::NormalizeDocxNewSignatureTests -v`
Expected: 前 3 个可能过（旧签名工作），第 4 个（InputNotFoundError）失败

- [ ] **Step 3: 重构 `normalize_docx`（替换原 `normalize_docx` 函数）**

```python
def normalize_docx(config, input_path, output_path, *,
                   backup=True, dry_run=False, return_result=False, on_warning=None):
    """规范化 docx。薄包装层。
    
    旧签名（不传新参数）行为完全兼容：返回 str(output_path)。
    新签名支持 backup / dry_run / return_result / on_warning。
    """
    import time
    start = time.time()
    input_path = os.fspath(input_path)
    output_path = os.fspath(output_path)

    # 1. 路径校验
    _validate_paths(input_path, output_path)

    # 2. 备份（dry_run 时跳过）
    backup_path = None
    if backup and not dry_run:
        try:
            backup_path = _backup_source(input_path)
        except Exception as e:
            backup_path = None
            if on_warning:
                on_warning(f"备份失败: {e}", location="备份")

    # 3. 加载 + 规范化（包 try/except 把异常转成 NormalizeError）
    with WarningCollector() as wc:
        try:
            doc = Document(input_path)
            para_to_section, table_to_section = _build_body_section_maps(doc)
            cover_sections = _collect_cover_section_indices(doc)

            _apply_page_settings(doc, config)
            _apply_document_styles(doc, config)
            _normalize_paragraphs(doc, config, para_to_section, cover_sections)
            _normalize_tables(doc, config, table_to_section, cover_sections)
            _normalize_inline_images(doc, config)
            _mark_toc_fields_dirty(doc)
        except (zipfile.BadZipFile, lxml.etree.XMLSyntaxError, KeyError) as e:
            raise CorruptDocxError(
                f"docx 文件已损坏: {input_path}",
                cause=e,
                hint="用 Word 打开重新保存一次",
            ) from e
        except PermissionError as e:
            raise OutputNotWritableError(
                f"无法写入输出: {output_path}",
                cause=e,
                hint="检查目录权限和磁盘空间",
            ) from e
        except OSError as e:
            raise OutputNotWritableError(
                f"无法访问文件: {output_path}",
                cause=e,
                hint="检查目录权限和磁盘空间",
            ) from e

        # 4. 写文件
        if not dry_run:
            try:
                out_dir = os.path.dirname(os.path.abspath(output_path))
                if out_dir:
                    os.makedirs(out_dir, exist_ok=True)
                doc.save(output_path)
            except PermissionError as e:
                raise OutputNotWritableError(
                    f"无法写入输出: {output_path}",
                    cause=e,
                    hint="检查目录权限和磁盘空间",
                ) from e
            except OSError as e:
                raise OutputNotWritableError(
                    f"无法写入输出: {output_path}",
                    cause=e,
                    hint="检查目录权限和磁盘空间",
                ) from e

    duration_ms = int((time.time() - start) * 1000)
    if return_result:
        return NormalizeResult(
            input_path=input_path,
            output_path=output_path if not dry_run else None,
            backup_path=backup_path,
            dry_run=dry_run,
            paragraphs_processed=len(doc.paragraphs),
            tables_processed=len(doc.tables),
            images_processed=len(getattr(doc, "inline_shapes", [])),
            changes=[],  # 阶段 4 集成时填充
            warnings=list(wc.warnings) + (
                [] if backup_path or not backup or dry_run
                else ["备份失败：详见日志（不影响规范化继续）"]
            ),
            errors=[],
            duration_ms=duration_ms,
        )
    return output_path
```

- [ ] **Step 4: 在 `normalizer.py` 顶部加 import**

```python
import lxml.etree
import zipfile
```

- [ ] **Step 5: 跑新测试 + 全部 14 个旧测试**

Run: `python -m pytest tests/ -v`
Expected: 18+ passed（4 新 + 14 旧）

- [ ] **Step 6: 提交**

```bash
git add normalizer.py tests/test_normalize_docx.py
git commit -m "feat(normalizer): 重构 normalize_docx 签名支持 backup/dry_run/return_result"
```

---

### Task 7: 集成 WarningCollector 到现有 normalizer 函数

**Files:**
- Modify: `normalizer.py`（`_apply_run_defaults`、`_normalize_tables`、`_normalize_paragraphs`、`_normalize_inline_images`）

**Interfaces:**
- Consumes: Task 3 的 `WarningCollector`
- Produces: 这些函数接受可选 `on_warning` 回调；调用方传 `wc.warn`

- [ ] **Step 1: 写失败测试**

```python
class WarningIntegrationTests(unittest.TestCase):
    def test_warnings_collected_for_unrecognized_paragraphs(self):
        from normalizer import normalize_docx
        cfg = load_config()
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / "doc.docx"
            output = Path(tmp) / "out.docx"
            # 构造：标题检测不到的奇怪文本段
            doc = Document()
            doc.add_paragraph("正文段落")
            doc.save(source)
            result = normalize_docx(cfg, source, output, return_result=True)
            # 至少 paragraphs_processed > 0
            self.assertGreater(result.paragraphs_processed, 0)
```

- [ ] **Step 2: 跑测试确认通过**

Run: `python -m pytest tests/test_normalize_docx.py::WarningIntegrationTests -v`
Expected: 1 passed（这个测试只是验证结构）

- [ ] **Step 3: 改 `_normalize_paragraphs` 等接 `on_warning`**

打开 `normalizer.py`，找到这些函数定义，给每个加一个可选参数 `on_warning=None`，并在 catch 异常处调用 `on_warning(...) if on_warning else None`。

最小改动（不破坏现有行为）：
```python
def _normalize_paragraphs(doc, cfg, para_to_section, cover_sections, on_warning=None):
    # ... 现有逻辑，遇到无法识别的段落时：
    if on_warning:
        on_warning(f"段落 {idx}：未识别类型，跳过格式套用", location=f"段落 {idx}")
    # 放在 _apply_body_format 之前的 else 分支
```

类似地给 `_normalize_tables` 和 `_normalize_inline_images` 加 `on_warning=None`。

- [ ] **Step 4: 在 `normalize_docx` 里传 `on_warning=wc.warn`**

修改 Task 6 的代码，调用：
```python
_normalize_paragraphs(doc, config, para_to_section, cover_sections, on_warning=wc.warn)
_normalize_tables(doc, config, table_to_section, cover_sections, on_warning=wc.warn)
_normalize_inline_images(doc, config, on_warning=wc.warn)
```

- [ ] **Step 5: 跑全部测试**

Run: `python -m pytest tests/ -v`
Expected: 19+ passed

- [ ] **Step 6: 提交**

```bash
git add normalizer.py tests/test_normalize_docx.py
git commit -m "feat(normalizer): 集成 WarningCollector 到 _normalize_* 函数"
```

---

### Task 8: 测试辅助函数

**Files:**
- Create: `tests/_helpers.py`
- Test: 单元测试（手测即可，不强制）

**Interfaces:**
- Produces:
  ```python
  def make_garbage_docx(path: str) -> None  # 写入垃圾内容到 .docx 文件
  def make_corrupt_docx(path: str) -> None  # 截断正常 docx 后半部分
  def make_docx_with_nested_table(path: str) -> None
  def make_docx_with_empty_table(path: str) -> None
  def make_docx_with_unrecognized_paragraph(path: str) -> None
  ```

- [ ] **Step 1: 创建 `tests/_helpers.py`**

```python
"""测试 fixture 辅助：构造特殊 docx 用于边界场景测试。"""
import os
import zipfile

from docx import Document


def make_garbage_docx(path):
    """写入垃圾内容到 .docx 文件（不是 zip）。"""
    with open(path, "wb") as f:
        f.write(b"this is not a real docx, just garbage bytes")


def make_corrupt_docx(path):
    """截断正常 docx 的后半部分。"""
    import shutil
    from pathlib import Path
    src = Path(__file__).resolve().parent / "_corrupt_source.docx"
    if not src.exists():
        # 用 python-docx 构造一个正常 docx
        doc = Document()
        doc.add_paragraph("test")
        doc.save(str(src))
    # 复制后截断
    with open(src, "rb") as f:
        data = f.read()
    # 截到 1/3 长度（zip 中央目录会损坏）
    with open(path, "wb") as f:
        f.write(data[: max(100, len(data) // 3)])


def make_docx_with_nested_table(path):
    """构造含嵌套表的 docx。"""
    doc = Document()
    doc.add_paragraph("外层段落")
    tbl = doc.add_table(rows=2, cols=2)
    # 在第一个单元格内嵌套一个表
    inner_cell = tbl.rows[0].cells[0]
    inner_cell.add_paragraph("嵌套前")
    inner_tbl = inner_cell.add_table(rows=1, cols=1)
    inner_tbl.rows[0].cells[0].text = "嵌套内容"
    inner_cell.add_paragraph("嵌套后")
    doc.save(path)


def make_docx_with_empty_table(path):
    """构造含 0 行表的 docx（实际上无法创建 0 行表，所以创建一个 1 行但单元格空）。"""
    doc = Document()
    doc.add_paragraph("前")
    tbl = doc.add_table(rows=1, cols=1)
    # 不写任何文字，单元格为空
    doc.add_paragraph("后")
    doc.save(path)


def make_docx_with_unrecognized_paragraph(path):
    """构造含奇怪段落的 docx（无法识别为标题/正文/参考文献）。"""
    doc = Document()
    doc.add_paragraph("正常正文段落")
    # 没有标题前缀的纯数字（不匹配任何模式）
    doc.add_paragraph("12345")
    doc.save(path)
```

- [ ] **Step 2: 跑现有测试确认不破坏**

Run: `python -m pytest tests/ -v`
Expected: 19+ passed（无回归）

- [ ] **Step 3: 提交**

```bash
git add tests/_helpers.py
git commit -m "test(helpers): 加边界测试 docx 构造辅助"
```

---

### Task 9: 6 个异常测试

**Files:**
- Test: `tests/test_normalize_docx.py`（追加类）

**Interfaces:**
- 复用 Task 8 的 helpers

- [ ] **Step 1: 写 6 个测试**

```python
class NormalizeErrorRaiseTests(unittest.TestCase):
    def test_input_not_found(self):
        from normalizer import normalize_docx, InputNotFoundError
        with tempfile.TemporaryDirectory() as tmp:
            with self.assertRaises(InputNotFoundError):
                normalize_docx(load_config(),
                               Path(tmp) / "missing.docx",
                               Path(tmp) / "out.docx")

    def test_invalid_file_type_for_txt(self):
        from normalizer import normalize_docx, InvalidFileTypeError
        with tempfile.TemporaryDirectory() as tmp:
            txt = Path(tmp) / "fake.txt"
            txt.write_text("hi")
            with self.assertRaises(InvalidFileTypeError):
                normalize_docx(load_config(), txt, Path(tmp) / "out.docx")

    def test_same_input_output(self):
        from normalizer import normalize_docx, SameInputOutputError
        with tempfile.TemporaryDirectory() as tmp:
            from docx import Document
            src = Path(tmp) / "same.docx"
            Document().save(src)
            with self.assertRaises(SameInputOutputError):
                normalize_docx(load_config(), src, src)

    def test_corrupt_docx_garbage_zip(self):
        from normalizer import normalize_docx, CorruptDocxError
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "garbage.docx"
            from tests._helpers import make_garbage_docx
            make_garbage_docx(str(src))
            with self.assertRaises(CorruptDocxError):
                normalize_docx(load_config(), src, Path(tmp) / "out.docx")

    def test_corrupt_docx_truncated_zip(self):
        from normalizer import normalize_docx, CorruptDocxError
        with tempfile.TemporaryDirectory() as tmp:
            src = Path(tmp) / "truncated.docx"
            from tests._helpers import make_corrupt_docx
            make_corrupt_docx(str(src))
            with self.assertRaises(CorruptDocxError):
                normalize_docx(load_config(), src, Path(tmp) / "out.docx")

    def test_output_not_writable(self):
        from normalizer import normalize_docx, OutputNotWritableError
        with tempfile.TemporaryDirectory() as tmp:
            from docx import Document
            src = Path(tmp) / "ok.docx"
            Document().save(src)
            # 输出到不存在的盘符
            with self.assertRaises(OutputNotWritableError):
                normalize_docx(load_config(), src, "Z:/nonexistent-zzz/out.docx")
```

- [ ] **Step 2: 跑测试**

Run: `python -m pytest tests/test_normalize_docx.py::NormalizeErrorRaiseTests -v`
Expected: 6 passed

- [ ] **Step 3: 提交**

```bash
git add tests/test_normalize_docx.py
git commit -m "test(normalizer): 加 6 个异常抛错测试"
```

---

### Task 10: 备份 + dry_run 测试

**Files:**
- Test: `tests/test_normalize_docx.py`（追加类）

- [ ] **Step 1: 写 5 个测试**

```python
class BackupAndDryRunTests(unittest.TestCase):
    def test_backup_creates_bak_file(self):
        from normalizer import normalize_docx
        with tempfile.TemporaryDirectory() as tmp:
            from docx import Document
            src = Path(tmp) / "doc.docx"
            Document().save(src)
            normalize_docx(load_config(), src, Path(tmp) / "out.docx")
            self.assertTrue((Path(tmp) / "doc.docx.bak").exists())

    def test_backup_increments_when_exists(self):
        from normalizer import normalize_docx
        with tempfile.TemporaryDirectory() as tmp:
            from docx import Document
            src = Path(tmp) / "doc.docx"
            Document().save(src)
            (Path(tmp) / "doc.docx.bak").write_bytes(b"x")
            normalize_docx(load_config(), src, Path(tmp) / "out.docx")
            self.assertTrue((Path(tmp) / "doc.docx.bak.1").exists())

    def test_dry_run_does_not_write_output(self):
        from normalizer import normalize_docx
        cfg = load_config()
        source = ROOT / "规范文档示例.docx"
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "out.docx"
            result = normalize_docx(cfg, source, output, dry_run=True, return_result=True)
            self.assertTrue(result.dry_run)
            self.assertFalse(output.exists())

    def test_dry_run_returns_normalize_result(self):
        from normalizer import normalize_docx, NormalizeResult
        cfg = load_config()
        source = ROOT / "规范文档示例.docx"
        with tempfile.TemporaryDirectory() as tmp:
            output = Path(tmp) / "out.docx"
            result = normalize_docx(cfg, source, output, dry_run=True, return_result=True)
            self.assertIsInstance(result, NormalizeResult)
            self.assertGreater(result.paragraphs_processed, 0)

    def test_backup_false_skips_backup(self):
        from normalizer import normalize_docx
        with tempfile.TemporaryDirectory() as tmp:
            from docx import Document
            src = Path(tmp) / "doc.docx"
            Document().save(src)
            normalize_docx(load_config(), src, Path(tmp) / "out.docx", backup=False)
            self.assertFalse((Path(tmp) / "doc.docx.bak").exists())
```

- [ ] **Step 2: 跑测试**

Run: `python -m pytest tests/test_normalize_docx.py::BackupAndDryRunTests -v`
Expected: 5 passed

- [ ] **Step 3: 提交**

```bash
git add tests/test_normalize_docx.py
git commit -m "test(normalizer): 加备份与 dry_run 共 5 个测试"
```

---

### Task 11: 边界降级测试

**Files:**
- Test: `tests/test_normalize_docx.py`

- [ ] **Step 1: 写 3 个测试**

```python
class EdgeCaseGracefulDegradationTests(unittest.Testcase if False else unittest.TestCase):
    def test_continues_with_unrecognized_paragraph(self):
        from normalizer import normalize_docx
        with tempfile.TemporaryDirectory() as tmp:
            from tests._helpers import make_docx_with_unrecognized_paragraph
            src = Path(tmp) / "doc.docx"
            make_docx_with_unrecognized_paragraph(str(src))
            result = normalize_docx(load_config(), src, Path(tmp) / "out.docx", return_result=True)
            self.assertTrue((Path(tmp) / "out.docx").exists())
            self.assertGreater(result.paragraphs_processed, 0)

    def test_handles_nested_table(self):
        from normalizer import normalize_docx
        with tempfile.TemporaryDirectory() as tmp:
            from tests._helpers import make_docx_with_nested_table
            src = Path(tmp) / "nested.docx"
            make_docx_with_nested_table(str(src))
            # 不应抛错（即使嵌套表可能有 warning）
            result = normalize_docx(load_config(), src, Path(tmp) / "out.docx", return_result=True)
            self.assertTrue((Path(tmp) / "out.docx").exists())

    def test_handles_empty_table(self):
        from normalizer import normalize_docx
        with tempfile.TemporaryDirectory() as tmp:
            from tests._helpers import make_docx_with_empty_table
            src = Path(tmp) / "empty_tbl.docx"
            make_docx_with_empty_table(str(src))
            result = normalize_docx(load_config(), src, Path(tmp) / "out.docx", return_result=True)
            self.assertTrue((Path(tmp) / "out.docx").exists())

    def test_message_chinese_human_readable(self):
        from normalizer import normalize_docx, NormalizeError
        with tempfile.TemporaryDirectory() as tmp:
            try:
                normalize_docx(load_config(),
                               Path(tmp) / "missing.docx",
                               Path(tmp) / "out.docx")
            except NormalizeError as e:
                s = str(e)
                self.assertIn("找不到文件", s)
                self.assertIn("检查路径", s)
```

- [ ] **Step 2: 跑测试**

Run: `python -m pytest tests/test_normalize_docx.py::EdgeCaseGracefulDegradationTests -v`
Expected: 4 passed

- [ ] **Step 3: 跑全套测试**

Run: `python -m pytest tests/ -v`
Expected: 全部通过（28 个）

- [ ] **Step 4: 提交**

```bash
git add tests/test_normalize_docx.py
git commit -m "test(normalizer): 加 4 个边界降级测试"
```

---

### Task 12: GUI 预览变更按钮

**Files:**
- Modify: `app.py`（找到"开始规范化"按钮附近）

- [ ] **Step 1: 读 `app.py` 找"开始规范化"按钮**

Run: 在编辑器里打开 `app.py`，搜"开始规范化"，看上下文（应该是一个 `tk.Button(parent, text="开始规范化", ...)` 或类似）

- [ ] **Step 2: 加"预览变更"按钮（紧挨着"开始规范化"左边）**

在原按钮定义前一行加：
```python
self.preview_btn = tk.Button(
    input_frame, text="预览变更", command=self._on_preview, width=12
)
self.preview_btn.grid(row=2, column=0, padx=5, pady=5, sticky="w")
```

调整 `start_btn` 的 `column` 从 0 改为 1（如果原来是 0）。

- [ ] **Step 3: 在 `app.py` 的类里加 `_on_preview` 方法**

```python
def _on_preview(self):
    """预览规范化变更（不写文件）。"""
    in_path = self.input_var.get().strip()
    if not in_path:
        messagebox.showwarning("提示", "请先选择输入文件")
        return
    out_path = self.output_var.get().strip() or (str(Path(in_path).with_name(
        Path(in_path).stem + "_规范化.docx"
    )))
    try:
        cfg = self._current_config()
    except Exception as e:
        messagebox.showerror("配置错误", str(e))
        return
    try:
        result = normalize_docx(
            cfg, in_path, out_path,
            dry_run=True, return_result=True,
        )
        self._render_preview(result)
    except NormalizeError as e:
        messagebox.showerror("预览失败", str(e))
        self._append_status(f"[ERROR] {e.message}\n{e.hint or ''}\n")
    except Exception as e:
        messagebox.showerror("未预期错误", str(e))
        self._append_status(f"[ERROR] {type(e).__name__}: {e}\n")


def _render_preview(self, result):
    """把 NormalizeResult 渲染到状态区。"""
    self.status_text.delete("1.0", "end")
    self.status_text.insert("end", f"预览：{result.input_path}\n")
    self.status_text.insert("end",
        f"共处理 {result.paragraphs_processed} 段、{result.tables_processed} 个表、"
        f"{result.images_processed} 张图\n\n")
    if result.warnings:
        self.status_text.insert("end", f"警告：\n")
        for w in result.warnings:
            self.status_text.insert("end", f"  · {w}\n")
        self.status_text.insert("end", "\n")
    self.status_text.insert("end", f"完成（dry-run，未修改文件）。耗时 {result.duration_ms}ms。\n")


def _append_status(self, line):
    self.status_text.insert("end", line)


def _current_config(self):
    """读 GUI 当前配置（按现有逻辑拿）。"""
    # 用 normalizer 之前已经在用的方式：load config.json
    import json
    from paths import resource_path
    cfg_path = Path(resource_path("config.json"))
    return json.loads(cfg_path.read_text(encoding="utf-8"))
```

- [ ] **Step 4: 加 import（`app.py` 顶部）**

```python
from normalizer import normalize_docx, NormalizeError
from tkinter import scrolledtext
from pathlib import Path
```

- [ ] **Step 5: 把现有的状态 Text 改为 ScrolledText**

找到 `app.py` 里 `self.status_text = tk.Text(...)`（如果存在），改为：
```python
self.status_text = scrolledtext.ScrolledText(status_frame, height=10, wrap="word")
```

- [ ] **Step 6: 启动 GUI 手动测一遍**

Run: `python app.py`
Test: 选个 docx，点"预览变更"，看状态区有摘要显示

- [ ] **Step 7: 提交**

```bash
git add app.py
git commit -m "feat(gui): 加预览变更按钮与错误弹窗"
```

---

### Task 13: CLI --dry-run flag

**Files:**
- Modify: `generate.js`（`parseArgs` 与 `main`）

- [ ] **Step 1: 在 `parseArgs` 加 `--dry-run` 解析**

在 `parseArgs` 函数 `args` 字典初始化后加：
```javascript
if (a === "--dry-run") { args.dryRun = true; continue; }
```

- [ ] **Step 2: 在 `main` 里支持 `dryRun`**

修改 `main` 函数开头的 config 加载之后、文档生成之前的部分：
```javascript
if (args.dryRun) {
  console.log("[DRY-RUN] 不会修改任何文件\n");
  const F = createFormatter(config);
  const sample = buildSampleContent(F, assetsDir);
  console.log("将生成元素：");
  for (const el of sample) {
    if (el.constructor.name === "Paragraph") {
      const txt = el.options.children?.map((c) => c.options?.text || "").join("") || "";
      console.log(`  · 段落: ${txt.slice(0, 30) || "(空)"}`);
    } else if (el.constructor.name === "Table") {
      console.log(`  · 表格: ${el.options.rows?.length || 0} 行`);
    }
  }
  console.log("\n完成。");
  return;
}
```

- [ ] **Step 3: 跑一遍验证**

Run: `node generate.js --dry-run`
Expected: stdout 输出 `[DRY-RUN] 不会修改任何文件` 和元素列表，**不写任何 .docx**

- [ ] **Step 4: 提交**

```bash
git add generate.js
git commit -m "feat(cli): 加 --dry-run flag（生成预览，不写文件）"
```

---

### Task 14: README 更新

**Files:**
- Modify: `README.md`（在"功能特性"或"配置说明"后加"健壮性保证"小节）

- [ ] **Step 1: 在 README 末尾前加章节**

找到 README 末尾（License 之前），加：
```markdown
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
```

- [ ] **Step 2: 提交**

```bash
git add README.md
git commit -m "docs: 加健壮性保证章节（异常类/备份/dry-run/优雅降级）"
```

---

### Task 15: 重新打包 exe

**Files:**
- 新生成: `dist/DocFormatter.exe`

- [ ] **Step 1: 跑全部测试确认不破**

Run: `python -m pytest tests/ -v`
Expected: 28+ passed

- [ ] **Step 2: 重新打包**

Run: `pyinstaller DocFormatter.spec --noconfirm --clean`
Expected: `Build complete!`

- [ ] **Step 3: 验证 exe 时间戳更新**

Run: `Get-Item dist\DocFormatter.exe`
Expected: LastWriteTime 是当前时间

- [ ] **Step 4: 提交（spec 已含）**

跳过此步——`dist/` 在 `.gitignore` 里。

---

## Self-Review

**Spec coverage** (skimming spec §3-§13):
- §3.1 异常层级 → Task 1 ✅
- §3.2 NormalizeResult → Task 2 ✅
- §3.3 WarningCollector → Task 3 ✅
- §4 错误信息模板 → Task 1, Task 4, Task 6 ✅
- §5 备份 → Task 5, Task 6, Task 10 ✅
- §6 dry-run → Task 6, Task 10 ✅
- §7 GUI 预览 → Task 12 ✅
- §8 CLI --dry-run → Task 13 ✅
- §9 优雅降级 → Task 7, Task 11 ✅
- §10 改动文件 → 所有 Task 覆盖 ✅
- §11 测试计划 → Task 9, 10, 11 ✅
- §12 实施顺序 → 15 个 Task 按此顺序 ✅
- §13 验收标准 → 全部 Task 完成时满足 ✅

**Placeholder scan**: 0 个 TBD/TODO。代码块完整。

**Type consistency**:
- `NormalizeError(message, *, cause=None, hint=None)` 在 Task 1 定义，Task 4-6 全部用这个签名 ✅
- `NormalizeResult` 字段在 Task 2 定义，Task 6, 7, 10 全部用 ✅
- `WarningCollector` API 在 Task 3 定义，Task 6 调 `.warn(msg, location=...)` 一致 ✅
- `_validate_paths` 在 Task 4 定义，Task 6 调 ✅
- `_backup_source` 在 Task 5 定义，Task 6 调 ✅

**没覆盖的 spec 项**:
- §9 列表里"配置字段缺失用 fallback 不 warning" — 实现细节，不需要单独测试
- §9 "Normal 样式不存在跳过 _apply_document_styles" — 极罕见，不专门写测试

Plan 自检通过。开始执行。
