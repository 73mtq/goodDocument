"""normalizer 异常类、结果数据类、警告收集器。"""
from dataclasses import dataclass, field
from typing import Optional


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


class WarningCollector:
    """规范化过程中收集非致命警告。"""

    def __init__(self):
        self.warnings = []

    def warn(self, msg, *, location=None):
        prefix = f"[{location}] " if location else ""
        self.warnings.append(f"{prefix}{msg}")

    def __enter__(self):
        return self

    def __exit__(self, *args):
        return None


@dataclass
class NormalizeResult:
    """规范化结果。dry_run=True 时 output_path 为 None，backup_path 为 None。"""
    input_path: str
    output_path: "Optional[str]" = None
    backup_path: "Optional[str]" = None
    dry_run: bool = False
    paragraphs_processed: int = 0
    tables_processed: int = 0
    images_processed: int = 0
    changes: list = field(default_factory=list)
    warnings: list = field(default_factory=list)
    errors: list = field(default_factory=list)
    duration_ms: int = 0
