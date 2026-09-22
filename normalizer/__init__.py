"""normalizer 包：规范化现有 Word 文档。

公共 API：
- normalize_docx
- inspect_docx / DocxInspection
- NormalizeError 及 5 个子类
- NormalizeResult
- WarningCollector
- HEADING_PATTERNS
"""
from .errors import (
    NormalizeError,
    InputNotFoundError,
    InvalidFileTypeError,
    CorruptDocxError,
    OutputNotWritableError,
    SameInputOutputError,
    WarningCollector,
    NormalizeResult,
)
from .core import normalize_docx
from .normalize import (
    DocxInspection,
    inspect_docx,
    HEADING_PATTERNS,
)


__all__ = [
    "normalize_docx",
    "inspect_docx",
    "DocxInspection",
    "NormalizeError",
    "InputNotFoundError",
    "InvalidFileTypeError",
    "CorruptDocxError",
    "OutputNotWritableError",
    "SameInputOutputError",
    "NormalizeResult",
    "WarningCollector",
    "HEADING_PATTERNS",
]
