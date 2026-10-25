"""normalizer 包：规范化现有 Word 文档。

公共 API：
- normalize_docx
- inspect_docx / DocxInspection
- NormalizeError 及 5 个子类
- NormalizeResult
- WarningCollector
- HEADING_PATTERNS
"""
from .core import normalize_docx
from .errors import (
    CorruptDocxError,
    InputNotFoundError,
    InvalidFileTypeError,
    NormalizeError,
    NormalizeResult,
    OutputNotWritableError,
    SameInputOutputError,
    WarningCollector,
)
from .normalize import (
    HEADING_PATTERNS,
    DocxInspection,
    inspect_docx,
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
