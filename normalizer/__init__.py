"""normalizer 包：规范化现有 Word 文档。

公共 API（re-export）：
- normalize_docx
- inspect_docx / DocxInspection
- NormalizeError / 5 子类 / NormalizeResult / WarningCollector
- 内部辅助（供 tests 验证）
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
from .paths import (
    _same_path,
    _is_docx_path,
    _validate_paths,
    _backup_source,
)
from .normalize import (
    _paragraph_text,
    _detect_heading_level,
    _is_toc_style,
    _is_caption_para,
    _find_toc_title_index,
    _section_has_page_number,
    _build_body_section_maps,
    _collect_cover_section_indices,
    _apply_run_defaults,
    _apply_page_settings,
    _apply_document_styles,
    _apply_heading_format,
    _apply_body_format,
    _apply_caption_format,
    _apply_reference_format,
    _normalize_paragraphs,
    _normalize_tables,
    _normalize_inline_images,
    _mark_toc_fields_dirty,
    _has_header_footer,
    _has_complex_objects,
)


__all__ = [
    # 公共 API
    "normalize_docx",
    "inspect_docx",
    "DocxInspection",
    # 异常
    "NormalizeError",
    "InputNotFoundError",
    "InvalidFileTypeError",
    "CorruptDocxError",
    "OutputNotWritableError",
    "SameInputOutputError",
    # 数据类
    "NormalizeResult",
    # 警告
    "WarningCollector",
    # 常量
    "HEADING_PATTERNS",
]
