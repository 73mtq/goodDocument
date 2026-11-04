"""normalizer 主入口。"""

from __future__ import annotations

import os
import time
from typing import Optional

from formatter import apply_document_styles

from .errors import (
    NormalizeResult,
    OutputNotWritableError,
    WarningCollector,
)
from .normalize import (
    _apply_page_settings,
    _build_body_section_maps,
    _collect_cover_section_indices,
    _format_headers,
    _mark_toc_fields_dirty,
    _normalize_inline_images,
    _normalize_paragraphs,
    _normalize_tables,
    _open_document,
)
from .paths import _backup_source, _validate_paths


def normalize_docx(
    config,
    input_path,
    output_path,
    assets_dir=None,
    *,
    backup=False,
    dry_run=False,
    return_result=False,
    on_warning=None,
):
    """规范化 docx。薄包装层。

    旧签名（不传新参数）行为完全兼容：返回 str(output_path)。
    新签名支持 backup / dry_run / return_result / on_warning。
    抛 NormalizeError 子类替代原本的 ValueError / FileNotFoundError / KeyError。
    dry_run=True 且 return_result=False 时返回 None（未写文件）。

    注意:backup 默认 False (0.2.1 起) — 早期版本默认 True,会一直产生 .bak 累计。
          需要备份时显式传 backup=True。
    """
    del assets_dir  # Kept for a stable public interface.
    start = time.time()
    input_path = os.fspath(input_path)
    output_path = os.fspath(output_path)

    # 1. 路径校验
    _validate_paths(input_path, output_path)

    with WarningCollector() as wc:

        def emit_warning(msg, *, location=None):
            wc.warn(msg, location=location)
            if on_warning:
                on_warning(msg, location=location)

        # 2. 备份（dry_run 时跳过；备份失败归为 warning）
        backup_path: Optional[str] = None
        if backup and not dry_run:
            backup_path = _backup_source(input_path)
            if backup_path is None:
                emit_warning("备份失败（不影响规范化继续）", location="备份")

        # 3. 加载 + 规范化
        doc = _open_document(input_path)
        changes = []
        try:
            para_to_section, table_to_section = _build_body_section_maps(doc)
            cover_sections = _collect_cover_section_indices(doc)

            _apply_page_settings(doc, config)
            apply_document_styles(doc, config)
            _format_headers(doc, config, cover_sections)
            paragraphs_processed = _normalize_paragraphs(
                doc, config, para_to_section, cover_sections, on_warning=emit_warning, on_change=changes.append
            )
            tables_processed = _normalize_tables(
                doc, config, table_to_section, cover_sections, on_warning=emit_warning, on_change=changes.append
            )
            images_processed = _normalize_inline_images(doc, config, on_warning=emit_warning, on_change=changes.append)
            _mark_toc_fields_dirty(doc)
        except PermissionError as e:
            raise OutputNotWritableError(
                f"无法访问文件: {output_path}",
                cause=e,
                hint="检查目录权限和磁盘空间",
            ) from e

        # 4. 写文件（dry_run 时跳过）
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
            paragraphs_processed=paragraphs_processed,
            tables_processed=tables_processed,
            images_processed=images_processed,
            changes=list(changes),
            warnings=list(wc.warnings),
            errors=[],
            duration_ms=duration_ms,
        )
    if dry_run:
        return None
    return output_path
