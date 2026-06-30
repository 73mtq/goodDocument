"""normalizer 路径校验与备份。"""
from __future__ import annotations

import os
import shutil
from typing import Optional

from .errors import (
    InputNotFoundError,
    InvalidFileTypeError,
    SameInputOutputError,
)


def _same_path(a, b) -> bool:
    try:
        return (
            os.path.abspath(os.fspath(a)).lower()
            == os.path.abspath(os.fspath(b)).lower()
        )
    except TypeError:
        return False


def _is_docx_path(path) -> bool:
    return str(path).lower().endswith(".docx")


def _validate_paths(input_path, output_path) -> None:
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


def _backup_source(input_path) -> "Optional[str]":
    """把输入文件复制到同目录的 .bak（不覆盖已有 .bak）。失败返回 None。"""
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
