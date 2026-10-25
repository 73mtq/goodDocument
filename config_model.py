"""Configuration loading, saving, and validation."""

import json
import os
from pathlib import Path

from paths import templates_dir

REQUIRED_SECTIONS = (
    "page",
    "pageNumber",
    "body",
    "headings",
    "figure",
    "table",
    "references",
    "output",
)

SIZE_NAME_TO_PT = {
    "初号": 42, "小初": 36,
    "一号": 26, "小一": 24,
    "二号": 22, "小二": 18,
    "三号": 16, "小三": 15,
    "四号": 14, "小四": 12,
    "五号": 10.5, "小五": 9,
    "六号": 7.5, "小六": 6.5,
    "七号": 5.5, "八号": 5,
}

ALIGN_VALUES = ("left", "center", "right", "justify", "both")
CELL_ALIGN_VALUES = ALIGN_VALUES + ("smart",)
PAGE_SIZES = ("a4", "letter")
FIGURE_MODES = ("fit", "fixed")
CAPTION_POSITIONS = ("above", "below")


def _is_valid_size(value):
    if value is None:
        return False
    if isinstance(value, (int, float)):
        return value > 0
    text = str(value).strip()
    if text in SIZE_NAME_TO_PT:
        return True
    try:
        return float(text) > 0
    except ValueError:
        return False


def _is_number(value):
    return isinstance(value, (int, float)) and not isinstance(value, bool)


def _check_number(mapping, key, label, minimum=None, maximum=None):
    if key not in mapping:
        return
    value = mapping[key]
    if value is None:
        return
    if not _is_number(value):
        raise ValueError(label + " 必须是数字，当前为: " + str(value))
    if minimum is not None and value < minimum:
        raise ValueError(label + " 过小，当前为: " + str(value))
    if maximum is not None and value > maximum:
        raise ValueError(label + " 过大，当前为: " + str(value))


def _check_align(mapping, key, label):
    if key not in mapping:
        return
    value = str(mapping[key]).lower()
    if value not in ALIGN_VALUES:
        raise ValueError(label + " 只能是 left/center/right/justify，当前为: " + str(mapping[key]))


def _check_size_value(mapping, key, label):
    if key not in mapping:
        return
    if not _is_valid_size(mapping[key]):
        raise ValueError(label + " 非法，当前为: " + str(mapping[key]))


def _check_cell_align(mapping, key, label):
    if key not in mapping:
        return
    value = str(mapping[key]).lower()
    if value not in CELL_ALIGN_VALUES:
        raise ValueError(
            label + " 只能是 left/center/right/justify/smart，当前为: " + str(mapping[key]))


def load_config_file(path):
    with open(path, "r", encoding="utf-8") as f:
        cfg = json.load(f)
    validate_config(cfg)
    return cfg


def save_config_file(config, path):
    p = Path(path)
    with open(p, "w", encoding="utf-8") as f:
        json.dump(config, f, ensure_ascii=False, indent=2)


def validate_config(config):
    if not isinstance(config, dict):
        raise ValueError("配置必须是 JSON 对象")

    missing = [name for name in REQUIRED_SECTIONS if name not in config]
    if missing:
        raise ValueError("配置缺少必要区段: " + ", ".join(missing))

    page = config.get("page", {})
    if not isinstance(page, dict):
        raise ValueError("page 必须是对象")
    if "size" in page and str(page["size"]).strip().lower() not in PAGE_SIZES:
        raise ValueError("page.size 只能是 A4 或 Letter，当前为: " + str(page["size"]))
    margins = page.get("margins")
    if not isinstance(margins, dict):
        raise ValueError("配置缺少 page.margins")
    for key in ("top", "bottom", "left", "right"):
        if key not in margins:
            raise ValueError("配置缺少 page.margins." + key)
    for key in ("top", "bottom", "left", "right", "header", "footer"):
        _check_number(margins, key, "page.margins." + key, minimum=0, maximum=15)

    body = config.get("body", {})
    if not isinstance(body, dict):
        raise ValueError("body 必须是对象")
    for key in ("font", "asciiFont", "size"):
        if key not in body:
            raise ValueError("配置缺少 body." + key)
    _check_size_value(body, "size", "body.size")
    _check_number(body, "lineSpacing", "body.lineSpacing", minimum=0.5, maximum=5)
    _check_number(body, "firstLineIndentChars", "body.firstLineIndentChars", minimum=0, maximum=10)
    _check_number(body, "spaceAfterPt", "body.spaceAfterPt", minimum=0, maximum=100)
    _check_align(body, "align", "body.align")

    headings = config.get("headings", {})
    if not isinstance(headings, dict):
        raise ValueError("headings 必须是对象")
    for key in ("h1", "h2", "h3", "h4"):
        if key not in headings:
            raise ValueError("配置缺少 headings." + key)
        item = headings[key] or {}
        if not isinstance(item, dict):
            raise ValueError("headings." + key + " 必须是对象")
        _check_size_value(item, "size", "headings." + key + ".size")
        _check_number(item, "lineSpacing", "headings." + key + ".lineSpacing", minimum=0.5, maximum=5)
        _check_number(item, "spaceBeforePt", "headings." + key + ".spaceBeforePt", minimum=0, maximum=200)
        _check_number(item, "spaceAfterPt", "headings." + key + ".spaceAfterPt", minimum=0, maximum=200)
        _check_number(item, "spaceBeforeLines", "headings." + key + ".spaceBeforeLines", minimum=0, maximum=5)
        _check_number(item, "spaceAfterLines", "headings." + key + ".spaceAfterLines", minimum=0, maximum=5)
        _check_align(item, "align", "headings." + key + ".align")

    figure = config.get("figure", {})
    if not isinstance(figure, dict):
        raise ValueError("figure 必须是对象")
    size_mode = figure.get("sizeMode")
    if size_mode is not None and str(size_mode).lower() not in FIGURE_MODES:
        raise ValueError(
            "figure.sizeMode 只能是 fit（等比缩放）或 fixed（统一尺寸），当前为: " + str(size_mode))
    for key in ("maxWidthCm", "maxHeightCm", "fixedWidthCm", "fixedHeightCm"):
        if key in figure and figure[key] is not None:
            _check_number(figure, key, "figure." + key, minimum=0, maximum=40)
    _check_align(figure, "align", "figure.align")
    _check_align(figure, "captionAlign", "figure.captionAlign")
    _check_align(figure, "noteAlign", "figure.noteAlign")
    if "captionPosition" in figure and str(figure["captionPosition"]).lower() not in CAPTION_POSITIONS:
        raise ValueError("figure.captionPosition 只能是 above 或 below，当前为: " + str(figure["captionPosition"]))
    _check_size_value(figure, "captionSize", "figure.captionSize")
    _check_size_value(figure, "noteSize", "figure.noteSize")

    table = config.get("table", {})
    if not isinstance(table, dict):
        raise ValueError("table 必须是对象")
    _check_align(table, "align", "table.align")
    _check_align(table, "captionAlign", "table.captionAlign")
    _check_align(table, "noteAlign", "table.noteAlign")
    _check_cell_align(table, "cellAlign", "table.cellAlign")
    _check_number(table, "smartAlignThresholdChars", "table.smartAlignThresholdChars", minimum=0, maximum=50)
    if "captionPosition" in table and str(table["captionPosition"]).lower() not in CAPTION_POSITIONS:
        raise ValueError("table.captionPosition 只能是 above 或 below，当前为: " + str(table["captionPosition"]))
    _check_size_value(table, "captionSize", "table.captionSize")
    _check_size_value(table, "noteSize", "table.noteSize")
    _check_size_value(table, "cellSize", "table.cellSize")
    _check_number(table, "borderSize", "table.borderSize", minimum=0, maximum=100)

    references = config.get("references", {})
    if not isinstance(references, dict):
        raise ValueError("references 必须是对象")
    _check_size_value(references, "size", "references.size")
    _check_number(references, "lineSpacing", "references.lineSpacing", minimum=0.5, maximum=5)
    _check_number(references, "hangingIndentPt", "references.hangingIndentPt", minimum=0, maximum=200)
    _check_align(references, "align", "references.align")

    page_number = config.get("pageNumber", {})
    if not isinstance(page_number, dict):
        raise ValueError("pageNumber 必须是对象")
    _check_align(page_number, "align", "pageNumber.align")
    _check_size_value(page_number, "size", "pageNumber.size")

    header = config.get("header")
    if header is not None:
        if not isinstance(header, dict):
            raise ValueError("header 必须是对象")
        _check_align(header, "align", "header.align")
        _check_size_value(header, "size", "header.size")

    return True


# ----- 排版预设（config/templates/*.json） -----

def list_presets():
    """扫描模板目录，返回可用预设列表（校验失败的模板跳过，不阻断启动）。

    每项：{"name", "description", "path", "config"}；name 取 _preset.name，
    缺省用文件名（去扩展名）。
    """
    directory = templates_dir()
    presets = []
    if not os.path.isdir(directory):
        return presets
    for path in sorted(Path(directory).glob("*.json")):
        try:
            cfg = load_config_file(path)
        except (OSError, ValueError):
            continue
        meta = cfg.get("_preset") or {}
        presets.append({
            "name": str(meta.get("name") or path.stem),
            "description": str(meta.get("description") or ""),
            "path": str(path),
            "config": cfg,
        })
    return presets


def load_preset(name):
    """按名称载入预设（含已校验的完整配置），找不到抛 ValueError。"""
    for preset in list_presets():
        if preset["name"] == name:
            return preset
    raise ValueError("找不到排版预设: " + str(name))
