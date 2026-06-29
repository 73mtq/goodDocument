"""Configuration loading, saving, and validation."""

import json
from pathlib import Path


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

    margins = config.get("page", {}).get("margins")
    if not isinstance(margins, dict):
        raise ValueError("配置缺少 page.margins")
    for key in ("top", "bottom", "left", "right"):
        if key not in margins:
            raise ValueError("配置缺少 page.margins." + key)

    for key in ("font", "asciiFont", "size"):
        if key not in config.get("body", {}):
            raise ValueError("配置缺少 body." + key)

    headings = config.get("headings", {})
    for key in ("h1", "h2", "h3", "h4"):
        if key not in headings:
            raise ValueError("配置缺少 headings." + key)

    return True
