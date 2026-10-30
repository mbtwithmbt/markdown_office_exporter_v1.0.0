"""Configuration loading and validation."""

from __future__ import annotations

from copy import deepcopy
from pathlib import Path
from typing import Any, Mapping

import yaml


DEFAULT_CONFIG: dict[str, Any] = {
    "docx": {
        "page": {
            "size": "A4",
            "orientation": "portrait",
            "margins_cm": {
                "top": 2.54,
                "right": 2.54,
                "bottom": 2.54,
                "left": 2.54,
            },
        },
        "body": {
            "font": "Microsoft YaHei",
            "font_size_pt": 11,
            "color": "000000",
            "alignment": "left",
            "space_before_pt": 0,
            "space_after_pt": 6,
            "line_spacing": 1.15,
        },
        "headings": {
            "1": {
                "font": "Microsoft YaHei",
                "font_size_pt": 20,
                "color": "1F4E79",
                "bold": True,
                "space_before_pt": 16,
                "space_after_pt": 8,
            },
            "2": {
                "font": "Microsoft YaHei",
                "font_size_pt": 16,
                "color": "2F5597",
                "bold": True,
                "space_before_pt": 12,
                "space_after_pt": 6,
            },
            "3": {
                "font": "Microsoft YaHei",
                "font_size_pt": 13,
                "color": "1F4E79",
                "bold": True,
                "space_before_pt": 10,
                "space_after_pt": 4,
            },
        },
        "lists": {
            "left_indent_cm": 0.74,
            "hanging_cm": 0.32,
            "space_after_pt": 3,
        },
        "table": {
            "style": "Table Grid",
            "font": "Microsoft YaHei",
            "font_size_pt": 10,
            "indent_dxa": 120,
            "header_fill": "D9EAF7",
            "header_font_color": "1F1F1F",
            "border_color": "A6A6A6",
            "cell_margin_dxa": {
                "top": 80,
                "right": 120,
                "bottom": 80,
                "left": 120,
            },
            "vertical_alignment": "center",
            "repeat_header": True,
        },
        "code": {
            "font": "Consolas",
            "font_size_pt": 9,
            "fill": "F3F4F6",
            "color": "1F2937",
        },
        "blockquote": {
            "left_indent_cm": 0.74,
            "color": "666666",
            "italic": True,
        },
        "header": {"text": "", "font_size_pt": 9, "color": "808080"},
        "footer": {
            "text": "",
            "show_page_number": False,
            "font_size_pt": 9,
            "color": "808080",
        },
    },
    "excel": {
        "merge_mode": "merge",
        "sheet_name_mode": "title",
        "font": "Microsoft YaHei",
        "font_size_pt": 10,
        "header": {
            "fill": "4472C4",
            "font_color": "FFFFFF",
            "bold": True,
        },
        "body": {"font_color": "000000"},
        "border_color": "D9E2F3",
        "alignment": {
            "horizontal": "left",
            "vertical": "center",
            "wrap_text": True,
        },
        "freeze_header": True,
        "auto_filter": True,
        "hide_gridlines": True,
        "min_column_width": 10,
        "max_column_width": 45,
        "row_height": 22,
    },
}


def deep_merge(base: Mapping[str, Any], override: Mapping[str, Any]) -> dict[str, Any]:
    """Recursively merge *override* into *base* and return a new dictionary."""

    result = deepcopy(dict(base))
    for key, value in override.items():
        if (
            key in result
            and isinstance(result[key], Mapping)
            and isinstance(value, Mapping)
        ):
            result[key] = deep_merge(result[key], value)
        else:
            result[key] = deepcopy(value)
    return result


def load_config(
    config_path: str | Path | None = None,
    overrides: Mapping[str, Any] | None = None,
) -> dict[str, Any]:
    """Load YAML/JSON-compatible configuration over the built-in defaults."""

    config = deepcopy(DEFAULT_CONFIG)
    if config_path:
        path = Path(config_path)
        if not path.is_file():
            raise FileNotFoundError(f"配置文件不存在: {path}")
        with path.open("r", encoding="utf-8") as stream:
            loaded = yaml.safe_load(stream) or {}
        if not isinstance(loaded, Mapping):
            raise ValueError("配置文件根节点必须是对象/映射")
        config = deep_merge(config, loaded)
    if overrides:
        config = deep_merge(config, overrides)

    merge_mode = config["excel"]["merge_mode"]
    if merge_mode not in {"merge", "expand"}:
        raise ValueError("excel.merge_mode 只能是 'merge' 或 'expand'")
    sheet_name_mode = config["excel"]["sheet_name_mode"]
    if sheet_name_mode not in {"title", "sequential"}:
        raise ValueError("excel.sheet_name_mode 只能是 'title' 或 'sequential'")
    return config
