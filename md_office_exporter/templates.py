"""Preset template management for DOCX/Excel export configuration files.

Templates are plain YAML files stored in ``config_repo``. Each template may
override any subset of the built-in configuration (see :mod:`.config`); the
exporters deep-merge the template over the defaults. Built-in seed templates
are written once, on first use, and afterwards behave like user uploads: they
can be downloaded, replaced or deleted.
"""

from __future__ import annotations

from pathlib import Path
from typing import Any, Mapping

import yaml

from .config import DEFAULT_CONFIG, load_config
from .storage import RepositoryStorage


SEED_MARKER = ".seeded"

BUILTIN_TEMPLATE_LABEL = "内置默认"

SEED_TEMPLATES: dict[str, dict[str, Any]] = {
    "默认模板": {},
    "正式公文": {
        "docx": {
            "body": {
                "font": "SimSun",
                "font_size_pt": 12,
                "color": "000000",
                "alignment": "justify",
                "line_spacing": 1.5,
            },
            "headings": {
                "1": {
                    "font": "SimHei",
                    "font_size_pt": 22,
                    "color": "000000",
                    "bold": True,
                    "space_before_pt": 18,
                    "space_after_pt": 12,
                },
                "2": {
                    "font": "SimHei",
                    "font_size_pt": 16,
                    "color": "000000",
                    "bold": True,
                    "space_before_pt": 14,
                    "space_after_pt": 8,
                },
                "3": {
                    "font": "SimHei",
                    "font_size_pt": 14,
                    "color": "000000",
                    "bold": True,
                    "space_before_pt": 12,
                    "space_after_pt": 6,
                },
            },
            "table": {
                "font": "SimSun",
                "header_fill": "F2F2F2",
                "header_font_color": "000000",
            },
            "header": {"text": ""},
            "footer": {"text": "", "show_page_number": True},
        },
    },
    "清爽绿表格": {
        "excel": {
            "header": {"fill": "548235", "font_color": "FFFFFF", "bold": True},
            "border_color": "C6E0B4",
            "freeze_header": True,
            "auto_filter": True,
        },
    },
}

ALLOWED_TOP_LEVEL_KEYS = {"docx", "excel"}


def template_yaml_text(config: Mapping[str, Any]) -> str:
    """Serialize a config mapping to readable YAML text."""

    return yaml.safe_dump(
        dict(config),
        allow_unicode=True,
        sort_keys=True,
        default_flow_style=False,
    )


def validate_template_bytes(content: bytes) -> dict[str, Any]:
    """Parse and validate an uploaded template file, returning its mapping.

    Raises ``ValueError`` with a Chinese message when the file is not usable.
    """

    try:
        text = content.decode("utf-8")
    except UnicodeDecodeError as exc:
        raise ValueError("模板文件必须是 UTF-8 编码的 YAML 文本") from exc
    try:
        loaded = yaml.safe_load(text)
    except yaml.YAMLError as exc:
        raise ValueError(f"YAML 解析失败：{exc}") from exc
    if loaded is None:
        raise ValueError("模板文件为空")
    if not isinstance(loaded, Mapping):
        raise ValueError("模板根节点必须是对象/映射")
    unknown = [key for key in loaded if key not in ALLOWED_TOP_LEVEL_KEYS]
    if unknown:
        raise ValueError(
            "模板只支持顶层键 docx / excel，未知键：" + "、".join(map(str, unknown))
        )
    try:
        load_config(overrides=dict(loaded))
    except ValueError as exc:
        raise ValueError(f"模板配置校验失败：{exc}") from exc
    return dict(loaded)


def ensure_seed_templates(storage: RepositoryStorage) -> list[Path]:
    """Write the built-in seed templates once; never overwrite existing files."""

    marker = storage.config_repo / SEED_MARKER
    if marker.exists():
        return []
    created: list[Path] = []
    for name, overrides in SEED_TEMPLATES.items():
        path = storage.config_repo / f"{name}.yaml"
        if path.exists():
            continue
        config = load_config(overrides=overrides) if overrides else DEFAULT_CONFIG
        path.write_text(template_yaml_text(config), encoding="utf-8")
        created.append(path)
    marker.write_text("", encoding="utf-8")
    return created


def list_templates(storage: RepositoryStorage) -> list[str]:
    """Return template display names (file stems) in case-insensitive order."""

    names = [path.stem for path in storage.list_repo_files("config")]
    return sorted(names, key=str.casefold)


def template_file(storage: RepositoryStorage, name: str) -> Path:
    """Resolve a template display name to its YAML file."""

    for path in storage.list_repo_files("config"):
        if path.stem == name:
            return path
    raise FileNotFoundError(f"模板不存在：{name}")


def save_template(storage: RepositoryStorage, name: str, content: bytes) -> Path:
    """Validate and persist an uploaded template under the given name."""

    validate_template_bytes(content)
    return storage.save_config(name, content)


def delete_template(storage: RepositoryStorage, name: str) -> None:
    """Delete one template by display name."""

    template_file(storage, name).unlink()
