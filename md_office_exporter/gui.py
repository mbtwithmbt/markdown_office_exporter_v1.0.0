"""Lightweight desktop interface helpers for Markdown Office exports."""

from __future__ import annotations

from pathlib import Path


EXPORT_EXTENSIONS = {"docx": ".docx", "excel": ".xlsx"}


def suggested_output_path(source: str | Path, export_type: str) -> Path:
    """Return *source* with the extension required by *export_type*."""

    if export_type not in EXPORT_EXTENSIONS:
        raise ValueError("导出类型只能是 docx 或 excel")
    return Path(source).with_suffix(EXPORT_EXTENSIONS[export_type])


def validate_export_request(
    markdown_path: str | Path,
    output_path: str | Path,
    export_type: str,
    config_path: str | Path | None = None,
) -> tuple[Path, Path, Path | None]:
    """Validate GUI fields and return normalized paths."""

    if export_type not in EXPORT_EXTENSIONS:
        raise ValueError("导出类型只能是 docx 或 excel")
    if not str(markdown_path).strip():
        raise ValueError("请选择 Markdown 文件")
    source = Path(markdown_path)
    if not source.is_file():
        raise ValueError(f"Markdown 文件不存在：{source}")
    if not str(output_path).strip():
        raise ValueError("请选择输出文件")
    output = Path(output_path)
    expected_suffix = EXPORT_EXTENSIONS[export_type]
    if output.suffix.lower() != expected_suffix:
        raise ValueError(f"输出文件扩展名必须是 {expected_suffix}")

    resolved_config = None
    if config_path is not None and str(config_path).strip():
        resolved_config = Path(config_path)
        if not resolved_config.is_file():
            raise ValueError(f"配置文件不存在：{resolved_config}")
    return source, output, resolved_config
