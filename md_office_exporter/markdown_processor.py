"""Sequential orchestration for MinerU parsing and Office exports."""

from __future__ import annotations

from collections.abc import Iterable, Mapping
from dataclasses import dataclass, field
from pathlib import Path
from typing import Literal, Protocol

from .docx_exporter import convert_markdown_to_docx
from .excel_exporter import export_markdown_tables_to_excel


ExportType = Literal["docx", "excel"]
MergeMode = Literal["merge", "expand"]


class MinerUParser(Protocol):
    def parse_file(self, path: Path, backend: str = "hybrid-engine") -> Mapping[str, str]: ...


class ProcessorStorage(Protocol):
    def save_markdown(self, source_name: str, text: str) -> Path: ...

    def remove_temp(self, path: Path) -> None: ...

    def output_path(self, markdown: Path, export_type: ExportType) -> Path: ...


@dataclass(slots=True)
class OperationResult:
    """Outcome of processing one uploaded or Markdown file."""

    source: str
    success: bool
    outputs: list[Path] = field(default_factory=list)
    error: str = ""


def parse_uploaded_files(
    paths: Iterable[Path],
    client: MinerUParser,
    storage: ProcessorStorage,
    backend: str = "hybrid-engine",
) -> list[OperationResult]:
    """Parse each temporary upload, saving returned Markdown independently."""

    results: list[OperationResult] = []
    for path in paths:
        source = Path(path)
        outputs: list[Path] = []
        result = OperationResult(source.name, False, outputs)
        try:
            parsed = client.parse_file(source, backend)
            for name, text in parsed.items():
                outputs.append(storage.save_markdown(name, text))
            result.success = True
        except Exception as exc:
            result.error = str(exc)
        finally:
            try:
                storage.remove_temp(source)
            except Exception as cleanup_error:
                cleanup_message = f"cleanup failed: {cleanup_error}"
                result.success = False
                result.error = (
                    f"{result.error}; {cleanup_message}"
                    if result.error
                    else cleanup_message
                )
        results.append(result)
    return results

def export_markdown_files(
    paths: Iterable[Path],
    storage: ProcessorStorage,
    export_type: ExportType,
    config_path: str | Path | None = None,
    merge_mode: MergeMode = "merge",
    tables_only: bool = False,
) -> list[OperationResult]:
    """Export Markdown files sequentially, isolating individual failures."""

    if export_type not in {"docx", "excel"}:
        raise ValueError("export_type must be 'docx' or 'excel'")
    if export_type == "excel" and merge_mode not in {"merge", "expand"}:
        raise ValueError("merge_mode must be 'merge' or 'expand'")

    results: list[OperationResult] = []
    for path in paths:
        markdown = Path(path)
        try:
            output = storage.output_path(markdown, export_type)
            if export_type == "docx":
                convert_markdown_to_docx(
                    markdown,
                    output,
                    config_path,
                    tables_only=tables_only,
                )
            else:
                export_markdown_tables_to_excel(
                    markdown,
                    output,
                    merge_mode=merge_mode,
                    config_path=config_path,
                    config={"excel": {"sheet_name_mode": "sequential"}},
                )
            results.append(OperationResult(markdown.name, True, [output]))
        except Exception as exc:
            results.append(OperationResult(markdown.name, False, error=str(exc)))
    return results
