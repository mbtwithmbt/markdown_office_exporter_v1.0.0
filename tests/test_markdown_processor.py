from pathlib import Path
from unittest.mock import Mock, call, patch

import pytest

from md_office_exporter.storage import RepositoryStorage


def test_parse_uploaded_files_processes_inputs_in_order_and_saves_every_markdown(tmp_path):
    from md_office_exporter.markdown_processor import parse_uploaded_files

    storage = RepositoryStorage(tmp_path / "storage")
    first = storage.save_upload("first.pdf", b"first")
    second = storage.save_upload("second.pdf", b"second")
    client = Mock()
    client.parse_file.side_effect = [
        {"first.md": "# First"},
        {"second.md": "# Second", "appendix.md": "# Appendix"},
    ]

    results = parse_uploaded_files([first, second], client, storage)

    assert [result.source for result in results] == ["first.pdf", "second.pdf"]
    assert [result.success for result in results] == [True, True]
    assert [path.name for path in results[0].outputs] == ["first.md"]
    assert [path.name for path in results[1].outputs] == ["second.md", "appendix.md"]
    assert (storage.markdown_repo / "first.md").read_text(encoding="utf-8") == "# First"
    assert (storage.markdown_repo / "appendix.md").read_text(encoding="utf-8") == "# Appendix"
    assert client.parse_file.call_args_list == [
        call(first, "hybrid-engine"),
        call(second, "hybrid-engine"),
    ]


def test_parse_uploaded_files_continues_after_parse_failure_and_removes_all_temp_files(tmp_path):
    from md_office_exporter.markdown_processor import parse_uploaded_files

    storage = RepositoryStorage(tmp_path / "storage")
    first = storage.save_upload("broken.pdf", b"broken")
    second = storage.save_upload("working.pdf", b"working")
    client = Mock()
    client.parse_file.side_effect = [RuntimeError("service unavailable"), {"working.md": "ok"}]

    results = parse_uploaded_files([first, second], client, storage, backend="pipeline")

    assert [result.success for result in results] == [False, True]
    assert results[0].error == "service unavailable"
    assert results[1].outputs[0].name == "working.md"
    assert not first.exists()
    assert not second.exists()
    assert client.parse_file.call_args_list == [call(first, "pipeline"), call(second, "pipeline")]


def test_parse_uploaded_files_reports_cleanup_failure_and_continues(tmp_path):
    from md_office_exporter.markdown_processor import parse_uploaded_files

    storage = RepositoryStorage(tmp_path / "storage")
    first = storage.save_upload("first.pdf", b"first")
    second = storage.save_upload("second.pdf", b"second")
    client = Mock()
    client.parse_file.side_effect = [{"first.md": "one"}, {"second.md": "two"}]
    original_remove_temp = storage.remove_temp

    def remove_temp(path):
        if Path(path) == first:
            raise OSError("locked")
        original_remove_temp(path)

    storage.remove_temp = remove_temp

    results = parse_uploaded_files([first, second], client, storage)

    assert [result.success for result in results] == [False, True]
    assert [path.name for path in results[0].outputs] == ["first.md"]
    assert "cleanup failed: locked" in results[0].error
    assert not second.exists()


def test_parse_uploaded_files_keeps_outputs_saved_before_a_later_save_failure(tmp_path):
    from md_office_exporter.markdown_processor import parse_uploaded_files

    storage = RepositoryStorage(tmp_path / "storage")
    uploaded = storage.save_upload("report.pdf", b"report")
    client = Mock()
    client.parse_file.return_value = {"first.md": "first", "second.md": "second"}
    original_save_markdown = storage.save_markdown

    def save_markdown(name, text):
        if name == "second.md":
            raise OSError("disk full")
        return original_save_markdown(name, text)

    storage.save_markdown = save_markdown

    results = parse_uploaded_files([uploaded], client, storage)

    assert results[0].success is False
    assert results[0].error == "disk full"
    assert [path.name for path in results[0].outputs] == ["first.md"]


def test_parse_uploaded_files_preserves_primary_base_exception_when_cleanup_fails(tmp_path):
    from md_office_exporter.markdown_processor import parse_uploaded_files

    storage = RepositoryStorage(tmp_path / "storage")
    uploaded = storage.save_upload("interrupt.pdf", b"interrupt")
    client = Mock()
    client.parse_file.side_effect = KeyboardInterrupt()
    storage.remove_temp = Mock(side_effect=OSError("locked"))

    with pytest.raises(KeyboardInterrupt):
        parse_uploaded_files([uploaded], client, storage)


def test_operation_results_have_independent_output_lists(tmp_path):
    from md_office_exporter.markdown_processor import OperationResult

    first = OperationResult("first", True)
    second = OperationResult("second", True)
    first.outputs.append(tmp_path / "one.md")

    assert second.outputs == []


def test_export_markdown_files_routes_docx_and_forwards_config_path(tmp_path):
    from md_office_exporter.markdown_processor import export_markdown_files

    storage = RepositoryStorage(tmp_path / "storage")
    markdown = storage.save_markdown("report.md", "# Report")
    config_path = tmp_path / "settings.yaml"
    expected_output = storage.output_path(markdown, "docx")

    with patch("md_office_exporter.markdown_processor.convert_markdown_to_docx") as converter:
        results = export_markdown_files([markdown], storage, "docx", config_path=config_path)

    assert results[0].success is True
    assert results[0].outputs == [expected_output]
    converter.assert_called_once_with(
        markdown, expected_output, config_path, tables_only=False
    )


def test_export_markdown_files_forwards_docx_tables_only(tmp_path):
    from md_office_exporter.markdown_processor import export_markdown_files

    storage = RepositoryStorage(tmp_path / "storage")
    markdown = storage.save_markdown("report.md", "# Report")
    expected_output = storage.output_path(markdown, "docx")

    with patch("md_office_exporter.markdown_processor.convert_markdown_to_docx") as converter:
        results = export_markdown_files(
            [markdown], storage, "docx", tables_only=True
        )

    assert results[0].success is True
    converter.assert_called_once_with(
        markdown, expected_output, None, tables_only=True
    )


@pytest.mark.parametrize("merge_mode", ["merge", "expand"])
def test_export_markdown_files_routes_excel_with_sequential_sheet_config(tmp_path, merge_mode):
    from md_office_exporter.markdown_processor import export_markdown_files

    storage = RepositoryStorage(tmp_path / "storage")
    markdown = storage.save_markdown("table.md", "| A |\n| - |\n| 1 |")
    config_path = tmp_path / "settings.yaml"
    expected_output = storage.output_path(markdown, "excel")

    with patch("md_office_exporter.markdown_processor.export_markdown_tables_to_excel") as exporter:
        results = export_markdown_files(
            [markdown], storage, "excel", config_path=config_path, merge_mode=merge_mode
        )

    assert results[0].success is True
    assert results[0].outputs == [expected_output]
    exporter.assert_called_once_with(
        markdown,
        expected_output,
        merge_mode=merge_mode,
        config_path=config_path,
        config={"excel": {"sheet_name_mode": "sequential"}},
    )


def test_export_markdown_files_continues_after_first_export_failure(tmp_path):
    from md_office_exporter.markdown_processor import export_markdown_files

    storage = RepositoryStorage(tmp_path / "storage")
    first = storage.save_markdown("first.md", "# First")
    second = storage.save_markdown("second.md", "# Second")

    with patch(
        "md_office_exporter.markdown_processor.convert_markdown_to_docx",
        side_effect=[RuntimeError("bad markdown"), None],
    ) as converter:
        results = export_markdown_files([first, second], storage, "docx")

    assert [result.success for result in results] == [False, True]
    assert results[0].error == "bad markdown"
    assert results[1].outputs == [storage.output_path(second, "docx")]
    assert converter.call_count == 2


def test_export_markdown_files_rejects_invalid_type_before_processing(tmp_path):
    from md_office_exporter.markdown_processor import export_markdown_files

    storage = RepositoryStorage(tmp_path / "storage")

    with pytest.raises(ValueError, match="export_type"):
        export_markdown_files([], storage, "pdf")


def test_export_markdown_files_rejects_invalid_excel_merge_mode_before_processing(tmp_path):
    from md_office_exporter.markdown_processor import export_markdown_files

    storage = RepositoryStorage(tmp_path / "storage")

    with pytest.raises(ValueError, match="merge_mode"):
        export_markdown_files([], storage, "excel", merge_mode="invalid")
