from __future__ import annotations

from pathlib import Path
import unittest
from unittest.mock import patch

from md_office_exporter.desktop_gui import perform_export


class ExportDispatchTests(unittest.TestCase):
    def setUp(self) -> None:
        self.markdown = Path("report.md")
        self.config = Path("format.yaml")

    def test_dispatches_docx_export(self) -> None:
        output = Path("report.docx")
        with patch(
            "md_office_exporter.desktop_gui.convert_markdown_to_docx"
        ) as export:
            export.return_value = output
            result = perform_export(
                self.markdown,
                output,
                "docx",
                config_path=self.config,
            )

        self.assertEqual(output, result)
        export.assert_called_once_with(self.markdown, output, self.config)

    def test_dispatches_excel_export_with_merge_mode(self) -> None:
        output = Path("report.xlsx")
        with patch(
            "md_office_exporter.desktop_gui.export_markdown_tables_to_excel"
        ) as export:
            export.return_value = output
            result = perform_export(
                self.markdown,
                output,
                "excel",
                merge_mode="expand",
            )

        self.assertEqual(output, result)
        export.assert_called_once_with(
            self.markdown,
            output,
            merge_mode="expand",
            config_path=None,
        )

    def test_rejects_invalid_excel_merge_mode_before_export(self) -> None:
        with patch(
            "md_office_exporter.desktop_gui.export_markdown_tables_to_excel"
        ) as export:
            with self.assertRaisesRegex(ValueError, "Excel 合并方式"):
                perform_export(
                    self.markdown,
                    Path("report.xlsx"),
                    "excel",
                    merge_mode="invalid",
                )

        export.assert_not_called()

    def test_rejects_unknown_export_type(self) -> None:
        with self.assertRaisesRegex(ValueError, "导出类型"):
            perform_export(
                self.markdown,
                Path("report.pdf"),
                "pdf",
            )


if __name__ == "__main__":
    unittest.main()
