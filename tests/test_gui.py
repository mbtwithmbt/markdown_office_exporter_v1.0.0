from __future__ import annotations

from pathlib import Path
import tempfile
import unittest

from md_office_exporter.gui import suggested_output_path, validate_export_request


class GuiRequestHelperTests(unittest.TestCase):
    def setUp(self) -> None:
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.markdown = self.root / "report.md"
        self.markdown.write_text("# Report\n", encoding="utf-8")
        self.config = self.root / "format.yaml"
        self.config.write_text("docx: {}\n", encoding="utf-8")

    def tearDown(self) -> None:
        self.temp_dir.cleanup()

    def test_suggests_extension_for_export_type(self) -> None:
        self.assertEqual(
            self.root / "report.docx",
            suggested_output_path(self.markdown, "docx"),
        )
        self.assertEqual(
            self.root / "custom.xlsx",
            suggested_output_path(self.root / "custom.docx", "excel"),
        )

    def test_rejects_unknown_type_when_suggesting_path(self) -> None:
        with self.assertRaisesRegex(ValueError, "导出类型"):
            suggested_output_path(self.markdown, "pdf")

    def test_accepts_valid_request_and_optional_config(self) -> None:
        self.assertEqual(
            (self.markdown, self.root / "report.docx", None),
            validate_export_request(
                self.markdown,
                self.root / "report.docx",
                "docx",
            ),
        )
        self.assertEqual(
            (self.markdown, self.root / "report.xlsx", self.config),
            validate_export_request(
                self.markdown,
                self.root / "report.xlsx",
                "excel",
                self.config,
            ),
        )

    def test_rejects_empty_markdown_path(self) -> None:
        with self.assertRaisesRegex(ValueError, "请选择 Markdown 文件"):
            validate_export_request("  ", self.root / "report.docx", "docx")

    def test_rejects_missing_markdown_file(self) -> None:
        with self.assertRaisesRegex(ValueError, "Markdown 文件不存在"):
            validate_export_request(
                self.root / "missing.md",
                self.root / "report.docx",
                "docx",
            )

    def test_rejects_empty_output_path(self) -> None:
        with self.assertRaisesRegex(ValueError, "请选择输出文件"):
            validate_export_request(self.markdown, "  ", "docx")

    def test_rejects_missing_selected_config_file(self) -> None:
        with self.assertRaisesRegex(ValueError, "配置文件不存在"):
            validate_export_request(
                self.markdown,
                self.root / "report.docx",
                "docx",
                self.root / "missing.yaml",
            )

    def test_rejects_wrong_output_extension(self) -> None:
        with self.assertRaisesRegex(ValueError, r"\.xlsx"):
            validate_export_request(
                self.markdown,
                self.root / "report.docx",
                "excel",
            )

    def test_rejects_unsupported_export_type(self) -> None:
        with self.assertRaisesRegex(ValueError, "导出类型"):
            validate_export_request(
                self.markdown,
                self.root / "report.pdf",
                "pdf",
            )


if __name__ == "__main__":
    unittest.main()
