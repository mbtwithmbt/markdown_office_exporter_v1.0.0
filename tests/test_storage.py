import tempfile
import unittest
from io import BytesIO
from pathlib import Path
from zipfile import ZipFile

from md_office_exporter.storage import RepositoryStorage, safe_filename


class RepositoryStorageTests(unittest.TestCase):
    def test_safe_filename_rejects_windows_reserved_device_names(self):
        names = (
            "CON", "prN", "Aux.txt", "NUL.md", "com1", "CoM9.log",
            "lpt1", "LpT9.csv",
        )
        for name in names:
            with self.subTest(name=name):
                with self.assertRaises(ValueError):
                    safe_filename(name)

    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.storage = RepositoryStorage(self.root)

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_constructor_creates_repository_directories(self):
        self.assertTrue(self.storage.markdown_repo.is_dir())
        self.assertTrue(self.storage.docx_repo.is_dir())
        self.assertTrue(self.storage.excel_repo.is_dir())
        self.assertTrue(self.storage.config_repo.is_dir())
        self.assertTrue(self.storage.temp_uploads.is_dir())

    def test_safe_filename_discards_posix_traversal(self):
        self.assertEqual(safe_filename("../report.pdf"), "report.pdf")

    def test_safe_filename_discards_windows_traversal(self):
        self.assertEqual(safe_filename(r"folder\report.pdf"), "report.pdf")

    def test_safe_filename_replaces_invalid_characters_and_trims_suffix(self):
        name = safe_filename('bad:name\x01. ')
        self.assertNotRegex(name, r'[<>:"/\\|?*\x00-\x1f]')
        self.assertFalse(name.endswith((".", " ")))

    def test_safe_filename_rejects_blank_and_dot_names(self):
        for name in ("", "   ", ".", ".."):
            with self.subTest(name=name):
                with self.assertRaises(ValueError):
                    safe_filename(name)

    def test_save_upload_overwrites_existing_upload(self):
        result = self.storage.save_upload("folder\\report.pdf", b"first")
        overwritten = self.storage.save_upload("report.pdf", b"second")
        self.assertEqual(result, self.storage.temp_uploads / "report.pdf")
        self.assertEqual(overwritten, result)
        self.assertEqual(result.read_bytes(), b"second")

    def test_save_markdown_writes_utf8_and_overwrites(self):
        result = self.storage.save_markdown("notes.txt", "第一版")
        overwritten = self.storage.save_markdown("notes.md", "第二版\n")
        self.assertEqual(result, self.storage.markdown_repo / "notes.md")
        self.assertEqual(overwritten, result)
        self.assertEqual(result.read_text(encoding="utf-8"), "第二版\n")

    def test_list_markdown_in_case_insensitive_name_order(self):
        self.storage.save_markdown("zeta", "z")
        self.storage.save_markdown("Alpha", "a")
        (self.storage.markdown_repo / "ignore.txt").write_text("x", encoding="utf-8")
        self.assertEqual(
            [path.name for path in self.storage.list_markdown()],
            ["Alpha.md", "zeta.md"],
        )

    def test_output_path_routes_to_requested_export_repository(self):
        markdown = self.storage.markdown_repo / "report.md"
        self.assertEqual(self.storage.output_path(markdown, "docx"), self.storage.docx_repo / "report.docx")
        self.assertEqual(self.storage.output_path(markdown, "excel"), self.storage.excel_repo / "report.xlsx")

    def test_output_path_rejects_unknown_export_type_in_chinese(self):
        with self.assertRaisesRegex(ValueError, "不支持"):
            self.storage.output_path(self.storage.markdown_repo / "report.md", "pdf")

    def test_remove_temp_deletes_only_files_inside_upload_directory(self):
        upload = self.storage.save_upload("delete-me.txt", b"data")
        self.storage.remove_temp(upload)
        self.storage.remove_temp(upload)
        self.assertFalse(upload.exists())
        outside = self.root / "outside.txt"
        outside.write_bytes(b"keep")
        with self.assertRaises(ValueError):
            self.storage.remove_temp(outside)
        self.assertTrue(outside.exists())


class RepositoryFileManagementTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.root = Path(self.temp_dir.name)
        self.storage = RepositoryStorage(self.root)
        self.storage.save_markdown("alpha", "# A")
        self.storage.save_markdown("Beta", "# B")
        (self.storage.markdown_repo / "ignored.txt").write_text("x", encoding="utf-8")
        (self.storage.docx_repo / "report.docx").write_bytes(b"docx")
        (self.storage.excel_repo / "report.xlsx").write_bytes(b"xlsx")
        self.storage.save_config("正式公文", b"docx: {}")

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_list_repo_files_filters_by_repo_suffix(self):
        self.assertEqual(
            [path.name for path in self.storage.list_repo_files("markdown")],
            ["alpha.md", "Beta.md"],
        )
        self.assertEqual(
            [path.name for path in self.storage.list_repo_files("docx")],
            ["report.docx"],
        )
        self.assertEqual(
            [path.name for path in self.storage.list_repo_files("excel")],
            ["report.xlsx"],
        )
        self.assertEqual(
            [path.name for path in self.storage.list_repo_files("config")],
            ["正式公文.yaml"],
        )

    def test_list_repo_files_rejects_unknown_repo(self):
        with self.assertRaisesRegex(ValueError, "未知"):
            self.storage.list_repo_files("pdf")

    def test_repo_file_resolves_existing_file_and_rejects_traversal(self):
        path = self.storage.repo_file("docx", "report.docx")
        self.assertEqual(path, self.storage.docx_repo / "report.docx")
        for bad in ("../config_repo/正式公文.yaml", "nested/../../report.docx", "ghost.docx"):
            with self.subTest(bad=bad):
                with self.assertRaises((ValueError, FileNotFoundError)):
                    self.storage.repo_file("docx", bad)

    def test_upload_markdown_saves_bytes_with_md_suffix(self):
        path = self.storage.upload_markdown("notes.txt", "# 备注\n".encode("utf-8"))
        self.assertEqual(path, self.storage.markdown_repo / "notes.md")
        self.assertEqual(path.read_text(encoding="utf-8"), "# 备注\n")

    def test_save_config_forces_yaml_suffix_and_overwrites(self):
        first = self.storage.save_config("mine.json", b"docx: {}")
        second = self.storage.save_config("mine", b"docx: {font: X}")
        self.assertEqual(first, self.storage.config_repo / "mine.yaml")
        self.assertEqual(second, first)
        self.assertEqual(b"docx: {font: X}", first.read_bytes())

    def test_delete_repo_files_reports_mixed_outcomes(self):
        results = self.storage.delete_repo_files(
            "markdown", ["alpha.md", "missing.md"]
        )
        self.assertEqual(
            [("alpha.md", True, "已删除"), ("missing.md", False, "文件不存在：missing.md")],
            results,
        )
        self.assertFalse((self.storage.markdown_repo / "alpha.md").exists())
        self.assertTrue((self.storage.markdown_repo / "Beta.md").exists())

    def test_delete_repo_files_rejects_traversal_names(self):
        results = self.storage.delete_repo_files("markdown", ["../outside.md"])
        self.assertFalse(results[0][1])
        self.assertIn("非法", results[0][2])

    def test_zip_repo_files_packs_requested_files(self):
        data = self.storage.zip_repo_files("markdown", ["alpha.md", "Beta.md"])
        with ZipFile(BytesIO(data)) as archive:
            self.assertEqual(["alpha.md", "Beta.md"], archive.namelist())
            self.assertEqual(b"# A", archive.read("alpha.md"))

    def test_zip_repo_files_rejects_unknown_member(self):
        with self.assertRaises(FileNotFoundError):
            self.storage.zip_repo_files("markdown", ["missing.md"])


if __name__ == "__main__":
    unittest.main()
