import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from md_office_exporter import storage
from md_office_exporter.storage import RepositoryStorage, safe_filename


class StorageAtomicTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.storage = RepositoryStorage(Path(self.temp_dir.name))

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_reserved_windows_device_names_are_rejected(self):
        for name in ("CON", "prN", "Aux.txt", "NUL.md", "NUL.tar.gz", "COM1.backup.zip", "COM¹.log", "lPt².backup", "LPT³.csv", "com1", "CoM9.log", "lpt1", "LpT9.csv"):
            with self.subTest(name=name):
                with self.assertRaises(ValueError):
                    safe_filename(name)

    def test_output_path_rejects_derived_windows_device_name(self):
        with self.assertRaises(ValueError):
            self.storage.output_path(self.storage.markdown_repo / "CON .md", "docx")
        with self.assertRaises(ValueError):
            self.storage.output_path(self.storage.markdown_repo / "NUL.tar.gz.foo", "docx")
        with self.assertRaises(ValueError):
            self.storage.output_path(self.storage.markdown_repo / "COM¹.backup.zip", "docx")

    def test_upload_replace_failure_preserves_target_and_cleans_temp(self):
        target = self.storage.save_upload("report.pdf", b"old upload")
        with patch.object(storage.os, "replace", side_effect=OSError("replace failed")):
            with self.assertRaises(OSError):
                self.storage.save_upload("report.pdf", b"new upload")
        self.assertEqual(target.read_bytes(), b"old upload")
        self.assertEqual(list(self.storage.temp_uploads.glob(".report.pdf.*.tmp")), [])

    def test_upload_pre_replace_failure_preserves_target_and_cleans_temp(self):
        target = self.storage.save_upload("report.pdf", b"old upload")
        with patch.object(storage.os, "fsync", side_effect=OSError("write preparation failed")):
            with self.assertRaises(OSError):
                self.storage.save_upload("report.pdf", b"new upload")
        self.assertEqual(target.read_bytes(), b"old upload")
        self.assertEqual(list(self.storage.temp_uploads.glob(".report.pdf.*.tmp")), [])

    def test_markdown_pre_replace_failure_preserves_target_and_cleans_temp(self):
        target = self.storage.save_markdown("notes.txt", "old markdown")
        with patch.object(storage.os, "fsync", side_effect=OSError("write preparation failed")):
            with self.assertRaises(OSError):
                self.storage.save_markdown("notes.md", "new markdown")
        self.assertEqual(target.read_text(encoding="utf-8"), "old markdown")
        self.assertEqual(list(self.storage.markdown_repo.glob(".notes.md.*.tmp")), [])

    def test_markdown_replace_failure_preserves_target_and_cleans_temp(self):
        target = self.storage.save_markdown("notes.txt", "old markdown")
        with patch.object(storage.os, "replace", side_effect=OSError("replace failed")):
            with self.assertRaises(OSError):
                self.storage.save_markdown("notes.md", "new markdown")
        self.assertEqual(target.read_text(encoding="utf-8"), "old markdown")
        self.assertEqual(list(self.storage.markdown_repo.glob(".notes.md.*.tmp")), [])
