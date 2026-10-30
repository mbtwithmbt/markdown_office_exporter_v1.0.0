import tempfile
import unittest
from pathlib import Path

import yaml

from md_office_exporter.config import DEFAULT_CONFIG, load_config
from md_office_exporter.storage import RepositoryStorage
from md_office_exporter.templates import (
    SEED_TEMPLATES,
    ensure_seed_templates,
    list_templates,
    save_template,
    template_file,
    template_yaml_text,
    validate_template_bytes,
    delete_template,
)


class TemplateTests(unittest.TestCase):
    def setUp(self):
        self.temp_dir = tempfile.TemporaryDirectory()
        self.storage = RepositoryStorage(Path(self.temp_dir.name))

    def tearDown(self):
        self.temp_dir.cleanup()

    def test_ensure_seed_templates_writes_each_seed_once(self):
        created = ensure_seed_templates(self.storage)
        self.assertEqual(
            sorted(path.name for path in created),
            sorted(f"{name}.yaml" for name in SEED_TEMPLATES),
        )
        self.assertEqual([], ensure_seed_templates(self.storage))

    def test_ensure_seed_templates_does_not_resurrect_deleted_files(self):
        ensure_seed_templates(self.storage)
        delete_template(self.storage, "默认模板")
        self.assertEqual([], ensure_seed_templates(self.storage))
        self.assertNotIn("默认模板", list_templates(self.storage))

    def test_ensure_seed_templates_never_overwrites_user_edits(self):
        path = self.storage.config_repo / "默认模板.yaml"
        path.write_text("docx:\n  body:\n    font: Custom\n", encoding="utf-8")
        ensure_seed_templates(self.storage)
        self.assertIn(b"font: Custom", path.read_bytes())

    def test_seed_template_is_a_valid_full_config(self):
        ensure_seed_templates(self.storage)
        path = template_file(self.storage, "默认模板")
        loaded = yaml.safe_load(path.read_text(encoding="utf-8"))
        self.assertEqual(DEFAULT_CONFIG, load_config(overrides=loaded))

    def test_list_templates_returns_sorted_stems(self):
        ensure_seed_templates(self.storage)
        save_template(self.storage, "aaa", b"docx: {}")
        self.assertEqual(
            ["aaa", "正式公文", "清爽绿表格", "默认模板"],
            list_templates(self.storage),
        )

    def test_template_file_resolves_by_stem(self):
        path = save_template(self.storage, "我的模板", b"excel: {}")
        self.assertEqual(path, template_file(self.storage, "我的模板"))
        with self.assertRaises(FileNotFoundError):
            template_file(self.storage, "不存在")

    def test_save_template_persists_yaml_and_is_usable_by_loaders(self):
        save_template(self.storage, "my", b"docx:\n  body:\n    font: SimSun\n")
        config = load_config(config_path=template_file(self.storage, "my"))
        self.assertEqual("SimSun", config["docx"]["body"]["font"])
        self.assertEqual(DEFAULT_CONFIG["docx"]["body"]["font_size_pt"],
                         config["docx"]["body"]["font_size_pt"])

    def test_validate_template_bytes_rejects_broken_inputs(self):
        cases = {
            "non-utf8": b"\xff\xfe\x00docx",
            "bad yaml": b"docx: [unclosed",
            "empty file": b"",
            "null root": b"",
            "non mapping": b"- docx\n- excel\n",
            "unknown key": b"pdf: {}",
            "invalid enum": b"excel:\n  merge_mode: both\n",
        }
        cases["null root"] = b"# only a comment\n"
        for name, payload in cases.items():
            with self.subTest(name=name):
                with self.assertRaises(ValueError):
                    validate_template_bytes(payload)

    def test_validate_template_bytes_accepts_partial_overrides(self):
        parsed = validate_template_bytes(
            b"docx:\n  body:\n    font: SimSun\nexcel:\n  row_height: 30\n"
        )
        self.assertEqual({"docx": {"body": {"font": "SimSun"}},
                          "excel": {"row_height": 30}}, parsed)

    def test_delete_template_removes_file(self):
        ensure_seed_templates(self.storage)
        delete_template(self.storage, "清爽绿表格")
        self.assertNotIn("清爽绿表格", list_templates(self.storage))

    def test_template_yaml_text_round_trips_config(self):
        text = template_yaml_text(DEFAULT_CONFIG)
        self.assertEqual(DEFAULT_CONFIG, yaml.safe_load(text))


if __name__ == "__main__":
    unittest.main()
