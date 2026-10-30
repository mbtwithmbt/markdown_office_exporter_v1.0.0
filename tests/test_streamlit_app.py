import os
from pathlib import Path
from tempfile import TemporaryDirectory
import unittest
from unittest.mock import patch

from streamlit.testing.v1 import AppTest

from md_office_exporter.markdown_processor import OperationResult


PROJECT_ROOT = Path(__file__).resolve().parents[1]
APP_PATH = PROJECT_ROOT / 'streamlit_app.py'
PARSE_LABEL = '\u5f00\u59cb\u89e3\u6790'
EXPORT_LABEL = '\u6279\u91cf\u5bfc\u51fa'
FILE_SELECT_LABEL = '选择文件（用于打包下载或批量删除）'
DELETE_LABEL = '批量删除选中文件'


class TrackingClient:
    instances = []

    def __init__(self, base_url, timeout):
        self.base_url = base_url
        self.timeout = timeout
        self.closed = False
        self.__class__.instances.append(self)

    def __enter__(self):
        return self

    def __exit__(self, exc_type, exc_value, traceback):
        self.closed = True


class StreamlitAppTests(unittest.TestCase):
    def setUp(self) -> None:
        self.data_root = TemporaryDirectory()
        self.previous_data_root = os.environ.get('MARKDOWN_OFFICE_DATA_ROOT')
        os.environ['MARKDOWN_OFFICE_DATA_ROOT'] = self.data_root.name
        markdown_repo = Path(self.data_root.name) / 'markdown_repo'
        markdown_repo.mkdir(parents=True)
        (markdown_repo / 'sample.md').write_text('# Sample', encoding='utf-8')
        self.app = AppTest.from_file(APP_PATH)

    def tearDown(self) -> None:
        if self.previous_data_root is None:
            os.environ.pop('MARKDOWN_OFFICE_DATA_ROOT', None)
        else:
            os.environ['MARKDOWN_OFFICE_DATA_ROOT'] = self.previous_data_root
        self.data_root.cleanup()

    def _button(self, label):
        return next(button for button in self.app.button if button.label == label)

    def _buttons(self, label):
        return [button for button in self.app.button if button.label == label]

    def _widget(self, kind, label):
        return next(widget for widget in getattr(self.app, kind) if widget.label == label)

    def _widget_by_key(self, kind, key):
        return next(widget for widget in getattr(self.app, kind) if widget.key == key)

    def test_default_page_exposes_pipeline_controls_without_errors(self) -> None:
        self.app.run(timeout=10)

        self.assertFalse(self.app.exception)
        self.assertGreaterEqual(len(self.app.file_uploader), 1)
        self.assertGreaterEqual(len(self.app.button), 2)
        self.assertIn('MinerU URL', [widget.label for widget in self.app.text_input])
        self.assertIn('\u5065\u5eb7\u68c0\u67e5', [widget.label for widget in self.app.button])
        self.assertIn('Markdown \u6587\u4ef6', [widget.label for widget in self.app.multiselect])
        self.assertIn('\u5bfc\u51fa\u7c7b\u578b', [widget.label for widget in self.app.selectbox])
        self.assertIn(EXPORT_LABEL, [widget.label for widget in self.app.button])

    def test_directory_mode_enables_directory_upload(self) -> None:
        self.app.run(timeout=10)
        self.app.radio[0].set_value('\u6587\u4ef6\u5939').run(timeout=10)

        self.assertFalse(self.app.exception)
        self.assertTrue(self.app.file_uploader[0].accept_directory)

    def test_multi_file_mode_enables_multiple_uploads(self) -> None:
        self.app.run(timeout=10)
        self.app.radio[0].set_value('\u591a\u6587\u4ef6').run(timeout=10)

        self.assertFalse(self.app.exception)
        self.assertTrue(self.app.file_uploader[0].accept_multiple_files)
        self.assertFalse(self.app.file_uploader[0].accept_directory)

    def test_parse_runs_one_file_at_a_time_and_persists_partial_results(self) -> None:
        parse_calls = []
        TrackingClient.instances.clear()

        def fake_parse(paths, client, storage, backend):
            path_list = list(paths)
            parse_calls.append([path.name for path in path_list])
            path = path_list[0]
            try:
                if path.name == 'bad.pdf':
                    return [OperationResult(path.name, False, error='parse failed')]
                output = storage.save_markdown(path.name, '# Parsed')
                return [OperationResult(path.name, True, [output])]
            finally:
                storage.remove_temp(path)

        with (
            patch('md_office_exporter.mineru_client.MinerUClient', TrackingClient),
            patch(
                'md_office_exporter.markdown_processor.parse_uploaded_files',
                side_effect=fake_parse,
            ),
        ):
            self.app.run(timeout=10)
            self.app.radio[0].set_value('\u591a\u6587\u4ef6').run(timeout=10)
            self.app.file_uploader[0].set_value(
                [
                    ('bad.pdf', b'bad', 'application/pdf'),
                    ('good.pdf', b'good', 'application/pdf'),
                ]
            ).run(timeout=10)
            self._button(PARSE_LABEL).click().run(timeout=10)

        self.assertEqual([['bad.pdf'], ['good.pdf']], parse_calls)
        self.assertEqual(1, len(TrackingClient.instances))
        self.assertTrue(TrackingClient.instances[0].closed)
        frame = self.app.dataframe[0].value
        self.assertEqual(['error', 'success'], frame['status'].tolist())
        self.assertEqual(['parse failed', ''], frame['error'].tolist())

        self.app.run(timeout=10)
        persisted = self.app.dataframe[0].value
        self.assertEqual(['bad.pdf', 'good.pdf'], persisted['source'].tolist())
        self.assertEqual(['error', 'success'], persisted['status'].tolist())

    def test_export_download_and_results_persist_across_rerun(self) -> None:
        export_calls = []

        def fake_export(
            paths,
            storage,
            export_type,
            config_path=None,
            merge_mode='merge',
            tables_only=False,
        ):
            markdown_paths = list(paths)
            export_calls.append(
                (
                    [path.name for path in markdown_paths],
                    export_type,
                    merge_mode,
                    config_path,
                    tables_only,
                )
            )
            output = storage.output_path(markdown_paths[0], export_type)
            output.write_bytes(b'fake docx')
            return [OperationResult(markdown_paths[0].name, True, [output])]

        with patch(
            'md_office_exporter.markdown_processor.export_markdown_files',
            side_effect=fake_export,
        ):
            self.app.run(timeout=10)
            self.app.multiselect[0].set_value(['sample.md']).run(timeout=10)
            self._button(EXPORT_LABEL).click().run(timeout=10)

        self.assertEqual(
            [(['sample.md'], 'docx', 'merge', None, False)], export_calls
        )
        export_downloads = [
            button
            for button in self.app.download_button
            if button.label == 'Download sample.docx'
        ]
        self.assertEqual(1, len(export_downloads))
        self.assertEqual('success', self.app.dataframe[0].value['status'].iloc[0])

        self.app.run(timeout=10)
        export_downloads = [
            button
            for button in self.app.download_button
            if button.label == 'Download sample.docx'
        ]
        self.assertEqual(1, len(export_downloads))
        self.assertEqual('sample.md', self.app.dataframe[0].value['source'].iloc[0])

    def test_export_with_preset_template_and_tables_only(self) -> None:
        export_calls = []

        def fake_export(
            paths,
            storage,
            export_type,
            config_path=None,
            merge_mode='merge',
            tables_only=False,
        ):
            markdown_paths = list(paths)
            export_calls.append((export_type, config_path, tables_only))
            output = storage.output_path(markdown_paths[0], export_type)
            output.write_bytes(b'fake docx')
            return [OperationResult(markdown_paths[0].name, True, [output])]

        with patch(
            'md_office_exporter.markdown_processor.export_markdown_files',
            side_effect=fake_export,
        ):
            self.app.run(timeout=10)
            self._widget('selectbox', '预设模板（docx / excel 导出共用）').set_value(
                '正式公文'
            ).run(timeout=10)
            self._widget('radio', 'DOCX 导出范围').set_value('仅表格').run(timeout=10)
            self.app.multiselect[0].set_value(['sample.md']).run(timeout=10)
            self._button(EXPORT_LABEL).click().run(timeout=10)

        template_path = (
            Path(self.data_root.name) / 'config_repo' / '正式公文.yaml'
        )
        self.assertEqual(
            [('docx', template_path, True)], export_calls
        )
        download_labels = [button.label for button in self.app.download_button]
        self.assertIn('下载模板 正式公文.yaml', download_labels)

    def test_file_manager_lists_markdown_files_with_download_buttons(self) -> None:
        self.app.run(timeout=10)

        repo_labels = [radio.label for radio in self.app.radio]
        self.assertIn('选择仓库', repo_labels)
        frame = self.app.dataframe[-1].value
        self.assertIn('sample.md', frame['文件名'].tolist())
        download_labels = [button.label for button in self.app.download_button]
        self.assertIn('下载 sample.md', download_labels)

    def test_file_manager_uploads_markdown_files(self) -> None:
        self.app.run(timeout=10)

        uploader = self._widget('file_uploader', '上传 Markdown 文件到 markdown_repo')
        uploader.set_value([('added.md', '# 新增'.encode('utf-8'), 'text/markdown')]).run(timeout=10)

        self.assertFalse(self.app.exception)
        uploaded = Path(self.data_root.name) / 'markdown_repo' / 'added.md'
        self.assertTrue(uploaded.exists())
        self.assertEqual('# 新增', uploaded.read_text(encoding='utf-8'))

    def test_file_manager_zip_download_and_batch_delete(self) -> None:
        (Path(self.data_root.name) / 'markdown_repo' / 'second.md').write_text(
            '# Second', encoding='utf-8'
        )
        self.app.run(timeout=10)

        self._widget('multiselect', FILE_SELECT_LABEL).set_value(
            ['sample.md', 'second.md']
        ).run(timeout=10)

        zip_button = next(
            button
            for button in self.app.download_button
            if button.label == '打包下载 2 个文件（ZIP）'
        )
        self.assertEqual('repo_zip_markdown', zip_button.key)

        self._button(DELETE_LABEL).click().run(timeout=10)
        self.assertFalse(
            (Path(self.data_root.name) / 'markdown_repo' / 'sample.md').exists()
        )
        self.assertFalse(
            (Path(self.data_root.name) / 'markdown_repo' / 'second.md').exists()
        )
        self.assertEqual(
            [], self._widget('multiselect', FILE_SELECT_LABEL).value
        )

    def test_template_upload_in_files_tab_saves_preset(self) -> None:
        self.app.run(timeout=10)

        self._widget('radio', '选择仓库').set_value('模板仓库（config_repo）').run(timeout=10)
        self._widget_by_key('file_uploader', 'files_template_file').set_value(
            ('custom.yaml', b'docx:\n  body:\n    font: KaiTi\n', 'text/yaml')
        ).run(timeout=10)
        self._widget_by_key('text_input', 'files_template_name').set_value(
            '楷体正文'
        ).run(timeout=10)
        self._buttons('保存为预设模板')[-1].click().run(timeout=10)

        saved = Path(self.data_root.name) / 'config_repo' / '楷体正文.yaml'
        self.assertTrue(saved.exists())
        self.app.run(timeout=10)
        template_options = self._widget(
            'selectbox', '预设模板（docx / excel 导出共用）'
        ).options
        self.assertIn('楷体正文', template_options)

    def test_duplicate_upload_names_are_saved_and_processed_one_at_a_time(self) -> None:
        observed_payloads = []
        TrackingClient.instances.clear()

        def fake_parse(paths, client, storage, backend):
            path = list(paths)[0]
            observed_payloads.append(path.read_bytes())
            try:
                output = storage.save_markdown(
                    f'{len(observed_payloads)}-{path.name}',
                    '# Parsed',
                )
                return [OperationResult(path.name, True, [output])]
            finally:
                storage.remove_temp(path)

        with (
            patch('md_office_exporter.mineru_client.MinerUClient', TrackingClient),
            patch(
                'md_office_exporter.markdown_processor.parse_uploaded_files',
                side_effect=fake_parse,
            ),
        ):
            self.app.run(timeout=10)
            self.app.radio[0].set_value('\u591a\u6587\u4ef6').run(timeout=10)
            self.app.file_uploader[0].set_value(
                [
                    ('duplicate.pdf', b'first payload', 'application/pdf'),
                    ('duplicate.pdf', b'second payload', 'application/pdf'),
                ]
            ).run(timeout=10)
            self._button(PARSE_LABEL).click().run(timeout=10)

        self.assertEqual([b'first payload', b'second payload'], observed_payloads)
        self.assertEqual(1, len(TrackingClient.instances))
        self.assertTrue(TrackingClient.instances[0].closed)
        self.assertEqual(
            ['success', 'success'],
            self.app.dataframe[0].value['status'].tolist(),
        )

    def test_unexpected_processor_cleanup_failure_reports_orphan(self) -> None:
        def cleanup_failure(storage, path):
            raise OSError('cleanup denied')

        def unexpected_parse(paths, client, storage, backend):
            raise RuntimeError('processor exploded')

        with (
            patch(
                'md_office_exporter.mineru_client.MinerUClient',
                TrackingClient,
            ),
            patch(
                'md_office_exporter.markdown_processor.parse_uploaded_files',
                side_effect=unexpected_parse,
            ),
            patch(
                'md_office_exporter.storage.RepositoryStorage.remove_temp',
                new=cleanup_failure,
            ),
        ):
            self.app.run(timeout=10)
            self.app.file_uploader[0].set_value(
                ('orphan.pdf', b'content', 'application/pdf')
            ).run(timeout=10)
            self._button(PARSE_LABEL).click().run(timeout=10)

        warning_text = '\n'.join(warning.value for warning in self.app.warning)
        self.assertIn('orphan.pdf', warning_text)
        self.assertIn('cleanup denied', warning_text)


if __name__ == '__main__':
    unittest.main()
