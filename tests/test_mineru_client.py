"""Tests for the MinerU HTTP client."""

from __future__ import annotations

import tempfile
import unittest
from pathlib import Path
from unittest.mock import Mock, patch

import requests

from md_office_exporter.mineru_client import (
    MinerUClient,
    MinerUError,
    SUPPORTED_BACKENDS,
)


class MinerUClientTests(unittest.TestCase):
    def setUp(self) -> None:
        self.session = Mock()
        self.client = MinerUClient("http://mineru.test/api///", timeout=12, session=self.session)

    def response(self, payload: object) -> Mock:
        result = Mock()
        result.json.return_value = payload
        return result

    def sample_file(self) -> tempfile.TemporaryDirectory[str]:
        temp_dir = tempfile.TemporaryDirectory()
        Path(temp_dir.name, "sample.pdf").write_bytes(b"PDF")
        return temp_dir

    def test_constructor_strips_url_and_validates_inputs(self) -> None:
        self.assertEqual(self.client.base_url, "http://mineru.test/api")
        self.assertEqual(self.client.timeout, 12)
        self.assertEqual(
            SUPPORTED_BACKENDS,
            {"pipeline", "vlm-engine", "hybrid-engine", "vlm-http-client", "hybrid-http-client"},
        )
        with self.assertRaises(MinerUError):
            MinerUClient("   ", session=self.session)
        with self.assertRaises(MinerUError):
            MinerUClient("http://mineru.test", timeout=0, session=self.session)

    def test_health_check_uses_expected_request_and_returns_object(self) -> None:
        payload = {"status": "ok"}
        self.session.get.return_value = self.response(payload)

        self.assertEqual(self.client.health_check(), payload)
        self.session.get.assert_called_once_with("http://mineru.test/api/health", timeout=12)
        self.session.get.return_value.raise_for_status.assert_called_once_with()

    def test_parse_file_posts_expected_payload_and_closes_stream(self) -> None:
        self.session.post.return_value = self.response(
            {"status": "completed", "results": {"one": {"md_content": "# One"}, "two": {"md_content": "  "}, "three": {"md_content": "# Three"}}}
        )
        with self.sample_file() as folder:
            source = Path(folder, "sample.pdf")
            self.assertEqual(self.client.parse_file(source), {"one": "# One", "three": "# Three"})

        self.session.post.assert_called_once()
        kwargs = self.session.post.call_args.kwargs
        self.assertEqual(self.session.post.call_args.args, ("http://mineru.test/api/file_parse",))
        self.assertEqual(kwargs["files"]["files"][0], "sample.pdf")
        self.assertEqual(kwargs["files"]["files"][2], "application/octet-stream")
        self.assertTrue(kwargs["files"]["files"][1].closed)
        self.assertEqual(kwargs["data"], {"return_md": "true", "response_format_zip": "false", "return_original_file": "false", "backend": "hybrid-engine"})
        self.assertEqual(kwargs["timeout"], 12)

    def test_parse_file_rejects_missing_file_and_unknown_backend(self) -> None:
        with self.assertRaises(MinerUError):
            self.client.parse_file(Path("does-not-exist.pdf"))
        with self.sample_file() as folder:
            with self.assertRaises(MinerUError):
                self.client.parse_file(Path(folder, "sample.pdf"), backend="unknown")
        self.session.post.assert_not_called()

    def test_health_check_wraps_request_and_response_errors(self) -> None:
        for failure in (requests.ConnectionError("offline"), requests.HTTPError("bad status")):
            with self.subTest(failure=failure):
                self.session.reset_mock()
                self.session.get.side_effect = None
                if isinstance(failure, requests.HTTPError):
                    response = self.response({"status": "ok"})
                    response.raise_for_status.side_effect = failure
                    self.session.get.return_value = response
                else:
                    self.session.get.side_effect = failure
                with self.assertRaisesRegex(MinerUError, "健康检查") as caught:
                    self.client.health_check()
                self.assertIs(caught.exception.__cause__, failure)

        for payload in (ValueError("bad json"), [], None):
            with self.subTest(payload=payload):
                self.session.reset_mock()
                self.session.get.side_effect = None
                response = self.response({})
                if isinstance(payload, ValueError):
                    response.json.side_effect = payload
                else:
                    response.json.return_value = payload
                self.session.get.return_value = response
                with self.assertRaisesRegex(MinerUError, "健康检查"):
                    self.client.health_check()

    def test_parse_file_wraps_request_http_and_json_errors(self) -> None:
        cases = [requests.ConnectionError("offline"), requests.HTTPError("bad status"), ValueError("bad json"), []]
        with self.sample_file() as folder:
            source = Path(folder, "sample.pdf")
            for failure in cases:
                with self.subTest(failure=failure):
                    self.session.reset_mock()
                    self.session.post.side_effect = None
                    response = self.response({"status": "completed", "results": {}})
                    if isinstance(failure, requests.RequestException):
                        if isinstance(failure, requests.HTTPError):
                            response.raise_for_status.side_effect = failure
                            self.session.post.return_value = response
                        else:
                            self.session.post.side_effect = failure
                    elif isinstance(failure, ValueError):
                        response.json.side_effect = failure
                        self.session.post.return_value = response
                    else:
                        response.json.return_value = failure
                        self.session.post.return_value = response
                    with self.assertRaisesRegex(MinerUError, "解析") as caught:
                        self.client.parse_file(source)
                    if isinstance(failure, requests.RequestException):
                        self.assertIs(caught.exception.__cause__, failure)

    def test_parse_file_rejects_unsuccessful_or_invalid_results(self) -> None:
        invalid_payloads = [
            {"status": "failed", "error": "boom"},
            {"status": "completed"},
            {"status": "completed", "results": []},
            {"status": "completed", "results": {"one": {}}},
            {"status": "completed", "results": {"one": {"md_content": ""}}},
        ]
        with self.sample_file() as folder:
            source = Path(folder, "sample.pdf")
            for payload in invalid_payloads:
                with self.subTest(payload=payload):
                    self.session.reset_mock()
                    self.session.post.side_effect = None
                    self.session.post.return_value = self.response(payload)
                    expected = "md_content" if payload.get("status") == "completed" and isinstance(payload.get("results"), dict) else "解析"
                    with self.assertRaisesRegex(MinerUError, expected):
                        self.client.parse_file(source)

    @patch("md_office_exporter.mineru_client.requests.Session")
    def test_close_and_context_manager_close_only_owned_session(self, session_factory: Mock) -> None:
        owned = Mock()
        session_factory.return_value = owned
        client = MinerUClient()
        client.close()
        owned.close.assert_called_once_with()

        owned.reset_mock()
        with MinerUClient() as contextual_client:
            self.assertIs(contextual_client.session, owned)
        owned.close.assert_called_once_with()

    def test_close_does_not_close_injected_session_and_retains_falsey_session(self) -> None:
        class FalseySession:
            def __bool__(self) -> bool:
                return False

        injected = FalseySession()
        client = MinerUClient(session=injected)
        self.assertIs(client.session, injected)
        client.close()

    def test_parse_file_wraps_file_read_os_error(self) -> None:
        with self.sample_file() as folder:
            source = Path(folder, "sample.pdf")
            with patch.object(Path, "open", side_effect=OSError("unreadable")) as open_file:
                with self.assertRaisesRegex(MinerUError, "文件读取") as caught:
                    self.client.parse_file(source)
            open_file.assert_called_once_with("rb")
            self.assertIsInstance(caught.exception.__cause__, OSError)
        self.session.post.assert_not_called()


if __name__ == "__main__":
    unittest.main()
