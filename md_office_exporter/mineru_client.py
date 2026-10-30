"""Small HTTP client for a local or remote MinerU service."""

from __future__ import annotations

import os
from collections.abc import Mapping
from pathlib import Path
from typing import Any

import requests


DEFAULT_MINERU_BASE_URL = os.getenv("MINERU_BASE_URL", "http://127.0.0.1:8000")
SUPPORTED_BACKENDS = {
    "pipeline",
    "vlm-engine",
    "hybrid-engine",
    "vlm-http-client",
    "hybrid-http-client",
}


class MinerUError(RuntimeError):
    """Raised when MinerU cannot be reached or returns an unusable response."""


class MinerUClient:
    def __init__(
        self,
        base_url: str = DEFAULT_MINERU_BASE_URL,
        timeout: float = 300,
        session: requests.Session | Any | None = None,
    ) -> None:
        if not isinstance(base_url, str) or not base_url.strip():
            raise MinerUError("MinerU 服务地址不能为空")
        if timeout <= 0:
            raise MinerUError("MinerU 超时时间必须大于 0")

        self.base_url = base_url.rstrip("/")
        self.timeout = timeout
        self._owns_session = session is None
        self.session = requests.Session() if self._owns_session else session

    def close(self) -> None:
        if self._owns_session:
            self.session.close()

    def __enter__(self) -> "MinerUClient":
        return self

    def __exit__(self, exc_type: object, exc_value: object, traceback: object) -> None:
        self.close()

    def health_check(self) -> dict[str, Any]:
        try:
            response = self.session.get(f"{self.base_url}/health", timeout=self.timeout)
            response.raise_for_status()
            return self._response_object(response, "健康检查")
        except requests.RequestException as exc:
            raise MinerUError(f"MinerU 健康检查请求失败: {exc}") from exc

    def parse_file(self, path: str | Path, backend: str = "hybrid-engine") -> dict[str, str]:
        source = Path(path)
        if not source.is_file():
            raise MinerUError(f"要解析的文件不存在: {source}")
        if backend not in SUPPORTED_BACKENDS:
            raise MinerUError(f"不支持的 MinerU 后端: {backend}")

        try:
            with source.open("rb") as stream:
                response = self.session.post(
                    f"{self.base_url}/file_parse",
                    files={"files": (source.name, stream, "application/octet-stream")},
                    data={
                        "return_md": "true",
                        "response_format_zip": "false",
                        "return_original_file": "false",
                        "backend": backend,
                    },
                    timeout=self.timeout,
                )
                response.raise_for_status()
                payload = self._response_object(response, "文件解析")
        except requests.RequestException as exc:
            raise MinerUError(f"MinerU 文件解析请求失败: {exc}") from exc
        except OSError as exc:
            raise MinerUError(f"MinerU 文件读取失败: {exc}") from exc

        status = payload.get("status")
        if status != "completed":
            error = payload.get("error")
            raise MinerUError(f"MinerU 文件解析未完成 (status={status!r}, error={error!r})")

        results = payload.get("results")
        if not isinstance(results, Mapping):
            raise MinerUError("MinerU 文件解析响应中的 results 无效")

        markdown: dict[str, str] = {}
        for name, result in results.items():
            if isinstance(result, Mapping):
                content = result.get("md_content")
                if isinstance(content, str) and content.strip():
                    markdown[str(name)] = content
        if not markdown:
            raise MinerUError("MinerU 响应中没有可用的 md_content")
        return markdown

    @staticmethod
    def _response_object(response: Any, context: str) -> dict[str, Any]:
        try:
            payload = response.json()
        except (ValueError, requests.RequestException) as exc:
            raise MinerUError(f"MinerU {context}响应不是有效 JSON: {exc}") from exc
        if not isinstance(payload, dict):
            raise MinerUError(f"MinerU {context}响应必须是 JSON 对象")
        return payload
