"""Filesystem storage for uploaded Markdown documents and exported files."""

from io import BytesIO
from pathlib import Path
from zipfile import ZIP_DEFLATED, ZipFile
import os
import re
import tempfile


INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
WINDOWS_RESERVED_NAMES = {'con', 'prn', 'aux', 'nul', *(f'com{number}' for number in range(1, 10)), *(f'lpt{number}' for number in range(1, 10))}

REPO_KEYS = ("markdown", "docx", "excel", "config")
_REPO_FILE_SUFFIXES = {
    "markdown": {".md", ".markdown"},
    "docx": {".docx"},
    "excel": {".xlsx"},
    "config": {".yaml", ".yml"},
}


def safe_filename(name: str) -> str:
    """Return a single safe file name, rejecting empty or dot-only names."""
    filename = re.split(r"[\\/]", str(name))[-1]
    filename = INVALID_FILENAME_CHARS.sub("_", filename).rstrip(". ")
    if not filename or filename in {".", ".."}:
        raise ValueError("文件名不能为空")
    if filename.split(".", 1)[0].translate(str.maketrans("¹²³", "123")).casefold() in WINDOWS_RESERVED_NAMES:
        raise ValueError('Windows 保留设备名不能用作文件名')
    return filename


def _atomic_write(path: Path, content: bytes) -> None:
    temporary_path = None
    try:
        with tempfile.NamedTemporaryFile(mode='wb', dir=path.parent, prefix=f'.{path.name}.', suffix='.tmp', delete=False) as temporary_file:
            temporary_path = Path(temporary_file.name)
            temporary_file.write(content)
            temporary_file.flush()
            os.fsync(temporary_file.fileno())
        os.replace(temporary_path, path)
    except BaseException:
        if temporary_path is not None:
            temporary_path.unlink(missing_ok=True)
        raise


class RepositoryStorage:
    """Manage the application's on-disk repositories beneath one root directory."""

    def __init__(self, root: str | Path):
        self.root = Path(root).resolve()
        self.markdown_repo = self.root / "markdown_repo"
        self.docx_repo = self.root / "docx_repo"
        self.excel_repo = self.root / "excel_repo"
        self.config_repo = self.root / "config_repo"
        self.temp_uploads = self.root / "temp_uploads"
        self._repos = {
            "markdown": self.markdown_repo,
            "docx": self.docx_repo,
            "excel": self.excel_repo,
            "config": self.config_repo,
        }
        for directory in (
            self.markdown_repo,
            self.docx_repo,
            self.excel_repo,
            self.config_repo,
            self.temp_uploads,
        ):
            directory.mkdir(parents=True, exist_ok=True)

    def repo_path(self, repo_key: str) -> Path:
        """Return the directory backing *repo_key* ('markdown'/'docx'/'excel'/'config')."""

        if repo_key not in self._repos:
            raise ValueError(f"未知的仓库类型：{repo_key}")
        return self._repos[repo_key]

    def save_upload(self, name: str, content: bytes) -> Path:
        path = self.temp_uploads / safe_filename(name)
        _atomic_write(path, content)
        return path

    def save_markdown(self, source_name: str, text: str) -> Path:
        filename = Path(safe_filename(source_name)).with_suffix(".md").name
        path = self.markdown_repo / filename
        _atomic_write(path, text.encode("utf-8"))
        return path

    def upload_markdown(self, name: str, content: bytes) -> Path:
        """Store an uploaded Markdown file (bytes) directly in markdown_repo."""

        return self.save_markdown(name, content.decode("utf-8"))

    def save_config(self, name: str, content: bytes) -> Path:
        """Store an uploaded template config file in config_repo with a YAML suffix."""

        filename = Path(safe_filename(name)).with_suffix(".yaml").name
        path = self.config_repo / filename
        _atomic_write(path, content)
        return path

    def list_markdown(self) -> list[Path]:
        return sorted(self.markdown_repo.glob("*.md"), key=lambda path: path.name.casefold())

    def list_repo_files(self, repo_key: str) -> list[Path]:
        """List the managed files of one repository, case-insensitively by name."""

        repo = self.repo_path(repo_key)
        suffixes = _REPO_FILE_SUFFIXES[repo_key]
        files = [
            path
            for path in repo.iterdir()
            if path.is_file() and path.suffix.casefold() in suffixes
        ]
        return sorted(files, key=lambda path: path.name.casefold())

    def repo_file(self, repo_key: str, filename: str) -> Path:
        """Resolve a single file inside a repository, rejecting traversal."""

        repo = self.repo_path(repo_key)
        name = safe_filename(filename)
        if name != filename:
            raise ValueError(f"非法文件名：{filename}")
        path = (repo / name).resolve()
        if path.parent != repo.resolve():
            raise ValueError(f"非法文件名：{filename}")
        if not path.is_file():
            raise FileNotFoundError(f"文件不存在：{filename}")
        return path

    def delete_repo_files(self, repo_key: str, filenames: list[str]) -> list[tuple[str, bool, str]]:
        """Delete files from one repository, returning (name, ok, message) per file."""

        results: list[tuple[str, bool, str]] = []
        for filename in filenames:
            try:
                path = self.repo_file(repo_key, filename)
                path.unlink()
            except (ValueError, FileNotFoundError, OSError) as exc:
                results.append((filename, False, str(exc)))
            else:
                results.append((filename, True, "已删除"))
        return results

    def zip_repo_files(self, repo_key: str, filenames: list[str]) -> bytes:
        """Pack the named repository files into an in-memory ZIP archive."""

        buffer = BytesIO()
        with ZipFile(buffer, "w", ZIP_DEFLATED) as archive:
            for filename in filenames:
                path = self.repo_file(repo_key, filename)
                archive.write(path, arcname=path.name)
        return buffer.getvalue()

    def output_path(self, markdown: str | Path, export_type: str) -> Path:
        markdown = safe_filename(Path(markdown).name)
        filename = safe_filename(Path(markdown).with_suffix("").name)
        if export_type == "docx":
            return self.docx_repo / f"{filename}.docx"
        if export_type == "excel":
            return self.excel_repo / f"{filename}.xlsx"
        raise ValueError(f"不支持的导出类型：{export_type}")

    def remove_temp(self, path: str | Path) -> None:
        target = Path(path).resolve()
        uploads_root = self.temp_uploads.resolve()
        if target == uploads_root or not target.is_relative_to(uploads_root):
            raise ValueError("只能删除临时上传目录中的文件")
        target.unlink(missing_ok=True)
