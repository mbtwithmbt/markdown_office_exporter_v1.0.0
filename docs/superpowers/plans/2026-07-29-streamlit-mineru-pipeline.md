# Streamlit + MinerU Pipeline Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a local Streamlit application that uploads single, multiple, or directory files, parses them through MinerU into a Markdown repository, and batch-exports selected Markdown files to DOCX or Excel.

**Architecture:** Add isolated MinerU, storage, and processing modules inside the existing package, with a root Streamlit entrypoint that only coordinates UI state. Reuse the existing Office exporters, preserve Tkinter/CLI compatibility, and run batch work sequentially with per-file result isolation.

**Tech Stack:** Python 3.10+, Streamlit 1.54+, requests 2.32+, unittest/mock, existing python-docx/openpyxl/PyYAML/lxml.

---

## File map

- Modify `md_office_exporter/config.py`: add the backward-compatible Excel `sheet_name_mode` default and validation.
- Modify `md_office_exporter/excel_exporter.py`: select title or sequential worksheet names from configuration.
- Modify `config.example.yaml`: document `excel.sheet_name_mode`.
- Create `md_office_exporter/storage.py`: repository paths and safe file persistence.
- Create `md_office_exporter/mineru_client.py`: MinerU HTTP adapter and response validation.
- Create `md_office_exporter/markdown_processor.py`: sequential parse/export orchestration and per-file results.
- Create `streamlit_app.py`: Streamlit page and UI coordination.
- Modify `requirements.txt` and `pyproject.toml`: declare Streamlit and requests.
- Modify `README.md`: document the complete pipeline and both UI launch methods.
- Create `tests/test_storage.py`, `tests/test_mineru_client.py`, `tests/test_markdown_processor.py`, and `tests/test_streamlit_app.py`.
- Modify `tests/test_exporters.py`: cover sequential worksheet naming while retaining title naming coverage.

This directory is not a Git repository. Each task therefore ends with a full test checkpoint instead of a commit; repository initialization is outside scope.

### Task 1: Backward-compatible sequential Excel sheet names

**Files:**
- Modify: `md_office_exporter/config.py`
- Modify: `md_office_exporter/excel_exporter.py`
- Modify: `config.example.yaml`
- Modify: `tests/test_exporters.py`

- [ ] **Step 1: Write the failing sequential-name test**

Add a test that exports the existing two-table sample with a configuration override and expects `Sheet1` and `Sheet2`:

```python
def test_excel_sequential_sheet_names(self) -> None:
    output = self.root / "sequential.xlsx"
    export_markdown_tables_to_excel(
        self.markdown,
        output,
        config={"excel": {"sheet_name_mode": "sequential"}},
    )
    workbook = load_workbook(output)
    self.assertEqual(["Sheet1", "Sheet2"], workbook.sheetnames)
```

Also test `load_config(overrides={"excel": {"sheet_name_mode": "invalid"}})` raises a `ValueError` mentioning `sheet_name_mode`.

- [ ] **Step 2: Verify RED**

Run `python -m unittest tests.test_exporters -v`.

Expected: the sequential-name assertion fails because sheets still use Markdown titles, and invalid mode is not rejected.

- [ ] **Step 3: Implement configuration and exporter changes**

Add `"sheet_name_mode": "title"` under the Excel defaults. In `load_config`, validate:

```python
sheet_name_mode = config["excel"]["sheet_name_mode"]
if sheet_name_mode not in {"title", "sequential"}:
    raise ValueError("excel.sheet_name_mode 只能是 'title' 或 'sequential'")
```

In the export loop choose the title source without changing the existing `_sheet_name` sanitation:

```python
sheet_name_mode = resolved_config["excel"]["sheet_name_mode"]
for number, table in enumerate(tables, start=1):
    requested_name = (
        f"Sheet{number}" if sheet_name_mode == "sequential" else table.title
    )
    worksheet = workbook.create_sheet(
        _sheet_name(requested_name, used_names, number)
    )
```

Add `sheet_name_mode: title` with a comment describing both values to `config.example.yaml`.

- [ ] **Step 4: Verify GREEN and compatibility**

Run `python -m unittest tests.test_exporters -v`.

Expected: title-mode tests and the new `Sheet1/Sheet2` test all pass.

### Task 2: Repository storage layer

**Files:**
- Create: `md_office_exporter/storage.py`
- Create: `tests/test_storage.py`

- [ ] **Step 1: Write failing storage tests**

Tests must construct `RepositoryStorage(root=temp_path)` and verify:

```python
storage = RepositoryStorage(self.root)
self.assertTrue(storage.markdown_dir.is_dir())
self.assertTrue(storage.docx_dir.is_dir())
self.assertTrue(storage.excel_dir.is_dir())
self.assertTrue(storage.temp_dir.is_dir())

saved = storage.save_upload("../report.pdf", b"first")
self.assertEqual(storage.temp_dir / "report.pdf", saved)
storage.save_upload("report.pdf", b"second")
self.assertEqual(b"second", saved.read_bytes())

markdown = storage.save_markdown("folder\\report.pdf", "# parsed")
self.assertEqual(storage.markdown_dir / "report.md", markdown)
self.assertEqual([markdown], storage.list_markdown())
```

Also verify blank names and `.`/`..` raise `ValueError`, output paths use `.docx/.xlsx`, and `remove_temp` removes only files inside `temp_uploads`.

- [ ] **Step 2: Verify RED**

Run `python -m unittest tests.test_storage -v`.

Expected: import failure because `md_office_exporter.storage` does not exist.

- [ ] **Step 3: Implement storage**

Create:

```python
from pathlib import Path
import re

INVALID_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')

def safe_filename(name: str) -> str:
    basename = Path(str(name).replace("\\", "/")).name.strip()
    basename = INVALID_FILENAME_CHARS.sub("_", basename).rstrip(". ")
    if basename in {"", ".", ".."}:
        raise ValueError("文件名不能为空")
    return basename

class RepositoryStorage:
    def __init__(self, root: str | Path) -> None:
        self.root = Path(root).resolve()
        self.markdown_dir = self.root / "markdown_repo"
        self.docx_dir = self.root / "docx_repo"
        self.excel_dir = self.root / "excel_repo"
        self.temp_dir = self.root / "temp_uploads"
        for directory in (
            self.markdown_dir, self.docx_dir, self.excel_dir, self.temp_dir
        ):
            directory.mkdir(parents=True, exist_ok=True)

    def save_upload(self, name: str, content: bytes) -> Path:
        target = self.temp_dir / safe_filename(name)
        target.write_bytes(content)
        return target

    def save_markdown(self, source_name: str, content: str) -> Path:
        target = self.markdown_dir / Path(safe_filename(source_name)).with_suffix(".md")
        target.write_text(content, encoding="utf-8")
        return target

    def list_markdown(self) -> list[Path]:
        return sorted(self.markdown_dir.glob("*.md"), key=lambda p: p.name.casefold())

    def output_path(self, markdown: Path, export_type: str) -> Path:
        if export_type == "docx":
            return self.docx_dir / markdown.with_suffix(".docx").name
        if export_type == "excel":
            return self.excel_dir / markdown.with_suffix(".xlsx").name
        raise ValueError("导出类型只能是 docx 或 excel")

    def remove_temp(self, path: Path) -> None:
        resolved = path.resolve()
        if not resolved.is_relative_to(self.temp_dir.resolve()):
            raise ValueError("只能删除 temp_uploads 中的文件")
        resolved.unlink(missing_ok=True)
```

- [ ] **Step 4: Verify GREEN**

Run `python -m unittest tests.test_storage -v`.

Expected: all storage and safety tests pass.

### Task 3: MinerU HTTP client

**Files:**
- Create: `md_office_exporter/mineru_client.py`
- Create: `tests/test_mineru_client.py`

- [ ] **Step 1: Write failing client tests**

Use `unittest.mock.Mock` for an injected `requests.Session`. Cover:

```python
client = MinerUClient("http://127.0.0.1:8000/", session=session)
self.assertEqual(health_payload, client.health_check())
session.get.assert_called_once_with(
    "http://127.0.0.1:8000/health", timeout=300.0
)

result = client.parse_file(sample_pdf, backend="hybrid-engine")
self.assertEqual({"sample": "# OCR"}, result)
```

Inspect `session.post.call_args` to assert URL, timeout, `files["files"][0]`, and lowercase form values. Add individual tests for a missing file, invalid backend, nonpositive timeout, requests exception, HTTP failure, JSON decode failure, `status != "completed"`, missing `results`, and missing/blank `md_content`.

- [ ] **Step 2: Verify RED**

Run `python -m unittest tests.test_mineru_client -v`.

Expected: import failure because the MinerU client does not exist.

- [ ] **Step 3: Implement the client**

Define:

```python
DEFAULT_MINERU_BASE_URL = os.getenv(
    "MINERU_BASE_URL", "http://127.0.0.1:8000"
)
SUPPORTED_BACKENDS = {
    "pipeline", "vlm-engine", "hybrid-engine",
    "vlm-http-client", "hybrid-http-client",
}

class MinerUError(RuntimeError):
    pass

class MinerUClient:
    def __init__(
        self,
        base_url: str = DEFAULT_MINERU_BASE_URL,
        timeout: float = 300,
        session: requests.Session | None = None,
    ) -> None:
        if not base_url.strip():
            raise ValueError("MinerU API 地址不能为空")
        if timeout <= 0:
            raise ValueError("超时时间必须大于 0")
        self.base_url = base_url.rstrip("/")
        self.timeout = float(timeout)
        self.session = session or requests.Session()
```

`health_check` and `parse_file` must wrap `requests.RequestException` and JSON errors in `MinerUError`. `parse_file` opens the input with a context manager, posts one `files` part, validates `status == "completed"`, and returns every nonblank `md_content` entry. If no valid Markdown exists, raise `MinerUError("MinerU 响应中没有可用的 md_content")`.

- [ ] **Step 4: Verify GREEN**

Run `python -m unittest tests.test_mineru_client -v`.

Expected: all network, response, and validation tests pass without a live MinerU server.

### Task 4: Sequential parsing and batch export processor

**Files:**
- Create: `md_office_exporter/markdown_processor.py`
- Create: `tests/test_markdown_processor.py`

- [ ] **Step 1: Write failing processor tests**

Define expectations around a shared result model:

```python
@dataclass(slots=True)
class OperationResult:
    source: str
    success: bool
    outputs: list[Path] = field(default_factory=list)
    error: str = ""
```

Tests use temporary `RepositoryStorage`, a mock MinerU client, and patched exporter functions. Verify:

- Two uploads are processed in order and their Markdown paths are returned.
- When the first client call raises `MinerUError`, the second still succeeds.
- Every input temp file is removed in `finally`.
- DOCX export writes to `docx_repo`, forwards the optional config path, and overwrites the same name.
- Excel export writes to `excel_repo`, forwards `merge` or `expand`, and passes `config={"excel": {"sheet_name_mode": "sequential"}}`.
- A failing export becomes `success=False` without stopping later Markdown files.

- [ ] **Step 2: Verify RED**

Run `python -m unittest tests.test_markdown_processor -v`.

Expected: import failure because the processor does not exist.

- [ ] **Step 3: Implement the processor**

Create `parse_uploaded_files(paths, client, storage, backend)` and `export_markdown_files(paths, storage, export_type, config_path=None, merge_mode="merge")`. Both functions return `list[OperationResult]` and catch exceptions inside the loop.

The parse loop must use:

```python
try:
    parsed = client.parse_file(path, backend=backend)
    outputs = [storage.save_markdown(name, text) for name, text in parsed.items()]
    results.append(OperationResult(path.name, True, outputs))
except Exception as exc:
    results.append(OperationResult(path.name, False, error=str(exc)))
finally:
    storage.remove_temp(path)
```

The export loop calls the existing exporters. Excel must validate merge mode and call:

```python
export_markdown_tables_to_excel(
    markdown,
    output,
    merge_mode=merge_mode,
    config_path=config_path,
    config={"excel": {"sheet_name_mode": "sequential"}},
)
```

- [ ] **Step 4: Verify GREEN**

Run `python -m unittest tests.test_markdown_processor -v`.

Expected: parsing isolation, cleanup, repository routing, and both export modes pass.

### Task 5: Streamlit application and dependencies

**Files:**
- Create: `streamlit_app.py`
- Create: `tests/test_streamlit_app.py`
- Modify: `requirements.txt`
- Modify: `pyproject.toml`

- [ ] **Step 1: Declare dependencies and install them**

Add identical runtime requirements to both dependency files:

```text
requests>=2.32
streamlit>=1.54
```

Run `python -m pip install -e .`.

Expected: installation succeeds and `python -c "import requests, streamlit"` exits 0.

- [ ] **Step 2: Write the failing Streamlit page test**

Use `streamlit.testing.v1.AppTest` with a temporary data root set through `MARKDOWN_OFFICE_DATA_ROOT`:

```python
app = AppTest.from_file("streamlit_app.py")
app.run(timeout=10)
self.assertFalse(app.exception)
self.assertGreaterEqual(len(app.file_uploader), 1)
self.assertGreaterEqual(len(app.button), 2)

app.radio[0].set_value("文件夹").run(timeout=10)
self.assertTrue(app.file_uploader[0].accept_directory)
```

Also assert the page contains the MinerU URL input, health-check button, Markdown multiselect, export-type selector, and batch-export button.

- [ ] **Step 3: Verify RED**

Run `python -m unittest tests.test_streamlit_app -v`.

Expected: failure because `streamlit_app.py` does not exist.

- [ ] **Step 4: Implement the Streamlit page**

At module load:

```python
PROJECT_ROOT = Path(__file__).resolve().parent
DATA_ROOT = Path(os.getenv("MARKDOWN_OFFICE_DATA_ROOT", PROJECT_ROOT))
storage = RepositoryStorage(DATA_ROOT)
st.set_page_config(page_title="Markdown Office Exporter", page_icon="📄", layout="wide")
```

Sidebar widgets return URL, backend, and timeout. Health check constructs `MinerUClient` only after its button is clicked.

Upload mode must map exactly:

```python
UPLOAD_MODES = {
    "单文件": False,
    "多文件": True,
    "文件夹": "directory",
}
uploaded = st.file_uploader(
    "选择待解析文件",
    type=["pdf", "png", "jpg", "jpeg", "webp", "tif", "tiff", "bmp", "docx", "pptx", "xlsx"],
    accept_multiple_files=UPLOAD_MODES[upload_mode],
    key=f"source_upload_{upload_mode}",
)
```

Normalize the single-file return value to a list. On “开始解析”, save every `UploadedFile.getvalue()`, call `parse_uploaded_files`, update a progress bar per result, and display a dataframe-like result list with source/status/output/error.

The export section lists `storage.list_markdown()`, uses `st.multiselect` for any quantity, accepts an optional YAML uploader, and shows `merge/expand` only for Excel. Save an uploaded YAML config into the temporary repository for the duration of export and remove it afterward. On “批量导出”, call `export_markdown_files` and display per-file results. Render one download button per successful output using `Path.read_bytes()`.

- [ ] **Step 5: Verify GREEN**

Run `python -m unittest tests.test_streamlit_app -v`.

Expected: the app runs without exceptions and directory-upload metadata is enabled when selected.

### Task 6: Documentation and complete verification

**Files:**
- Modify: `README.md`
- Verify: all source and test files

- [ ] **Step 1: Update README**

Document:

- `streamlit run streamlit_app.py` as the primary interface.
- `MINERU_BASE_URL`, default URL, backend choices, and 300-second default timeout.
- the four auto-created storage directories.
- single/multiple/directory upload, Markdown management, batch exports, overwrite behavior, and no-table error behavior.
- retained Tkinter and CLI launch instructions.
- `Sheet1/Sheet2` behavior in Streamlit versus title naming in the legacy API.

Remove the incorrect Tkinter claim that drag-and-drop is supported.

- [ ] **Step 2: Run all tests**

Run `python -m unittest discover -s tests -v`.

Expected: the existing 17 tests plus all new tests pass with no warnings or errors.

- [ ] **Step 3: Run syntax and import checks**

Run:

```powershell
python -m compileall -q md_office_exporter streamlit_app.py run_gui.pyw
python -c "import streamlit_app; import md_office_exporter.desktop_gui; print('imports OK')"
```

Expected: compilation succeeds and output contains `imports OK`.

- [ ] **Step 4: Run real Office smoke exports**

Using `examples/sample.md` and a temporary directory, call the processor for DOCX, Excel merge, and Excel expand. Open the outputs with python-docx/openpyxl and assert:

- DOCX contains both sample tables.
- merge workbook uses `Sheet1/Sheet2` and has merged ranges in `Sheet2`.
- expand workbook uses `Sheet1/Sheet2`, has no merged ranges, and repeated covered-cell values are present.

- [ ] **Step 5: Check optional live MinerU status**

Construct the client with the configured/default URL and call `health_check` once with a short diagnostic timeout. A successful response is reported as live integration evidence. Connection failure is reported as an environmental limitation, not a product test failure, because all HTTP behavior is covered by mocks and the UI exposes the same health check.

- [ ] **Step 6: Provide launch handoff**

Report exact test counts, live MinerU status, and these commands:

```powershell
streamlit run streamlit_app.py
python -m md_office_exporter.desktop_gui
```

Also link the absolute `streamlit_app.py`, `markdown_repo`, `docx_repo`, and `excel_repo` paths.
