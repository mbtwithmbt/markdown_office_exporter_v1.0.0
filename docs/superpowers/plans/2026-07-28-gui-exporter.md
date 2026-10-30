# Markdown Office Exporter GUI Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a lightweight Tkinter window for manually exporting one Markdown file to DOCX or Excel with optional YAML configuration and Excel merged-cell mode selection.

**Architecture:** Keep all parsing and Office generation in the existing public exporter functions. Add a focused GUI module containing testable path/validation/dispatch helpers plus a thin `ExporterApp`, and add a `.pyw` launcher for double-click use on Windows.

**Tech Stack:** Python 3.10+, standard-library Tkinter, pathlib, unittest/mock, existing python-docx/openpyxl/PyYAML/lxml exporters.

---

## File map

- Create `md_office_exporter/gui.py`: pure request helpers, exporter dispatch, Tkinter application, and module entry point.
- Create `run_gui.pyw`: console-free Windows launcher that calls `md_office_exporter.gui.main`.
- Create `tests/test_gui.py`: headless tests for output-path derivation, validation, and exporter dispatch.
- Modify `README.md`: add concise GUI launch and field usage instructions.

The directory is not a Git repository, so commit steps are intentionally replaced with explicit test checkpoints; initializing version control is outside this personal-tool change.

### Task 1: Test and implement request helpers

**Files:**
- Create: `tests/test_gui.py`
- Create: `md_office_exporter/gui.py`

- [ ] **Step 1: Write failing path and validation tests**

Create tests using `tempfile.TemporaryDirectory` that assert:

```python
from md_office_exporter.gui import suggested_output_path, validate_export_request

self.assertEqual(
    self.root / "report.docx",
    suggested_output_path(self.root / "report.md", "docx"),
)
self.assertEqual(
    self.root / "custom.xlsx",
    suggested_output_path(self.root / "custom.docx", "excel"),
)
```

Create a real `report.md` and optional `format.yaml`, then verify `validate_export_request` accepts valid DOCX/Excel requests and raises `ValueError` with field-specific Chinese messages for an empty path, a missing Markdown file, a missing selected configuration file, an unsupported export type, and an incorrect output suffix.

- [ ] **Step 2: Run the new tests and verify they fail**

Run:

```powershell
python -m unittest tests.test_gui -v
```

Expected: failure because `md_office_exporter.gui` does not exist.

- [ ] **Step 3: Implement minimal pure helpers**

In `md_office_exporter/gui.py`, define:

```python
EXPORT_EXTENSIONS = {"docx": ".docx", "excel": ".xlsx"}

def suggested_output_path(source: str | Path, export_type: str) -> Path:
    if export_type not in EXPORT_EXTENSIONS:
        raise ValueError("导出类型只能是 docx 或 excel")
    return Path(source).with_suffix(EXPORT_EXTENSIONS[export_type])

def validate_export_request(
    markdown_path: str | Path,
    output_path: str | Path,
    export_type: str,
    config_path: str | Path | None = None,
) -> tuple[Path, Path, Path | None]:
    # Validate supported type, non-empty paths, existing input/config files,
    # and the exact .docx/.xlsx output suffix; return normalized Path objects.
```

Whitespace-only values must count as empty. The config value `None` or `""` means use defaults and is valid.

- [ ] **Step 4: Run helper tests and verify they pass**

Run:

```powershell
python -m unittest tests.test_gui -v
```

Expected: all path and validation tests pass without creating a Tk root window.

### Task 2: Test and implement exporter dispatch

**Files:**
- Modify: `tests/test_gui.py`
- Modify: `md_office_exporter/gui.py`

- [ ] **Step 1: Write failing dispatch tests**

Patch the exporter symbols inside `md_office_exporter.gui` and assert exact calls:

```python
with patch("md_office_exporter.gui.convert_markdown_to_docx") as export_docx:
    perform_export(markdown, output, "docx", config_path=config)
    export_docx.assert_called_once_with(markdown, output, config)

with patch("md_office_exporter.gui.export_markdown_tables_to_excel") as export_excel:
    perform_export(markdown, output, "excel", config_path=None, merge_mode="expand")
    export_excel.assert_called_once_with(
        markdown, output, merge_mode="expand", config_path=None
    )
```

Also assert an invalid merge mode raises `ValueError` before the Excel exporter is called.

- [ ] **Step 2: Run dispatch tests and verify they fail**

Run:

```powershell
python -m unittest tests.test_gui -v
```

Expected: failure because `perform_export` is not defined.

- [ ] **Step 3: Implement the dispatch helper**

Add:

```python
def perform_export(
    markdown_path: Path,
    output_path: Path,
    export_type: str,
    config_path: Path | None = None,
    merge_mode: str = "merge",
) -> Path:
    if export_type == "docx":
        return convert_markdown_to_docx(markdown_path, output_path, config_path)
    if export_type == "excel":
        if merge_mode not in {"merge", "expand"}:
            raise ValueError("Excel 合并方式只能是 merge 或 expand")
        return export_markdown_tables_to_excel(
            markdown_path,
            output_path,
            merge_mode=merge_mode,
            config_path=config_path,
        )
    raise ValueError("导出类型只能是 docx 或 excel")
```

- [ ] **Step 4: Run GUI helper tests**

Run:

```powershell
python -m unittest tests.test_gui -v
```

Expected: all helper and dispatch tests pass.

### Task 3: Build the Tkinter window and Windows launcher

**Files:**
- Modify: `md_office_exporter/gui.py`
- Create: `run_gui.pyw`

- [ ] **Step 1: Implement `ExporterApp` layout**

Use `tkinter` and `ttk` to create a single-page form with these variables and controls:

```python
self.markdown_var = tk.StringVar()
self.config_var = tk.StringVar()
self.export_type_var = tk.StringVar(value="docx")
self.output_var = tk.StringVar()
self.merge_mode_var = tk.StringVar(value="merge")
self.status_var = tk.StringVar(value="请选择 Markdown 文件")
```

Add labeled entries and browse buttons for Markdown, optional YAML, and output. Add DOCX/Excel radio buttons, merge/expand radio buttons, an export button, and a wrapping status label. Configure column 1 to expand and use a modest default window size suitable for long Windows paths.

- [ ] **Step 2: Implement browse and mode-change behavior**

`browse_markdown` uses `filedialog.askopenfilename` with Markdown filters and then sets the same-directory suggested output. `browse_config` accepts YAML/YML files. `browse_output` uses `asksaveasfilename` with the current type's extension and file filter.

`on_export_type_changed` rewrites only the current output suffix and sets both merge radio buttons to `disabled` for DOCX or `normal` for Excel. It must not discard a manually chosen output directory or stem.

- [ ] **Step 3: Implement export behavior and error handling**

`export_file` performs the following sequence:

1. Call `validate_export_request`.
2. If the target exists, call `messagebox.askyesno`; return without writing when declined.
3. Disable the export button and show “正在导出…”.
4. Call `perform_export` with the selected merge mode.
5. On success, set a status containing the resolved path and show `messagebox.showinfo`.
6. On any exception, set a failure status and show `messagebox.showerror` with the original exception text.
7. Restore the button state in `finally`.

Add:

```python
def main() -> int:
    root = tk.Tk()
    ExporterApp(root)
    root.mainloop()
    return 0

if __name__ == "__main__":
    raise SystemExit(main())
```

- [ ] **Step 4: Add the double-click launcher**

Create `run_gui.pyw`:

```python
from md_office_exporter.gui import main

raise SystemExit(main())
```

- [ ] **Step 5: Verify import and syntax without opening a window**

Run:

```powershell
python -m py_compile md_office_exporter\gui.py run_gui.pyw
python -c "import md_office_exporter.gui; print('GUI import OK')"
```

Expected: compilation succeeds and output is `GUI import OK`.

### Task 4: Document use and run full verification

**Files:**
- Modify: `README.md`
- Test: `tests/test_gui.py`
- Test: `tests/test_exporters.py`

- [ ] **Step 1: Add concise GUI instructions**

Document both launch methods:

```powershell
python -m md_office_exporter.gui
```

and double-clicking `run_gui.pyw`. Explain that YAML is optional, DOCX ignores the Excel merge selector, `merge` preserves merged ranges, and `expand` repeats the merged value in every covered cell.

- [ ] **Step 2: Run all unit tests**

Run:

```powershell
python -m unittest discover -s tests -v
```

Expected: the existing four exporter tests and all new GUI helper tests pass.

- [ ] **Step 3: Run an end-to-end export smoke test through dispatch**

Use `examples/sample.md` and a temporary output directory to call `perform_export` three times: DOCX, Excel `merge`, and Excel `expand`. Load the generated files with `python-docx` and `openpyxl`; assert that the DOCX contains tables, the merged workbook contains merged ranges, and the expanded workbook contains no merged ranges.

- [ ] **Step 4: Report the personal-use launch path**

Confirm that the user can launch by double-clicking the absolute file `C:\Users\86182\Downloads\markdown_office_exporter_v1.0.0\run_gui.pyw`, and list the exact automated verification results. Note that a visible interactive window cannot be fully clicked through in a headless test, while all nonvisual behavior and real Office output paths have been exercised.
