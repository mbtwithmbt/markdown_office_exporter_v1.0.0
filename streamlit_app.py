"""Streamlit interface for the MinerU to Markdown to Office pipeline."""

from __future__ import annotations

import os
from datetime import datetime
from pathlib import Path

import streamlit as st

from md_office_exporter.markdown_processor import (
    export_markdown_files,
    parse_uploaded_files,
)
from md_office_exporter.mineru_client import (
    DEFAULT_MINERU_BASE_URL,
    SUPPORTED_BACKENDS,
    MinerUClient,
)
from md_office_exporter.storage import RepositoryStorage
from md_office_exporter.templates import (
    BUILTIN_TEMPLATE_LABEL,
    ensure_seed_templates,
    list_templates,
    save_template,
    template_file,
)


PROJECT_ROOT = Path(__file__).resolve().parent
DATA_ROOT = Path(os.getenv("MARKDOWN_OFFICE_DATA_ROOT", PROJECT_ROOT))
storage = RepositoryStorage(DATA_ROOT)
ensure_seed_templates(storage)

if 'parse_results' not in st.session_state:
    st.session_state.parse_results = []
if 'export_results' not in st.session_state:
    st.session_state.export_results = []

UPLOAD_MODES = {
    "\u5355\u6587\u4ef6": False,
    "\u591a\u6587\u4ef6": True,
    "\u6587\u4ef6\u5939": "directory",
}
UPLOAD_TYPES = [
    "pdf",
    "png",
    "jpg",
    "jpeg",
    "webp",
    "tif",
    "tiff",
    "bmp",
    "docx",
    "pptx",
    "xlsx",
]
BACKEND_CHOICES = [
    "pipeline",
    "vlm-engine",
    "hybrid-engine",
    "vlm-http-client",
    "hybrid-http-client",
]
REPO_LABELS = {
    "Markdown 仓库（markdown_repo）": "markdown",
    "DOCX 仓库（docx_repo）": "docx",
    "Excel 仓库（excel_repo）": "excel",
    "模板仓库（config_repo）": "config",
}
FILE_MIME_TYPES = {
    ".md": "text/markdown",
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".yaml": "text/yaml",
    ".yml": "text/yaml",
}


def _file_mime(path: Path) -> str:
    return FILE_MIME_TYPES.get(path.suffix.casefold(), "application/octet-stream")


def _format_size(size_bytes: int) -> str:
    size = float(size_bytes)
    for unit in ("B", "KB", "MB", "GB"):
        if size < 1024 or unit == "GB":
            return f"{size:.1f} {unit}" if unit != "B" else f"{int(size)} B"
        size /= 1024
    return f"{size:.1f} GB"


def render_template_saver(key_prefix: str) -> None:
    """Render the upload-and-save preset template controls."""

    with st.expander("上传并保存新模板（YAML）"):
        upload = st.file_uploader(
            "模板 YAML 文件（可覆盖 docx / excel 部分配置）",
            type=["yaml", "yml"],
            accept_multiple_files=False,
            key=f"{key_prefix}_template_file",
        )
        template_name = st.text_input(
            "模板名称（保存为 config_repo/名称.yaml）",
            value=Path(upload.name).stem if upload is not None else "",
            key=f"{key_prefix}_template_name",
        )
        if st.button("保存为预设模板", key=f"{key_prefix}_template_save"):
            if upload is None or not template_name.strip():
                st.warning("请先选择模板文件并填写模板名称。")
            else:
                try:
                    saved = save_template(
                        storage,
                        template_name.strip(),
                        upload.getvalue(),
                    )
                except ValueError as exc:
                    st.error(f"模板保存失败：{exc}")
                else:
                    st.success(
                        f"模板已保存：{saved.name}。同名模板会被覆盖；"
                        "导出时请在“预设模板”下拉框选择它。"
                    )


def render_repo_manager() -> None:
    """Render upload, listing, download, zip and delete controls per repository."""

    repo_label = st.radio(
        "选择仓库",
        options=list(REPO_LABELS),
        horizontal=True,
        key="repo_choice",
    )
    repo_key = REPO_LABELS[repo_label]

    if repo_key == "markdown":
        uploads = st.file_uploader(
            "上传 Markdown 文件到 markdown_repo",
            type=["md"],
            accept_multiple_files=True,
            key="markdown_repo_upload",
        )
        if uploads:
            saved = []
            for item in uploads:
                saved.append(storage.upload_markdown(item.name, item.getvalue()).name)
            st.success("已上传并覆盖同名文件：" + "、".join(saved))
    elif repo_key == "config":
        render_template_saver("files")

    files = storage.list_repo_files(repo_key)
    st.subheader(f"文件列表（{len(files)} 个文件）")
    if not files:
        st.info("当前仓库暂无文件。")
        return

    st.dataframe(
        [
            {
                "文件名": path.name,
                "大小": _format_size(path.stat().st_size),
                "修改时间": datetime.fromtimestamp(
                    path.stat().st_mtime
                ).strftime("%Y-%m-%d %H:%M"),
            }
            for path in files
        ],
        width="stretch",
        hide_index=True,
    )

    st.subheader("单个下载")
    for path in files:
        try:
            content = path.read_bytes()
        except OSError as exc:
            st.error(f"无法读取 {path.name}: {exc}")
            continue
        st.download_button(
            label=f"下载 {path.name}",
            data=content,
            file_name=path.name,
            mime=_file_mime(path),
            key=f"repo_download_{repo_key}_{path.name}",
        )

    st.subheader("批量操作")
    selection_key = f"repo_selection_{repo_key}"
    if st.button("批量删除选中文件", key=f"repo_delete_{repo_key}"):
        chosen = list(st.session_state.get(selection_key, []))
        if not chosen:
            st.warning("请先在下方选择要删除的文件。")
        else:
            results = storage.delete_repo_files(repo_key, chosen)
            deleted = [name for name, ok, _ in results if ok]
            failed = [(name, message) for name, ok, message in results if not ok]
            if deleted:
                st.success("已删除：" + "、".join(deleted))
            for name, message in failed:
                st.error(f"{name}: {message}")
            st.session_state[selection_key] = []

    selected = st.multiselect(
        "选择文件（用于打包下载或批量删除）",
        options=[path.name for path in files],
        key=selection_key,
    )
    if selected:
        try:
            zip_bytes = storage.zip_repo_files(repo_key, selected)
        except (ValueError, FileNotFoundError) as exc:
            st.error(f"打包失败：{exc}")
        else:
            st.download_button(
                label=f"打包下载 {len(selected)} 个文件（ZIP）",
                data=zip_bytes,
                file_name=f"{repo_key}_repo.zip",
                mime="application/zip",
                key=f"repo_zip_{repo_key}",
            )


st.set_page_config(
    page_title='Markdown Office Exporter',
    page_icon="\U0001f4c4",
    layout="wide",
)

st.title("Markdown Office Exporter")
st.markdown(
    "\u4e0a\u4f20\u6587\u4ef6 \u2192 MinerU \u89e3\u6790 \u2192 Markdown \u4ed3\u5e93 "
    "\u2192 DOCX / Excel \u6279\u91cf\u5bfc\u51fa\u3002"
)

with st.sidebar:
    st.header("MinerU")
    mineru_url = st.text_input("MinerU URL", value=DEFAULT_MINERU_BASE_URL)
    backend = st.selectbox(
        "MinerU backend",
        options=BACKEND_CHOICES,
        index=BACKEND_CHOICES.index("hybrid-engine"),
    )
    timeout = st.number_input(
        "\u8d85\u65f6\u65f6\u95f4\uff08\u79d2\uff09",
        min_value=1,
        value=300,
        step=1,
    )
    if st.button("\u5065\u5eb7\u68c0\u67e5"):
        try:
            with MinerUClient(mineru_url, timeout=float(timeout)) as client:
                health = client.health_check()
            health_summary = {
                key: health[key]
                for key in ("protocol", "window", "concurrency")
                if key in health
            }
            st.json(health_summary or health)
        except Exception as exc:
            st.error(f"MinerU health check failed: {exc}")

export_tab, files_tab = st.tabs(["解析与导出", "文件管理"])

with export_tab:
    st.header("1. MinerU \u89e3\u6790")
    upload_mode = st.radio(
        "\u4e0a\u4f20\u65b9\u5f0f",
        options=list(UPLOAD_MODES),
        horizontal=True,
    )
    uploaded = st.file_uploader(
        "\u9009\u62e9\u5f85\u89e3\u6790\u6587\u4ef6",
        type=UPLOAD_TYPES,
        accept_multiple_files=UPLOAD_MODES[upload_mode],
        key=f"source_upload_{upload_mode}",
    )

    if uploaded is None:
        uploaded_files = []
    elif isinstance(uploaded, (list, tuple)):
        uploaded_files = list(uploaded)
    else:
        uploaded_files = [uploaded]

    if st.button('\u5f00\u59cb\u89e3\u6790', type='primary'):
        if not uploaded_files:
            st.warning('\u8bf7\u5148\u9009\u62e9\u5f85\u89e3\u6790\u6587\u4ef6\u3002')
        else:
            progress = st.progress(0)
            result_rows = []
            total = max(len(uploaded_files), 1)
            completed = 0
            cleanup_failures = []
            try:
                with MinerUClient(mineru_url, timeout=float(timeout)) as client:
                    for source in uploaded_files:
                        try:
                            path = storage.save_upload(source.name, source.getvalue())
                        except Exception as exc:
                            result_rows.append(
                                {
                                    'source': source.name,
                                    'status': 'error',
                                    'outputs': [],
                                    'output': '',
                                    'error': str(exc),
                                }
                            )
                        else:
                            try:
                                per_file_results = parse_uploaded_files(
                                    [path],
                                    client,
                                    storage,
                                    backend=str(backend),
                                )
                            except Exception as exc:
                                result_rows.append(
                                    {
                                        'source': source.name,
                                        'status': 'error',
                                        'outputs': [],
                                        'output': '',
                                        'error': str(exc),
                                    }
                                )
                                try:
                                    storage.remove_temp(path)
                                except Exception as cleanup_exc:
                                    cleanup_failures.append(f'{path}: {cleanup_exc}')
                            else:
                                for result in per_file_results:
                                    result_rows.append(
                                        {
                                            'source': result.source,
                                            'status': (
                                                'success' if result.success else 'error'
                                            ),
                                            'outputs': [
                                                str(output) for output in result.outputs
                                            ],
                                            'output': ', '.join(
                                                output.name for output in result.outputs
                                            ),
                                            'error': result.error,
                                        }
                                    )
                        finally:
                            completed += 1
                            progress.progress(min(completed / total, 1.0))
            except Exception as exc:
                st.error(f'MinerU client failed: {exc}')
                for source in uploaded_files[completed:]:
                    result_rows.append(
                        {
                            'source': source.name,
                            'status': 'error',
                            'outputs': [],
                            'output': '',
                            'error': str(exc),
                        }
                    )
                    completed += 1
                    progress.progress(min(completed / total, 1.0))
            if cleanup_failures:
                st.warning(
                    'Temporary upload cleanup failed; orphan files: '
                    + '; '.join(cleanup_failures)
                )
            progress.progress(1.0)
            st.session_state.parse_results = result_rows

    if st.session_state.parse_results:
        st.dataframe(
            [
                {
                    key: row[key]
                    for key in ('source', 'status', 'output', 'error')
                }
                for row in st.session_state.parse_results
            ],
            width='stretch',
            hide_index=True,
        )

    st.divider()
    st.header("2. Markdown \u7ba1\u7406\u4e0e\u5bfc\u51fa")

    render_template_saver("export")

    template_names = list_templates(storage)
    template_choice = st.selectbox(
        "预设模板（docx / excel 导出共用）",
        options=[BUILTIN_TEMPLATE_LABEL, *template_names],
        key="template_choice",
    )
    if template_choice != BUILTIN_TEMPLATE_LABEL:
        try:
            preset_path = template_file(storage, template_choice)
        except FileNotFoundError as exc:
            st.error(f"{exc}")
        else:
            st.download_button(
                label=f"下载模板 {preset_path.name}",
                data=preset_path.read_bytes(),
                file_name=preset_path.name,
                mime="text/yaml",
                key=f"template_download_{template_choice}",
            )

    markdown_paths = storage.list_markdown()
    markdown_by_name = {path.name: path for path in markdown_paths}
    selected_names = st.multiselect(
        "Markdown \u6587\u4ef6",
        options=list(markdown_by_name),
    )
    export_kind = st.selectbox("\u5bfc\u51fa\u7c7b\u578b", options=["DOCX", "Excel"])

    merge_mode = "merge"
    docx_scope = "整个文档"
    if export_kind == "DOCX":
        docx_scope = st.radio(
            "DOCX 导出范围",
            options=["整个文档", "仅表格"],
            horizontal=True,
        )
    else:
        merge_mode = st.radio(
            "Excel \u5408\u5e76\u5355\u5143\u683c",
            options=["merge", "expand"],
            horizontal=True,
        )

    if st.button('\u6279\u91cf\u5bfc\u51fa', type='primary'):
        if not selected_names:
            st.warning('\u8bf7\u5148\u9009\u62e9 Markdown \u6587\u4ef6\u3002')
        else:
            config_path = None
            if template_choice != BUILTIN_TEMPLATE_LABEL:
                try:
                    config_path = template_file(storage, template_choice)
                except FileNotFoundError as exc:
                    st.error(f"导出中止，模板文件读取失败：{exc}")
                    st.stop()
            try:
                export_results = export_markdown_files(
                    [markdown_by_name[name] for name in selected_names],
                    storage,
                    'docx' if export_kind == 'DOCX' else 'excel',
                    config_path=config_path,
                    merge_mode=merge_mode,
                    tables_only=export_kind == 'DOCX' and docx_scope == '仅表格',
                )
                st.session_state.export_results = [
                    {
                        'source': result.source,
                        'status': 'success' if result.success else 'error',
                        'outputs': [str(output) for output in result.outputs],
                        'output': ', '.join(output.name for output in result.outputs),
                        'error': result.error,
                    }
                    for result in export_results
                ]
            except Exception as exc:
                st.error(f'Export failed: {exc}')
                st.session_state.export_results = []

    if st.session_state.export_results:
        st.dataframe(
            [
                {
                    key: row[key]
                    for key in ('source', 'status', 'output', 'error')
                }
                for row in st.session_state.export_results
            ],
            width='stretch',
            hide_index=True,
        )

        for result_index, result in enumerate(st.session_state.export_results):
            if result['status'] != 'success':
                continue
            for output_index, output_string in enumerate(result['outputs']):
                output = Path(output_string)
                try:
                    content = output.read_bytes()
                except OSError as exc:
                    st.error(f'Cannot read {output.name}: {exc}')
                    continue
                st.download_button(
                    label=f'Download {output.name}',
                    data=content,
                    file_name=output.name,
                    mime=_file_mime(output),
                    key=f'download_{result_index}_{output_index}_{output.name}',
                )

with files_tab:
    st.header("3. \u6587\u4ef6\u7ba1\u7406")
    st.caption(
        "支持 Markdown / DOCX / Excel / 模板四个仓库的文件浏览、单个下载、"
        "批量打包下载与批量删除；Markdown 与模板仓库支持直接上传。"
    )
    render_repo_manager()

if set(BACKEND_CHOICES) != SUPPORTED_BACKENDS:
    st.error("Configured MinerU backend choices do not match the client.")
