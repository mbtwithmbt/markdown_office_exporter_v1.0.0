# Streamlit + MinerU 文档处理流水线设计

## 目标

在现有 Markdown → DOCX/Excel 导出库和 Tkinter 窗口之外，新增一个本地运行的 Streamlit 主界面，打通“上传文件 → MinerU 解析 → 保存 Markdown → 批量导出 DOCX/Excel”完整流程。第一版优先保证功能可用、错误清晰和代码可测试，不实现并发任务、用户系统、云存储或数据库。

## 保留与兼容

- 保留现有 `md_office_exporter` 导出函数、命令行和 Tkinter 窗口。
- 新功能复用 `convert_markdown_to_docx` 与 `export_markdown_tables_to_excel`，不复制解析或 Office 生成逻辑。
- Streamlit 流水线生成的 Excel 默认使用 `Sheet1`、`Sheet2` 顺序命名；现有调用方式继续使用 Markdown 标题命名，避免破坏旧行为。
- Streamlit 仓库导出采用同名自动覆盖；Tkinter 继续保留覆盖确认。

## 架构与文件职责

### `md_office_exporter/mineru_client.py`

封装 MinerU HTTP API，只负责远程通信与响应校验：

- `MinerUClient(base_url, timeout=300)`：规范化服务地址。
- `health_check()`：调用 `GET /health` 并返回 JSON 对象。
- `parse_file(path, backend="hybrid-engine")`：以 multipart/form-data 调用 `POST /file_parse`，固定发送 `return_md=true`、`response_format_zip=false`、`return_original_file=false`。
- 将响应中的 `results[*].md_content` 规范化为 `{名称: Markdown 文本}`。
- 网络错误、HTTP 错误、非 JSON、非完成状态或缺少 Markdown 内容统一转换为带上下文的 `MinerUError`。

MinerU 地址默认读取 `MINERU_BASE_URL`，未设置时为 `http://127.0.0.1:8000`。Streamlit 侧边栏可以临时覆盖地址、后端和超时。

### `md_office_exporter/storage.py`

管理项目内四个仓库目录：

- `markdown_repo/`
- `docx_repo/`
- `excel_repo/`
- `temp_uploads/`

启动时自动创建目录。所有外部文件名先转换为安全的相对文件名，禁止 `..` 或绝对路径逃逸。提供保存上传文件、保存 Markdown、列出 Markdown、删除临时文件和生成目标路径的小函数。同名文件直接覆盖。

### `md_office_exporter/markdown_processor.py`

编排本地业务逻辑：

- 将一个或多个上传文件逐个交给 MinerU，并把所有解析结果保存到 Markdown 仓库。
- 批量导出选中的 Markdown 为 DOCX 或 Excel。
- 单文件失败记录为结果对象，不中断后续文件。
- DOCX 接受可选 YAML 配置路径。
- Excel 接受 `merge/expand`，并在 Streamlit 路径中启用顺序 Sheet 命名。
- Markdown 无表格时返回明确错误，不创建伪造表格。

### `streamlit_app.py`

仅负责页面状态和用户交互，不直接拼装 HTTP 请求或实现导出逻辑。

## Streamlit 交互

页面由三个区域组成。

### 1. 工具说明与 MinerU 设置

页面顶部说明完整处理流程。侧边栏包含 MinerU 地址、解析后端和超时输入，并提供健康检查按钮。健康检查显示服务返回的协议版本、处理窗口和并发信息；失败时显示具体错误。

### 2. 上传并解析

用户选择上传模式：单文件、多文件或整个文件夹。文件类型允许 PDF、常见图片、DOCX、PPTX 和 XLSX。目录模式使用 Streamlit 官方支持的 `accept_multiple_files="directory"`。

点击“开始解析”后：

1. 将每个上传对象写入 `temp_uploads`。
2. 逐个调用 MinerU 同步接口。
3. 将返回的 Markdown 以原文件主名保存到 `markdown_repo`；API 返回多个结果时分别保存。
4. 删除本轮临时文件。
5. 用进度条和结果表展示每个文件的成功路径或失败原因。

### 3. Markdown 管理与批量导出

页面实时列出 `markdown_repo` 中的 `.md` 文件，并允许多选。

- DOCX：选择可选 YAML 配置文件，批量导出到 `docx_repo`。
- Excel：选择 `merge` 或 `expand`，批量导出到 `excel_repo`。
- 输出名称与 Markdown 主名一致并自动覆盖。
- 完成后显示每个输出路径；提供单文件下载按钮。

## Excel Sheet 命名

为 Excel 导出配置新增 `excel.sheet_name_mode`：

- `title`：现有默认值，使用最近的 Markdown 标题并处理重名。
- `sequential`：使用 `Sheet1`、`Sheet2` 等名称。

Streamlit 的 Markdown 处理器强制覆盖为 `sequential`；CLI、Python API 和 Tkinter 未指定时仍使用 `title`。

## 错误与安全处理

- API 地址、后端和超时在调用前校验。
- 上传文件名只保留安全名称；同批重名文件按输入顺序覆盖，结果中明确显示最终路径。
- HTTP 超时、连接失败、服务端错误和返回格式异常均显示原始上下文，但不泄露文件内容。
- 批量处理隔离单文件异常，最终给出成功/失败统计。
- YAML 无效、Markdown 不存在、Excel 无表格、输出文件被占用等导出错误按文件展示。
- 不执行上传文件；仅保存并发送给配置的 MinerU 服务。

## 依赖与启动

新增运行依赖：

- `streamlit>=1.54`
- `requests>=2.32`

Streamlit 启动命令：

```powershell
streamlit run streamlit_app.py
```

保留 `python -m md_office_exporter.desktop_gui` 和双击 `run_gui.pyw`。

## 测试与验收

- MinerU 客户端：Mock HTTP 测试健康检查、multipart 参数、完成响应、HTTP/JSON/状态/内容错误。
- 存储：临时目录测试自动建目录、安全文件名、保存、列举和覆盖。
- 处理器：Mock MinerU 与导出函数，测试批量成功、部分失败、DOCX 配置和 Excel 两种模式。
- Excel：新增顺序 Sheet 命名测试，并确保旧标题命名测试继续通过。
- Streamlit：使用 `streamlit.testing.v1.AppTest` 验证页面无异常启动、存在上传与管理控件，目录上传模式被启用。
- 回归：运行现有全部测试，并真实生成 DOCX、merge Excel、expand Excel。

验收时如果本机没有运行 MinerU 服务，API 使用 Mock 测试完成；同时通过页面健康检查明确显示服务未连接。只在实际服务可用时执行真实 OCR 请求。
