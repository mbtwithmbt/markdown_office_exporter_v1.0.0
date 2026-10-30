# Markdown Office Exporter

Markdown Office Exporter 是一个本地文档处理工具：通过 MinerU 将 PDF、图片和 Office 文件解析为 Markdown，并批量导出 DOCX 或 Excel。项目也保留了原有的 Tkinter 桌面界面、命令行和 Python API。

## 安装

需要 Python 3.10 或更高版本。建议使用虚拟环境：

```powershell
python -m venv .venv
.venv\Scripts\Activate.ps1
python -m pip install -e .
```

也可以只安装运行依赖：

```powershell
python -m pip install -r requirements.txt
```

主要依赖包括 `streamlit>=1.54`、`requests>=2.32`、`python-docx>=1.1`、`openpyxl>=3.1`、`PyYAML>=6.0` 和 `lxml>=5.0`。

## 启动主界面（Streamlit）

在项目根目录运行：

```powershell
streamlit run streamlit_app.py
```

应用启动时会在项目根目录自动创建并使用以下五个目录：

- `markdown_repo/`：保存 MinerU 解析结果和待导出的 Markdown。
- `docx_repo/`：保存 DOCX 导出结果。
- `excel_repo/`：保存 Excel 导出结果。
- `config_repo/`：保存导出模板（YAML 配置文件），首次启动时自动写入三个内置预设。
- `temp_uploads/`：暂存上传文件；处理完成后自动清理。

同名 Markdown 和导出文件会直接覆盖。请在重新解析或导出前自行备份需要保留的旧版本。

### MinerU 设置

MinerU 服务地址默认读取环境变量 `MINERU_BASE_URL`；未设置时使用 `http://127.0.0.1:8000`。默认请求超时为 300 秒。Streamlit 侧边栏可以临时覆盖服务地址、解析后端和超时，并可执行健康检查。

支持五种后端：

- `pipeline`
- `vlm-engine`
- `hybrid-engine`（默认）
- `vlm-http-client`
- `hybrid-http-client`

健康检查调用 MinerU 的 `GET /health`。上传解析调用 `POST /file_parse`，因此使用主界面前需要另行启动或配置可访问的 MinerU 服务。

### 上传、解析与 Markdown 管理

主界面分为「解析与导出」和「文件管理」两个标签页。解析标签页支持三种上传方式：单文件、多文件和整个文件夹。可上传 PDF、常见图片、DOCX、PPTX 和 XLSX。点击“开始解析”后，文件会先写入 `temp_uploads/`，再逐个交给 MinerU；成功返回的 Markdown 保存到 `markdown_repo/`。单个文件失败不会中断同批其他文件，界面会显示每项的成功路径或具体错误。

Markdown 管理区会实时列出 `markdown_repo/` 中的 `.md` 文件，并支持多选后批量导出：

- DOCX：
  - 导出范围可选“整个文档”或“仅表格”。仅表格模式只输出文档中的全部表格，没有表格时报告错误。
  - 可选择预设模板（见下文）。
- Excel：可选 `merge` 或 `expand` 模式，同样支持预设模板。
  - `merge` 保留 HTML 表格 `rowspan` / `colspan` 对应的真实合并单元格。
  - `expand` 不创建合并区域，而将原合并区域的值重复填入每个覆盖单元格。

如果选中的 Markdown 没有表格，Excel 导出与 DOCX 仅表格导出会返回清晰的“Markdown 文件中未找到表格”错误，不会生成伪造表格。导出结果按 Markdown 主文件名写入对应仓库，同名文件自动覆盖。

### 导出模板

DOCX 和 Excel 导出共用一套模板体系，模板是保存在 `config_repo/` 中的 YAML 文件，内容为内置默认配置的局部覆盖（顶层只允许 `docx` 和 `excel` 两个键）。首次启动会自动生成三个内置预设：

- `默认模板`：完整的内置默认配置，可作为自定义模板的起点。
- `正式公文`：宋体正文、黑体标题、页脚页码的公文风格 DOCX 覆盖。
- `清爽绿表格`：绿色表头的 Excel 覆盖。

导出区可随时“选择预设模板”并在选中后直接下载该模板 YAML；也可以在“上传并保存新模板”中上传自定义 YAML、填写模板名称后保存为新的预设（同名覆盖）。模板上传时会校验 UTF-8 编码、YAML 语法、顶层键和枚举取值，校验失败会给出具体原因。

### 文件管理

「文件管理」标签页支持对四个仓库的统一管理：

- Markdown / DOCX / Excel / 模板仓库的文件浏览（文件名、大小、修改时间）。
- 每个文件的单独下载按钮。
- 多选后一键打包下载 ZIP。
- 多选后批量删除（删除失败的单个文件会逐条提示）。
- Markdown 仓库支持直接上传 `.md` 文件（例如已有 Markdown 想用内置导出重新排版时）。
- 模板仓库支持上传 YAML 保存为预设模板。

### Excel Sheet 命名差异

Streamlit 批量导出的 Excel 使用顺序名称 `Sheet1`、`Sheet2` 等。为兼容既有行为，Python API、CLI 和 Tkinter 桌面界面默认仍根据每个表格前最近的 Markdown 标题命名 Sheet；标题为空时使用 `Table_1`、`Table_2`，并自动处理非法字符、长度和重名。

## 传统桌面界面（Tkinter）

命令行启动：

```powershell
python -m md_office_exporter.desktop_gui
```

Windows 也可以双击 `run_gui.pyw`，无需保留命令行窗口。桌面界面通过文件选择器选择单个 Markdown、可选 YAML 配置和输出路径，并支持 DOCX 或 Excel（`merge` / `expand`）导出；它不支持拖放上传。覆盖已有文件时会请求确认。

## 命令行

```powershell
# Markdown -> DOCX
python -m md_office_exporter docx input.md output.docx -c config.example.yaml

# Markdown -> DOCX，仅导出表格
python -m md_office_exporter docx input.md tables.docx --tables-only

# Markdown 表格 -> Excel，保留合并单元格
python -m md_office_exporter excel input.md tables.xlsx -c config.example.yaml --merge-mode merge

# Markdown 表格 -> Excel，展开并重复填充值
python -m md_office_exporter excel input.md tables_expanded.xlsx -c config.example.yaml --merge-mode expand

# 同时生成 DOCX 和 Excel
python -m md_office_exporter all input.md --docx output.docx --excel tables.xlsx -c config.example.yaml --merge-mode merge
```

安装项目后，也可以将 `python -m md_office_exporter` 替换为 `md-office-export`。

## Python API

```python
from md_office_exporter import (
    convert_markdown_to_docx,
    export_markdown_tables_to_excel,
)

convert_markdown_to_docx(
    'input.md',
    'output.docx',
    config_path='config.example.yaml',  # 可选
)

export_markdown_tables_to_excel(
    'input.md',
    'tables.xlsx',
    merge_mode='merge',                 # 或 'expand'
    config_path='config.example.yaml',  # 可选
)
```

配置也可以作为字典传入，并会覆盖内置默认配置：

```python
convert_markdown_to_docx(
    'input.md',
    'output.docx',
    config={
        'docx': {
            'body': {'font': 'SimSun', 'font_size_pt': 12},
            'page': {'orientation': 'landscape'},
        }
    },
)
```

`convert_markdown_to_docx` 还支持 `tables_only=True`，只导出文档中的全部表格：

```python
convert_markdown_to_docx('input.md', 'tables.docx', tables_only=True)
```

## 合并单元格的 Markdown 写法

标准 Markdown 管道表格没有合并单元格语法。普通表格可以继续使用管道语法；需要合并时，请使用 Markdown 中嵌入的 HTML 表格：

```html
<table>
  <tr>
    <th rowspan=2>部门</th>
    <th colspan=2>季度</th>
  </tr>
  <tr><th>Q1</th><th>Q2</th></tr>
  <tr><td rowspan=2>研发</td><td>10</td><td>12</td></tr>
  <tr><td colspan=2>合计 22</td></tr>
</table>
```

`merge` 模式生成真实合并区域；`expand` 模式不合并，并把各个锚点值填入原合并区域覆盖的每个单元格。

## 当前支持范围

DOCX 支持常见块元素，包括标题、段落、粗体、斜体、行内代码、链接、有序/无序列表、引用、代码块、水平线、GFM 管道表格和 HTML 表格。合并单元格只来自 HTML `rowspan` / `colspan`；管道表格始终是普通网格。图片语法当前导出为图片占位文本，不下载远程图片。输入 Markdown 按 UTF-8 读取。

## 测试

完整测试套件（包括 pytest 风格的处理器测试和 Streamlit `AppTest`）：

```powershell
python -m pip install pytest
```

`pytest` 仅用于测试和开发，不是项目运行时依赖。安装后运行：

```powershell
python -m pytest -q
```

语法编译检查：

```powershell
python -m compileall -q md_office_exporter streamlit_app.py run_gui.pyw
```

示例文件位于 `examples/sample.md`。
