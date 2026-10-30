# Markdown Office Exporter 试验窗口设计

## 目标

为现有函数化导出能力增加一个轻量 Windows 桌面窗口，方便手动选择 Markdown 文件并试验 DOCX 或 Excel 导出。窗口不重复实现文档解析或导出逻辑，只调用现有公共函数。

## 范围

- 选择一个 Markdown 输入文件。
- 可选选择一个 YAML 配置文件；未选择时使用程序内置默认配置。
- 在 DOCX 和 Excel 两种导出类型之间切换。
- 自动根据 Markdown 文件名和导出类型生成同目录输出路径，同时允许手动修改。
- Excel 模式支持 `merge`（保留合并单元格）与 `expand`（拆分并将锚点值填入区域内所有单元格）。
- 目标文件存在时确认是否覆盖。
- 在窗口内反馈成功路径或中文错误信息。

不包含格式可视化编辑、批量文件、导出历史、导出后自动打开或网页界面。

## 结构

新增 `md_office_exporter/gui.py`，负责窗口布局、字段状态、输入校验、路径推导和调用导出函数。新增根目录 `run_gui.pyw`，作为 Windows 下可双击且不显示控制台的启动入口。现有 `docx_exporter.py`、`excel_exporter.py` 和解析器保持为业务逻辑唯一来源。

窗口调用关系：

- DOCX：`convert_markdown_to_docx(markdown_path, output_path, config_path)`
- Excel：`export_markdown_tables_to_excel(markdown_path, output_path, merge_mode, config_path)`

## 界面与数据流

窗口采用单页表单，依次包含：Markdown 文件、可选 YAML 配置、导出类型、输出文件、Excel 合并方式、导出按钮和状态文本。

选择 Markdown 后，输出路径设为输入文件同目录、同名并使用当前类型对应的 `.docx` 或 `.xlsx` 后缀。切换导出类型时更新输出后缀；用户手动选择的目录和文件主名保留。DOCX 模式禁用 Excel 合并方式，Excel 模式启用并默认选择 `merge`。

点击导出后先校验 Markdown 文件、可选配置文件和输出扩展名。若目标文件已存在，通过确认框询问是否覆盖。导出期间禁用导出按钮；结束后恢复。成功时显示绝对输出路径，失败时捕获异常并显示中文错误提示，窗口保持可用。

## 错误处理

- 输入或配置路径不存在：阻止导出并指出对应字段。
- 输出路径为空或扩展名与类型不符：阻止导出并提示正确扩展名。
- 没有可导出的表格、配置无效、文件被占用或保存失败：保留底层异常信息，并用错误弹窗展示。
- 用户取消文件选择或覆盖确认：不改变现有文件，不视为程序错误。

## 测试与验收

将可独立验证的路径推导和输入校验逻辑写成小函数，避免测试依赖图形显示器。新增测试覆盖：

- Markdown 路径到 `.docx` / `.xlsx` 路径的推导。
- 切换类型时输出扩展名变化。
- 缺失输入、缺失配置和错误扩展名的校验。
- GUI 模块可导入且启动入口语法有效。

继续运行现有导出测试，确保 DOCX、Excel `merge`、Excel `expand` 和表格解析行为不回归。验收方式是在 Windows 上双击 `run_gui.pyw` 或执行 `python -m md_office_exporter.gui`，分别生成 DOCX、合并模式 Excel 和展开模式 Excel。
