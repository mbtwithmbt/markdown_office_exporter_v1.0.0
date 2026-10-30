"""Command-line interface."""

from __future__ import annotations

import argparse
from pathlib import Path

from .docx_exporter import convert_markdown_to_docx
from .excel_exporter import export_markdown_tables_to_excel


def _build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="md-office-export",
        description="Markdown 导出为 DOCX，并将表格导出为 XLSX。",
    )
    subparsers = parser.add_subparsers(dest="command", required=True)

    docx_parser = subparsers.add_parser("docx", help="Markdown 转 DOCX")
    docx_parser.add_argument("input", type=Path, help="输入 Markdown 文件")
    docx_parser.add_argument("output", type=Path, help="输出 DOCX 文件")
    docx_parser.add_argument("-c", "--config", type=Path, help="YAML 配置文件")
    docx_parser.add_argument(
        "--tables-only",
        action="store_true",
        help="仅导出表格，跳过其他内容",
    )

    excel_parser = subparsers.add_parser("excel", help="提取表格到 Excel")
    excel_parser.add_argument("input", type=Path, help="输入 Markdown 文件")
    excel_parser.add_argument("output", type=Path, help="输出 XLSX 文件")
    excel_parser.add_argument("-c", "--config", type=Path, help="YAML 配置文件")
    excel_parser.add_argument(
        "--merge-mode",
        choices=("merge", "expand"),
        help="merge=保留合并；expand=拆分并填充值",
    )

    all_parser = subparsers.add_parser("all", help="同时导出 DOCX 和 XLSX")
    all_parser.add_argument("input", type=Path, help="输入 Markdown 文件")
    all_parser.add_argument("--docx", required=True, type=Path, help="输出 DOCX")
    all_parser.add_argument("--excel", required=True, type=Path, help="输出 XLSX")
    all_parser.add_argument("-c", "--config", type=Path, help="YAML 配置文件")
    all_parser.add_argument(
        "--merge-mode",
        choices=("merge", "expand"),
        help="Excel 合并单元格模式",
    )
    all_parser.add_argument(
        "--tables-only",
        action="store_true",
        help="DOCX 仅导出表格，跳过其他内容",
    )
    return parser


def main() -> int:
    args = _build_parser().parse_args()
    if args.command == "docx":
        result = convert_markdown_to_docx(
            args.input,
            args.output,
            args.config,
            tables_only=args.tables_only,
        )
        print(f"DOCX 已生成: {result}")
    elif args.command == "excel":
        result = export_markdown_tables_to_excel(
            args.input,
            args.output,
            merge_mode=args.merge_mode,
            config_path=args.config,
        )
        print(f"Excel 已生成: {result}")
    else:
        docx_result = convert_markdown_to_docx(
            args.input,
            args.docx,
            args.config,
            tables_only=args.tables_only,
        )
        excel_result = export_markdown_tables_to_excel(
            args.input,
            args.excel,
            merge_mode=args.merge_mode,
            config_path=args.config,
        )
        print(f"DOCX 已生成: {docx_result}")
        print(f"Excel 已生成: {excel_result}")
    return 0
