"""Lightweight desktop GUI for Markdown Office exports."""

from __future__ import annotations

from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from .docx_exporter import convert_markdown_to_docx
from .excel_exporter import export_markdown_tables_to_excel
from .gui import EXPORT_EXTENSIONS, suggested_output_path, validate_export_request


def perform_export(
    markdown_path: Path,
    output_path: Path,
    export_type: str,
    config_path: Path | None = None,
    merge_mode: str = "merge",
) -> Path:
    """Dispatch to the correct exporter based on type."""

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


class ExporterApp:
    """Single-page Tkinter form for exporting Markdown to DOCX or Excel."""

    def __init__(self, master: tk.Tk) -> None:
        self.master = master
        master.title("Markdown Office Exporter")
        master.geometry("640x320")
        master.minsize(480, 280)

        # Variables
        self.markdown_var = tk.StringVar()
        self.config_var = tk.StringVar()
        self.export_type_var = tk.StringVar(value="docx")
        self.output_var = tk.StringVar()
        self.merge_mode_var = tk.StringVar(value="merge")
        self.status_var = tk.StringVar(value="请选择 Markdown 文件")

        # Layout
        frame = ttk.Frame(master, padding="12")
        frame.grid(sticky="nsew")
        master.columnconfigure(0, weight=1)
        master.rowconfigure(0, weight=1)
        frame.columnconfigure(1, weight=1)

        row = 0
        ttk.Label(frame, text="Markdown 文件：").grid(row=row, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.markdown_var).grid(
            row=row, column=1, sticky="ew", padx=4
        )
        ttk.Button(frame, text="浏览…", command=self.browse_markdown).grid(
            row=row, column=2, sticky="w"
        )

        row += 1
        ttk.Label(frame, text="配置文件：").grid(row=row, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.config_var).grid(
            row=row, column=1, sticky="ew", padx=4
        )
        ttk.Button(frame, text="浏览…", command=self.browse_config).grid(
            row=row, column=2, sticky="w"
        )

        row += 1
        ttk.Label(frame, text="输出文件：").grid(row=row, column=0, sticky="w")
        ttk.Entry(frame, textvariable=self.output_var).grid(
            row=row, column=1, sticky="ew", padx=4
        )
        ttk.Button(frame, text="浏览…", command=self.browse_output).grid(
            row=row, column=2, sticky="w"
        )

        row += 1
        type_frame = ttk.Frame(frame)
        type_frame.grid(row=row, column=0, columnspan=3, sticky="w", pady=4)
        ttk.Label(type_frame, text="导出类型：").pack(side=tk.LEFT)
        for val, label in (
            ("docx", "DOCX"),
            ("excel", "Excel"),
        ):
            ttk.Radiobutton(
                type_frame,
                text=label,
                variable=self.export_type_var,
                value=val,
                command=self.on_export_type_changed,
            ).pack(side=tk.LEFT, padx=4)

        row += 1
        merge_frame = ttk.Frame(frame)
        merge_frame.grid(row=row, column=0, columnspan=3, sticky="w", pady=4)
        ttk.Label(merge_frame, text="Excel 合并方式：").pack(side=tk.LEFT)
        for val, label in (
            ("merge", "合并 (merge)"),
            ("expand", "展开 (expand)"),
        ):
            ttk.Radiobutton(
                merge_frame,
                text=label,
                variable=self.merge_mode_var,
                value=val,
            ).pack(side=tk.LEFT, padx=4)
        self.merge_radios: list[ttk.Radiobutton] = []
        for child in merge_frame.winfo_children():
            if isinstance(child, ttk.Radiobutton):
                self.merge_radios.append(child)
        self._set_merge_state("disabled")

        row += 1
        self.export_button = ttk.Button(
            frame, text="导出", command=self.export_file
        )
        self.export_button.grid(row=row, column=0, columnspan=3, pady=8)

        row += 1
        self.status_label = ttk.Label(frame, textvariable=self.status_var, wraplength=600)
        self.status_label.grid(row=row, column=0, columnspan=3, sticky="ew")

    # ------------------------------------------------------------------ #
    # Browse helpers
    # ------------------------------------------------------------------ #

    def browse_markdown(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 Markdown 文件",
            filetypes=[
                ("Markdown 文件", "*.md *.markdown"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.markdown_var.set(path)
            export_type = self.export_type_var.get()
            suggested = suggested_output_path(Path(path), export_type)
            self.output_var.set(str(suggested))

    def browse_config(self) -> None:
        path = filedialog.askopenfilename(
            title="选择 YAML 配置文件",
            filetypes=[
                ("YAML 文件", "*.yaml *.yml"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.config_var.set(path)

    def browse_output(self) -> None:
        export_type = self.export_type_var.get()
        ext = EXPORT_EXTENSIONS[export_type]
        path = filedialog.asksaveasfilename(
            title="选择输出文件",
            defaultextension=ext,
            filetypes=[
                (f"{ext.upper()[1:]} 文件", f"*{ext}"),
                ("所有文件", "*.*"),
            ],
        )
        if path:
            self.output_var.set(path)

    # ------------------------------------------------------------------ #
    # Mode change
    # ------------------------------------------------------------------ #

    def on_export_type_changed(self) -> None:
        export_type = self.export_type_var.get()
        if export_type == "docx":
            self._set_merge_state("disabled")
        else:
            self._set_merge_state("normal")

        output = self.output_var.get().strip()
        if output:
            try:
                new_path = suggested_output_path(Path(output), export_type)
                self.output_var.set(str(new_path))
            except ValueError:
                pass  # keep existing path if suggestion fails

    def _set_merge_state(self, state: str) -> None:
        for radio in self.merge_radios:
            radio.configure(state=state)

    # ------------------------------------------------------------------ #
    # Export action
    # ------------------------------------------------------------------ #

    def export_file(self) -> None:
        try:
            source, output, config = validate_export_request(
                self.markdown_var.get(),
                self.output_var.get(),
                self.export_type_var.get(),
                self.config_var.get() or None,
            )
        except ValueError as exc:
            self.status_var.set(f"验证失败：{exc}")
            messagebox.showerror("输入错误", str(exc))
            return

        if output.exists():
            if not messagebox.askyesno(
                "文件已存在",
                f"目标文件已存在，是否覆盖？\n{output}",
            ):
                return

        self.export_button.configure(state="disabled")
        self.status_var.set("正在导出…")
        self.master.update_idletasks()

        try:
            result = perform_export(
                source,
                output,
                self.export_type_var.get(),
                config_path=config,
                merge_mode=self.merge_mode_var.get(),
            )
            self.status_var.set(f"导出完成：{result}")
            messagebox.showinfo("成功", f"文件已保存到：\n{result}")
        except Exception as exc:
            self.status_var.set(f"导出失败：{exc}")
            messagebox.showerror("导出失败", str(exc))
        finally:
            self.export_button.configure(state="normal")


def main() -> int:
    root = tk.Tk()
    ExporterApp(root)
    root.mainloop()
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
