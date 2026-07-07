from __future__ import annotations

import threading
import tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox, ttk
from typing import List

import customtkinter as ctk

from .comparer import CompareOptions, ContentComparer, DiffItem
from .constants import APP_NAME, APP_VERSION, COMPARE_MODES, SUPPORTED_EXTENSIONS
from .excel_exporter import export_diffs_to_excel
from .file_readers import FileReadError, get_file_info, read_file
from .normalizer import NormalizeOptions, clean_ignore_fields


class FileDocCompareApp(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_NAME} v{APP_VERSION}")
        self.geometry("1280x820")
        self.minsize(1120, 720)

        ctk.set_appearance_mode("System")
        ctk.set_default_color_theme("blue")

        self.actual_path_var = tk.StringVar(value="")
        self.document_path_var = tk.StringVar(value="")
        self.actual_status_var = tk.StringVar(value="未选择文件")
        self.document_status_var = tk.StringVar(value="未选择文件")
        self.compare_mode_var = tk.StringVar(value=COMPARE_MODES[2])
        self.ignore_spaces_var = tk.BooleanVar(value=True)
        self.ignore_newlines_var = tk.BooleanVar(value=False)
        self.ignore_case_var = tk.BooleanVar(value=False)
        self.ignore_fields_var = tk.StringVar(value="时间戳, timestamp, version, 版本号")
        self.search_var = tk.StringVar(value="")
        self.summary_var = tk.StringVar(value="请选择文件后开始比对")

        self.all_diffs: List[DiffItem] = []
        self.visible_diffs: List[DiffItem] = []
        self.is_working = False

        self._build_ui()
        self._bind_events()

    def _build_ui(self):
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(3, weight=1)

        header = ctk.CTkFrame(self, corner_radius=0)
        header.grid(row=0, column=0, sticky="ew")
        header.grid_columnconfigure(0, weight=1)

        title = ctk.CTkLabel(
            header,
            text=f"{APP_NAME}",
            font=ctk.CTkFont(size=22, weight="bold"),
            anchor="w",
        )
        title.grid(row=0, column=0, padx=18, pady=(14, 2), sticky="w")

        subtitle = ctk.CTkLabel(
            header,
            text="本地读取、本地比对、不上传、不修改源文件",
            text_color=("gray35", "gray70"),
            anchor="w",
        )
        subtitle.grid(row=1, column=0, padx=18, pady=(0, 12), sticky="w")

        file_frame = ctk.CTkFrame(self)
        file_frame.grid(row=1, column=0, padx=16, pady=(14, 8), sticky="ew")
        file_frame.grid_columnconfigure(0, weight=1)
        file_frame.grid_columnconfigure(1, weight=1)

        self._build_file_panel(
            parent=file_frame,
            column=0,
            title="待比对文件（实际配置文件 / 导出文件）",
            path_var=self.actual_path_var,
            status_var=self.actual_status_var,
            button_text="选择待比对文件",
            command=lambda: self.select_file("actual"),
        )
        self._build_file_panel(
            parent=file_frame,
            column=1,
            title="参考文档（策划文档 / 配置说明）",
            path_var=self.document_path_var,
            status_var=self.document_status_var,
            button_text="选择参考文档",
            command=lambda: self.select_file("document"),
        )

        option_frame = ctk.CTkFrame(self)
        option_frame.grid(row=2, column=0, padx=16, pady=8, sticky="ew")
        option_frame.grid_columnconfigure(7, weight=1)

        ctk.CTkLabel(option_frame, text="比对模式").grid(row=0, column=0, padx=(12, 6), pady=12, sticky="w")
        mode_menu = ctk.CTkOptionMenu(option_frame, values=COMPARE_MODES, variable=self.compare_mode_var, width=170)
        mode_menu.grid(row=0, column=1, padx=6, pady=12, sticky="w")

        ctk.CTkCheckBox(option_frame, text="忽略空格", variable=self.ignore_spaces_var, width=90).grid(
            row=0, column=2, padx=6, pady=12, sticky="w"
        )
        ctk.CTkCheckBox(option_frame, text="忽略换行", variable=self.ignore_newlines_var, width=90).grid(
            row=0, column=3, padx=6, pady=12, sticky="w"
        )
        ctk.CTkCheckBox(option_frame, text="忽略大小写", variable=self.ignore_case_var, width=105).grid(
            row=0, column=4, padx=6, pady=12, sticky="w"
        )

        ctk.CTkLabel(option_frame, text="忽略字段").grid(row=0, column=5, padx=(12, 6), pady=12, sticky="w")
        ignore_entry = ctk.CTkEntry(option_frame, textvariable=self.ignore_fields_var, placeholder_text="字段1, 字段2")
        ignore_entry.grid(row=0, column=6, padx=6, pady=12, sticky="ew")
        option_frame.grid_columnconfigure(6, weight=1)

        self.compare_button = ctk.CTkButton(option_frame, text="开始比对", width=110, command=self.start_compare)
        self.compare_button.grid(row=0, column=8, padx=(8, 12), pady=12, sticky="e")

        result_frame = ctk.CTkFrame(self)
        result_frame.grid(row=3, column=0, padx=16, pady=(8, 10), sticky="nsew")
        result_frame.grid_columnconfigure(0, weight=1)
        result_frame.grid_rowconfigure(2, weight=1)

        toolbar = ctk.CTkFrame(result_frame, fg_color="transparent")
        toolbar.grid(row=0, column=0, sticky="ew", padx=10, pady=(10, 4))
        toolbar.grid_columnconfigure(1, weight=1)

        ctk.CTkLabel(toolbar, textvariable=self.summary_var, anchor="w").grid(row=0, column=0, padx=(0, 12), sticky="w")
        search_entry = ctk.CTkEntry(toolbar, textvariable=self.search_var, placeholder_text="搜索差异结果")
        search_entry.grid(row=0, column=1, padx=6, sticky="ew")
        ctk.CTkButton(toolbar, text="复制结果", width=100, command=self.copy_visible_diffs).grid(row=0, column=2, padx=6)
        ctk.CTkButton(toolbar, text="导出 Excel", width=110, command=self.export_excel).grid(row=0, column=3, padx=(6, 0))

        tip = ctk.CTkLabel(
            result_frame,
            text="提示：推荐使用“策划文档 vs 配置文件”；它会优先比对配置值、奖励内容、档位、开关状态。复杂自然语言仍建议配合人工复核。",
            text_color=("gray35", "gray70"),
            anchor="w",
        )
        tip.grid(row=1, column=0, sticky="ew", padx=10, pady=(0, 6))

        table_frame = ctk.CTkFrame(result_frame)
        table_frame.grid(row=2, column=0, padx=10, pady=(0, 10), sticky="nsew")
        table_frame.grid_columnconfigure(0, weight=1)
        table_frame.grid_rowconfigure(0, weight=1)

        columns = ("index", "diff_type", "document_content", "file_content", "location", "remark")
        self.tree = ttk.Treeview(table_frame, columns=columns, show="headings", height=18)
        self.tree.heading("index", text="序号")
        self.tree.heading("diff_type", text="差异类型")
        self.tree.heading("document_content", text="文档内容")
        self.tree.heading("file_content", text="文件内容")
        self.tree.heading("location", text="所在位置")
        self.tree.heading("remark", text="备注")

        self.tree.column("index", width=60, anchor="center", stretch=False)
        self.tree.column("diff_type", width=100, anchor="center", stretch=False)
        self.tree.column("document_content", width=320, anchor="w")
        self.tree.column("file_content", width=320, anchor="w")
        self.tree.column("location", width=180, anchor="w")
        self.tree.column("remark", width=240, anchor="w")

        y_scroll = ttk.Scrollbar(table_frame, orient="vertical", command=self.tree.yview)
        x_scroll = ttk.Scrollbar(table_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=y_scroll.set, xscrollcommand=x_scroll.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        y_scroll.grid(row=0, column=1, sticky="ns")
        x_scroll.grid(row=1, column=0, sticky="ew")

    def _build_file_panel(self, parent, column, title, path_var, status_var, button_text, command):
        panel = ctk.CTkFrame(parent)
        panel.grid(row=0, column=column, padx=10, pady=10, sticky="nsew")
        panel.grid_columnconfigure(0, weight=1)

        ctk.CTkLabel(panel, text=title, font=ctk.CTkFont(size=15, weight="bold"), anchor="w").grid(
            row=0, column=0, padx=12, pady=(12, 6), sticky="ew"
        )
        ctk.CTkEntry(panel, textvariable=path_var, state="readonly").grid(
            row=1, column=0, padx=12, pady=6, sticky="ew"
        )
        ctk.CTkLabel(panel, textvariable=status_var, text_color=("gray35", "gray70"), anchor="w").grid(
            row=2, column=0, padx=12, pady=(2, 8), sticky="ew"
        )
        ctk.CTkButton(panel, text=button_text, command=command).grid(row=3, column=0, padx=12, pady=(2, 12), sticky="ew")

    def _bind_events(self):
        self.search_var.trace_add("write", lambda *_: self.apply_search_filter())

    def select_file(self, role: str):
        filetypes = [
            ("支持的文件", " ".join(f"*{ext}" for ext in sorted(SUPPORTED_EXTENSIONS))),
            ("所有文件", "*.*"),
        ]
        path = filedialog.askopenfilename(title="选择文件", filetypes=filetypes)
        if not path:
            return

        if role == "actual":
            self.actual_path_var.set(path)
            self.actual_status_var.set(f"已选择：{get_file_info(path)}")
        else:
            self.document_path_var.set(path)
            self.document_status_var.set(f"已选择：{get_file_info(path)}")

    def start_compare(self):
        if self.is_working:
            return

        actual_path = self.actual_path_var.get().strip()
        document_path = self.document_path_var.get().strip()
        if not actual_path or not document_path:
            messagebox.showwarning("提示", "请先选择待比对文件和参考文档")
            return

        self.is_working = True
        self.compare_button.configure(state="disabled", text="比对中...")
        self.summary_var.set("正在读取文件并比对...")
        self.clear_table()

        options = self.get_compare_options()
        thread = threading.Thread(
            target=self._compare_worker,
            args=(actual_path, document_path, options),
            daemon=True,
        )
        thread.start()

    def get_compare_options(self) -> CompareOptions:
        normalize_options = NormalizeOptions(
            ignore_spaces=self.ignore_spaces_var.get(),
            ignore_newlines=self.ignore_newlines_var.get(),
            ignore_case=self.ignore_case_var.get(),
        )
        ignore_fields = clean_ignore_fields(self.ignore_fields_var.get(), normalize_options)
        return CompareOptions(
            mode=self.compare_mode_var.get(),
            normalize=normalize_options,
            ignore_fields=ignore_fields,
        )

    def _compare_worker(self, actual_path: str, document_path: str, options: CompareOptions):
        try:
            actual = read_file(actual_path)
            document = read_file(document_path)
            comparer = ContentComparer()
            diffs = comparer.compare(actual=actual, document=document, options=options)
            self.after(0, lambda: self._on_compare_success(actual, document, diffs))
        except FileReadError as exc:
            message = str(exc)
            self.after(0, lambda msg=message: self._on_compare_error(msg))
        except Exception as exc:
            message = f"比对失败：{exc}"
            self.after(0, lambda msg=message: self._on_compare_error(msg))

    def _on_compare_success(self, actual, document, diffs: List[DiffItem]):
        self.is_working = False
        self.compare_button.configure(state="normal", text="开始比对")
        self.all_diffs = diffs
        self.visible_diffs = diffs
        self.actual_status_var.set(
            f"读取成功：{actual.ext or '无扩展名'} / 行数 {len(actual.lines)} / 字段 {len(actual.fields)}"
        )
        self.document_status_var.set(
            f"读取成功：{document.ext or '无扩展名'} / 行数 {len(document.lines)} / 字段 {len(document.fields)}"
        )
        self.apply_search_filter()
        if not diffs:
            self.summary_var.set("比对完成：未发现差异")
        else:
            counts = self._count_by_type(diffs)
            self.summary_var.set(
                f"比对完成：共 {len(diffs)} 条差异 | 缺失 {counts.get('缺失', 0)} | 多余 {counts.get('多余', 0)} | 值不一致 {counts.get('值不一致', 0)}"
            )

    def _on_compare_error(self, message: str):
        self.is_working = False
        self.compare_button.configure(state="normal", text="开始比对")
        self.summary_var.set("比对失败")
        messagebox.showerror("错误", message)

    @staticmethod
    def _count_by_type(diffs: List[DiffItem]):
        result = {}
        for item in diffs:
            result[item.diff_type] = result.get(item.diff_type, 0) + 1
        return result

    def apply_search_filter(self):
        keyword = self.search_var.get().strip().lower()
        if not keyword:
            self.visible_diffs = self.all_diffs
        else:
            self.visible_diffs = [item for item in self.all_diffs if keyword in self._diff_to_text(item).lower()]
        self.refresh_table(self.visible_diffs)

    def refresh_table(self, diffs: List[DiffItem]):
        self.clear_table()
        for item in diffs:
            self.tree.insert(
                "",
                "end",
                values=(
                    item.index,
                    item.diff_type,
                    self._shorten(item.document_content),
                    self._shorten(item.file_content),
                    self._shorten(item.location, 120),
                    self._shorten(item.remark, 160),
                ),
            )

    def clear_table(self):
        for item in self.tree.get_children():
            self.tree.delete(item)

    @staticmethod
    def _shorten(text: str, max_len: int = 220) -> str:
        text = "" if text is None else str(text).replace("\n", " ⏎ ")
        return text if len(text) <= max_len else text[:max_len] + "..."

    @staticmethod
    def _diff_to_text(item: DiffItem) -> str:
        return "\t".join([
            str(item.index),
            item.diff_type,
            item.document_content,
            item.file_content,
            item.location,
            item.remark,
        ])

    def copy_visible_diffs(self):
        if not self.visible_diffs:
            messagebox.showinfo("提示", "当前没有可复制的差异结果")
            return

        headers = ["序号", "差异类型", "文档内容", "文件内容", "所在位置", "备注"]
        rows = ["\t".join(headers)]
        for item in self.visible_diffs:
            rows.append(self._diff_to_text(item))

        text = "\n".join(rows)
        self.clipboard_clear()
        self.clipboard_append(text)
        messagebox.showinfo("提示", f"已复制 {len(self.visible_diffs)} 条差异结果")

    def export_excel(self):
        if not self.visible_diffs:
            messagebox.showinfo("提示", "当前没有可导出的差异结果")
            return

        output_path = filedialog.asksaveasfilename(
            title="导出差异报告",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="差异比对结果.xlsx",
        )
        if not output_path:
            return

        try:
            export_diffs_to_excel(self.visible_diffs, output_path)
            messagebox.showinfo("导出成功", f"已导出：{Path(output_path)}")
        except Exception as exc:
            messagebox.showerror("导出失败", str(exc))
