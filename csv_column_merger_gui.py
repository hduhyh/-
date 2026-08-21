# -*- coding: utf-8 -*-
"""Tkinter GUI for extracting template columns from CSV/XLSX files."""

from __future__ import annotations

import json
import os
import queue
import re
import threading
import ctypes
from pathlib import Path
from typing import Any, Dict

import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from merger_core import MergeResult, merge_folder


APP_NAME = "CSV列名提取合并"
INVALID_WINDOWS_FILENAME_CHARS = re.compile(r'[<>:"/\\|?*]')
DEFAULT_SETTINGS: Dict[str, Any] = {
    "input_dir": "",
    "output_dir": "",
    "template_file": "",
    "output_filename": "merged_extract_output.csv",
    "recursive": False,
    "excel_all_sheets": False,
}


def enable_windows_high_dpi() -> None:
    """Enable crisp rendering on scaled Windows displays before Tk creates a window."""

    if os.name != "nt":
        return

    try:
        # Windows 10 Creators Update and newer: per-monitor DPI awareness v2.
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except (AttributeError, OSError):
        pass

    try:
        # Windows 8.1 fallback.
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except (AttributeError, OSError):
        pass

    try:
        # Windows 7 fallback.
        ctypes.windll.user32.SetProcessDPIAware()
    except (AttributeError, OSError):
        pass


def settings_path() -> Path:
    app_data = os.environ.get("APPDATA")
    if app_data:
        return Path(app_data) / "CSVColumnMerger" / "settings.json"
    return Path.home() / ".csv_column_merger" / "settings.json"


def load_settings() -> Dict[str, Any]:
    settings = dict(DEFAULT_SETTINGS)
    path = settings_path()
    try:
        loaded = json.loads(path.read_text(encoding="utf-8"))
        if isinstance(loaded, dict):
            settings.update({key: loaded[key] for key in settings if key in loaded})
    except (OSError, ValueError, TypeError):
        pass
    return settings


def save_settings(settings: Dict[str, Any]) -> None:
    path = settings_path()
    temp_path = path.with_suffix(".tmp")
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        temp_path.write_text(
            json.dumps(settings, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
        os.replace(temp_path, path)
    except OSError:
        try:
            temp_path.unlink(missing_ok=True)
        except OSError:
            pass


class MergerApplication:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("820x600")
        self.root.minsize(720, 520)

        settings = load_settings()
        self.input_dir = tk.StringVar(value=str(settings["input_dir"]))
        self.output_dir = tk.StringVar(value=str(settings["output_dir"]))
        self.template_file = tk.StringVar(value=str(settings["template_file"]))
        self.output_filename = tk.StringVar(value=str(settings["output_filename"]))
        self.recursive = tk.BooleanVar(value=bool(settings["recursive"]))
        self.excel_all_sheets = tk.BooleanVar(value=bool(settings["excel_all_sheets"]))
        self.status_text = tk.StringVar(value="请选择路径和模板文件")

        self.events: "queue.Queue[tuple]" = queue.Queue()
        self.is_running = False
        self._build_ui()
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)
        self.root.after(100, self._poll_events)

    def _build_ui(self) -> None:
        outer = ttk.Frame(self.root, padding=18)
        outer.pack(fill="both", expand=True)
        outer.columnconfigure(1, weight=1)
        outer.rowconfigure(8, weight=1)

        title = ttk.Label(outer, text=APP_NAME, font=("Microsoft YaHei UI", 17, "bold"))
        title.grid(row=0, column=0, columnspan=3, sticky="w", pady=(0, 6))
        description = ttk.Label(
            outer,
            text="按模板列名批量提取输入文件夹中的 CSV / XLSX 数据，并合并为一个 CSV。",
        )
        description.grid(row=1, column=0, columnspan=3, sticky="w", pady=(0, 18))

        self._add_path_row(
            outer,
            row=2,
            label="输入文件夹",
            variable=self.input_dir,
            command=self._choose_input_dir,
        )
        self._add_path_row(
            outer,
            row=3,
            label="输出文件夹",
            variable=self.output_dir,
            command=self._choose_output_dir,
        )
        self._add_path_row(
            outer,
            row=4,
            label="模板文件",
            variable=self.template_file,
            command=self._choose_template,
        )

        ttk.Label(outer, text="输出文件名").grid(row=5, column=0, sticky="w", padx=(0, 12), pady=7)
        ttk.Entry(outer, textvariable=self.output_filename).grid(
            row=5, column=1, columnspan=2, sticky="ew", pady=7
        )

        option_frame = ttk.Frame(outer)
        option_frame.grid(row=6, column=0, columnspan=3, sticky="ew", pady=(8, 12))
        ttk.Checkbutton(option_frame, text="递归扫描子文件夹", variable=self.recursive).pack(
            side="left", padx=(0, 24)
        )
        ttk.Checkbutton(
            option_frame,
            text="读取 XLSX 的全部工作表（默认仅第一个）",
            variable=self.excel_all_sheets,
        ).pack(side="left")

        action_frame = ttk.Frame(outer)
        action_frame.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(0, 10))
        action_frame.columnconfigure(1, weight=1)
        self.run_button = ttk.Button(
            action_frame,
            text="开始提取并合并",
            command=self._start_merge,
            style="Primary.TButton",
            width=20,
        )
        self.run_button.grid(row=0, column=0, sticky="w", padx=(0, 16))
        self.progress = ttk.Progressbar(action_frame, mode="determinate", maximum=100)
        self.progress.grid(row=0, column=1, sticky="ew")

        log_frame = ttk.LabelFrame(outer, text="处理日志", padding=8)
        log_frame.grid(row=8, column=0, columnspan=3, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_box = scrolledtext.ScrolledText(
            log_frame,
            height=12,
            wrap="word",
            state="disabled",
            font=("Consolas", 10),
        )
        self.log_box.grid(row=0, column=0, sticky="nsew")

        ttk.Label(outer, textvariable=self.status_text).grid(
            row=9, column=0, columnspan=3, sticky="w", pady=(10, 0)
        )

    @staticmethod
    def _add_path_row(
        parent: ttk.Frame,
        *,
        row: int,
        label: str,
        variable: tk.StringVar,
        command: Any,
    ) -> None:
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=(0, 12), pady=7)
        ttk.Entry(parent, textvariable=variable).grid(row=row, column=1, sticky="ew", pady=7)
        ttk.Button(parent, text="浏览…", command=command, width=10).grid(
            row=row, column=2, sticky="e", padx=(10, 0), pady=7
        )

    @staticmethod
    def _existing_dir(value: str) -> str:
        path = Path(value).expanduser() if value else None
        if path and path.is_dir():
            return str(path)
        if path and path.is_file():
            return str(path.parent)
        return str(Path.home())

    def _choose_input_dir(self) -> None:
        selected = filedialog.askdirectory(
            title="选择输入文件夹",
            initialdir=self._existing_dir(self.input_dir.get()),
        )
        if selected:
            self.input_dir.set(selected)
            self._save_current_settings()

    def _choose_output_dir(self) -> None:
        selected = filedialog.askdirectory(
            title="选择输出文件夹",
            initialdir=self._existing_dir(self.output_dir.get()),
        )
        if selected:
            self.output_dir.set(selected)
            self._save_current_settings()

    def _choose_template(self) -> None:
        selected = filedialog.askopenfilename(
            title="选择列名模板",
            initialdir=self._existing_dir(self.template_file.get()),
            filetypes=(("模板文件", "*.xlsx *.csv"), ("Excel 工作簿", "*.xlsx"), ("CSV 文件", "*.csv")),
        )
        if selected:
            self.template_file.set(selected)
            self._save_current_settings()

    def _current_settings(self) -> Dict[str, Any]:
        return {
            "input_dir": self.input_dir.get().strip(),
            "output_dir": self.output_dir.get().strip(),
            "template_file": self.template_file.get().strip(),
            "output_filename": self.output_filename.get().strip(),
            "recursive": self.recursive.get(),
            "excel_all_sheets": self.excel_all_sheets.get(),
        }

    def _save_current_settings(self) -> None:
        save_settings(self._current_settings())

    def _validate_inputs(self) -> tuple[Path, Path, Path]:
        input_dir = Path(self.input_dir.get().strip())
        output_dir = Path(self.output_dir.get().strip())
        template_file = Path(self.template_file.get().strip())
        filename = self.output_filename.get().strip()

        if not input_dir.is_dir():
            raise ValueError("请选择有效的输入文件夹")
        if not output_dir.is_dir():
            raise ValueError("请选择有效的输出文件夹")
        if not template_file.is_file() or template_file.suffix.lower() not in {".xlsx", ".csv"}:
            raise ValueError("请选择有效的 CSV 或 XLSX 模板文件")
        if not filename:
            raise ValueError("请输入输出文件名")
        if Path(filename).name != filename or INVALID_WINDOWS_FILENAME_CHARS.search(filename):
            raise ValueError("输出文件名不能包含路径或字符 < > : \" / \\ | ? *")
        if not filename.lower().endswith(".csv"):
            filename += ".csv"
            self.output_filename.set(filename)

        return input_dir, output_dir / filename, template_file

    def _start_merge(self) -> None:
        if self.is_running:
            return

        try:
            input_dir, output_path, template_file = self._validate_inputs()
        except ValueError as exc:
            messagebox.showerror("参数有误", str(exc), parent=self.root)
            return

        if output_path.exists() and not messagebox.askyesno(
            "确认覆盖",
            f"输出文件已经存在：\n{output_path}\n\n是否覆盖？",
            parent=self.root,
        ):
            return

        self._save_current_settings()
        self.is_running = True
        self.run_button.configure(state="disabled")
        self.progress.configure(value=0)
        self.status_text.set("正在读取模板和源文件…")
        self._clear_log()
        self._append_log(f"输入文件夹：{input_dir}")
        self._append_log(f"模板文件：{template_file}")
        self._append_log(f"输出文件：{output_path}")

        # Tk variables must only be read on the UI thread.
        recursive = self.recursive.get()
        excel_all_sheets = self.excel_all_sheets.get()

        worker = threading.Thread(
            target=self._merge_worker,
            args=(input_dir, output_path, template_file, recursive, excel_all_sheets),
            daemon=True,
        )
        worker.start()

    def _merge_worker(
        self,
        input_dir: Path,
        output_path: Path,
        template_file: Path,
        recursive: bool,
        excel_all_sheets: bool,
    ) -> None:
        try:
            result = merge_folder(
                input_dir=input_dir,
                output_path=output_path,
                template_path=template_file,
                recursive=recursive,
                excel_all_sheets=excel_all_sheets,
                progress_callback=lambda current, total, message: self.events.put(
                    ("progress", current, total, message)
                ),
            )
            self.events.put(("done", result))
        except Exception as exc:
            self.events.put(("error", str(exc) or exc.__class__.__name__))

    def _poll_events(self) -> None:
        try:
            while True:
                event = self.events.get_nowait()
                if event[0] == "progress":
                    _, current, total, message = event
                    self.progress.configure(value=(current / total) * 100 if total else 0)
                    self.status_text.set(f"正在处理 {current}/{total}")
                    self._append_log(message)
                elif event[0] == "done":
                    self._finish_success(event[1])
                elif event[0] == "error":
                    self._finish_error(event[1])
        except queue.Empty:
            pass
        self.root.after(100, self._poll_events)

    def _finish_success(self, result: MergeResult) -> None:
        self.is_running = False
        self.run_button.configure(state="normal")
        self.progress.configure(value=100)
        failed_count = len(result.failed_files)
        self.status_text.set(
            f"完成：{result.successful_file_count}/{result.source_file_count} 个文件，{result.total_rows} 行"
        )
        self._append_log(f"输出完成：{result.output_path}")
        self._append_log(f"提取列数：{len(result.extracted_columns)}；总行数：{result.total_rows}")

        if failed_count:
            self._append_log("以下文件处理失败：")
            for name, error in result.failed_files:
                self._append_log(f"  - {name}: {error}")
            messagebox.showwarning(
                "部分完成",
                f"合并文件已生成，但有 {failed_count} 个源文件处理失败。\n请查看处理日志。",
                parent=self.root,
            )
        else:
            messagebox.showinfo(
                "处理完成",
                f"成功处理 {result.successful_file_count} 个文件、{result.total_rows} 行。\n\n{result.output_path}",
                parent=self.root,
            )

    def _finish_error(self, error: str) -> None:
        self.is_running = False
        self.run_button.configure(state="normal")
        self.progress.configure(value=0)
        self.status_text.set("处理失败")
        self._append_log(f"错误：{error}")
        messagebox.showerror("处理失败", error, parent=self.root)

    def _append_log(self, text: str) -> None:
        self.log_box.configure(state="normal")
        self.log_box.insert("end", text + "\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _clear_log(self) -> None:
        self.log_box.configure(state="normal")
        self.log_box.delete("1.0", "end")
        self.log_box.configure(state="disabled")

    def _on_close(self) -> None:
        if self.is_running:
            messagebox.showwarning("正在处理", "请等待当前任务完成后再关闭程序。", parent=self.root)
            return
        self._save_current_settings()
        self.root.destroy()


def main() -> None:
    enable_windows_high_dpi()
    root = tk.Tk()

    try:
        # Keep fonts and controls proportional after DPI awareness is enabled.
        root.tk.call("tk", "scaling", root.winfo_fpixels("1i") / 72.0)
    except tk.TclError:
        pass

    style = ttk.Style(root)
    try:
        style.theme_use("vista" if os.name == "nt" else "clam")
    except tk.TclError:
        pass
    root.option_add("*Font", ("Microsoft YaHei UI", 10))
    style.configure(".", font=("Microsoft YaHei UI", 10))
    style.configure(
        "Primary.TButton",
        font=("Microsoft YaHei UI", 12, "bold"),
        padding=(28, 14),
    )
    MergerApplication(root)
    root.mainloop()


if __name__ == "__main__":
    main()
