from __future__ import annotations

import json
import os
import shutil
import sys
import threading
import traceback
from argparse import Namespace
from dataclasses import replace
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import generate_shift as gas
import pythoncom


APP_NAME = "勤怠生成ツール"
APP_VERSION = "Ver0.0.2"


class ShiftGeneratorApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"{APP_NAME} {APP_VERSION}")
        self.root.geometry("820x620")
        self.root.minsize(740, 540)
        self.is_frozen = getattr(sys, "frozen", False)
        self.base_dir = Path(sys.executable).resolve().parent if self.is_frozen else Path(__file__).resolve().parent

        self.target_var = tk.StringVar()
        self.config_var = tk.StringVar(value=self._to_display_path(self._resolve_gui_config_path()))
        self.previous_var = tk.StringVar()
        self.report_output_var = tk.StringVar()
        self.status_var = tk.StringVar(value="対象の xls ファイルを選択してください。")
        self.progress_var = tk.DoubleVar(value=0)
        self.report_var = tk.StringVar(value="")
        self.open_report_var = tk.BooleanVar(value=False)
        self.open_workbook_var = tk.BooleanVar(value=True)

        self._build_ui()

    def _build_ui(self) -> None:
        container = ttk.Frame(self.root, padding=16)
        container.pack(fill=tk.BOTH, expand=True)
        container.columnconfigure(1, weight=1)
        container.rowconfigure(6, weight=1)

        ttk.Label(container, text="記載したい勤怠表 (.xls / .xlsx)").grid(row=0, column=0, sticky="w", pady=(0, 8))
        target_entry = ttk.Entry(container, textvariable=self.target_var)
        target_entry.grid(row=0, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
        ttk.Button(container, text="参照", command=self._browse_target).grid(row=0, column=2, sticky="ew", pady=(0, 8))

        ttk.Label(container, text="前月勤務表 (任意)").grid(row=1, column=0, sticky="w", pady=(0, 8))
        previous_entry = ttk.Entry(container, textvariable=self.previous_var)
        previous_entry.grid(row=1, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
        ttk.Button(container, text="参照", command=self._browse_previous).grid(row=1, column=2, sticky="ew", pady=(0, 8))

        ttk.Label(container, text="レポート保存先 (任意)").grid(row=2, column=0, sticky="w", pady=(0, 8))
        report_entry = ttk.Entry(container, textvariable=self.report_output_var)
        report_entry.grid(row=2, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
        ttk.Button(container, text="参照", command=self._browse_report_output).grid(row=2, column=2, sticky="ew", pady=(0, 8))

        ttk.Label(container, text="設定 JSON").grid(row=3, column=0, sticky="w", pady=(0, 8))
        config_entry = ttk.Entry(container, textvariable=self.config_var, state=("readonly" if self.is_frozen else "normal"))
        config_entry.grid(row=3, column=1, sticky="ew", padx=(8, 8), pady=(0, 8))
        ttk.Button(
            container,
            text="参照",
            command=self._browse_config,
            state=(tk.DISABLED if self.is_frozen else tk.NORMAL),
        ).grid(row=3, column=2, sticky="ew", pady=(0, 8))

        options = ttk.Frame(container)
        options.grid(row=4, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        ttk.Checkbutton(options, text="完了後にレポートを開く", variable=self.open_report_var).pack(side=tk.LEFT)
        ttk.Checkbutton(options, text="完了後に Excel を開く", variable=self.open_workbook_var).pack(side=tk.LEFT, padx=(12, 0))

        actions = ttk.Frame(container)
        actions.grid(row=5, column=0, columnspan=3, sticky="ew", pady=(0, 12))
        self.generate_button = ttk.Button(actions, text="生成実行", command=self._start_generate)
        self.generate_button.pack(side=tk.LEFT)
        ttk.Button(actions, text="レポートを開く", command=self._open_report).pack(side=tk.LEFT, padx=(8, 0))
        ttk.Button(actions, text="使い方を見る", command=self._open_help).pack(side=tk.LEFT, padx=(8, 0))

        log_frame = ttk.LabelFrame(container, text="実行ログ", padding=8)
        log_frame.grid(row=6, column=0, columnspan=3, sticky="nsew")
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)

        self.log_text = tk.Text(log_frame, wrap="word", height=18)
        self.log_text.grid(row=0, column=0, sticky="nsew")
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.grid(row=0, column=1, sticky="ns")
        self.log_text.configure(yscrollcommand=scrollbar.set)
        self.log_text.configure(state=tk.DISABLED)

        progress_bar = ttk.Progressbar(container, variable=self.progress_var, maximum=100, mode="determinate")
        progress_bar.grid(row=7, column=0, columnspan=3, sticky="ew", pady=(12, 4))

        status_bar = ttk.Label(container, textvariable=self.status_var, anchor="w")
        status_bar.grid(row=8, column=0, columnspan=3, sticky="ew")

    def _browse_target(self) -> None:
        file_path = filedialog.askopenfilename(
            title="対象の勤怠表を選択",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if file_path:
            self.target_var.set(self._to_display_path(Path(file_path)))
            if not self.report_output_var.get().strip():
                target_path = Path(file_path)
                self.report_output_var.set(self._to_display_path(target_path.with_name(target_path.stem + "_validation.html")))

    def _browse_config(self) -> None:
        if self.is_frozen:
            messagebox.showinfo("設定固定", "設定 JSON は実行フォルダ内の shift_config.json を使用します。")
            return
        file_path = filedialog.askopenfilename(
            title="設定 JSON を選択",
            filetypes=[("JSON files", "*.json"), ("All files", "*.*")],
        )
        if file_path:
            self.config_var.set(self._to_display_path(Path(file_path)))

    def _resolve_gui_config_path(self) -> Path:
        if not self.is_frozen:
            return gas.DEFAULT_CONFIG_PATH

        for file_name in gas.DEFAULT_CONFIG_CANDIDATE_FILENAMES:
            external_config = self.base_dir / file_name
            if external_config.exists():
                return external_config

        for file_name in gas.DEFAULT_CONFIG_CANDIDATE_FILENAMES:
            external_config = self.base_dir / file_name
            internal_config = self.base_dir / "_internal" / file_name
            if not external_config.exists() and internal_config.exists():
                try:
                    shutil.copy2(internal_config, external_config)
                    return external_config
                except OSError:
                    return internal_config

        return self.base_dir / gas.DEFAULT_CONFIG_CANDIDATE_FILENAMES[0]

    def _to_display_path(self, path: Path) -> str:
        try:
            return str(path.resolve().relative_to(self.base_dir.resolve()))
        except ValueError:
            return str(path.resolve())

    def _resolve_input_path(self, raw_path: str) -> Path:
        path = Path(raw_path)
        if path.is_absolute():
            return path.resolve()
        return (self.base_dir / path).resolve()

    def _browse_previous(self) -> None:
        file_path = filedialog.askopenfilename(
            title="前月勤務表を選択",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")],
        )
        if file_path:
            self.previous_var.set(self._to_display_path(Path(file_path)))

    def _browse_report_output(self) -> None:
        file_path = filedialog.asksaveasfilename(
            title="レポート保存先を選択",
            defaultextension=".html",
            filetypes=[("HTML files", "*.html"), ("All files", "*.*")],
        )
        if file_path:
            self.report_output_var.set(self._to_display_path(Path(file_path)))

    def _append_log(self, message: str) -> None:
        self.log_text.configure(state=tk.NORMAL)
        self.log_text.insert(tk.END, message + "\n")
        self.log_text.see(tk.END)
        self.log_text.configure(state=tk.DISABLED)

    def _set_status(self, message: str, percent: int | None = None) -> None:
        if percent is None:
            self.status_var.set(message)
            return
        clamped = max(0, min(100, int(percent)))
        self.progress_var.set(clamped)
        self.status_var.set(f"進捗 {clamped}%: {message}")

    def _queue_progress(self, percent: int, message: str) -> None:
        self.root.after(0, self._set_status, message, percent)

    def _set_running_state(self, is_running: bool) -> None:
        self.generate_button.configure(state=(tk.DISABLED if is_running else tk.NORMAL))

    def _start_generate(self) -> None:
        target_text = self.target_var.get().strip()
        if self.is_frozen:
            config_text = self._to_display_path(self._resolve_gui_config_path())
            self.config_var.set(config_text)
        else:
            config_text = self.config_var.get().strip()
        previous_text = self.previous_var.get().strip()
        report_output_text = self.report_output_var.get().strip()
        if not target_text:
            messagebox.showwarning("入力不足", "対象の xls ファイルを選択してください。")
            return
        if not config_text:
            messagebox.showwarning("入力不足", "設定 JSON を指定してください。")
            return

        target_path = self._resolve_input_path(target_text)
        config_path = self._resolve_input_path(config_text)
        previous_path = self._resolve_input_path(previous_text) if previous_text else None
        report_output_path = (
            self._resolve_input_path(report_output_text)
            if report_output_text
            else target_path.with_name(target_path.stem + "_validation.html")
        )
        if not target_path.exists():
            messagebox.showerror("ファイル未検出", f"対象ファイルが見つかりません。\n{target_path}")
            return
        if not config_path.exists():
            messagebox.showerror("ファイル未検出", f"設定 JSON が見つかりません。\n{config_path}")
            return
        if previous_path is not None and not previous_path.exists():
            messagebox.showerror("ファイル未検出", f"前月勤務表が見つかりません。\n{previous_path}")
            return
        if report_output_path.parent and not report_output_path.parent.exists():
            messagebox.showerror("フォルダ未検出", f"レポート保存先のフォルダが見つかりません。\n{report_output_path.parent}")
            return

        self.report_var.set("")
        self.progress_var.set(0)
        self._append_log(f"対象ファイル: {target_path}")
        self._append_log(f"設定 JSON: {config_path}")
        if previous_path is not None:
            self._append_log(f"前月勤務表: {previous_path}")
        self._append_log(f"レポート保存先: {report_output_path}")
        self._set_running_state(True)
        self._set_status("生成準備中です。", 0)

        worker = threading.Thread(
            target=self._run_generate,
            args=(config_path, target_path, previous_path, report_output_path),
            daemon=True,
        )
        worker.start()

    def _run_generate(
        self,
        config_path: Path,
        target_path: Path,
        previous_path: Path | None,
        report_output_path: Path,
    ) -> None:
        com_initialized = False
        try:
            pythoncom.CoInitialize()
            com_initialized = True
            self._queue_progress(10, "設定を読み込んでいます。")
            config = gas.load_config(config_path)
            args = Namespace(target=target_path, year=None, month=None, unit_name=None, days=None)
            self._queue_progress(20, "対象ファイルと設定を照合しています。")
            config = gas.with_generate_overrides(config, args)
            if previous_path is not None:
                self._queue_progress(30, "前月勤務を取り込んでいます。")
                previous_tail_length = max(
                    config.max_consecutive_work,
                    config.max_consecutive_night,
                    config.max_consecutive_rest,
                    config.max_consecutive_rest_with_special,
                )
                previous_tails = gas.read_previous_tail_from_workbook(
                    previous_path,
                    config.sheet_index,
                    config.employees,
                    config.shift_kinds,
                    previous_tail_length,
                )
                merged_employees = [
                    replace(
                        employee,
                        previous_tail=previous_tails.get(employee.employee_id, employee.previous_tail),
                    )
                    for employee in config.employees
                ]
                config = replace(config, employees=tuple(merged_employees))
            self._queue_progress(40, "勤務表を計算しています。しばらくお待ちください。")
            solve_result = gas.solve_schedule(config)
            schedule = solve_result.schedule
            self._queue_progress(75, "計算結果を検証しています。")
            validation = gas.validate_schedule(config, schedule)
            validation["partial_mode"] = solve_result.is_partial
            validation["partial_reason"] = solve_result.message
            validation["partial_summary_lines"] = solve_result.diagnostics.get("summary_lines", [])
            if validation["issues"] and not solve_result.is_partial:
                raise RuntimeError("ルール検証で問題が見つかりました: " + json.dumps(validation, ensure_ascii=False))
            self._queue_progress(88, "Excel に書き込んでいます。")
            gas.write_schedule_to_excel(config, schedule)
            self._queue_progress(96, "検証レポートを作成しています。")
            report_path = gas.write_validation_report(config, validation, report_output_path)
            self.root.after(0, self._on_generate_success, config, validation, report_path, target_path)
        except Exception as exc:
            details = "".join(traceback.format_exception(exc))
            summary = str(exc).strip() or "生成に失敗しました。"
            self.root.after(0, self._on_generate_error, summary, details)
        finally:
            if com_initialized:
                pythoncom.CoUninitialize()

    def _on_generate_success(
        self,
        config: gas.SchedulerConfig,
        validation: dict[str, object],
        report_path: Path,
        target_path: Path,
    ) -> None:
        self.report_var.set(str(report_path))
        completion_message = "途中案の出力が完了しました。" if validation.get("partial_mode") else "生成が完了しました。"
        self._set_status(completion_message, 100)
        self._set_running_state(False)
        self._append_log(completion_message)
        self._append_log(f"対象月: {config.year}年{config.month}月")
        self._append_log(f"勤怠表: {target_path}")
        self._append_log(f"レポート: {report_path}")
        if validation.get("partial_mode"):
            self._append_log(str(validation.get("partial_reason", "")))
            for line in validation.get("partial_summary_lines", []):
                self._append_log(f"  - {line}")
        self._append_log(f"夜勤ばらつき: {validation['night_spread']} 回")
        self._append_log(f"土日休みばらつき: {validation['weekend_rest_spread']} 回")
        if self.open_workbook_var.get():
            self._open_workbook(target_path)
        if self.open_report_var.get():
            self._open_report()

    def _on_generate_error(self, summary: str, details: str) -> None:
        current_progress = int(self.progress_var.get())
        self._set_running_state(False)
        self._append_log("生成に失敗しました。")
        self._append_log(details.rstrip())
        self._set_status(summary, current_progress)
        messagebox.showerror("生成失敗", summary)

    def _open_report(self) -> None:
        report_path = self.report_var.get().strip()
        if not report_path:
            messagebox.showinfo("レポート未作成", "まだレポートは作成されていません。")
            return
        report = Path(report_path)
        if not report.exists():
            messagebox.showerror("ファイル未検出", f"レポートが見つかりません。\n{report}")
            return
        os.startfile(report)

    def _resolve_help_path(self) -> Path | None:
        if getattr(sys, "frozen", False):
            base_dir = Path(sys.executable).resolve().parent
            candidates = [
                base_dir / "docs" / "guides" / "3分で使う_GUI版.html",
                base_dir / "docs" / "guides" / "3分で使う_GUI版.md",
                base_dir / "3分で使う_GUI版.html",
                base_dir / "3分で使う_GUI版.md",
                base_dir / "_internal" / "3分で使う_GUI版.html",
                base_dir / "_internal" / "3分で使う_GUI版.md",
            ]
        else:
            base_dir = Path(__file__).resolve().parent
            candidates = [
                base_dir / "docs" / "guides" / "3分で使う_GUI版.html",
                base_dir / "docs" / "guides" / "3分で使う_GUI版.md",
                base_dir / "exe" / "generate_akanecco_shift_gui" / "3分で使う_GUI版.html",
                base_dir / "exe" / "generate_akanecco_shift_gui" / "3分で使う_GUI版.md",
                base_dir / "3分で使う_GUI版.md",
            ]
        for candidate in candidates:
            if candidate.exists():
                return candidate
        return None

    def _open_help(self) -> None:
        help_path = self._resolve_help_path()
        if help_path is None:
            messagebox.showerror("ファイル未検出", "使い方ファイルが見つかりません。")
            return
        os.startfile(help_path)

    def _open_workbook(self, target_path: Path) -> None:
        if not target_path.exists():
            messagebox.showerror("ファイル未検出", f"勤怠表が見つかりません。\n{target_path}")
            return
        os.startfile(target_path)


def main() -> None:
    root = tk.Tk()
    style = ttk.Style(root)
    if "vista" in style.theme_names():
        style.theme_use("vista")
    app = ShiftGeneratorApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()