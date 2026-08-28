from __future__ import annotations

import subprocess
import sys
import tkinter as tk
from datetime import date
from pathlib import Path
from tkinter import filedialog, messagebox, ttk


DEFAULT_TEMPLATE = (
    Path.home()
    / "Desktop"
    / "WorkSpace"
    / "DATA"
    / "sales_template.xlsx"
)


class CollectorWindow(ttk.Frame):
    def __init__(self, master: tk.Tk) -> None:
        super().__init__(master, padding=24)
        self.master = master
        self.template_var = tk.StringVar(value=str(DEFAULT_TEMPLATE))
        yesterday = date.today().fromordinal(date.today().toordinal() - 1).isoformat()
        self.start_date_var = tk.StringVar(value=yesterday)
        self.end_date_var = tk.StringVar(value=yesterday)
        self.scope_var = tk.StringVar(value="all")
        self.store_name_var = tk.StringVar()
        self.status_var = tk.StringVar(value="기간과 대상을 선택한 뒤 실행하세요.")
        self.grid(sticky="nsew")
        self._build()
        self._update_scope()

    def _build(self) -> None:
        self.master.title("매출 데이터 수집")
        self.master.minsize(680, 390)
        self.columnconfigure(1, weight=1)

        ttk.Label(self, text="매출 데이터 수집", font=("Malgun Gothic", 16, "bold")).grid(
            row=0, column=0, columnspan=3, sticky="w", pady=(0, 18)
        )
        ttk.Label(self, text="원본 템플릿").grid(row=1, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.template_var).grid(row=1, column=1, sticky="ew", padx=8)
        ttk.Button(self, text="찾아보기", command=self._choose_template).grid(row=1, column=2)

        ttk.Label(self, text="조회 시작일").grid(row=2, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.start_date_var, width=16).grid(row=2, column=1, sticky="w", padx=8)
        ttk.Label(self, text="YYYY-MM-DD").grid(row=2, column=2, sticky="w")
        ttk.Label(self, text="조회 종료일").grid(row=3, column=0, sticky="w", pady=6)
        ttk.Entry(self, textvariable=self.end_date_var, width=16).grid(row=3, column=1, sticky="w", padx=8)
        ttk.Label(self, text="YYYY-MM-DD").grid(row=3, column=2, sticky="w")

        ttk.Label(self, text="대상 가맹점").grid(row=4, column=0, sticky="nw", pady=(12, 6))
        scope = ttk.Frame(self)
        scope.grid(row=4, column=1, columnspan=2, sticky="ew", padx=8, pady=(12, 6))
        ttk.Radiobutton(scope, text="전체 가맹점", variable=self.scope_var, value="all", command=self._update_scope).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(scope, text="특정 가맹점", variable=self.scope_var, value="specific", command=self._update_scope).grid(row=0, column=1, sticky="w", padx=(20, 0))
        ttk.Label(self, text="매장명").grid(row=5, column=0, sticky="w", pady=6)
        self.store_combo = ttk.Combobox(self, textvariable=self.store_name_var, state="normal")
        self.store_combo.grid(row=5, column=1, columnspan=2, sticky="ew", padx=8)

        ttk.Separator(self).grid(row=6, column=0, columnspan=3, sticky="ew", pady=20)
        ttk.Label(self, text="실행하면 별도 CMD 창에서 서비스 ID와 비밀번호를 입력합니다.", foreground="#4a5568").grid(row=7, column=0, columnspan=3, sticky="w")
        ttk.Button(self, text="드라이런 후 기록 시작", command=self._launch, width=24).grid(row=8, column=2, sticky="e", pady=(14, 8))
        ttk.Label(self, textvariable=self.status_var, foreground="#2b6cb0").grid(row=9, column=0, columnspan=3, sticky="w")

    def _choose_template(self) -> None:
        selected = filedialog.askopenfilename(filetypes=[("Excel workbook", "*.xlsx")])
        if selected:
            self.template_var.set(selected)

    def _update_scope(self) -> None:
        self.store_combo.configure(state="normal" if self.scope_var.get() == "specific" else "disabled")

    def _launch(self) -> None:
        template = Path(self.template_var.get().strip())
        if not template.is_file():
            messagebox.showerror("템플릿 확인", "유효한 원본 Excel 템플릿을 선택하세요.")
            return
        try:
            start = date.fromisoformat(self.start_date_var.get().strip())
            end = date.fromisoformat(self.end_date_var.get().strip())
        except ValueError:
            messagebox.showerror("날짜 형식", "날짜는 YYYY-MM-DD 형식으로 입력하세요.")
            return
        if end < start:
            messagebox.showerror("날짜 범위", "종료일은 시작일보다 빠를 수 없습니다.")
            return

        command = self._collector_command(template, start, end)
        if self.scope_var.get() == "all":
            command.append("--all-stores")
        else:
            store_name = self.store_name_var.get().strip()
            if not store_name:
                messagebox.showerror("매장 선택", "특정 가맹점의 정확한 매장명을 입력하세요.")
                return
            command.extend(["--store-name", store_name])
        try:
            subprocess.Popen(
                ["cmd.exe", "/k", subprocess.list2cmdline(command)],
                creationflags=getattr(subprocess, "CREATE_NEW_CONSOLE", 0),
            )
        except OSError as exc:
            messagebox.showerror("실행 실패", str(exc))
            return
        self.status_var.set("CMD 창에서 로그인 후 드라이런과 기록을 진행 중입니다.")

    @staticmethod
    def _date_arguments(start: date, end: date) -> list[str]:
        values: list[str] = []
        current = start
        while current <= end:
            values.extend(["--date", current.isoformat()])
            current = current.fromordinal(current.toordinal() + 1)
        return values

    def _collector_command(self, template: Path, start: date, end: date) -> list[str]:
        if getattr(sys, "frozen", False):
            executable = Path(sys.executable).with_name("Sales_Data_Collector.exe")
            if not executable.is_file():
                raise FileNotFoundError(f"Collector executable not found: {executable}")
            command = [str(executable)]
        else:
            command = [sys.executable, str(Path(__file__).with_name("main.py"))]
        return command + ["--production-write", "--template", str(template)] + self._date_arguments(start, end)


def main() -> None:
    root = tk.Tk()
    root.columnconfigure(0, weight=1)
    root.rowconfigure(0, weight=1)
    CollectorWindow(root)
    root.mainloop()


if __name__ == "__main__":
    main()
