"""엑셀 일괄 PDF 변환기 - GUI

드래그앤드롭으로 엑셀 파일/폴더를 추가하고 일괄로 PDF 변환한다.
"""

from __future__ import annotations

import json
import os
import queue
import subprocess
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

try:
    from tkinterdnd2 import DND_FILES, TkinterDnD

    _HAS_DND = True
except ImportError:
    _HAS_DND = False

from converter import (
    SUPPORTED_EXTS,
    BatchConverter,
    ConversionError,
    ConvertOptions,
    find_libreoffice,
)

APP_NAME = "엑셀 일괄 PDF 변환기"
CONFIG_DIR = os.path.join(os.environ.get("LOCALAPPDATA", os.path.expanduser("~")), "ExcelToPDF")
CONFIG_PATH = os.path.join(CONFIG_DIR, "settings.json")


def load_settings() -> dict:
    try:
        with open(CONFIG_PATH, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(data: dict):
    try:
        os.makedirs(CONFIG_DIR, exist_ok=True)
        with open(CONFIG_PATH, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def excel_available() -> bool:
    try:
        import winreg

        winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Excel.Application")
        return True
    except Exception:
        return False


def collect_files_from_paths(paths: list[str]) -> list[str]:
    result = []
    for p in paths:
        if os.path.isdir(p):
            for root, _dirs, names in os.walk(p):
                for name in names:
                    if name.lower().endswith(SUPPORTED_EXTS) and not name.startswith("~$"):
                        result.append(os.path.join(root, name))
        elif os.path.isfile(p):
            if p.lower().endswith(SUPPORTED_EXTS) and not os.path.basename(p).startswith("~$"):
                result.append(p)
    return result


def parse_dnd_paths(data: str) -> list[str]:
    """tkinterdnd2가 넘겨주는 '{공백 포함 경로} 경로2' 형식을 파싱."""
    paths = []
    buf = ""
    in_brace = False
    for ch in data:
        if ch == "{":
            in_brace = True
            buf = ""
        elif ch == "}":
            in_brace = False
            paths.append(buf)
            buf = ""
        elif ch == " " and not in_brace:
            if buf:
                paths.append(buf)
                buf = ""
        else:
            buf += ch
    if buf:
        paths.append(buf)
    return paths


class App:
    def __init__(self, root):
        self.root = root
        self.root.title(APP_NAME)
        self.root.geometry("820x560")
        self.root.minsize(680, 440)

        self.settings = load_settings()
        self.file_rows: dict[str, str] = {}  # path -> treeview item id
        self.out_paths: dict[str, str] = {}  # path -> 생성된 PDF 경로
        self._closing = False
        self.worker_thread: threading.Thread | None = None
        self.batch_converter: BatchConverter | None = None
        self.ui_queue: queue.Queue = queue.Queue()

        self._build_widgets()
        self._poll_queue()
        self._check_engines()

    # ---------- UI 구성 ----------
    def _build_widgets(self):
        top = ttk.Frame(self.root, padding=8)
        top.pack(fill="x")

        self.list_buttons = [
            ttk.Button(top, text="파일 추가", command=self.add_files),
            ttk.Button(top, text="폴더 추가", command=self.add_folder),
            ttk.Button(top, text="선택 삭제", command=self.remove_selected),
            ttk.Button(top, text="목록 지우기", command=self.clear_list),
        ]
        for i, b in enumerate(self.list_buttons):
            b.pack(side="left", padx=(0 if i == 0 else 4, 0))

        self.engine_var = tk.StringVar(value=self.settings.get("engine", "excel"))
        ttk.Label(top, text="변환 엔진:").pack(side="left", padx=(16, 4))
        self.engine_combo = ttk.Combobox(
            top, textvariable=self.engine_var, state="readonly", width=16,
            values=["excel (권장)", "libreoffice"],
        )
        self.engine_combo.pack(side="left")
        self.engine_combo.bind("<<ComboboxSelected>>", lambda e: self._on_engine_change())
        self._sync_engine_combo_display()

        # 드래그앤드롭 대상 + 목록
        mid = ttk.Frame(self.root, padding=(8, 0))
        mid.pack(fill="both", expand=True)

        columns = ("status", "output")
        self.tree = ttk.Treeview(mid, columns=columns, show="tree headings", selectmode="extended")
        self.tree.heading("#0", text="파일")
        self.tree.heading("status", text="상태")
        self.tree.heading("output", text="결과 / 메시지")
        self.tree.column("#0", width=380)
        self.tree.column("status", width=90, anchor="center")
        self.tree.column("output", width=300)
        self.tree.pack(side="left", fill="both", expand=True)

        scroll = ttk.Scrollbar(mid, orient="vertical", command=self.tree.yview)
        scroll.pack(side="left", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)

        self.tree.tag_configure("완료", foreground="#2e7d32")
        self.tree.tag_configure("실패", foreground="#c62828")
        self.tree.tag_configure("변환중", foreground="#1565c0")
        self.tree.tag_configure("취소됨", foreground="#757575")
        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Delete>", lambda e: self.remove_selected())

        # 목록이 비었을 때 가운데에 보이는 안내 (파일이 들어오면 숨김)
        self.empty_hint = tk.Label(
            self.tree,
            text=("엑셀 파일이나 폴더를 여기로 끌어다 놓으세요\n"
                  "(.xlsx / .xlsm / .xls / .xlsb / .csv · 폴더는 하위 폴더까지 찾음)\n\n"
                  "변환이 끝난 줄을 더블클릭하면 PDF가 열립니다")
            if _HAS_DND else "[파일 추가] 또는 [폴더 추가]로 엑셀 파일을 넣으세요",
            fg="#888", bg="white", font=("맑은 고딕", 10), justify="center",
        )
        self.empty_hint.place(relx=0.5, rely=0.5, anchor="center")

        if _HAS_DND:
            self.tree.drop_target_register(DND_FILES)
            self.tree.dnd_bind("<<Drop>>", self._on_drop)
            self.empty_hint.drop_target_register(DND_FILES)
            self.empty_hint.dnd_bind("<<Drop>>", self._on_drop)
        else:
            hint = ttk.Label(
                self.root,
                text="(드래그앤드롭 모듈(tkinterdnd2)이 없어 DnD가 비활성화되었습니다. '파일 추가' 버튼을 사용하세요)",
                foreground="#a33",
            )
            hint.pack(fill="x", padx=8)

        # 옵션
        opt = ttk.LabelFrame(self.root, text="옵션", padding=8)
        opt.pack(fill="x", padx=8, pady=4)

        self.output_mode = tk.StringVar(value=self.settings.get("output_mode", "same"))
        ttk.Radiobutton(opt, text="원본과 같은 폴더", variable=self.output_mode, value="same",
                         command=self._toggle_output_dir).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(opt, text="지정 폴더:", variable=self.output_mode, value="custom",
                         command=self._toggle_output_dir).grid(row=0, column=1, sticky="w")

        self.output_dir_var = tk.StringVar(value=self.settings.get("output_dir", ""))
        self.output_dir_entry = ttk.Entry(opt, textvariable=self.output_dir_var, width=45)
        self.output_dir_entry.grid(row=0, column=2, sticky="we", padx=4)
        self.browse_btn = ttk.Button(opt, text="찾아보기", command=self._browse_output_dir)
        self.browse_btn.grid(row=0, column=3)

        self.fit_to_page = tk.BooleanVar(value=self.settings.get("fit_to_page", True))
        ttk.Checkbutton(opt, text="한 페이지 너비에 맞춤", variable=self.fit_to_page).grid(
            row=1, column=0, columnspan=2, sticky="w", pady=(6, 0)
        )

        self.overwrite = tk.BooleanVar(value=self.settings.get("overwrite", True))
        ttk.Checkbutton(opt, text="동일 이름 PDF 덮어쓰기", variable=self.overwrite).grid(
            row=1, column=2, sticky="w", pady=(6, 0)
        )

        self.open_after = tk.BooleanVar(value=self.settings.get("open_after", True))
        ttk.Checkbutton(opt, text="완료 후 결과 폴더 열기", variable=self.open_after).grid(
            row=2, column=0, columnspan=2, sticky="w"
        )

        opt.columnconfigure(2, weight=1)
        self._toggle_output_dir()

        # 하단 진행/버튼
        bottom = ttk.Frame(self.root, padding=8)
        bottom.pack(fill="x")

        self.progress = ttk.Progressbar(bottom, mode="determinate")
        self.progress.pack(side="left", fill="x", expand=True, padx=(0, 8))

        self.status_label = ttk.Label(bottom, text="대기 중")
        self.status_label.pack(side="left", padx=(0, 8))

        self.cancel_btn = ttk.Button(bottom, text="취소", command=self.cancel_conversion, state="disabled")
        self.cancel_btn.pack(side="right")
        self.convert_btn = ttk.Button(bottom, text="PDF로 변환", command=self.start_conversion)
        self.convert_btn.pack(side="right", padx=8)

        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _sync_engine_combo_display(self):
        self.engine_combo.set("excel (권장)" if self.engine_var.get() == "excel" else "libreoffice")

    def _on_engine_change(self):
        val = self.engine_combo.get()
        self.engine_var.set("excel" if val.startswith("excel") else "libreoffice")

    def _toggle_output_dir(self):
        state = "normal" if self.output_mode.get() == "custom" else "disabled"
        self.output_dir_entry.configure(state=state)
        self.browse_btn.configure(state=state)

    def _browse_output_dir(self):
        d = filedialog.askdirectory()
        if d:
            self.output_dir_var.set(d)

    def _check_engines(self):
        has_excel = excel_available()
        has_lo = find_libreoffice() is not None
        if not has_excel and not has_lo:
            messagebox.showwarning(
                APP_NAME,
                "Microsoft Excel과 LibreOffice를 모두 찾을 수 없습니다.\n"
                "변환을 위해 둘 중 하나를 설치해 주세요.",
            )
        elif not has_excel:
            self.engine_var.set("libreoffice")
            self._sync_engine_combo_display()

    # ---------- 파일 목록 관리 ----------
    def _busy(self) -> bool:
        return bool(self.worker_thread and self.worker_thread.is_alive())

    def _refresh_hint(self):
        if self.file_rows:
            self.empty_hint.place_forget()
        else:
            self.empty_hint.place(relx=0.5, rely=0.5, anchor="center")

    def _on_double_click(self, event):
        item = self.tree.identify_row(event.y)
        if not item:
            return
        path = self.tree.item(item, "text")
        target = self.out_paths.get(path) or path
        if os.path.exists(target):
            os.startfile(target)  # noqa: S606

    def _add_paths(self, paths: list[str]):
        if self._busy():
            return
        files = collect_files_from_paths(paths)
        added = 0
        for f in files:
            if f in self.file_rows:
                continue
            item = self.tree.insert("", "end", text=f, values=("대기", ""))
            self.file_rows[f] = item
            added += 1
        if not files:
            messagebox.showinfo(APP_NAME, "지원되는 엑셀 파일(.xlsx/.xlsm/.xls/.xlsb/.csv)을 찾지 못했습니다.")
        self._refresh_hint()

    def _on_drop(self, event):
        paths = parse_dnd_paths(event.data)
        self._add_paths(paths)

    def add_files(self):
        paths = filedialog.askopenfilenames(
            title="엑셀 파일 선택",
            filetypes=[("Excel/CSV 파일", "*.xlsx *.xlsm *.xls *.xlsb *.csv"), ("모든 파일", "*.*")],
        )
        if paths:
            self._add_paths(list(paths))

    def add_folder(self):
        d = filedialog.askdirectory(title="폴더 선택")
        if d:
            self._add_paths([d])

    def clear_list(self):
        if self._busy():
            return
        self.tree.delete(*self.tree.get_children())
        self.file_rows.clear()
        self.out_paths.clear()
        self._refresh_hint()

    def remove_selected(self):
        if self._busy():
            return
        for item in self.tree.selection():
            path = self.tree.item(item, "text")
            self.file_rows.pop(path, None)
            self.out_paths.pop(path, None)
            self.tree.delete(item)
        self._refresh_hint()

    # ---------- 변환 ----------
    def start_conversion(self):
        if self.worker_thread and self.worker_thread.is_alive():
            return
        files = list(self.file_rows.keys())
        if not files:
            messagebox.showinfo(APP_NAME, "변환할 파일을 먼저 추가해 주세요.")
            return

        if self.output_mode.get() == "custom" and not self.output_dir_var.get().strip():
            messagebox.showwarning(APP_NAME, "출력 폴더를 지정해 주세요.")
            return

        options = ConvertOptions(
            fit_to_page=self.fit_to_page.get(),
            overwrite=self.overwrite.get(),
            output_mode=self.output_mode.get(),
            output_dir=self.output_dir_var.get().strip() or None,
        )
        self._save_current_settings()

        self.out_paths.clear()
        for path, item in self.file_rows.items():
            self.tree.item(item, values=("대기", ""), tags=())
        for b in self.list_buttons:
            b.configure(state="disabled")

        self.progress.configure(maximum=len(files), value=0)
        self.status_label.configure(text=f"0 / {len(files)}")
        self.convert_btn.configure(state="disabled")
        self.cancel_btn.configure(state="normal")

        self.batch_converter = BatchConverter(options, engine_name=self.engine_var.get())
        self._done_count = 0
        self._total_count = len(files)

        def worker():
            def on_progress(path, status, message, out_path=None):
                self.ui_queue.put((path, status, message, out_path))

            try:
                self.batch_converter.run(files, on_progress)
            except ConversionError as exc:
                self.ui_queue.put(("__FATAL__", "오류", str(exc), None))
            except Exception as exc:  # noqa: BLE001
                self.ui_queue.put(("__FATAL__", "오류", str(exc), None))
            finally:
                self.ui_queue.put(("__DONE__", "", "", None))

        self.worker_thread = threading.Thread(target=worker, daemon=True)
        self.worker_thread.start()

    def cancel_conversion(self):
        if self.batch_converter:
            self.batch_converter.cancel()
        self.cancel_btn.configure(state="disabled")
        self.status_label.configure(text="취소 중...")

    def _poll_queue(self):
        try:
            while True:
                path, status, message, out_path = self.ui_queue.get_nowait()
                if path == "__DONE__":
                    self._on_batch_done()
                    continue
                if path == "__FATAL__":
                    messagebox.showerror(APP_NAME, message)
                    continue
                item = self.file_rows.get(path)
                if item:
                    self.tree.item(item, values=(status, message), tags=(status,))
                    if status == "변환중":
                        self.tree.see(item)
                if out_path:
                    self.out_paths[path] = out_path
                if status in ("완료", "실패", "취소됨"):
                    self._done_count += 1
                    self.progress.configure(value=self._done_count)
                    self.status_label.configure(text=f"{self._done_count} / {self._total_count}")
        except queue.Empty:
            pass
        self.root.after(100, self._poll_queue)

    def _on_batch_done(self):
        if self._closing:
            self.root.destroy()
            return
        self.convert_btn.configure(state="normal")
        self.cancel_btn.configure(state="disabled")
        for b in self.list_buttons:
            b.configure(state="normal")

        failed = [
            (self.tree.item(item, "text"), self.tree.item(item, "values")[1])
            for item in self.tree.get_children() if self.tree.item(item, "values")[0] == "실패"
        ]
        ok_count = sum(
            1 for item in self.tree.get_children() if self.tree.item(item, "values")[0] == "완료"
        )
        self.status_label.configure(text=f"완료 {ok_count} · 실패 {len(failed)}")
        summary = f"변환 완료: 성공 {ok_count}건, 실패 {len(failed)}건"
        if failed:
            name, reason = failed[0]
            summary += f"\n\n실패 예: {os.path.basename(name)}\n{reason}"
            if len(failed) > 1:
                summary += f"\n\n(나머지 {len(failed) - 1}건의 사유는 목록의 '결과 / 메시지' 칸에 있습니다)"
            messagebox.showwarning(APP_NAME, summary)
        else:
            messagebox.showinfo(APP_NAME, summary)

        if self.open_after.get() and ok_count > 0:
            target = self.output_dir_var.get().strip() if self.output_mode.get() == "custom" else None
            if not target:
                first_path = next(iter(self.file_rows.keys()), None)
                target = os.path.dirname(first_path) if first_path else None
            if target and os.path.isdir(target):
                try:
                    os.startfile(target)  # noqa: S606 (Windows 전용, 사용자 지정/원본 폴더만 연다)
                except Exception:
                    pass

    def _save_current_settings(self):
        save_settings(
            {
                "engine": self.engine_var.get(),
                "output_mode": self.output_mode.get(),
                "output_dir": self.output_dir_var.get().strip(),
                "fit_to_page": self.fit_to_page.get(),
                "overwrite": self.overwrite.get(),
                "open_after": self.open_after.get(),
            }
        )

    def _on_close(self):
        self._save_current_settings()
        if self._busy():
            if not messagebox.askyesno(
                APP_NAME, "변환 중입니다. 지금 파일까지만 변환하고 종료할까요?"
            ):
                return
            # 작업 스레드가 Excel을 정상 종료할 때까지 기다렸다가 창을 닫는다.
            # (바로 닫으면 숨겨진 EXCEL.EXE 가 백그라운드에 남는다)
            self._closing = True
            self.batch_converter.cancel()
            self.cancel_btn.configure(state="disabled")
            self.convert_btn.configure(state="disabled")
            self.status_label.configure(text="종료 중… (Excel 정리)")
            return
        self.root.destroy()


def main():
    if _HAS_DND:
        root = TkinterDnD.Tk()
    else:
        root = tk.Tk()
    try:
        style = ttk.Style()
        if "vista" in style.theme_names():
            style.theme_use("vista")
    except Exception:
        pass
    App(root)
    root.mainloop()


if __name__ == "__main__":
    main()
