"""
ROHDE 송신기 로그 분석기
HTML 파라미터 스냅샷 또는 DMB 텍스트 로그를 파싱하여 Excel에 측정 데이터를 기록한다.
"""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

from dmb_excel import update_dmb_excel
from dmb_parser import detect_dmb_excel_kind, parse_dmb_log
from excel_handler import detect_excel_tx_kind, update_excel
from html_parser import detect_html_tx_kind, parse_html


# ── 색상 / 폰트 상수 ────────────────────────────────────────────────────────
BG_COLOR     = "#f0f2f5"
ACCENT_COLOR = "#006599"
BTN_COLOR    = "#0080bf"
BTN_HOVER    = "#005f8f"
TEXT_BG      = "#1e1e2e"
TEXT_FG      = "#cdd6f4"
FONT_NORMAL  = ("맑은 고딕", 10)
FONT_BOLD    = ("맑은 고딕", 10, "bold")
FONT_TITLE   = ("맑은 고딕", 14, "bold")
FONT_MONO    = ("Consolas", 9)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ROHDE 송신기 로그 분석기")
        self.geometry("720x640")
        self.minsize(600, 540)
        self.configure(bg=BG_COLOR)
        self.resizable(True, True)

        self._html_path = tk.StringVar()
        self._excel_path = tk.StringVar()
        self._txt_path = tk.StringVar()
        self._excel_a_path = tk.StringVar()
        self._excel_b_path = tk.StringVar()
        self._tx_mode = tk.StringVar(value="dtv")

        self._build_ui()
        self._tx_mode.trace_add("write", self._on_mode_change)
        self._on_mode_change()

    # ── UI 구성 ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        header = tk.Frame(self, bg=ACCENT_COLOR, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="ROHDE 송신기 로그 분석기",
            bg=ACCENT_COLOR, fg="white",
            font=FONT_TITLE,
        ).pack(side="left", padx=20, pady=14)

        main = tk.Frame(self, bg=BG_COLOR, padx=20, pady=16)
        main.pack(fill="both", expand=True)

        mode_frame = tk.LabelFrame(
            main, text=" 분석 모드 ", bg=BG_COLOR,
            font=FONT_BOLD, fg=ACCENT_COLOR, padx=12, pady=8,
        )
        mode_frame.pack(fill="x", pady=(0, 10))
        tk.Radiobutton(
            mode_frame, text="DTV (AMP 2개)",
            variable=self._tx_mode, value="dtv",
            bg=BG_COLOR, font=FONT_NORMAL, anchor="w",
        ).pack(side="left", padx=(4, 16))
        tk.Radiobutton(
            mode_frame, text="UHDTV (AMP 6개)",
            variable=self._tx_mode, value="uhdtv",
            bg=BG_COLOR, font=FONT_NORMAL, anchor="w",
        ).pack(side="left", padx=(0, 16))
        tk.Radiobutton(
            mode_frame, text="DMB (TX-A / TX-B)",
            variable=self._tx_mode, value="dmb",
            bg=BG_COLOR, font=FONT_NORMAL, anchor="w",
        ).pack(side="left", padx=4)

        file_frame = tk.LabelFrame(
            main, text=" 파일 선택 ", bg=BG_COLOR,
            font=FONT_BOLD, fg=ACCENT_COLOR, padx=12, pady=10,
        )
        file_frame.pack(fill="x", pady=(0, 12))

        self._rohde_files = tk.Frame(file_frame, bg=BG_COLOR)
        self._add_file_row(self._rohde_files, "HTML 파일 (로그):", self._html_path, "html", row=0)
        self._add_file_row(self._rohde_files, "Excel 파일 (결과지):", self._excel_path, "excel", row=1)

        self._dmb_files = tk.Frame(file_frame, bg=BG_COLOR)
        self._add_file_row(self._dmb_files, "로그 파일 (.txt):", self._txt_path, "dmb_txt", row=0)
        self._add_file_row(self._dmb_files, "Excel TX-A:", self._excel_a_path, "dmb_a", row=1)
        self._add_file_row(self._dmb_files, "Excel TX-B:", self._excel_b_path, "dmb_b", row=2)

        self._run_btn = tk.Button(
            main,
            text="▶  분석 및 Excel 저장",
            command=self._on_run,
            bg=BTN_COLOR, fg="white",
            font=("맑은 고딕", 11, "bold"),
            relief="flat", cursor="hand2",
            padx=20, pady=8,
        )
        self._run_btn.pack(pady=(0, 12))
        self._run_btn.bind("<Enter>", lambda e: self._run_btn.config(bg=BTN_HOVER))
        self._run_btn.bind("<Leave>", lambda e: self._run_btn.config(bg=BTN_COLOR))

        self._progress = ttk.Progressbar(main, mode="indeterminate", length=300)
        self._progress.pack(fill="x", pady=(0, 8))

        log_frame = tk.LabelFrame(
            main, text=" 처리 결과 ", bg=BG_COLOR,
            font=FONT_BOLD, fg=ACCENT_COLOR,
        )
        log_frame.pack(fill="both", expand=True)

        self._log_box = scrolledtext.ScrolledText(
            log_frame,
            bg=TEXT_BG, fg=TEXT_FG,
            font=FONT_MONO,
            state="disabled",
            relief="flat",
            wrap="word",
        )
        self._log_box.pack(fill="both", expand=True, padx=6, pady=6)

        self._log_box.tag_config("info",    foreground="#89b4fa")
        self._log_box.tag_config("success", foreground="#a6e3a1")
        self._log_box.tag_config("error",   foreground="#f38ba8")
        self._log_box.tag_config("detail",  foreground="#a6adc8")

    def _on_mode_change(self, *args):
        mode = self._tx_mode.get()
        if mode == "dmb":
            self._rohde_files.pack_forget()
            self._dmb_files.pack(fill="x")
        else:
            self._dmb_files.pack_forget()
            self._rohde_files.pack(fill="x")

    def _add_file_row(self, parent, label_text, str_var, kind, row):
        tk.Label(parent, text=label_text, bg=BG_COLOR, font=FONT_NORMAL, width=20, anchor="w"
                 ).grid(row=row, column=0, sticky="w", pady=4)

        entry = tk.Entry(
            parent, textvariable=str_var, state="readonly",
            relief="solid", font=FONT_NORMAL, width=46,
            readonlybackground="white",
        )
        entry.grid(row=row, column=1, padx=(6, 6), sticky="ew", pady=4)

        btn = tk.Button(
            parent, text="찾아보기…",
            command=lambda k=kind: self._browse(k),
            bg="#e0e4e8", relief="flat", font=FONT_NORMAL,
            cursor="hand2", padx=8,
        )
        btn.grid(row=row, column=2, pady=4)

        parent.columnconfigure(1, weight=1)

    # ── 파일 탐색 ───────────────────────────────────────────────────────────

    def _html_matches_mode(self, path: str) -> bool:
        try:
            file_kind = detect_html_tx_kind(path)
        except OSError as exc:
            messagebox.showerror("파일 오류", f"HTML 파일을 읽을 수 없습니다.\n{exc}")
            return False
        want_uhdtv = self._tx_mode.get() == "uhdtv"
        if want_uhdtv and file_kind == "dtv":
            messagebox.showwarning("경고", "UHDTV파일을 선택해 주세요")
            return False
        if not want_uhdtv and file_kind == "uhdtv":
            messagebox.showwarning("경고", "DTV파일을 선택해 주세요")
            return False
        return True

    def _excel_matches_mode(self, path: str) -> bool:
        try:
            file_kind = detect_excel_tx_kind(path)
        except Exception as exc:
            messagebox.showerror("파일 오류", f"Excel 파일을 읽을 수 없습니다.\n{exc}")
            return False
        want_uhdtv = self._tx_mode.get() == "uhdtv"
        if want_uhdtv and file_kind == "dtv":
            messagebox.showwarning("경고", "UHDTV용 Excel 파일을 선택해 주세요")
            return False
        if not want_uhdtv and file_kind == "uhdtv":
            messagebox.showwarning("경고", "DTV용 Excel 파일을 선택해 주세요")
            return False
        return True

    def _browse(self, kind: str):
        if kind == "html":
            path = filedialog.askopenfilename(
                title="HTML 파일 선택",
                filetypes=[("HTML 파일", "*.html *.htm"), ("모든 파일", "*.*")],
            )
            if path and self._html_matches_mode(path):
                self._html_path.set(path)
        elif kind == "excel":
            path = filedialog.askopenfilename(
                title="Excel 파일 선택",
                filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
            )
            if path and self._excel_matches_mode(path):
                self._excel_path.set(path)
        elif kind == "dmb_txt":
            path = filedialog.askopenfilename(
                title="DMB 로그 파일 선택",
                filetypes=[("텍스트", "*.txt"), ("모든 파일", "*.*")],
            )
            if path:
                self._txt_path.set(path)
        elif kind == "dmb_a":
            path = filedialog.askopenfilename(
                title="DMB TX-A Excel 선택",
                filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
            )
            if path:
                try:
                    k = detect_dmb_excel_kind(path)
                except Exception as exc:
                    messagebox.showerror("파일 오류", f"Excel 파일을 읽을 수 없습니다.\n{exc}")
                    return
                if k != "txa":
                    messagebox.showwarning("경고", "DMB TX-A Excel 파일을 선택해 주세요")
                    return
                self._excel_a_path.set(path)
        elif kind == "dmb_b":
            path = filedialog.askopenfilename(
                title="DMB TX-B Excel 선택",
                filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
            )
            if path:
                try:
                    k = detect_dmb_excel_kind(path)
                except Exception as exc:
                    messagebox.showerror("파일 오류", f"Excel 파일을 읽을 수 없습니다.\n{exc}")
                    return
                if k != "txb":
                    messagebox.showwarning("경고", "DMB TX-B Excel 파일을 선택해 주세요")
                    return
                self._excel_b_path.set(path)

    # ── 로그 출력 ───────────────────────────────────────────────────────────

    def _log(self, msg: str, tag: str = "detail"):
        self._log_box.configure(state="normal")
        self._log_box.insert("end", msg + "\n", tag)
        self._log_box.see("end")
        self._log_box.configure(state="disabled")

    def _log_clear(self):
        self._log_box.configure(state="normal")
        self._log_box.delete("1.0", "end")
        self._log_box.configure(state="disabled")

    # ── 실행 ────────────────────────────────────────────────────────────────

    def _on_run(self):
        mode = self._tx_mode.get()

        if mode == "dmb":
            txt_path = self._txt_path.get().strip()
            xa = self._excel_a_path.get().strip()
            xb = self._excel_b_path.get().strip()
            if not txt_path:
                messagebox.showwarning("파일 미선택", "로그 파일(.txt)을 선택해 주세요.")
                return
            if not xa or not xb:
                messagebox.showwarning("파일 미선택", "Excel TX-A와 TX-B 파일을 모두 선택해 주세요.")
                return
            try:
                if detect_dmb_excel_kind(xa) != "txa":
                    messagebox.showwarning("경고", "DMB TX-A Excel 파일을 선택해 주세요")
                    return
                if detect_dmb_excel_kind(xb) != "txb":
                    messagebox.showwarning("경고", "DMB TX-B Excel 파일을 선택해 주세요")
                    return
            except Exception as exc:
                messagebox.showerror("파일 오류", str(exc))
                return
        else:
            html_path = self._html_path.get().strip()
            excel_path = self._excel_path.get().strip()
            if not html_path:
                messagebox.showwarning("파일 미선택", "HTML 파일을 선택해 주세요.")
                return
            if not excel_path:
                messagebox.showwarning("파일 미선택", "Excel 파일을 선택해 주세요.")
                return
            if not self._html_matches_mode(html_path):
                return
            if not self._excel_matches_mode(excel_path):
                return

        self._run_btn.configure(state="disabled")
        self._log_clear()
        self._progress.start(12)

        if mode == "dmb":
            thread = threading.Thread(
                target=self._run_task_dmb,
                args=(self._txt_path.get().strip(), self._excel_a_path.get().strip(), self._excel_b_path.get().strip()),
                daemon=True,
            )
        else:
            thread = threading.Thread(
                target=self._run_task_rohde,
                args=(self._html_path.get().strip(), self._excel_path.get().strip()),
                daemon=True,
            )
        thread.start()

    def _run_task_dmb(self, txt_path: str, excel_a: str, excel_b: str):
        try:
            self.after(0, self._log, "━━━ DMB 로그 파싱 ━━━", "info")
            self.after(0, self._log, f"  파일: {txt_path}", "detail")

            parsed = parse_dmb_log(txt_path)
            co = parsed.get("created_on")
            if co:
                self.after(0, self._log, f"  로그 일자: {co.strftime('%Y-%m-%d')}", "success")

            def on_log(msg: str):
                self.after(0, self._log, msg, "detail")

            self.after(0, self._log, "━━━ Excel TX-A (TX-1) ━━━", "info")
            update_dmb_excel(excel_a, parsed, tx_num=1, log_callback=on_log)

            self.after(0, self._log, "━━━ Excel TX-B (TX-2) ━━━", "info")
            update_dmb_excel(excel_b, parsed, tx_num=2, log_callback=on_log)

            self.after(0, self._log, "━━━ 완료 ━━━", "success")
            self.after(0, self._log, f"  TX-A: {excel_a}", "success")
            self.after(0, self._log, f"  TX-B: {excel_b}", "success")
            self.after(
                0,
                messagebox.showinfo,
                "완료",
                f"Excel 두 파일이 저장되었습니다.\n\nTX-A:\n{excel_a}\n\nTX-B:\n{excel_b}",
            )

        except Exception as exc:
            import traceback
            tb = traceback.format_exc()
            self.after(0, self._log, f"오류 발생: {exc}", "error")
            self.after(0, self._log, tb, "error")
            self.after(0, messagebox.showerror, "오류", f"처리 중 오류가 발생했습니다.\n\n{exc}")

        finally:
            self.after(0, self._progress.stop)
            self.after(0, self._run_btn.configure, {"state": "normal"})

    def _run_task_rohde(self, html_path: str, excel_path: str):
        try:
            self.after(0, self._log, "━━━ HTML 파싱 시작 ━━━", "info")
            self.after(0, self._log, f"  파일: {html_path}")

            num_amps = 6 if self._tx_mode.get() == "uhdtv" else 2
            mode_label = "UHDTV (AMP 6)" if num_amps == 6 else "DTV (AMP 2)"
            self.after(0, self._log, f"  모드: {mode_label}", "info")

            parsed = parse_html(html_path, num_amplifiers=num_amps)

            created_on = parsed.get("created_on")
            if created_on:
                self.after(0, self._log, f"  측정 일시: {created_on.strftime('%Y-%m-%d %H:%M:%S')}", "success")

            fwd = parsed.get("forward_power")
            ref = parsed.get("reflected_power")
            if fwd is not None:
                self.after(0, self._log, f"  Forward Power: {fwd} W", "detail")
            if ref is not None:
                self.after(0, self._log, f"  Reflected Power: {ref} W", "detail")

            for n in range(1, num_amps + 1):
                amp = parsed.get(f"amp{n}", {})
                cnt = len([v for v in amp.values() if v is not None]) if isinstance(amp, dict) else 0
                self.after(0, self._log, f"  AMP{n} 추출 항목: {cnt}개", "detail")

            self.after(0, self._log, "━━━ Excel 갱신 시작 ━━━", "info")

            def on_log(msg: str):
                self.after(0, self._log, msg, "detail")

            update_excel(excel_path, parsed, log_callback=on_log)

            self.after(0, self._log, "━━━ 완료 ━━━", "success")
            self.after(0, self._log, f"  저장 위치: {excel_path}", "success")
            self.after(0, messagebox.showinfo, "완료", f"Excel 파일이 저장되었습니다.\n\n{excel_path}")

        except Exception as exc:
            import traceback
            tb = traceback.format_exc()
            self.after(0, self._log, f"오류 발생: {exc}", "error")
            self.after(0, self._log, tb, "error")
            self.after(0, messagebox.showerror, "오류", f"처리 중 오류가 발생했습니다.\n\n{exc}")

        finally:
            self.after(0, self._progress.stop)
            self.after(0, self._run_btn.configure, {"state": "normal"})


# ── 진입점 ──────────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app = App()
    app.mainloop()
