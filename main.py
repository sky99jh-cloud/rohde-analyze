"""
ROHDE 송신기 로그 분석기
HTML 파라미터 스냅샷을 파싱하여 Excel 파일에 측정 데이터를 기록한다.
"""

import threading
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext, ttk

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
        self.geometry("720x580")
        self.minsize(600, 500)
        self.configure(bg=BG_COLOR)
        self.resizable(True, True)

        self._html_path = tk.StringVar()
        self._excel_path = tk.StringVar()
        self._tx_mode = tk.StringVar(value="dtv")

        self._build_ui()

    # ── UI 구성 ─────────────────────────────────────────────────────────────

    def _build_ui(self):
        # 헤더
        header = tk.Frame(self, bg=ACCENT_COLOR, height=56)
        header.pack(fill="x")
        header.pack_propagate(False)
        tk.Label(
            header,
            text="ROHDE 송신기 로그 분석기",
            bg=ACCENT_COLOR, fg="white",
            font=FONT_TITLE,
        ).pack(side="left", padx=20, pady=14)

        # 메인 컨테이너
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
        ).pack(side="left", padx=(4, 20))
        tk.Radiobutton(
            mode_frame, text="UHDTV (AMP 6개)",
            variable=self._tx_mode, value="uhdtv",
            bg=BG_COLOR, font=FONT_NORMAL, anchor="w",
        ).pack(side="left", padx=4)

        # 파일 선택 프레임
        file_frame = tk.LabelFrame(
            main, text=" 파일 선택 ", bg=BG_COLOR,
            font=FONT_BOLD, fg=ACCENT_COLOR, padx=12, pady=10,
        )
        file_frame.pack(fill="x", pady=(0, 12))

        self._add_file_row(file_frame, "HTML 파일 (로그):", self._html_path, "html", row=0)
        self._add_file_row(file_frame, "Excel 파일 (결과지):", self._excel_path, "excel", row=1)

        # 실행 버튼
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

        # 진행 상태 바
        self._progress = ttk.Progressbar(main, mode="indeterminate", length=300)
        self._progress.pack(fill="x", pady=(0, 8))

        # 결과 로그
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

        # 태그 색상 설정
        self._log_box.tag_config("info",    foreground="#89b4fa")
        self._log_box.tag_config("success", foreground="#a6e3a1")
        self._log_box.tag_config("error",   foreground="#f38ba8")
        self._log_box.tag_config("detail",  foreground="#a6adc8")

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
        """현재 분석 모드와 HTML 종류가 일치하면 True. 불일치·읽기 실패 시 경고 후 False."""
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
        """현재 분석 모드와 Excel 양식이 일치하면 True. 불일치·읽기 실패 시 경고 후 False."""
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
        else:
            path = filedialog.askopenfilename(
                title="Excel 파일 선택",
                filetypes=[("Excel 파일", "*.xlsx *.xlsm"), ("모든 파일", "*.*")],
            )
            if path and self._excel_matches_mode(path):
                self._excel_path.set(path)

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
        html_path  = self._html_path.get().strip()
        excel_path = self._excel_path.get().strip()

        if not html_path:
            messagebox.showwarning("파일 미선택", "HTML 파일을 선택해 주세요.")
            return
        if not excel_path:
            messagebox.showwarning("파일 미선택", "Excel 파일을 선택해 주세요.")
            return
        # HTML/Excel은 찾아보기에서 검증하지만, 분석 모드를 바꾼 뒤에도 불일치 가능
        if not self._html_matches_mode(html_path):
            return
        if not self._excel_matches_mode(excel_path):
            return

        self._run_btn.configure(state="disabled")
        self._log_clear()
        self._progress.start(12)

        thread = threading.Thread(
            target=self._run_task,
            args=(html_path, excel_path),
            daemon=True,
        )
        thread.start()

    def _run_task(self, html_path: str, excel_path: str):
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
