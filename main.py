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


# ── R&S 프로모 스타일: 딥 블루 대기 + HUD 패널 + 시안 글로우 (frontend-design) ─
APP_BG = "#040a10"
PANEL_BG = "#0c1620"
PANEL_BORDER = "#1a4a62"
PANEL_GLOW = "#2ec4e8"
ACCENT_CYAN = "#5dd5ff"
ACCENT_CYAN_SOFT = "#3eb8dc"
ACCENT_ORANGE = "#ff8c42"
TEXT_PRIMARY = "#f2f8ff"
TEXT_MUTED = "#7a9bb8"
ENTRY_BG = "#050c14"
LOG_BG = "#020810"
LOG_FG = "#c5e4f5"
BTN_PRIMARY = "#0088aa"
BTN_HOVER = "#00a8cc"
BTN_SECONDARY_BG = "#0f2535"
BTN_SECONDARY_HOVER = "#153548"
RADIO_SELECT = "#143044"
FONT_UI = "Segoe UI"
FONT_NORMAL = (FONT_UI, 10)
FONT_BOLD = (FONT_UI, 10, "bold")
FONT_TITLE = (FONT_UI, 16, "bold")
FONT_SUB = (FONT_UI, 9)
FONT_MONO = ("Consolas", 9)


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("ROHDE 송신기 로그 분석기")
        self.geometry("800x700")
        self.minsize(620, 540)
        self.configure(bg=APP_BG)
        self.resizable(True, True)

        self._html_path = tk.StringVar()
        self._excel_path = tk.StringVar()
        self._txt_path = tk.StringVar()
        self._excel_a_path = tk.StringVar()
        self._excel_b_path = tk.StringVar()
        self._tx_mode = tk.StringVar(value="dtv")

        self._setup_ttk_styles()
        self._build_ui()
        self._tx_mode.trace_add("write", self._on_mode_change)
        self._on_mode_change()

    def _setup_ttk_styles(self) -> None:
        style = ttk.Style()
        style.theme_use("clam")
        style.configure(
            "TProgressbar",
            background=ACCENT_CYAN,
            troughcolor="#0f2838",
            borderwidth=0,
            lightcolor=ACCENT_CYAN_SOFT,
            darkcolor=ACCENT_CYAN,
        )

    # ── UI 구성 ─────────────────────────────────────────────────────────────

    def _paint_header(self, _event=None) -> None:
        c = self._header_canvas
        w = max(c.winfo_width(), 2)
        h = 100
        c.delete("all")
        # 좌·하단이 살짝 밝고 우·상단이 깊은 네이비 (캡처 느낌의 대각 기류)
        c_tl = (0, 12, 28)
        c_tr = (2, 8, 18)
        c_bl = (12, 55, 82)
        c_br = (4, 28, 45)

        def _pix(x: int, y: int) -> str:
            tx, ty = x / max(w - 1, 1), y / max(h - 1, 1)
            r = int(
                c_tl[0] * (1 - tx) * (1 - ty)
                + c_tr[0] * tx * (1 - ty)
                + c_bl[0] * (1 - tx) * ty
                + c_br[0] * tx * ty
            )
            g = int(
                c_tl[1] * (1 - tx) * (1 - ty)
                + c_tr[1] * tx * (1 - ty)
                + c_bl[1] * (1 - tx) * ty
                + c_br[1] * tx * ty
            )
            b = int(
                c_tl[2] * (1 - tx) * (1 - ty)
                + c_tr[2] * tx * (1 - ty)
                + c_bl[2] * (1 - tx) * ty
                + c_br[2] * tx * ty
            )
            return f"#{r:02x}{g:02x}{b:02x}"

        step = 8
        for y in range(0, h, step):
            y2 = min(y + step, h)
            for x in range(0, w, step):
                x2 = min(x + step, w)
                cx = min(x + step // 2, w - 1)
                cy = min(y + step // 2, h - 1)
                color = _pix(cx, cy)
                c.create_rectangle(x, y, x2, y2, fill=color, outline=color, width=0)

        # 기술 그리드 (저채도)
        for gx in range(0, w, 28):
            c.create_line(gx, 0, gx, h, fill="#1a3044", width=1)
        for gy in range(0, h, 24):
            c.create_line(0, gy, w, gy, fill="#152838", width=1)

        # 워터마크 타이포 (배경 레이어)
        c.create_text(
            w * 0.52,
            h * 0.55,
            text="UHF",
            anchor="center",
            fill="#0d1f2e",
            font=(FONT_UI, 42, "bold"),
        )
        c.create_text(
            w * 0.52,
            h * 0.92,
            text="medium power",
            anchor="s",
            fill="#0a1824",
            font=(FONT_UI, 11),
        )

        # 우상단 기술 아이콘 (동심원·타깃)
        cx, cy = w - 48, 34
        for ri, col in ((22, "#1e4058"), (15, "#2a5570"), (8, ACCENT_CYAN)):
            c.create_oval(cx - ri, cy - ri, cx + ri, cy + ri, outline=col, width=1)
        c.create_line(cx - 26, cy, cx + 26, cy, fill="#3a6a88", width=1)
        c.create_line(cx, cy - 26, cx, cy + 26, fill="#3a6a88", width=1)

        c.create_text(
            24, 28, text="ROHDE 송신기 로그 분석기",
            anchor="w", fill=TEXT_PRIMARY, font=FONT_TITLE,
        )
        c.create_text(
            24, 58, text="Parameter snapshot → Excel workbook",
            anchor="w", fill=TEXT_MUTED, font=FONT_SUB,
        )

        # 스펙트럼 바 (화이트 → 시안 → 오렌지·핑크)
        bar_w = 3
        for i in range(14):
            bh = 14 + (i * 11) % 42
            col = ("#f0f9ff", "#7dd3fc", ACCENT_CYAN, "#ffb37a", "#ff6eb4")[i % 5]
            x0 = w - 28 - i * 6
            if x0 < 300:
                break
            c.create_rectangle(x0, h - 8 - bh, x0 + bar_w, h - 6, fill=col, outline="")

        c.create_line(0, h - 1, w, h - 1, fill=PANEL_GLOW, width=1)

    def _build_ui(self) -> None:
        accent_strip = tk.Frame(self, bg=PANEL_GLOW, height=3)
        accent_strip.pack(fill="x")
        accent_strip.pack_propagate(False)

        self._header_canvas = tk.Canvas(self, height=100, highlightthickness=0, bd=0)
        self._header_canvas.pack(fill="x")
        self._header_canvas.bind("<Configure>", self._paint_header)

        main = tk.Frame(self, bg=APP_BG, padx=22, pady=18)
        main.pack(fill="both", expand=True)

        lf_kw = {
            "bg": PANEL_BG,
            "font": FONT_BOLD,
            "fg": ACCENT_CYAN,
            "highlightbackground": PANEL_GLOW,
            "highlightthickness": 1,
            "labelanchor": "nw",
        }
        rb_kw = {
            "bg": PANEL_BG,
            "fg": TEXT_PRIMARY,
            "font": FONT_NORMAL,
            "anchor": "w",
            "activebackground": PANEL_BG,
            "activeforeground": ACCENT_CYAN,
            "selectcolor": RADIO_SELECT,
            "highlightthickness": 0,
            "bd": 0,
        }

        mode_frame = tk.LabelFrame(main, text=" 분석 모드 ", padx=14, pady=10, **lf_kw)
        mode_frame.pack(fill="x", pady=(0, 10))
        tk.Radiobutton(
            mode_frame, text="DTV (AMP 2개)",
            variable=self._tx_mode, value="dtv",
            **rb_kw,
        ).pack(side="left", padx=(4, 16))
        tk.Radiobutton(
            mode_frame, text="UHDTV (AMP 6개)",
            variable=self._tx_mode, value="uhdtv",
            **rb_kw,
        ).pack(side="left", padx=(0, 16))
        tk.Radiobutton(
            mode_frame, text="DMB (TX-A / TX-B)",
            variable=self._tx_mode, value="dmb",
            **rb_kw,
        ).pack(side="left", padx=4)

        file_frame = tk.LabelFrame(main, text=" 파일 선택 ", padx=14, pady=12, **lf_kw)
        file_frame.pack(fill="x", pady=(0, 12))

        self._rohde_files = tk.Frame(file_frame, bg=PANEL_BG)
        self._add_file_row(self._rohde_files, "HTML 파일 (로그):", self._html_path, "html", row=0)
        self._add_file_row(self._rohde_files, "Excel 파일 (결과지):", self._excel_path, "excel", row=1)

        self._dmb_files = tk.Frame(file_frame, bg=PANEL_BG)
        self._add_file_row(self._dmb_files, "로그 파일 (.txt):", self._txt_path, "dmb_txt", row=0)
        self._add_file_row(self._dmb_files, "Excel TX-A:", self._excel_a_path, "dmb_a", row=1)
        self._add_file_row(self._dmb_files, "Excel TX-B:", self._excel_b_path, "dmb_b", row=2)

        self._run_btn = tk.Button(
            main,
            text="▶  분석 및 Excel 저장",
            command=self._on_run,
            bg=BTN_PRIMARY,
            fg="#ffffff",
            font=(FONT_UI, 11, "bold"),
            relief="flat",
            cursor="hand2",
            padx=24,
            pady=10,
            activebackground=BTN_HOVER,
            activeforeground="#ffffff",
            bd=0,
            highlightthickness=1,
            highlightbackground=ACCENT_ORANGE,
        )
        self._run_btn.pack(pady=(0, 12))
        self._run_btn.bind(
            "<Enter>",
            lambda e: self._run_btn.config(bg=BTN_HOVER, fg="#ffffff"),
        )
        self._run_btn.bind(
            "<Leave>",
            lambda e: self._run_btn.config(bg=BTN_PRIMARY, fg="#ffffff"),
        )

        self._progress = ttk.Progressbar(main, mode="indeterminate", length=320, style="TProgressbar")
        self._progress.pack(fill="x", pady=(0, 8))

        log_frame = tk.LabelFrame(main, text=" 처리 결과 ", padx=8, pady=8, **lf_kw)
        log_frame.pack(fill="both", expand=True)

        self._log_box = scrolledtext.ScrolledText(
            log_frame,
            bg=LOG_BG,
            fg=LOG_FG,
            font=FONT_MONO,
            state="disabled",
            relief="flat",
            wrap="word",
            insertbackground=ACCENT_CYAN,
            selectbackground="#1a4a62",
            selectforeground=TEXT_PRIMARY,
            highlightthickness=1,
            highlightbackground=PANEL_BORDER,
        )
        self._log_box.pack(fill="both", expand=True, padx=6, pady=6)

        self._log_box.tag_config("info", foreground=ACCENT_CYAN)
        self._log_box.tag_config("success", foreground="#6ee7b7")
        self._log_box.tag_config("error", foreground="#fb7185")
        self._log_box.tag_config("detail", foreground=TEXT_MUTED)

        self.after_idle(self._paint_header)

    def _on_mode_change(self, *args):
        mode = self._tx_mode.get()
        if mode == "dmb":
            self._rohde_files.pack_forget()
            self._dmb_files.pack(fill="x")
        else:
            self._dmb_files.pack_forget()
            self._rohde_files.pack(fill="x")

    def _add_file_row(self, parent, label_text, str_var, kind, row):
        tk.Label(
            parent,
            text=label_text,
            bg=PANEL_BG,
            fg=TEXT_PRIMARY,
            font=FONT_NORMAL,
            width=20,
            anchor="w",
        ).grid(row=row, column=0, sticky="w", pady=4)

        entry = tk.Entry(
            parent,
            textvariable=str_var,
            state="readonly",
            relief="flat",
            font=FONT_NORMAL,
            width=46,
            readonlybackground=ENTRY_BG,
            fg=TEXT_PRIMARY,
            highlightthickness=1,
            highlightbackground=PANEL_BORDER,
            highlightcolor=ACCENT_CYAN_SOFT,
            bd=0,
        )
        entry.grid(row=row, column=1, padx=(6, 6), sticky="ew", pady=4)

        btn = tk.Button(
            parent,
            text="찾아보기…",
            command=lambda k=kind: self._browse(k),
            bg=BTN_SECONDARY_BG,
            fg=TEXT_PRIMARY,
            relief="flat",
            font=FONT_NORMAL,
            cursor="hand2",
            padx=10,
            pady=4,
            activebackground=BTN_SECONDARY_HOVER,
            activeforeground=ACCENT_CYAN,
            bd=0,
            highlightthickness=1,
            highlightbackground=PANEL_BORDER,
        )
        btn.grid(row=row, column=2, pady=4)
        btn.bind("<Enter>", lambda e: btn.config(bg=BTN_SECONDARY_HOVER))
        btn.bind("<Leave>", lambda e: btn.config(bg=BTN_SECONDARY_BG))

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

    def _show_analysis_result(self, alerts: list[str], save_summary: str) -> None:
        """1년 평균 대비 임계값 이탈 시 경고, 없으면 정상 메시지 (FWD·REF·AMP Temp·특이사항 50%, 그 외 20%)."""
        if alerts:
            for a in alerts:
                self.after(0, self._log, f"  [편차] {a}", "error")
            body = (
                "다음 항목이 이전 시트(최근 1년) 평균 대비 임계값 이상 차이납니다 "
                "(FWD·REF·AMP Temp·특이사항: 50%, 그 외: 20%). "
                "해당 셀은 빨간색으로 표시되었습니다.\n\n"
                + "\n".join(alerts)
                + "\n\n"
                + save_summary
            )
            self.after(0, messagebox.showwarning, "분석 완료 — 편차 알림", body)
        else:
            self.after(
                0,
                messagebox.showinfo,
                "분석 완료",
                "모든 데이터들이 정상적입니다.\n\n" + save_summary,
            )

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
            _, alerts_a = update_dmb_excel(excel_a, parsed, tx_num=1, log_callback=on_log)

            self.after(0, self._log, "━━━ Excel TX-B (TX-2) ━━━", "info")
            _, alerts_b = update_dmb_excel(excel_b, parsed, tx_num=2, log_callback=on_log)

            self.after(0, self._log, "━━━ 완료 ━━━", "success")
            self.after(0, self._log, f"  TX-A: {excel_a}", "success")
            self.after(0, self._log, f"  TX-B: {excel_b}", "success")

            combined = [f"[TX-A] {x}" for x in alerts_a] + [f"[TX-B] {x}" for x in alerts_b]
            save_summary = f"TX-A:\n{excel_a}\n\nTX-B:\n{excel_b}"
            self.after(0, self._show_analysis_result, combined, save_summary)

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

            _, deviation_alerts = update_excel(excel_path, parsed, log_callback=on_log)

            self.after(0, self._log, "━━━ 완료 ━━━", "success")
            self.after(0, self._log, f"  저장 위치: {excel_path}", "success")
            self.after(0, self._show_analysis_result, deviation_alerts, f"저장 파일:\n{excel_path}")

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
