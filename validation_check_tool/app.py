"""GUI for the Line List <-> Instrument Datasheet Validation Check tool.

Replaces the old console input() column-mapping workflow with a visual,
editable mapping table: pick files, let the tool auto-detect every column,
review/fix anything wrong in one screen, then generate the highlighted report.
"""

from __future__ import annotations

import queue
import subprocess
import sys
import threading
from datetime import datetime
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk
from tkinter import ttk

import engine

CONFIG_PATH = str(Path(__file__).resolve().parent / "config.json")

COLORS = {
    "bg": "#F4F6FB",
    "card": "#FFFFFF",
    "header": "#1E3A5F",
    "header_text": "#FFFFFF",
    "accent": "#2E6BE6",
    "accent_dark": "#1F4FB8",
    "text": "#1F2937",
    "muted": "#6B7280",
    "border": "#E2E5EC",
    "ok": "#1E8E3E",
    "warn": "#B7791F",
    "bad": "#C53030",
    "row_ok": "#E6F4EA",
    "row_warn": "#FFF6E0",
    "row_bad": "#FBE7E7",
}

FIELDS_FOR_INSTRUMENT = ["tag", "line", *engine.PROCESS_FIELDS]
FIELDS_FOR_MASTER = ["line", *engine.PROCESS_FIELDS]


def setup_style(root: tk.Tk) -> None:
    root.configure(bg=COLORS["bg"])
    style = ttk.Style(root)
    style.theme_use("clam")

    style.configure(".", background=COLORS["bg"], foreground=COLORS["text"], font=("Segoe UI", 10))
    style.configure("TFrame", background=COLORS["bg"])
    style.configure("Card.TFrame", background=COLORS["card"])
    style.configure("Header.TFrame", background=COLORS["header"])
    style.configure("Header.TLabel", background=COLORS["header"], foreground=COLORS["header_text"],
                    font=("Segoe UI", 16, "bold"))
    style.configure("SubHeader.TLabel", background=COLORS["header"], foreground="#C9D6EA",
                     font=("Segoe UI", 9))
    style.configure("CardTitle.TLabel", background=COLORS["card"], foreground=COLORS["text"],
                    font=("Segoe UI", 11, "bold"))
    style.configure("Muted.TLabel", background=COLORS["card"], foreground=COLORS["muted"])
    style.configure("Card.TLabel", background=COLORS["card"], foreground=COLORS["text"])
    style.configure("Status.TLabel", background=COLORS["bg"], foreground=COLORS["muted"], font=("Segoe UI", 9))

    style.configure("Accent.TButton", background=COLORS["accent"], foreground="white",
                    font=("Segoe UI", 10, "bold"), padding=8, borderwidth=0)
    style.map("Accent.TButton", background=[("active", COLORS["accent_dark"]), ("disabled", "#A9BEEA")])

    style.configure("TButton", padding=6, font=("Segoe UI", 9))
    style.configure("Treeview", rowheight=26, font=("Segoe UI", 9), fieldbackground="white")
    style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"), background="#EEF1F8")
    style.configure("Horizontal.TProgressbar", background=COLORS["accent"])


def card(parent, title=None):
    outer = ttk.Frame(parent, style="TFrame")
    frame = ttk.Frame(outer, style="Card.TFrame", padding=16)
    frame.pack(fill="both", expand=True)
    if title:
        ttk.Label(frame, text=title, style="CardTitle.TLabel").pack(anchor="w", pady=(0, 10))
    return outer, frame


class MainApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Validation Check - Line List vs Instrument Datasheet")
        self.geometry("1020x900")
        self.minsize(900, 760)
        setup_style(self)

        self.master_file: str | None = None
        self.instrument_files: list[str] = []
        self.sheet_mappings: list[engine.SheetMapping] = []
        self.master_sm: engine.SheetMapping | None = None
        self.master_df = None
        self.df_cache: dict = {}
        self.mapping_confirmed = False
        self.last_output_path: str | None = None

        self._msg_queue: queue.Queue = queue.Queue()

        self._build_header()
        self._build_body()
        self._build_footer()
        self.after(100, self._poll_queue)

    # ---------------------------------------------------------- layout

    def _build_header(self):
        header = ttk.Frame(self, style="Header.TFrame")
        header.pack(fill="x")
        inner = ttk.Frame(header, style="Header.TFrame", padding=(20, 16))
        inner.pack(fill="x")
        ttk.Label(inner, text="Validation Check Excel Mapping", style="Header.TLabel").pack(anchor="w")
        ttk.Label(inner, text="Line List ↔ Instrument Datasheet Process Data 자동 비교 · 자동 컬럼 매핑",
                  style="SubHeader.TLabel").pack(anchor="w", pady=(2, 0))

    def _build_body(self):
        body = ttk.Frame(self, padding=16)
        body.pack(fill="both", expand=True)

        # Step 1 - master file
        outer1, c1 = card(body, "① Line List (Master) 파일")
        outer1.pack(fill="x", pady=(0, 12))
        row = ttk.Frame(c1, style="Card.TFrame")
        row.pack(fill="x")
        self.master_label = ttk.Label(row, text="선택된 파일 없음", style="Muted.TLabel")
        self.master_label.pack(side="left", fill="x", expand=True)
        ttk.Button(row, text="파일 선택", command=self.pick_master).pack(side="right")

        # Step 2 - instrument files
        outer2, c2 = card(body, "② Instrument Datasheet 파일 (여러 개 선택 가능, 모든 시트 자동 참조)")
        outer2.pack(fill="both", expand=False, pady=(0, 12))
        list_row = ttk.Frame(c2, style="Card.TFrame")
        list_row.pack(fill="both", expand=True)
        self.inst_listbox = tk.Listbox(list_row, height=5, activestyle="none",
                                        font=("Segoe UI", 9), bd=1, relief="solid",
                                        selectbackground=COLORS["accent"])
        self.inst_listbox.pack(side="left", fill="both", expand=True)
        btns = ttk.Frame(list_row, style="Card.TFrame")
        btns.pack(side="left", fill="y", padx=(10, 0))
        ttk.Button(btns, text="추가", command=self.add_instrument_files).pack(fill="x", pady=2)
        ttk.Button(btns, text="선택 제거", command=self.remove_selected_instrument).pack(fill="x", pady=2)
        ttk.Button(btns, text="전체 제거", command=self.clear_instrument_files).pack(fill="x", pady=2)

        # Step 3 - scan / mapping
        outer3, c3 = card(body, "③ 컬럼 자동 매핑")
        outer3.pack(fill="x", pady=(0, 12))
        row3 = ttk.Frame(c3, style="Card.TFrame")
        row3.pack(fill="x")
        ttk.Label(row3, text="모든 시트를 스캔해서 Tag / Line No / Process Data 컬럼을 자동으로 찾고,\n"
                              "결과를 표로 보여줘요. 잘못 잡힌 컬럼은 그 자리에서 바로 고칠 수 있어요.",
                  style="Card.TLabel").pack(side="left", fill="x", expand=True)
        ttk.Button(row3, text="스캔 & 매핑 확인", style="Accent.TButton",
                   command=self.scan_and_review).pack(side="right")
        self.mapping_status = ttk.Label(c3, text="", style="Muted.TLabel")
        self.mapping_status.pack(anchor="w", pady=(8, 0))

        # Step 4 - generate
        outer4, c4 = card(body, "④ Validation Report 생성")
        outer4.pack(fill="x")
        row4 = ttk.Frame(c4, style="Card.TFrame")
        row4.pack(fill="x")
        ttk.Label(row4, text="Line List 기준으로 Instrument를 매칭하고, Process Data가 다르면\n"
                              "엑셀에서 바로 하이라이트 처리된 리포트를 만들어요.",
                  style="Card.TLabel").pack(side="left", fill="x", expand=True)
        self.generate_btn = ttk.Button(row4, text="Report 생성", style="Accent.TButton",
                                        command=self.generate_report, state="disabled")
        self.generate_btn.pack(side="right")

        # Log panel
        outer5, c5 = card(body, "진행 로그")
        outer5.pack(fill="both", expand=True, pady=(12, 0))
        self.progress = ttk.Progressbar(c5, mode="determinate")
        self.progress.pack(fill="x", pady=(0, 8))
        log_frame = ttk.Frame(c5, style="Card.TFrame")
        log_frame.pack(fill="both", expand=True)
        self.log_text = tk.Text(log_frame, height=8, font=("Consolas", 9), bd=1, relief="solid",
                                 wrap="word", bg="#0F172A", fg="#E2E8F0", insertbackground="white")
        self.log_text.pack(fill="both", expand=True)
        self.log_text.tag_config("info", foreground="#93C5FD")
        self.log_text.tag_config("ok", foreground="#86EFAC")
        self.log_text.tag_config("err", foreground="#FCA5A5")
        self.log_text.configure(state="disabled")

    def _build_footer(self):
        footer = ttk.Frame(self, padding=(16, 6))
        footer.pack(fill="x")
        self.status_label = ttk.Label(footer, text="대기 중", style="Status.TLabel")
        self.status_label.pack(side="left")

    # ---------------------------------------------------------- logging helpers

    def log(self, message: str, level: str = "info"):
        self._msg_queue.put(("log", message, level))

    def set_status(self, text: str):
        self._msg_queue.put(("status", text, None))

    def _poll_queue(self):
        try:
            while True:
                kind, payload, level = self._msg_queue.get_nowait()
                if kind == "log":
                    self.log_text.configure(state="normal")
                    ts = datetime.now().strftime("%H:%M:%S")
                    self.log_text.insert("end", f"[{ts}] {payload}\n", level or "info")
                    self.log_text.see("end")
                    self.log_text.configure(state="disabled")
                elif kind == "status":
                    self.status_label.configure(text=payload)
        except queue.Empty:
            pass
        self.after(100, self._poll_queue)

    # ---------------------------------------------------------- file pickers

    def pick_master(self):
        path = filedialog.askopenfilename(title="Line List (Master) 파일 선택",
                                           filetypes=[("Excel files", "*.xlsx *.xls")])
        if not path:
            return
        self.master_file = path
        self.master_label.configure(text=Path(path).name, style="Card.TLabel")
        self.mapping_confirmed = False
        self.generate_btn.configure(state="disabled")
        self.log(f"Line List 파일 선택: {Path(path).name}")

    def add_instrument_files(self):
        paths = filedialog.askopenfilenames(title="Instrument Datasheet 파일 선택 (여러 개 가능)",
                                             filetypes=[("Excel files", "*.xlsx *.xls")])
        added = 0
        for p in paths:
            if p not in self.instrument_files:
                self.instrument_files.append(p)
                self.inst_listbox.insert("end", Path(p).name)
                added += 1
        if added:
            self.mapping_confirmed = False
            self.generate_btn.configure(state="disabled")
            self.log(f"Instrument Datasheet {added}개 추가됨")

    def remove_selected_instrument(self):
        sel = list(self.inst_listbox.curselection())
        for i in reversed(sel):
            self.inst_listbox.delete(i)
            del self.instrument_files[i]
        if sel:
            self.mapping_confirmed = False
            self.generate_btn.configure(state="disabled")

    def clear_instrument_files(self):
        self.inst_listbox.delete(0, "end")
        self.instrument_files.clear()
        self.mapping_confirmed = False
        self.generate_btn.configure(state="disabled")

    # ---------------------------------------------------------- scan / mapping

    def scan_and_review(self):
        if not self.master_file:
            messagebox.showwarning("파일 필요", "Line List (Master) 파일을 먼저 선택하세요.")
            return
        if not self.instrument_files:
            messagebox.showwarning("파일 필요", "Instrument Datasheet 파일을 1개 이상 추가하세요.")
            return

        self.set_status("스캔 중...")
        self.progress.configure(mode="indeterminate")
        self.progress.start(12)

        def work():
            saved_config = engine.load_config(CONFIG_PATH)
            self.log("Line List 스캔 중...")
            master_sm, master_df = engine.scan_master_file(self.master_file, saved_config)
            self.log("Instrument Datasheet 스캔 중 (모든 시트)...")
            sheet_mappings, df_cache = engine.scan_instrument_files(
                self.instrument_files, saved_config, progress=lambda m: self.log(m)
            )
            self.after(0, lambda: self._on_scan_done(master_sm, master_df, sheet_mappings, df_cache))

        threading.Thread(target=work, daemon=True).start()

    def _on_scan_done(self, master_sm, master_df, sheet_mappings, df_cache):
        self.progress.stop()
        self.progress.configure(mode="determinate", value=0)
        self.set_status("대기 중")
        self.master_sm = master_sm
        self.master_df = master_df
        self.sheet_mappings = sheet_mappings
        self.df_cache = df_cache
        self.log(f"스캔 완료: 시트 {len(sheet_mappings)}개 발견", "ok")
        MappingDialog(self, master_sm, master_df, sheet_mappings, df_cache, self._on_mapping_confirmed)

    def _on_mapping_confirmed(self):
        self.mapping_confirmed = True
        engine.save_config(CONFIG_PATH, [self.master_sm, *self.sheet_mappings])
        included = sum(1 for sm in self.sheet_mappings if sm.include)
        self.mapping_status.configure(
            text=f"매핑 확인 완료 · Instrument 시트 {included}/{len(self.sheet_mappings)}개 포함",
            style="Card.TLabel")
        self.generate_btn.configure(state="normal")
        self.log("매핑 확정 및 저장 완료", "ok")

    # ---------------------------------------------------------- report generation

    def generate_report(self):
        if not self.mapping_confirmed:
            return
        default_name = f"Validation_Report_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx"
        output_path = filedialog.asksaveasfilename(
            title="Validation Report 저장 위치", defaultextension=".xlsx",
            initialfile=default_name, filetypes=[("Excel files", "*.xlsx")])
        if not output_path:
            return

        self.generate_btn.configure(state="disabled")
        self.set_status("리포트 생성 중...")
        self.progress.configure(mode="determinate", value=0)

        def work():
            self.log("Line List 데이터 로딩...")
            master_lines = engine.load_master_lines(self.master_df, self.master_sm.mapping)
            self.log(f"Line 개수: {len(master_lines)}")

            if not master_lines:
                self.after(0, self._on_master_empty)
                return

            self.log("Instrument 데이터 로딩...")
            instrument_rows = engine.load_instrument_rows(self.sheet_mappings, self.df_cache)
            self.log(f"Instrument 개수: {len(instrument_rows)}")
            if not instrument_rows:
                self.log("경고: 매칭된 Instrument가 0개입니다. Instrument 시트의 Line/Tag 매핑을 확인하세요.", "err")

            def progress_cb(msg):
                self.log(msg)

            stats = engine.build_report(master_lines, instrument_rows, output_path, progress=progress_cb)
            self.after(0, lambda: self._on_report_done(stats, output_path))

        threading.Thread(target=work, daemon=True).start()

    def _on_master_empty(self):
        self.generate_btn.configure(state="normal")
        self.set_status("대기 중")
        self.log("Line List에서 Line No를 하나도 찾지 못했습니다.", "err")
        messagebox.showerror(
            "Line List 인식 실패",
            "Line List 파일에서 Line No를 하나도 찾지 못해서 리포트가 비어있게 됩니다.\n\n"
            "③ 스캔 & 매핑 확인 화면에서 'Line List (Master)' 행을 더블클릭해\n"
            "Line No 컬럼(그리고 Process Data 컬럼)이 실제 파일과 맞는지 확인해 주세요.\n"
            "오른쪽에 뜨는 미리보기 값으로 맞는 컬럼인지 확인할 수 있어요."
        )

    def _on_report_done(self, stats: engine.ReportStats, output_path: str):
        self.generate_btn.configure(state="normal")
        self.progress.configure(value=100)
        self.set_status("완료")
        self.last_output_path = output_path
        self.log(f"리포트 생성 완료: {output_path}", "ok")

        msg = (
            f"Total Lines: {stats.total_lines}\n"
            f"Instrument Count: {stats.instrument_count}\n"
            f"Matched: {stats.matched}\n"
            f"Missing Lines: {stats.missing_lines}\n"
            f"Fail Count: {stats.fail_count}\n\n"
            f"저장 위치:\n{output_path}"
        )
        if messagebox.askyesno("완료", msg + "\n\n파일이 있는 폴더를 열까요?"):
            self._open_folder(output_path)

    def _open_folder(self, path: str):
        folder = str(Path(path).parent)
        try:
            if sys.platform.startswith("win"):
                subprocess.run(["explorer", folder])
            elif sys.platform == "darwin":
                subprocess.run(["open", folder])
            else:
                subprocess.run(["xdg-open", folder])
        except Exception as e:
            self.log(f"폴더 열기 실패: {e}", "err")


class MappingDialog(tk.Toplevel):
    SOURCE_LABEL = {"auto": "자동 인식", "default": "기본값", "saved": "저장됨", "manual": "수동 편집", "missing": "미확인"}

    def __init__(self, parent, master_sm, master_df, sheet_mappings, df_cache, on_confirm):
        super().__init__(parent)
        self.parent = parent
        self.master_sm = master_sm
        self.master_df = master_df
        self.sheet_mappings = sheet_mappings
        self.df_cache = df_cache
        self.on_confirm = on_confirm

        self.title("컬럼 매핑 확인")
        self.geometry("1260x580")
        self.minsize(1000, 480)
        self.configure(bg=COLORS["bg"])
        self.transient(parent)
        self.grab_set()

        self._build()
        self._refresh_tree()

    def _build(self):
        top = ttk.Frame(self, padding=(16, 12, 16, 0))
        top.pack(fill="x")
        ttk.Label(top, text="자동으로 인식된 컬럼 매핑이에요. 색이 있는 행은 다시 확인해 주세요. "
                             "행을 더블클릭하면 직접 수정할 수 있어요.",
                  style="Card.TLabel", background=COLORS["bg"]).pack(anchor="w")

        legend = ttk.Frame(top)
        legend.pack(anchor="w", pady=(6, 0))
        for text, color in [("자동/저장됨 인식 완료", COLORS["row_ok"]),
                             ("기본값 사용(확인 필요)", COLORS["row_warn"]),
                             ("컬럼 미확인", COLORS["row_bad"])]:
            sw = tk.Label(legend, text="  ", bg=color, relief="solid", bd=1)
            sw.pack(side="left", padx=(0, 4))
            ttk.Label(legend, text=text, background=COLORS["bg"], style="Status.TLabel").pack(
                side="left", padx=(0, 14))

        process_headers = {
            "p_op": "Op.Press", "p_des": "Des.Press", "t_op": "Op.Temp",
            "t_op_max": "Max Op.Temp", "t_min": "Min Des.Temp", "t_max": "Max Des.Temp",
        }
        columns = ["include", "file", "sheet", "family", "tag", "line", *engine.PROCESS_FIELDS, "status"]
        headers = ["포함", "파일", "시트", "구분", "Tag", "Line No",
                   *[process_headers[f] for f in engine.PROCESS_FIELDS], "인식 상태"]

        tree_frame = ttk.Frame(self, padding=16)
        tree_frame.pack(fill="both", expand=True)
        self.tree = ttk.Treeview(tree_frame, columns=columns, show="headings", selectmode="browse")
        narrow = {"include", "tag", "line", *engine.PROCESS_FIELDS}
        for col, head in zip(columns, headers):
            self.tree.heading(col, text=head)
            width = 70 if col in narrow else 150
            self.tree.column(col, width=width, anchor="center" if col != "file" and col != "sheet" else "w")
        vsb = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=vsb.set)
        self.tree.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")

        self.tree.tag_configure("ok", background=COLORS["row_ok"])
        self.tree.tag_configure("warn", background=COLORS["row_warn"])
        self.tree.tag_configure("bad", background=COLORS["row_bad"])
        self.tree.tag_configure("master", font=("Segoe UI", 9, "bold"))

        self.tree.bind("<Double-1>", self._on_double_click)
        self.tree.bind("<Button-1>", self._on_click)

        bottom = ttk.Frame(self, padding=(16, 0, 16, 16))
        bottom.pack(fill="x")
        ttk.Button(bottom, text="전체 포함", command=lambda: self._set_all_include(True)).pack(side="left")
        ttk.Button(bottom, text="전체 제외", command=lambda: self._set_all_include(False)).pack(side="left", padx=(6, 0))
        ttk.Button(bottom, text="취소", command=self.destroy).pack(side="right")
        ttk.Button(bottom, text="확인 (매핑 저장)", style="Accent.TButton",
                   command=self._confirm).pack(side="right", padx=(0, 8))

    def _row_values(self, sm: engine.SheetMapping):
        is_master = sm.key == engine.MASTER_KEY
        include_text = "필수" if is_master else ("✔ 포함" if sm.include else "✘ 제외")
        tag = "-" if is_master else sm.mapping.get("tag", "")
        return (
            include_text,
            Path(sm.file).name,
            sm.sheet,
            sm.family or "미분류",
            tag,
            sm.mapping.get("line", ""),
            *[sm.mapping.get(f, "") for f in engine.PROCESS_FIELDS],
            self._status_text(sm),
        )

    def _status_text(self, sm: engine.SheetMapping):
        counts = {}
        for v in sm.source.values():
            counts[v] = counts.get(v, 0) + 1
        parts = [f"{self.SOURCE_LABEL.get(k, k)} {v}" for k, v in counts.items()]
        return ", ".join(parts)

    def _row_tag(self, sm: engine.SheetMapping):
        sources = set(sm.source.values())
        if "missing" in sources:
            return "bad"
        if "default" in sources:
            return "warn"
        return "ok"

    def _refresh_tree(self):
        self.tree.delete(*self.tree.get_children())
        self.tree.insert("", "end", iid="master", values=self._row_values(self.master_sm),
                          tags=(self._row_tag(self.master_sm), "master"))
        for i, sm in enumerate(self.sheet_mappings):
            self.tree.insert("", "end", iid=str(i), values=self._row_values(sm), tags=(self._row_tag(sm),))

    def _on_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
        col = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        if not row or col != "#1":
            return
        if row == "master":
            return
        sm = self.sheet_mappings[int(row)]
        sm.include = not sm.include
        self.tree.item(row, values=self._row_values(sm))

    def _on_double_click(self, event):
        col = self.tree.identify_column(event.x)
        if col == "#1":
            return
        row = self.tree.identify_row(event.y)
        if not row:
            return
        if row == "master":
            sm, df = self.master_sm, self.master_df
        else:
            sm = self.sheet_mappings[int(row)]
            df = self.df_cache.get((sm.file, sm.sheet))
        EditRowDialog(self, sm, df, lambda: self._refresh_row(row))

    def _refresh_row(self, row):
        sm = self.master_sm if row == "master" else self.sheet_mappings[int(row)]
        self.tree.item(row, values=self._row_values(sm), tags=(self._row_tag(sm), "master") if row == "master" else (self._row_tag(sm),))

    def _set_all_include(self, value: bool):
        for i, sm in enumerate(self.sheet_mappings):
            sm.include = value
            self.tree.item(str(i), values=self._row_values(sm))

    def _confirm(self):
        self.destroy()
        self.on_confirm()


class EditRowDialog(tk.Toplevel):
    def __init__(self, parent, sm: engine.SheetMapping, df, on_saved):
        super().__init__(parent)
        self.sm = sm
        self.df = df
        self.on_saved = on_saved
        self.entries: dict[str, tk.Entry] = {}
        self.previews: dict[str, ttk.Label] = {}

        self.title(f"매핑 편집 - {sm.sheet}")
        self.geometry("640x520")
        self.minsize(560, 320)
        self.configure(bg=COLORS["bg"])
        self.transient(parent)
        self.grab_set()

        fields = FIELDS_FOR_MASTER if sm.key == engine.MASTER_KEY else FIELDS_FOR_INSTRUMENT
        self._build(fields)

    def _build(self, fields):
        header = ttk.Frame(self, padding=(16, 16, 16, 0))
        header.pack(fill="x")
        ttk.Label(header, text=f"{Path(self.sm.file).name} / {self.sm.sheet}",
                  font=("Segoe UI", 10, "bold")).pack(anchor="w")
        ttk.Label(header, text="엑셀 컬럼 문자(A, B, AC ...)를 입력하면 오른쪽에 실제 헤더/샘플 값이 나와요.",
                  style="Status.TLabel").pack(anchor="w")

        # Buttons are packed to the bottom first so they always stay visible and
        # reachable, no matter how tall the (scrollable) field list grows.
        btns = ttk.Frame(self, padding=(16, 8, 16, 16))
        btns.pack(side="bottom", fill="x")
        ttk.Button(btns, text="자동 재탐지", command=self._auto_redetect).pack(side="left")
        ttk.Button(btns, text="취소", command=self.destroy).pack(side="right")
        ttk.Button(btns, text="저장", style="Accent.TButton", command=self._save).pack(side="right", padx=(0, 8))

        body = ttk.Frame(self, padding=(16, 12, 16, 0))
        body.pack(side="top", fill="both", expand=True)
        canvas = tk.Canvas(body, bg=COLORS["bg"], highlightthickness=0)
        vsb = ttk.Scrollbar(body, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        canvas.pack(side="left", fill="both", expand=True)
        vsb.pack(side="left", fill="y")

        grid = ttk.Frame(canvas)
        grid_window = canvas.create_window((0, 0), window=grid, anchor="nw")
        grid.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        canvas.bind("<Configure>", lambda e: canvas.itemconfigure(grid_window, width=e.width))

        def on_mousewheel(event):
            if event.num == 4:
                canvas.yview_scroll(-1, "units")
            elif event.num == 5:
                canvas.yview_scroll(1, "units")
            else:
                canvas.yview_scroll(-1 * (event.delta // 120 or (1 if event.delta > 0 else -1)), "units")

        def bind_wheel(_e=None):
            canvas.bind_all("<MouseWheel>", on_mousewheel)
            canvas.bind_all("<Button-4>", on_mousewheel)
            canvas.bind_all("<Button-5>", on_mousewheel)

        def unbind_wheel(_e=None):
            canvas.unbind_all("<MouseWheel>")
            canvas.unbind_all("<Button-4>")
            canvas.unbind_all("<Button-5>")

        self.bind("<Enter>", bind_wheel)
        self.bind("<Leave>", unbind_wheel)
        self.bind("<Destroy>", unbind_wheel)

        for r, f in enumerate(fields):
            ttk.Label(grid, text=engine.FIELD_LABELS[f], width=22).grid(row=r, column=0, sticky="w", pady=4)
            entry = tk.Entry(grid, width=6, justify="center", font=("Consolas", 10))
            entry.insert(0, self.sm.mapping.get(f, ""))
            entry.grid(row=r, column=1, padx=(6, 10), pady=4)
            entry.bind("<KeyRelease>", lambda e, field=f: self._update_preview(field))
            self.entries[f] = entry

            preview = ttk.Label(grid, text="", style="Status.TLabel", wraplength=380, justify="left")
            preview.grid(row=r, column=2, sticky="w", pady=4)
            self.previews[f] = preview
            self._update_preview(f)

        grid.columnconfigure(2, weight=1)

    def _update_preview(self, field):
        col = self.entries[field].get().strip().upper()
        if not col or self.df is None:
            self.previews[field].configure(text="")
            return
        try:
            text = engine.column_preview(self.df, col)
        except Exception:
            text = "(잘못된 컬럼)"
        self.previews[field].configure(text=text)

    def _auto_redetect(self):
        if self.df is None:
            return
        auto_map = engine.auto_detect_mapping(self.df)
        default = engine.MASTER_DEFAULT if self.sm.key == engine.MASTER_KEY else engine.DEFAULT_MAP.get(self.sm.family, {})
        mapping, _ = engine.resolve_mapping(default, auto_map)
        for f, entry in self.entries.items():
            entry.delete(0, "end")
            entry.insert(0, mapping.get(f, ""))
            self._update_preview(f)

    def _save(self):
        for f, entry in self.entries.items():
            val = entry.get().strip().upper()
            if not val:
                messagebox.showwarning("입력 필요", f"{engine.FIELD_LABELS[f]} 컬럼을 입력하세요.")
                return
            self.sm.mapping[f] = val
            self.sm.source[f] = "manual"
        self.destroy()
        self.on_saved()


def main():
    app = MainApp()
    app.mainloop()


if __name__ == "__main__":
    main()
