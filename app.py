# -*- coding: utf-8 -*-
"""
PIT Audit Pro (v1.1) — Phần mềm chuyên dụng Đối chiếu & Tính Thuế TNCN:
- Nạp nhiều file XML / cả thư mục XML QTT TNCN nhiều năm (có cả tờ khai chính thức & bổ sung).
- Lọc theo khoảng năm [Từ năm ... đến năm ...], tự động chốt tờ khai bổ sung cuối cùng của từng năm
  và kế thừa phụ lục 05-3 từ lần kê khai trước.
- Xuất từng năm ra file `QTT_TNCN_YYYY.xlsx` chuẩn 100% mẫu (TK, 05-1, 05-2, 05-3) giống tax_xml_to_excel.
- Đối chiếu Người phụ thuộc (file TMS .xlsx/.csv) TRỰC TIẾP trên các cột bên phải (L..S) của sheet `05-3`
  (kèm sheet tổng hợp lỗi `DoiChieu_NPT` và bản sạch `DS_NPT_TMS`).
- Tùy chọn quét thêm sổ chi tiết tài khoản chi phí (`TK 334, 353, 641, 642, 622, 627, 154...`).
"""

from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import Dict, List, Optional, Tuple

from pit_workflow import export_yearly_tncn_with_audit, reconcile_existing_tncn_excel
from tax_parser import TaxDeclaration, parse_tax_xml


APP_VERSION = "v1.3"


class PitAuditProApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"PIT Audit Pro — Đối chiếu & Tính Thuế TNCN Chuyên sâu ({APP_VERSION})")
        self.root.geometry("1140x760")
        self.root.minsize(980, 640)

        self.declarations: List[TaxDeclaration] = []
        self.file_paths_set: set[str] = set()
        self.tms_files: List[str] = []
        self.ledger_files: List[str] = []

        self._setup_style()
        self._build_ui()

    def _setup_style(self) -> None:
        style = ttk.Style()
        try:
            style.theme_use("vista")
        except tk.TclError:
            style.theme_use("clam")
        style.configure("Header.TLabel", font=("Segoe UI", 13, "bold"), foreground="#1F4E78")
        style.configure("Primary.TButton", font=("Segoe UI", 10, "bold"))
        style.configure("Treeview.Heading", font=("Segoe UI", 9, "bold"))

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=(12, 10, 12, 4))
        top.pack(fill=tk.X)
        ttk.Label(
            top,
            text=f"PIT AUDIT PRO — ĐỐI CHIẾU & TÍNH THUẾ TNCN CHUYÊN SÂU ({APP_VERSION})",
            style="Header.TLabel",
        ).pack(side=tk.LEFT)

        # 1. Chọn nhiều file XML / thư mục XML nhiều năm (có cả tờ khai bổ sung)
        f1 = ttk.LabelFrame(self.root, text="1. Chọn tờ khai Quyết toán TNCN XML nhiều năm (hỗ trợ cả tờ khai chính thức & bổ sung)", padding=8)
        f1.pack(fill=tk.X, padx=12, pady=4)

        ttk.Button(f1, text="Chọn nhiều file XML...", command=self.add_files).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(f1, text="Chọn thư mục (quét cả thư mục con)...", command=self.add_folder).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(f1, text="Xóa mục đang chọn", command=self.remove_selected).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(f1, text="Xóa tất cả", command=self.clear_all).pack(side=tk.LEFT)

        self.lbl_count = ttk.Label(f1, text="Tổng: 0 file TNCN", font=("Segoe UI", 9, "bold"))
        self.lbl_count.pack(side=tk.RIGHT)

        # 1b. Bộ lọc khoảng năm Từ năm ... đến năm ...
        f_yr = ttk.LabelFrame(self.root, text="2. Lọc kỳ / năm quyết toán cần xuất (tự động chốt tờ khai bổ sung cuối cùng mỗi năm & kế thừa 05-3)", padding=8)
        f_yr.pack(fill=tk.X, padx=12, pady=4)

        self.var_filter_year = tk.BooleanVar(value=False)
        ttk.Checkbutton(
            f_yr,
            text="Chỉ lấy từ năm:",
            variable=self.var_filter_year,
            command=self._toggle_year_filter,
        ).pack(side=tk.LEFT, padx=(0, 4))

        self.var_year_from = tk.StringVar(value="2020")
        self.ent_year_from = ttk.Entry(f_yr, textvariable=self.var_year_from, width=7, state=tk.DISABLED)
        self.ent_year_from.pack(side=tk.LEFT, padx=(0, 6))
        self.var_year_from.trace_add("write", lambda *_: self.refresh_table())

        ttk.Label(f_yr, text="đến năm:").pack(side=tk.LEFT, padx=(0, 4))
        self.var_year_to = tk.StringVar(value="2025")
        self.ent_year_to = ttk.Entry(f_yr, textvariable=self.var_year_to, width=7, state=tk.DISABLED)
        self.ent_year_to.pack(side=tk.LEFT, padx=(0, 10))
        self.var_year_to.trace_add("write", lambda *_: self.refresh_table())

        ttk.Label(
            f_yr,
            text="(Bỏ tick để xuất tất cả các năm có trong danh sách; mỗi năm xuất 1 file QTT_TNCN_YYYY.xlsx chuẩn 4 sheet TK, 05-1, 05-2, 05-3)",
            font=("Segoe UI", 8, "italic"),
            foreground="#555555",
        ).pack(side=tk.LEFT)

        # 2. Chọn 1 hoặc nhiều file Người phụ thuộc hệ thống TMS để đối chiếu trực tiếp trên 05-3
        f_tms = ttk.LabelFrame(self.root, text="3. File Người phụ thuộc kết xuất từ TMS (.xlsx / .csv) — Hỗ trợ chọn nhiều file, đối chiếu trực tiếp trên 05-3", padding=8)
        f_tms.pack(fill=tk.X, padx=12, pady=4)

        self.var_tms_path = tk.StringVar(value="")
        ttk.Label(f_tms, text="File NPT TMS:").pack(side=tk.LEFT, padx=(0, 6))
        ttk.Entry(f_tms, textvariable=self.var_tms_path, state="readonly").pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(f_tms, text="Thêm file NPT (.xlsx/.csv)...", command=self.pick_tms_file).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(f_tms, text="Xóa tất cả", command=self.clear_tms_files).pack(side=tk.LEFT)

        # 3. Chọn nhiều file / cả thư mục Sổ chi tiết tài khoản & Bảng tính lương/thuế nhiều năm
        f_led = ttk.LabelFrame(
            self.root,
            text="4. Sổ chi tiết tài khoản (TK 334, 353, 641, 642, 622, 627, 338, 3335...) & Bảng lương nhiều năm (tự động tách theo từng năm)",
            padding=8,
        )
        f_led.pack(fill=tk.X, padx=12, pady=4)

        led_btns = ttk.Frame(f_led)
        led_btns.pack(fill=tk.X, pady=(0, 4))
        ttk.Button(led_btns, text="Thêm nhiều file (.xlsx)...", command=self.add_ledger_files).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(led_btns, text="Chọn thư mục sổ chi tiết / bảng lương (quét cả thư mục con)...", command=self.add_ledger_folder).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(led_btns, text="Xóa mục đang chọn", command=self.remove_selected_ledgers).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(led_btns, text="Xóa tất cả", command=self.clear_ledgers).pack(side=tk.LEFT)
        self.lbl_ledgers = ttk.Label(led_btns, text="Tổng: 0 file sổ chi tiết / bảng lương", font=("Segoe UI", 9, "bold"))
        self.lbl_ledgers.pack(side=tk.RIGHT)

        led_list_frame = ttk.Frame(f_led)
        led_list_frame.pack(fill=tk.X)
        self.lst_ledgers = tk.Listbox(led_list_frame, height=4, selectmode=tk.EXTENDED, font=("Segoe UI", 8))
        led_sb = ttk.Scrollbar(led_list_frame, orient=tk.VERTICAL, command=self.lst_ledgers.yview)
        self.lst_ledgers.configure(yscrollcommand=led_sb.set)
        self.lst_ledgers.pack(side=tk.LEFT, fill=tk.X, expand=True)
        led_sb.pack(side=tk.RIGHT, fill=tk.Y)

        # Bảng xem trước danh sách file XML TNCN
        tbl_frame = ttk.Frame(self.root, padding=(12, 4, 12, 4))
        tbl_frame.pack(fill=tk.BOTH, expand=True)

        cols = ("file_name", "ma_tkhai", "mst", "ten_nnt", "year", "so_lan", "ngay_lap", "status")
        self.tree = ttk.Treeview(tbl_frame, columns=cols, show="headings", selectmode="extended")
        headers: Dict[str, Tuple[str, int, str]] = {
            "file_name": ("Tên file XML", 230, "w"),
            "ma_tkhai": ("Mã TK", 80, "center"),
            "mst": ("MST", 115, "center"),
            "ten_nnt": ("Tên đơn vị chi trả", 250, "w"),
            "year": ("Năm", 70, "center"),
            "so_lan": ("Lần BS", 70, "center"),
            "ngay_lap": ("Ngày lập", 95, "center"),
            "status": ("Trạng thái", 130, "w"),
        }
        for c, (txt, w, anc) in headers.items():
            self.tree.heading(c, text=txt)
            self.tree.column(c, width=w, anchor=anc)  # type: ignore[arg-type]

        vsb = ttk.Scrollbar(tbl_frame, orient=tk.VERTICAL, command=self.tree.yview)
        hsb = ttk.Scrollbar(tbl_frame, orient=tk.HORIZONTAL, command=self.tree.xview)
        self.tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        self.tree.grid(row=0, column=0, sticky="nsew")
        vsb.grid(row=0, column=1, sticky="ns")
        hsb.grid(row=1, column=0, sticky="ew")
        tbl_frame.rowconfigure(0, weight=1)
        tbl_frame.columnconfigure(0, weight=1)

        # Thanh đáy
        bottom = ttk.Frame(self.root, padding=(12, 8, 12, 12))
        bottom.pack(fill=tk.X)

        self.var_open_after = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom, text="Mở thư mục sau khi xuất", variable=self.var_open_after).pack(side=tk.LEFT, padx=(0, 12))

        self.lbl_status = ttk.Label(bottom, text="Sẵn sàng", font=("Segoe UI", 9, "italic"))
        self.lbl_status.pack(side=tk.LEFT, fill=tk.X, expand=True)

        self.btn_export = ttk.Button(
            bottom,
            text="XUẤT FILE QTT TNCN THEO NĂM & ĐỐI CHIẾU",
            style="Primary.TButton",
            command=self.run_export,
        )
        self.btn_export.pack(side=tk.RIGHT, padx=(8, 0))

        self.btn_reconcile_excel = ttk.Button(
            bottom,
            text="ĐỐI CHIẾU NPT TRÊN EXCEL ĐÃ CÓ",
            command=self.run_reconcile_existing,
        )
        self.btn_reconcile_excel.pack(side=tk.RIGHT)

    def _toggle_year_filter(self) -> None:
        state = tk.NORMAL if self.var_filter_year.get() else tk.DISABLED
        self.ent_year_from.config(state=state)
        self.ent_year_to.config(state=state)
        self.refresh_table()

    def _get_year_bounds(self) -> Tuple[Optional[int], Optional[int]]:
        if not self.var_filter_year.get():
            return None, None
        y_from: Optional[int] = None
        y_to: Optional[int] = None
        s1 = self.var_year_from.get().strip()
        s2 = self.var_year_to.get().strip()
        if s1.isdigit():
            y_from = int(s1)
        if s2.isdigit():
            y_to = int(s2)
        return y_from, y_to

    def _matches_year(self, d: TaxDeclaration) -> bool:
        y_from, y_to = self._get_year_bounds()
        if y_from is not None and d.year is not None and d.year < y_from:
            return False
        if y_to is not None and d.year is not None and d.year > y_to:
            return False
        return True

    def _ingest_xml_paths(self, paths: List[str]) -> None:
        new_paths = [os.path.normpath(p) for p in paths if p.lower().endswith(".xml") and os.path.normpath(p) not in self.file_paths_set]
        if not new_paths:
            return
        self.lbl_status.config(text=f"Đang đọc {len(new_paths)} file XML...")
        self.root.update_idletasks()
        added = 0
        for p in new_paths:
            d = parse_tax_xml(p)
            if d.error is None and d.tax_type != "TNCN":
                continue  # Chỉ nạp tờ khai Quyết toán TNCN
            self.file_paths_set.add(p)
            self.declarations.append(d)
            added += 1
        self.refresh_table()
        self.lbl_status.config(text=f"Đã nạp thêm {added} file QTT TNCN.")

    def add_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Chọn nhiều file XML Tờ khai Quyết toán TNCN (05/QTT-TNCN)",
            filetypes=[("XML files", "*.xml"), ("All files", "*.*")],
        )
        if paths:
            self._ingest_xml_paths(list(paths))

    def add_folder(self) -> None:
        folder = filedialog.askdirectory(title="Chọn thư mục chứa hồ sơ XML (tự động quét thư mục con & lọc tờ khai QTT TNCN)")
        if not folder:
            return
        found: List[str] = []
        for root_dir, _, files in os.walk(folder):
            for fn in files:
                if fn.lower().endswith(".xml"):
                    found.append(os.path.join(root_dir, fn))
        self._ingest_xml_paths(found)

    def remove_selected(self) -> None:
        selected = set(self.tree.selection())
        if not selected:
            return
        kept: List[TaxDeclaration] = []
        for d in self.declarations:
            if d.file_path in selected:
                self.file_paths_set.discard(d.file_path)
            else:
                kept.append(d)
        self.declarations = kept
        self.refresh_table()

    def clear_all(self) -> None:
        self.declarations.clear()
        self.file_paths_set.clear()
        self.refresh_table()

    def pick_tms_file(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Chọn 1 hoặc nhiều file Người phụ thuộc kết xuất từ hệ thống TMS (.xlsx hoặc .csv)",
            filetypes=[("Excel / CSV", "*.xlsx *.csv"), ("Excel Workbook", "*.xlsx"), ("CSV", "*.csv"), ("All files", "*.*")],
        )
        for p in paths:
            norm = os.path.normpath(p)
            if norm not in self.tms_files:
                self.tms_files.append(norm)
        if self.tms_files:
            self.var_tms_path.set("; ".join(os.path.basename(x) for x in self.tms_files) + f" ({len(self.tms_files)} file)")

    def clear_tms_files(self) -> None:
        self.tms_files.clear()
        self.var_tms_path.set("")

    def _add_ledger_paths(self, paths: List[str]) -> None:
        added = 0
        for p in paths:
            if not p.lower().endswith(".xlsx") or os.path.basename(p).startswith("~$"):
                continue
            norm = os.path.normpath(p)
            if norm not in self.ledger_files:
                self.ledger_files.append(norm)
                self.lst_ledgers.insert(tk.END, norm)
                added += 1
        self.lbl_ledgers.config(text=f"Tổng: {len(self.ledger_files)} file sổ chi tiết / bảng lương")
        if added:
            self.lbl_status.config(text=f"Đã thêm {added} file sổ chi tiết / bảng lương.")

    def add_ledger_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Chọn nhiều file Excel sổ chi tiết tài khoản & bảng tính lương/thuế nhiều năm",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if paths:
            self._add_ledger_paths(list(paths))

    def add_ledger_folder(self) -> None:
        folder = filedialog.askdirectory(title="Chọn thư mục chứa sổ chi tiết tài khoản & bảng lương nhiều năm (quét cả thư mục con)")
        if not folder:
            return
        found: List[str] = []
        for root_dir, _, files in os.walk(folder):
            for fn in files:
                if fn.lower().endswith(".xlsx") and not fn.startswith("~$"):
                    found.append(os.path.join(root_dir, fn))
        self._add_ledger_paths(found)

    def remove_selected_ledgers(self) -> None:
        for idx in reversed(self.lst_ledgers.curselection()):
            self.lst_ledgers.delete(idx)
            del self.ledger_files[idx]
        self.lbl_ledgers.config(text=f"Tổng: {len(self.ledger_files)} file sổ chi tiết / bảng lương")

    def clear_ledgers(self) -> None:
        self.ledger_files.clear()
        self.lst_ledgers.delete(0, tk.END)
        self.lbl_ledgers.config(text="Tổng: 0 file sổ chi tiết / bảng lương")

    def refresh_table(self) -> None:
        for item in self.tree.get_children():
            self.tree.delete(item)
        shown = 0
        for d in sorted(self.declarations, key=lambda x: (x.year or 0, int(x.so_lan or 0), x.file_name)):
            if not self._matches_year(d):
                continue
            shown += 1
            st = "Hợp lệ" if d.error is None else f"Lỗi: {d.error}"
            nl = d.ngay_lap.strftime("%d/%m/%Y") if d.ngay_lap else ""
            lan_txt = "Chính thức" if not d.so_lan or str(d.so_lan) == "0" else f"BS lần {d.so_lan}"
            self.tree.insert(
                "",
                tk.END,
                iid=d.file_path,
                values=(d.file_name, d.ma_tkhai, d.mst, d.ten_nnt, d.year or "", lan_txt, nl, st),
            )
        self.lbl_count.config(text=f"Hiển thị: {shown} / {len(self.declarations)} file TNCN")

    def run_export(self) -> None:
        valid = [d for d in self.declarations if d.error is None and self._matches_year(d)]
        if not valid:
            messagebox.showwarning("Chú ý", "Chưa có tờ khai QTT TNCN hợp lệ nào trong khoảng năm đã chọn.")
            return

        out_dir = filedialog.askdirectory(title="Chọn thư mục lưu các file QTT_TNCN_YYYY.xlsx")
        if not out_dir:
            return

        y_from, y_to = self._get_year_bounds()
        valid_tms_list = [p for p in self.tms_files if os.path.isfile(p)]
        ledgers = list(self.ledger_files)
        open_after = self.var_open_after.get()

        self.btn_export.config(state=tk.DISABLED)
        self.btn_reconcile_excel.config(state=tk.DISABLED)
        self.lbl_status.config(text="Đang xuất file QTT TNCN từng năm & đối chiếu trực tiếp trên 05-3...")
        self.root.update_idletasks()

        def worker() -> None:
            try:
                created, stats_by_year = export_yearly_tncn_with_audit(
                    self.declarations,
                    out_dir,
                    year_from=y_from,
                    year_to=y_to,
                    tms_npt_paths=valid_tms_list,
                    ledger_paths=ledgers,
                )

                def on_done() -> None:
                    self.btn_export.config(state=tk.NORMAL)
                    self.btn_reconcile_excel.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Xuất file & đối chiếu thành công!")
                    lines = ["Đã xuất thành công các file QTT TNCN chuẩn mẫu (TK, 05-1, 05-2, 05-3):"] + created
                    if stats_by_year:
                        lines.append("\nKẾT QUẢ ĐỐI CHIẾU NGƯỜI PHỤ THUỘC TRỰC TIẾP TRÊN 05-3:")
                        for yr, st in sorted(stats_by_year.items()):
                            lines.append(
                                f"• Năm {yr}: Tổng {st['total_rows']} dòng | Khớp: {st['matched']} | "
                                f"Cảnh báo độ tuổi: {st.get('age_warnings', 0)} | Trùng NPT: {st.get('duplicates', 0)} | "
                                f"Loại trừ TMS: {st['total_diff_months']} tháng ({st['total_disallowed_amount']:,.0f} đ)"
                            )
                    messagebox.showinfo("Hoàn tất", "\n".join(lines))
                    if open_after:
                        try:
                            os.startfile(out_dir)
                        except Exception:
                            pass

                self.root.after(0, on_done)
            except Exception as exc:
                def on_err() -> None:
                    self.btn_export.config(state=tk.NORMAL)
                    self.btn_reconcile_excel.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Có lỗi xảy ra!")
                    messagebox.showerror("Lỗi", f"Không thể xuất file QTT TNCN:\n{exc}")

                self.root.after(0, on_err)

        threading.Thread(target=worker, daemon=True).start()

    def run_reconcile_existing(self) -> None:
        excel_path = filedialog.askopenfilename(
            title="Bước 1/2: Chọn file QTT_TNCN_YYYY.xlsx đã xuất sẵn (có sheet 05-3)",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not excel_path:
            return

        tms_path = self.tms_files[0] if (self.tms_files and os.path.isfile(self.tms_files[0])) else ""
        if not tms_path:
            tms_path = filedialog.askopenfilename(
                title="Bước 2/2: Chọn file Người phụ thuộc kết xuất từ TMS (.xlsx / .csv)",
                filetypes=[("Excel / CSV", "*.xlsx *.csv"), ("All files", "*.*")],
            )
            if not tms_path:
                return
            norm_t = os.path.normpath(tms_path)
            if norm_t not in self.tms_files:
                self.tms_files.append(norm_t)
            self.var_tms_path.set("; ".join(os.path.basename(x) for x in self.tms_files) + f" ({len(self.tms_files)} file)")

        open_after = self.var_open_after.get()
        self.btn_export.config(state=tk.DISABLED)
        self.btn_reconcile_excel.config(state=tk.DISABLED)
        self.lbl_status.config(text="Đang đối chiếu NPT trực tiếp trên sheet 05-3...")
        self.root.update_idletasks()

        def worker() -> None:
            try:
                st = reconcile_existing_tncn_excel(excel_path, tms_path)

                def on_done() -> None:
                    self.btn_export.config(state=tk.NORMAL)
                    self.btn_reconcile_excel.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Đối chiếu NPT thành công!")
                    msg = (
                        f"Đã đối chiếu trực tiếp Người phụ thuộc trên sheet '05-3' của file:\n{excel_path}\n\n"
                        f"• Tổng số dòng 05-3: {st['total_rows']}\n"
                        f"• Số dòng KHỚP: {st['matched']}\n"
                        f"• Số dòng CHÊNH LỆCH / CẦN KIỂM TRA: {st['issues']}\n"
                        f"• Tổng số tháng loại trừ: {st['total_diff_months']} tháng\n"
                        f"• Tổng tiền giảm trừ loại trừ: {st['total_disallowed_amount']:,.0f} đ"
                    )
                    messagebox.showinfo("Hoàn tất đối chiếu NPT", msg)
                    if open_after:
                        try:
                            os.startfile(excel_path)
                        except Exception:
                            pass

                self.root.after(0, on_done)
            except Exception as exc:
                def on_err() -> None:
                    self.btn_export.config(state=tk.NORMAL)
                    self.btn_reconcile_excel.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Lỗi khi đối chiếu NPT!")
                    messagebox.showerror("Lỗi", f"Không thể đối chiếu NPT:\n{exc}")

                self.root.after(0, on_err)

        threading.Thread(target=worker, daemon=True).start()


def main() -> None:
    root = tk.Tk()
    PitAuditProApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
