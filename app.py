# -*- coding: utf-8 -*-
"""
Giao diện Desktop cho phần mềm chuyên dụng Đối chiếu & Tính Thuế TNCN (pit_audit_pro v1.0).
Đầu vào hỗ trợ:
1. Tờ khai Quyết toán TNCN: File XML (05/QTT-TNCN) hoặc file Excel (QTT_TNCN_YYYY.xlsx)
2. File Người phụ thuộc kết xuất từ hệ thống TMS (.xlsx / .csv) - Tùy chọn
3. Các file Sổ chi tiết tài khoản (TK 334, 353, 641, 642, 622, 627, 154, 338, 3335...) - Tùy chọn
"""

from __future__ import annotations

import os
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from typing import List

from ledger_payroll_scanner import LedgerAuditSummary, scan_accounting_ledgers
from npt_tms_engine import audit_dependents, parse_tms_npt_file
from pit_excel_exporter import export_pit_audit_workbook
from pit_tax_rules import get_policy_for_year
from xml_05qtt_parser import load_pit_declaration


APP_VERSION = "v1.0"


class PitAuditProApp:
    def __init__(self, root: tk.Tk) -> None:
        self.root = root
        self.root.title(f"PIT Audit Pro — Phần mềm Đối chiếu & Tính Thuế TNCN Chuyên sâu ({APP_VERSION})")
        self.root.geometry("980x620")
        self.root.minsize(880, 540)

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

    def _build_ui(self) -> None:
        top = ttk.Frame(self.root, padding=(14, 12, 14, 6))
        top.pack(fill=tk.X)
        ttk.Label(
            top,
            text=f"PHẦN MỀM ĐỐI CHIẾU & TÍNH THUẾ TNCN CHUYÊN SÂU ({APP_VERSION})",
            style="Header.TLabel",
        ).pack(side=tk.LEFT)

        # 1. File tờ khai 05/QTT-TNCN (XML hoặc Excel)
        f1 = ttk.LabelFrame(self.root, text="1. Tờ khai Quyết toán Thuế TNCN (File XML 05/QTT-TNCN hoặc Excel QTT_TNCN_YYYY.xlsx)", padding=10)
        f1.pack(fill=tk.X, padx=14, pady=5)
        self.var_decl_path = tk.StringVar()
        ttk.Entry(f1, textvariable=self.var_decl_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(f1, text="Chọn file 05/QTT-TNCN...", command=self.pick_decl_file).pack(side=tk.LEFT)

        # 2. File NPT hệ thống TMS (tùy chọn)
        f2 = ttk.LabelFrame(self.root, text="2. File Người phụ thuộc kết xuất từ TMS (.xlsx / .csv) — Tùy chọn đối chiếu 05-3 & tính lại 05-1", padding=10)
        f2.pack(fill=tk.X, padx=14, pady=5)
        self.var_tms_path = tk.StringVar()
        ttk.Entry(f2, textvariable=self.var_tms_path).pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 8))
        ttk.Button(f2, text="Chọn file NPT TMS...", command=self.pick_tms_file).pack(side=tk.LEFT, padx=(0, 6))
        ttk.Button(f2, text="Bỏ chọn", command=lambda: self.var_tms_path.set("")).pack(side=tk.LEFT)

        # 3. Các file sổ chi tiết tài khoản (TK 334, 353, 641, 642, 622, 627, 154, 338, 3335...)
        f3 = ttk.LabelFrame(self.root, text="3. Sổ chi tiết tài khoản kế toán (TK 334, 353, 641, 642, 622, 627, 338, 3335...) — Tùy chọn quét rủi ro chi phí", padding=10)
        f3.pack(fill=tk.BOTH, expand=True, padx=14, pady=5)

        btn_row = ttk.Frame(f3)
        btn_row.pack(fill=tk.X, pady=(0, 6))
        ttk.Button(btn_row, text="Thêm file sổ chi tiết (.xlsx)...", command=self.add_ledger_files).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="Xóa file đang chọn", command=self.remove_selected_ledger).pack(side=tk.LEFT, padx=(0, 8))
        ttk.Button(btn_row, text="Xóa tất cả", command=self.clear_ledgers).pack(side=tk.LEFT)
        self.lbl_ledger_cnt = ttk.Label(btn_row, text="Đã chọn: 0 file sổ chi tiết")
        self.lbl_ledger_cnt.pack(side=tk.RIGHT)

        self.lst_ledgers = tk.Listbox(f3, height=6, selectmode=tk.EXTENDED, font=("Segoe UI", 9))
        self.lst_ledgers.pack(fill=tk.BOTH, expand=True)

        # Khung nhật ký kết quả
        log_frame = ttk.LabelFrame(self.root, text="4. Nhật ký phân tích & Tóm tắt chính sách áp dụng", padding=8)
        log_frame.pack(fill=tk.X, padx=14, pady=5)
        self.txt_log = tk.Text(log_frame, height=6, font=("Consolas", 9), state=tk.DISABLED)
        self.txt_log.pack(fill=tk.X)

        # Thanh đáy
        bottom = ttk.Frame(self.root, padding=(14, 8, 14, 12))
        bottom.pack(fill=tk.X)
        self.var_open_after = tk.BooleanVar(value=True)
        ttk.Checkbutton(bottom, text="Mở file Excel kết quả sau khi chạy", variable=self.var_open_after).pack(side=tk.LEFT)

        self.lbl_status = ttk.Label(bottom, text="Sẵn sàng", font=("Segoe UI", 9, "italic"))
        self.lbl_status.pack(side=tk.LEFT, padx=16)

        self.btn_run = ttk.Button(
            bottom,
            text="CHẠY ĐỐI CHIẾU & XUẤT BÁO CÁO TNCN",
            style="Primary.TButton",
            command=self.run_audit,
        )
        self.btn_run.pack(side=tk.RIGHT)

    def _log(self, msg: str) -> None:
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.insert(tk.END, msg + "\n")
        self.txt_log.see(tk.END)
        self.txt_log.config(state=tk.DISABLED)

    def pick_decl_file(self) -> None:
        p = filedialog.askopenfilename(
            title="Chọn file XML 05/QTT-TNCN hoặc file Excel QTT_TNCN_YYYY.xlsx",
            filetypes=[("XML / Excel", "*.xml *.xlsx"), ("XML", "*.xml"), ("Excel", "*.xlsx")],
        )
        if p:
            self.var_decl_path.set(os.path.normpath(p))

    def pick_tms_file(self) -> None:
        p = filedialog.askopenfilename(
            title="Chọn file dữ liệu Người phụ thuộc kết xuất từ TMS (.xlsx / .csv)",
            filetypes=[("Excel / CSV", "*.xlsx *.csv"), ("All files", "*.*")],
        )
        if p:
            self.var_tms_path.set(os.path.normpath(p))

    def add_ledger_files(self) -> None:
        paths = filedialog.askopenfilenames(
            title="Chọn các file Excel sổ chi tiết tài khoản (TK 334, 353, 641, 642, 622, 627, 338, 3335...)",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        for p in paths:
            norm = os.path.normpath(p)
            if norm not in self.ledger_files:
                self.ledger_files.append(norm)
                self.lst_ledgers.insert(tk.END, norm)
        self.lbl_ledger_cnt.config(text=f"Đã chọn: {len(self.ledger_files)} file sổ chi tiết")

    def remove_selected_ledger(self) -> None:
        for idx in reversed(self.lst_ledgers.curselection()):
            self.lst_ledgers.delete(idx)
            del self.ledger_files[idx]
        self.lbl_ledger_cnt.config(text=f"Đã chọn: {len(self.ledger_files)} file sổ chi tiết")

    def clear_ledgers(self) -> None:
        self.lst_ledgers.delete(0, tk.END)
        self.ledger_files.clear()
        self.lbl_ledger_cnt.config(text="Đã chọn: 0 file sổ chi tiết")

    def run_audit(self) -> None:
        decl_path = self.var_decl_path.get().strip()
        if not decl_path or not os.path.isfile(decl_path):
            messagebox.showwarning("Thiếu dữ liệu", "Vui lòng chọn file Tờ khai Quyết toán TNCN (XML hoặc Excel) trước.")
            return

        out_path = filedialog.asksaveasfilename(
            title="Lưu báo cáo đối chiếu & tính lại Thuế TNCN",
            defaultextension=".xlsx",
            initialfile="BaoCao_KiemTra_TNCN.xlsx",
            filetypes=[("Excel Workbook", "*.xlsx")],
        )
        if not out_path:
            return

        tms_path = self.var_tms_path.get().strip()
        ledgers = list(self.ledger_files)
        open_after = self.var_open_after.get()

        self.btn_run.config(state=tk.DISABLED)
        self.lbl_status.config(text="Đang xử lý đối chiếu & tính lại thuế TNCN...")
        self.txt_log.config(state=tk.NORMAL)
        self.txt_log.delete("1.0", tk.END)
        self.txt_log.config(state=tk.DISABLED)

        def worker() -> None:
            try:
                decl = load_pit_declaration(decl_path)
                policy = get_policy_for_year(decl.year)
                self.root.after(0, lambda: self._log(
                    f"[1/4] Đã đọc tờ khai năm {decl.year}: {len(decl.rows_051)} dòng 05-1, {len(decl.rows_052)} dòng 05-2, {len(decl.rows_053)} dòng 05-3."
                ))
                self.root.after(0, lambda: self._log(
                    f"      Chính sách áp dụng: {policy.legal_basis} | Bản thân: {policy.self_deduction_monthly:,}đ/th | NPT: {policy.dependent_deduction_monthly:,}đ/th | Biểu: {'5 bậc' if decl.year >= 2026 else '7 bậc'}"
                ))

                tms_rows = parse_tms_npt_file(tms_path) if (tms_path and os.path.isfile(tms_path)) else []
                npt_res, npt_per_nnt = audit_dependents(decl.rows_053, tms_rows)
                disallowed_m = sum(x.disallowed_months for x in npt_res)
                self.root.after(0, lambda: self._log(
                    f"[2/4] Đối chiếu NPT (05-3 vs TMS): {len(tms_rows)} dòng TMS | Tổng số tháng NPT bị loại trừ: {disallowed_m} tháng."
                ))

                ledger_sum = scan_accounting_ledgers(ledgers) if ledgers else LedgerAuditSummary()
                self.root.after(0, lambda: self._log(
                    f"[3/4] Quét sổ kế toán ({len(ledger_sum.files_scanned)} file): Phát hiện {len(ledger_sum.flagged_entries)} bút toán chi phí nghi vấn chịu thuế TNCN."
                ))

                export_pit_audit_workbook(
                    out_path, decl, npt_res, npt_per_nnt, tms_rows, ledger_summary=ledger_sum
                )

                def on_done() -> None:
                    self._log(f"[4/4] Đã xuất hoàn tất báo cáo 7 sheet: {out_path}")
                    self.btn_run.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Hoàn tất!")
                    messagebox.showinfo("Hoàn tất", f"Đã xuất thành công báo cáo đối chiếu & tính lại Thuế TNCN:\n{out_path}")
                    if open_after:
                        try:
                            os.startfile(out_path)
                        except Exception:
                            pass

                self.root.after(0, on_done)
            except Exception as exc:
                def on_err() -> None:
                    self.btn_run.config(state=tk.NORMAL)
                    self.lbl_status.config(text="Có lỗi xảy ra!")
                    self._log(f"[LỖI] {exc}")
                    messagebox.showerror("Lỗi", f"Không thể hoàn tất đối chiếu:\n{exc}")

                self.root.after(0, on_err)

        threading.Thread(target=worker, daemon=True).start()


def main() -> None:
    root = tk.Tk()
    PitAuditProApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
