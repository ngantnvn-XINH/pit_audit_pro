# -*- coding: utf-8 -*-
"""
Luồng xử lý chính của PIT Audit Pro (v1.1):
1. Nhận danh sách nhiều file XML tờ khai QTT TNCN nhiều năm (có cả các lần bổ sung).
2. Lọc theo khoảng năm [Từ năm .. Đến năm], tự động chọn tờ khai bổ sung cuối cùng của từng năm
   và kế thừa phụ lục 05-3 từ lần trước (dùng đúng chuẩn `filter_latest_per_year` của tax_xml_to_excel).
3. Xuất mỗi năm ra 1 file `QTT_TNCN_YYYY.xlsx` chuẩn 100% mẫu `TK`, `05-1`, `05-2`, `05-3`.
4. Đối chiếu Người phụ thuộc (TMS) TRỰC TIẾP trên các cột bên phải (L..S) của sheet `05-3`,
   kèm sheet tổng hợp lỗi `DoiChieu_NPT` và sheet dữ liệu sạch `DS_NPT_TMS`.
5. (Tùy chọn) Quét thêm sổ chi tiết tài khoản chi phí (TK 334, 353, 641, 642, 622, 627, 154...)
   và gắn sheet `QuetRuiRo_ChiPhi` vào file kết quả.
"""

from __future__ import annotations

import os
import re
from typing import Any, Dict, List, Optional, Sequence, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from excel_exporter import filter_latest_per_year, write_single_year_tncn_file
from ledger_payroll_scanner import FlaggedCostEntry, LedgerAuditSummary, scan_accounting_ledgers_by_year
from npt_tms_engine import TmsNptRow, audit_dependents, parse_tms_npt_file, parse_ym
from pit_tax_rules import excel_formula_progressive_pit_annual, get_policy_for_year
from tax_parser import TaxDeclaration
from tms_extra_engine import (
    MultiIncomeRecord,
    NptDirectoryRecord,
    NptIncomeRecord,
    check_npt_hkd_in_period,
    check_npt_income_threshold,
    evaluate_multi_income_for_person,
    parse_multi_income_files,
    parse_npt_directory_files,
    parse_npt_income_files,
)
from xml_05qtt_parser import Row053, clean_mst


SHEET_NPT_SUMMARY = "DoiChieu_NPT"
SHEET_NPT_SYSTEM = "DS_NPT_TMS"
SHEET_COST_RISK = "QuetRuiRo_ChiPhi"


def _clean_text(val: Any) -> str:
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val).replace("\xa0", " ").replace("\u200b", "")).strip()


def _clean_id(val: Any) -> str:
    s = _clean_text(val)
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s.replace(" ", "").upper()


def _write_tms_sheet(wb: openpyxl.Workbook, tms_rows: List[TmsNptRow]) -> None:
    if SHEET_NPT_SYSTEM in wb.sheetnames:
        del wb[SHEET_NPT_SYSTEM]
    ws = wb.create_sheet(title=SHEET_NPT_SYSTEM)
    ws.views.sheetView[0].showGridLines = True

    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_reg = Font(name="Calibri", size=9)
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    thin = Side(style="thin", color="BFBFBF")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center")

    headers = [
        "STT", "Ngày tạo", "Ngày TĐ", "MST NNT", "Tên NNT", "MST NPT", "Tên NPT",
        "Ngày sinh", "CMND/CCCD", "Hộ chiếu", "Quan hệ", "Từ tháng (TMS)", "Đến tháng (TMS)",
        "Trạng thái MST NPT", "Mã TT", "Mẫu BK/TK", "GKS/Số",
    ]
    ws.row_dimensions[1].height = 24
    for ci, h in enumerate(headers, start=1):
        c = ws.cell(1, ci, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    for idx, rec in enumerate(tms_rows, start=1):
        r = idx + 1
        tf = f"{rec.tu_thang[1]:02d}/{rec.tu_thang[0]}" if rec.tu_thang else ""
        tt = f"{rec.den_thang[1]:02d}/{rec.den_thang[0]}" if rec.den_thang else ""
        vals = [
            idx, rec.ngay_tao, rec.ngay_td, rec.mst_nnt, rec.ten_nnt, rec.mst_npt, rec.ten_npt,
            rec.ngay_sinh, rec.cmnd, rec.ho_chieu, rec.quan_he, tf, tt, rec.trang_thai, rec.ma_tt, rec.mau_tk, rec.gks_so,
        ]
        for ci, v in enumerate(vals, start=1):
            c = ws.cell(r, ci, v)
            c.font = f_reg
            c.border = b_thin
            if ci in (4, 6, 9, 10, 15, 17):
                c.number_format = "@"
                c.alignment = al_l
            elif ci in (1, 2, 3, 8, 11, 12, 13, 16):
                c.alignment = al_c
            else:
                c.alignment = al_l

    for ci, w in enumerate([6, 12, 12, 14, 24, 14, 24, 12, 15, 13, 10, 12, 12, 32, 8, 11, 13], start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A2"


def reconcile_053_inplace(
    wb: openpyxl.Workbook,
    tms_rows: List[TmsNptRow],
    tax_year: Optional[int] = None,
) -> Dict[str, Any]:
    """
    Đối chiếu Người phụ thuộc TRỰC TIẾP trên các cột L..S của sheet '05-3' trong workbook `wb`,
    đồng thời tạo sheet tổng hợp 'DoiChieu_NPT' và bản sạch TMS 'DS_NPT_TMS'.
    """
    if "05-3" not in wb.sheetnames:
        raise ValueError("Workbook không có sheet '05-3' để đối chiếu Người phụ thuộc.")

    if tms_rows:
        _write_tms_sheet(wb, tms_rows)
    ws_053 = wb["05-3"]

    # Đọc toàn bộ dòng dữ liệu từ sheet '05-3' (giữ đúng số dòng gốc)
    row_indices: List[int] = []
    parsed_053: List[Row053] = []
    detected_year = tax_year

    for r in range(4, ws_053.max_row + 1):
        ten_nnt = _clean_text(ws_053.cell(r, 2).value)
        raw_mst_nnt = ws_053.cell(r, 3).value
        mst_nnt = clean_mst(raw_mst_nnt)
        ten_npt = _clean_text(ws_053.cell(r, 4).value)
        if not mst_nnt and not ten_npt:
            continue
        if "TỔNG CỘNG" in (str(ws_053.cell(r, 1).value or "") + " " + ten_nnt).upper():
            continue
        raw_mst_npt = ws_053.cell(r, 6).value
        mst_npt = clean_mst(raw_mst_npt)

        # Tự động cập nhật lại MST trên sheet 05-3 nếu bị thiếu số 0 ở đầu (định dạng Text @)
        if mst_nnt and str(raw_mst_nnt or "").strip() != mst_nnt:
            c3 = ws_053.cell(r, 3, mst_nnt)
            c3.number_format = "@"
        if mst_npt and str(raw_mst_npt or "").strip() != mst_npt:
            c6 = ws_053.cell(r, 6, mst_npt)
            c6.number_format = "@"

        tu_t = _clean_text(ws_053.cell(r, 10).value)
        den_t = _clean_text(ws_053.cell(r, 11).value)
        if detected_year is None:
            ym = parse_ym(tu_t) or parse_ym(den_t)
            if ym:
                detected_year = ym[0]
        row_indices.append(r)
        parsed_053.append(
            Row053(
                stt=len(parsed_053) + 1,
                ho_ten_nnt=ten_nnt,
                mst_nnt=mst_nnt,
                ho_ten_npt=ten_npt,
                ngay_sinh=_clean_text(ws_053.cell(r, 5).value),
                mst_npt=mst_npt,
                loai_gt=_clean_text(ws_053.cell(r, 7).value),
                so_gt=_clean_id(ws_053.cell(r, 8).value),
                quan_he=_clean_text(ws_053.cell(r, 9).value),
                tu_thang=tu_t,
                den_thang=den_t,
            )
        )

    policy = get_policy_for_year(detected_year)
    results, _ = audit_dependents(parsed_053, tms_rows, tax_year=detected_year)

    # ===== GHI TRỰC TIẾP 11 CỘT ĐỐI CHIẾU (L..V) NGAY TRÊN SHEET '05-3' =====
    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_code = Font(name="Calibri", size=9, italic=True, color="1F4E78")
    f_reg = Font(name="Calibri", size=9)
    f_bold = Font(name="Calibri", size=9, bold=True)

    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_code = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_ok = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_err = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    thin = Side(style="thin", color="000000")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")
    NUM = "#,##0"

    if "L1:V1" not in [str(m) for m in ws_053.merged_cells.ranges]:
        ws_053.merge_cells("L1:V1")
    ws_053["L1"] = "ĐỐI CHIẾU NGƯỜI PHỤ THUỘC (ĐỘ TUỔI • TRÙNG ĐĂNG KÝ • HỆ THỐNG TMS)"
    for cc in range(12, 23):
        c1 = ws_053.cell(1, cc)
        c1.font = f_hdr
        c1.fill = fill_hdr
        c1.alignment = al_c
        c1.border = b_thin

    col_defs = [
        (12, "Số tháng\n05-3", "[DC1]"),
        (13, "Tuổi\nNPT", "[DC2]"),
        (14, "Kiểm tra độ tuổi\n(TT 111 & BLLĐ)", "[DC3]"),
        (15, "Kiểm tra trùng NPT\n(05-3 & TMS)", "[DC4]"),
        (16, "Từ tháng\n(TMS)", "[DC5]"),
        (17, "Đến tháng\n(TMS)", "[DC6]"),
        (18, "Trạng thái MST NPT\n(TMS)", "[DC7]"),
        (19, "Số tháng\nhợp lệ (TMS)", "[DC8]"),
        (20, "Số tháng\nchênh lệch", "[DC9]"),
        (21, "Tiền giảm trừ\nloại trừ (đồng)", "[DC10]"),
        (22, "Kết quả đối chiếu\ntổng hợp", "[DC11]"),
    ]
    for cc, title, code in col_defs:
        c2 = ws_053.cell(2, cc, title)
        c2.font = f_hdr
        c2.fill = fill_hdr
        c2.alignment = al_c
        c2.border = b_thin

        c3 = ws_053.cell(3, cc, code)
        c3.font = f_code
        c3.fill = fill_code
        c3.alignment = al_c
        c3.border = b_thin

    matched_cnt = 0
    issue_cnt = 0
    age_warn_cnt = 0
    dup_cnt = 0
    total_diff_months = 0
    issue_items: List[Tuple[int, Any]] = []

    for r_idx, res in zip(row_indices, results):
        ws_053.cell(r_idx, 12, res.decl_months).alignment = al_c
        ws_053.cell(r_idx, 13, res.age if res.age is not None else "").alignment = al_c
        ws_053.cell(r_idx, 14, res.age_status).alignment = al_l
        ws_053.cell(r_idx, 15, res.dup_status).alignment = al_l
        ws_053.cell(r_idx, 16, res.tms_from).alignment = al_c
        ws_053.cell(r_idx, 17, res.tms_to).alignment = al_c
        ws_053.cell(r_idx, 18, res.tms_status).alignment = al_l
        ws_053.cell(r_idx, 19, res.valid_months).alignment = al_c
        ws_053.cell(r_idx, 20, f"=MAX(0,L{r_idx}-S{r_idx})").alignment = al_c
        ws_053.cell(r_idx, 21, f"=T{r_idx}*'{SHEET_NPT_SUMMARY}'!$C$4").alignment = al_r
        ws_053.cell(r_idx, 21).number_format = NUM
        ws_053.cell(r_idx, 22, res.status).alignment = al_l

        if not res.is_age_ok:
            age_warn_cnt += 1
        if res.is_dup:
            dup_cnt += 1

        if res.is_ok:
            matched_cnt += 1
            r_fill = fill_ok
        else:
            issue_cnt += 1
            r_fill = fill_err if (res.disallowed_months > 0 or res.is_dup) else fill_warn
            issue_items.append((r_idx, res))
        total_diff_months += res.disallowed_months

        for cc in range(12, 23):
            cell = ws_053.cell(r_idx, cc)
            cell.font = f_bold if cc in (15, 22) else f_reg
            cell.border = b_thin
            if cc == 14:
                cell.fill = fill_ok if res.is_age_ok else fill_warn
            elif cc == 15:
                cell.fill = fill_err if res.is_dup else fill_ok
            elif cc in (20, 21, 22):
                cell.fill = r_fill

    for cc, w in {12: 10, 13: 8, 14: 30, 15: 28, 16: 11, 17: 12, 18: 26, 19: 11, 20: 11, 21: 17, 22: 36}.items():
        ws_053.column_dimensions[get_column_letter(cc)].width = w

    # ===== TẠO SHEET TỔNG HỢP 'DoiChieu_NPT' =====
    if SHEET_NPT_SUMMARY in wb.sheetnames:
        del wb[SHEET_NPT_SUMMARY]
    ws_dc = wb.create_sheet(title=SHEET_NPT_SUMMARY)
    ws_dc.views.sheetView[0].showGridLines = True

    f_title = Font(name="Times New Roman", size=13, bold=True, color="1F4E78")
    f_sub = Font(name="Times New Roman", size=10, italic=True, color="333333")
    f_sec = Font(name="Times New Roman", size=11, bold=True, color="1F4E78")
    f_tb = Font(name="Times New Roman", size=10, bold=True)
    f_tr = Font(name="Times New Roman", size=10)
    fill_ctrl = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")

    ws_dc.merge_cells("A1:P1")
    ws_dc["A1"] = f"BẢNG TỔNG HỢP ĐỐI CHIẾU NGƯỜI PHỤ THUỘC (ĐỘ TUỔI • TRÙNG NPT • TMS) — NĂM {policy.year}"
    ws_dc["A1"].font = f_title
    ws_dc["A1"].alignment = al_c
    ws_dc.row_dimensions[1].height = 24

    ws_dc.merge_cells("A2:P2")
    ws_dc["A2"] = f"(Căn cứ: {policy.legal_basis} | Mức giảm trừ NPT mặc định: {policy.dependent_deduction_monthly:,} đ/tháng)"
    ws_dc["A2"].font = f_sub
    ws_dc["A2"].alignment = al_c
    ws_dc.row_dimensions[2].height = 18

    last_r = max(row_indices) if row_indices else 4
    kpi_rows = [
        (4, "Mức giảm trừ gia cảnh NPT (đồng/tháng — sửa ô vàng để cập nhật toàn bộ):", policy.dependent_deduction_monthly, True),
        (5, "Tổng số dòng NPT kê khai trên bảng kê 05-3:", f"=COUNTA('05-3'!V4:V{last_r})", False),
        (6, "Số dòng KHỚP hoàn toàn (Độ tuổi • Không trùng • TMS):", f'=COUNTIF(\'05-3\'!V4:V{last_r},"KHỚP")', False),
        (7, "Số dòng PHÁT HIỆN CHÊNH LỆCH / CẢNH BÁO:", f'=COUNTIF(\'05-3\'!V4:V{last_r},"<>KHỚP")-COUNTBLANK(\'05-3\'!V4:V{last_r})', False),
        (8, "Trong đó — Số dòng CẢNH BÁO ĐỘ TUỔI NPT:", age_warn_cnt, False),
        (9, "Trong đó — Số dòng PHÁT HIỆN TRÙNG ĐĂNG KÝ NPT:", dup_cnt, False),
        (10, "Tổng số tháng giảm trừ chênh lệch (theo TMS):", f"=SUM('05-3'!T4:T{last_r})", False),
        (11, "Tổng số tiền giảm trừ NPT không đủ điều kiện (đồng):", f"=SUM('05-3'!U4:U{last_r})", False),
    ]
    for r_k, lbl, val, is_inp in kpi_rows:
        ws_dc.row_dimensions[r_k].height = 20
        ws_dc.merge_cells(start_row=r_k, start_column=1, end_row=r_k, end_column=2)
        ws_dc.cell(r_k, 1, lbl).font = f_tb if (is_inp or r_k == 11) else f_tr
        ws_dc.cell(r_k, 1).alignment = al_l
        for cc in (1, 2):
            ws_dc.cell(r_k, cc).fill = fill_ctrl
            ws_dc.cell(r_k, cc).border = b_thin
        c_v = ws_dc.cell(r_k, 3, val)
        c_v.font = f_tb
        c_v.number_format = NUM
        c_v.alignment = al_r
        c_v.fill = fill_input if is_inp else (fill_err if r_k in (7, 9, 10, 11) else (fill_warn if r_k == 8 else fill_ok))
        c_v.border = b_thin

    ws_dc.merge_cells("A13:P13")
    ws_dc["A13"] = "DANH SÁCH CHI TIẾT CÁC TRƯỜNG HỢP CHÊNH LỆCH / CẢNH BÁO ĐỘ TUỔI / TRÙNG NPT"
    ws_dc["A13"].font = f_sec
    ws_dc.row_dimensions[13].height = 22

    dc_hdrs = [
        "STT", "Dòng 05-3", "Họ tên NNT", "MST NNT", "Họ tên NPT", "MST NPT",
        "Ngày sinh", "Tuổi", "Quan hệ", "Từ – Đến\n(05-3)", "Từ – Đến\n(TMS)",
        "Kiểm tra độ tuổi", "Kiểm tra trùng NPT", "Số tháng\nchênh lệch",
        "Tiền giảm trừ\nloại trừ (đồng)", "Kết quả đối chiếu tổng hợp",
    ]
    ws_dc.row_dimensions[14].height = 26
    for cc, h in enumerate(dc_hdrs, start=1):
        c = ws_dc.cell(14, cc, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    if not issue_items:
        ws_dc.merge_cells("A15:P15")
        ws_dc["A15"] = "Toàn bộ Người phụ thuộc trên bảng kê 05-3 đều hợp lệ về độ tuổi, không trùng lặp và khớp 100% với hệ thống TMS."
        ws_dc["A15"].font = f_tb
        ws_dc["A15"].alignment = al_c
        ws_dc["A15"].fill = fill_ok
        for cc in range(1, 17):
            ws_dc.cell(15, cc).border = b_thin
    else:
        for idx_i, (r_src, res) in enumerate(issue_items, start=1):
            rr = 14 + idx_i
            s = res.src
            ws_dc.row_dimensions[rr].height = 22
            vals = [
                idx_i, r_src, s.ho_ten_nnt, s.mst_nnt, s.ho_ten_npt, s.mst_npt,
                s.ngay_sinh, res.age if res.age is not None else "", s.quan_he,
                f"{s.tu_thang} - {s.den_thang}".strip(" -"),
                f"{res.tms_from} - {res.tms_to}".strip(" -"),
                res.age_status, res.dup_status,
                f"='05-3'!T{r_src}", f"='05-3'!U{r_src}", res.status,
            ]
            for cc, v in enumerate(vals, start=1):
                c = ws_dc.cell(rr, cc, v)
                c.font = f_tr
                c.border = b_thin
                if cc in (4, 6):
                    c.number_format = "@"
                    c.alignment = al_l
                elif cc in (1, 2, 7, 8, 9, 10, 11, 14):
                    c.alignment = al_c
                elif cc == 15:
                    c.number_format = NUM
                    c.alignment = al_r
                else:
                    c.alignment = al_l
                if cc == 12 and not res.is_age_ok:
                    c.fill = fill_warn
                elif cc == 13 and res.is_dup:
                    c.fill = fill_err
                elif cc in (14, 15, 16):
                    c.fill = fill_err if (res.disallowed_months > 0 or res.is_dup) else fill_warn

        r_tot = 14 + len(issue_items) + 1
        ws_dc.row_dimensions[r_tot].height = 22
        ws_dc.merge_cells(start_row=r_tot, start_column=1, end_row=r_tot, end_column=13)
        ws_dc.cell(r_tot, 1, "TỔNG CỘNG").alignment = al_c
        ws_dc.cell(r_tot, 14, f"=SUM(N15:N{r_tot - 1})").alignment = al_c
        ws_dc.cell(r_tot, 15, f"=SUM(O15:O{r_tot - 1})").alignment = al_r
        ws_dc.cell(r_tot, 15).number_format = NUM
        for cc in range(1, 17):
            c = ws_dc.cell(r_tot, cc)
            c.font = f_tb
            c.fill = fill_code
            c.border = b_thin

    for cc, w in enumerate([6, 9, 22, 13, 22, 13, 11, 7, 10, 16, 16, 28, 28, 11, 17, 36], start=1):
        ws_dc.column_dimensions[get_column_letter(cc)].width = w
    ws_dc.freeze_panes = "A15"

    return {
        "total_rows": len(parsed_053),
        "matched": matched_cnt,
        "issues": issue_cnt,
        "age_warnings": age_warn_cnt,
        "duplicates": dup_cnt,
        "total_diff_months": total_diff_months,
        "total_disallowed_amount": total_diff_months * policy.dependent_deduction_monthly,
    }


def _append_cost_risk_sheet(wb: openpyxl.Workbook, flagged: List[FlaggedCostEntry]) -> None:
    if SHEET_COST_RISK in wb.sheetnames:
        del wb[SHEET_COST_RISK]
    ws = wb.create_sheet(title=SHEET_COST_RISK)
    ws.views.sheetView[0].showGridLines = True

    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_reg = Font(name="Calibri", size=9)
    f_bold = Font(name="Calibri", size=9, bold=True)
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_err = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    thin = Side(style="thin", color="000000")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")

    hdrs = [
        "STT", "File nguồn", "Ngày CT", "Số CT", "TK Nợ", "TK Có", "Đối tượng",
        "Diễn giải chứng từ", "Số tiền (đồng)", "Nhóm rủi ro TNCN", "Mức độ", "Căn cứ pháp lý / Gợi ý kiểm tra",
    ]
    ws.row_dimensions[1].height = 26
    for ci, h in enumerate(hdrs, start=1):
        c = ws.cell(1, ci, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    for idx, fe in enumerate(flagged, start=1):
        r = idx + 1
        vals = [idx, fe.source_file, fe.ngay_ct, fe.so_ct, fe.tk_no, fe.tk_co, fe.doi_tuong, fe.dien_giai, round(fe.so_tien), fe.nhom_rui_ro, fe.muc_do, fe.can_cu_phap_ly]
        for ci, v in enumerate(vals, start=1):
            c = ws.cell(r, ci, v)
            c.font = f_bold if ci == 11 else f_reg
            c.border = b_thin
            if ci in (1, 3, 4, 5, 6, 11):
                c.alignment = al_c
            elif ci == 9:
                c.number_format = "#,##0"
                c.alignment = al_r
            else:
                c.alignment = al_l
            if ci == 11:
                c.fill = fill_err if fe.muc_do == "CAO" else fill_warn

    for ci, w in enumerate([6, 18, 12, 12, 9, 9, 20, 38, 16, 30, 11, 42], start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w
    ws.freeze_panes = "A2"


def _append_payroll_reconcile_sheet(wb: openpyxl.Workbook, yr: int, ls: LedgerAuditSummary) -> None:
    sname = "DoiChieu_QuyLuong_334"
    if sname in wb.sheetnames:
        del wb[sname]
    ws = wb.create_sheet(title=sname)
    ws.views.sheetView[0].showGridLines = True

    f_title = Font(name="Times New Roman", size=12, bold=True, color="1F4E78")
    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_reg = Font(name="Calibri", size=9)
    f_bold = Font(name="Calibri", size=9, bold=True)
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    thin = Side(style="thin", color="000000")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")

    ws.merge_cells("A1:F1")
    ws["A1"] = f"ĐỐI CHIẾU QUỸ LƯƠNG (TK 334), BẢO HIỂM (TK 338) & THUẾ TNCN (TK 3335) — NĂM {yr}"
    ws["A1"].font = f_title
    ws["A1"].alignment = al_c
    ws.row_dimensions[1].height = 24

    hdrs = ["STT", "Chỉ tiêu đối chiếu", "Số liệu trên Sổ kế toán", "Số liệu trên 05/QTT-TNCN (TK)", "Chênh lệch (Sổ KT - TK)", "Ghi chú nghiệp vụ"]
    ws.row_dimensions[3].height = 24
    for ci, h in enumerate(hdrs, start=1):
        c = ws.cell(3, ci, h)
        c.font = f_hdr
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    has_tk = "TK" in wb.sheetnames
    tnct_ref = "=TK!D41+TK!D43+TK!D45" if has_tk else 0
    thue_ref = "=TK!D47" if has_tk else 0

    rows_spec = [
        (1, "Phát sinh Có TK 334 (Lương, thưởng phải trả) vs Tổng TNCT trên TK", round(ls.tk334_credit_total), tnct_ref, "Chênh lệch thường do lương tháng 12 chi trả sang năm sau hoặc khoản không chịu thuế"),
        (2, "Phát sinh Nợ TK 334 (Thực chi trả NLĐ) vs Tổng TNCT trên TK", round(ls.tk334_debit_total), tnct_ref, "Thuế TNCN tính theo thời điểm thực chi trả (Khoản 2 Điều 8 TT 111/2013)"),
        (3, "Bảo hiểm bắt buộc khấu trừ lương (TK 3383/3384/3386)", round(ls.tk338_insurance_total), 0, "Đối chiếu tổng BHXH/BHYT/BHTN phần NLĐ đóng trên 05-1"),
        (4, "Thuế TNCN phải nộp phát sinh (Có TK 3335) vs Tổng thuế khấu trừ trên TK", round(ls.tk3335_credit_total), thue_ref, "Kiểm tra số thuế TNCN đã khấu trừ trên sổ sách khớp tờ khai QTT"),
        (5, "Tổng chi Quỹ khen thưởng, phúc lợi (Nợ TK 353)", round(ls.tk353_debit_total), 0, "Rà soát các khoản chi phúc lợi đích danh cá nhân có nguy cơ bỏ sót TNCT"),
    ]
    for r_i, (stt, lbl, acc_v, tk_v, note) in enumerate(rows_spec, start=4):
        ws.row_dimensions[r_i].height = 22
        ws.cell(r_i, 1, stt).alignment = al_c
        ws.cell(r_i, 2, lbl).alignment = al_l
        ws.cell(r_i, 3, acc_v).number_format = "#,##0"
        ws.cell(r_i, 3).alignment = al_r
        ws.cell(r_i, 4, tk_v).number_format = "#,##0"
        ws.cell(r_i, 4).alignment = al_r
        ws.cell(r_i, 5, f"=C{r_i}-D{r_i}").number_format = "#,##0"
        ws.cell(r_i, 5).alignment = al_r
        ws.cell(r_i, 6, note).alignment = al_l
        for ci in range(1, 7):
            c = ws.cell(r_i, ci)
            c.font = f_bold if ci == 5 else f_reg
            c.border = b_thin

    for ci, w in enumerate([6, 44, 20, 22, 20, 44], start=1):
        ws.column_dimensions[get_column_letter(ci)].width = w


def add_quick_test_formulas_051_052(
    wb: openpyxl.Workbook,
    tax_year: Optional[int] = None,
) -> Dict[str, int]:
    """
    Gắn trực tiếp khối công thức TEST NHANH cảnh báo rủi ro khấu trừ / tính sai thuế TNCN
    ở bên phải các cột gốc của sheet '05-1' (cột X..AD) và '05-2' (cột M..Q).
    - Giữ nguyên bản 100% toàn bộ các cột gốc của bảng kê (A..W trên 05-1; A..L trên 05-2).
    - Tự động áp dụng đúng mức giảm trừ gia cảnh & biểu lũy tiến theo năm (7 bậc <=2025; 5 bậc >=2026).
    """
    yr = tax_year or 2024
    policy = get_policy_for_year(yr)
    self_yr = policy.self_deduction_annual
    dep_mo = policy.dependent_deduction_monthly
    bracket_lbl = "5 bậc" if yr >= 2026 else "7 bậc"

    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_code = Font(name="Calibri", size=9, italic=True, color="1F4E78")
    f_reg = Font(name="Calibri", size=9)
    f_bold = Font(name="Calibri", size=9, bold=True)
    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_code = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_tot = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    thin = Side(style="thin", color="000000")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")
    NUM = "#,##0"

    counts = {"rows_051": 0, "rows_052": 0}

    # ===== 1. SHEET '05-1' (Cột gốc A..W; thêm Test nhanh X..AD = cột 24..30) =====
    if "05-1" in wb.sheetnames:
        ws1 = wb["05-1"]
        if "X1:AD1" not in [str(m) for m in ws1.merged_cells.ranges]:
            ws1.merge_cells("X1:AD1")
        ws1["X1"] = f"TEST NHANH RỦI RO TÍNH THUẾ & KHẤU TRỪ (05-1 — NĂM {yr}: {bracket_lbl}, GTBT {self_yr:,.0f})"
        for cc in range(24, 31):
            c1 = ws1.cell(1, cc)
            c1.font = f_hdr
            c1.fill = fill_hdr
            c1.alignment = al_c
            c1.border = b_thin

        defs_051 = [
            (24, "TN tính thuế\n(Test nhanh)", "[T1]", 16),
            (25, "Chênh lệch TNTT\n(Test - Khai)", "[T2]", 15),
            (26, f"Thuế TNCN phải nộp\n(Test {bracket_lbl})", "[T3]", 16),
            (27, "Chênh lệch Thuế\n(Test - Khai/KT)", "[T4]", 16),
            (28, "Thuế khấu trừ thiếu\n(=MAX(0, Test-R))", "[T5]", 16),
            (29, "Kiểm tra GTGC\n(M vs Bản thân+NPT)", "[T6]", 18),
            (30, "Cảnh báo rủi ro\n(05-1)", "[T7]", 34),
        ]
        for cc, title, code, w in defs_051:
            c2 = ws1.cell(2, cc, title)
            c2.font = f_hdr
            c2.fill = fill_hdr
            c2.alignment = al_c
            c2.border = b_thin
            c3 = ws1.cell(3, cc, code)
            c3.font = f_code
            c3.fill = fill_code
            c3.alignment = al_c
            c3.border = b_thin
            ws1.column_dimensions[get_column_letter(cc)].width = w

        data_rows_1: List[int] = []
        tot_row_1: Optional[int] = None
        for r in range(4, ws1.max_row + 1):
            a_txt = _clean_text(ws1.cell(r, 1).value).upper()
            b_txt = _clean_text(ws1.cell(r, 2).value)
            if "TỔNG CỘNG" in a_txt or "TỔNG CỘNG" in b_txt.upper():
                tot_row_1 = r
                continue
            if not b_txt and not _clean_text(ws1.cell(r, 3).value):
                continue
            data_rows_1.append(r)
            uq_expr = f'OR(UPPER(TRIM(F{r}))="[X]",UPPER(TRIM(F{r}))="X",UPPER(TRIM(F{r}))="TRUE",F{r}=1)'
            ws1.cell(r, 24, f"=MAX(0,H{r}-J{r}-K{r}-M{r}-N{r}-O{r}-P{r})").number_format = NUM
            ws1.cell(r, 25, f"=X{r}-Q{r}").number_format = NUM
            ws1.cell(r, 26, excel_formula_progressive_pit_annual(f"X{r}", year=yr)).number_format = NUM
            ws1.cell(r, 27, f"=IF({uq_expr},Z{r}-T{r},Z{r}-R{r})").number_format = NUM
            ws1.cell(r, 28, f"=MAX(0,Z{r}-R{r})").number_format = NUM
            ws1.cell(
                r,
                29,
                f'=IF(M{r}>({self_yr}+L{r}*12*{dep_mo})+1000,"VƯỢT MỨC GTGC",'
                f'IF(AND({uq_expr},M{r}<{self_yr}),"THIẾU GT BẢN THÂN","Hợp lệ"))',
            )
            ws1.cell(
                r,
                30,
                f'=IF(ABS(Y{r})>1000,"SAI TN TÍNH THUẾ",'
                f'IF(AND({uq_expr},ABS(Z{r}-T{r})>1000),"TÍNH SAI THUẾ (ỦY QUYỀN QT)",'
                f'IF(AC{r}<>"Hợp lệ",AC{r},'
                f'IF(AND(H{r}>{self_yr},R{r}=0,Z{r}>0),"RỦI RO: TN CAO CHƯA KHẤU TRỪ",'
                f'IF(Z{r}-R{r}>1000,"RỦI RO: KHẤU TRỪ THIẾU (KT KỲ LV)","Hợp lệ")))))',
            )
            for cc in range(24, 31):
                c = ws1.cell(r, cc)
                c.font = f_bold if cc == 30 else f_reg
                c.border = b_thin
                c.alignment = al_r if cc <= 28 else (al_c if cc == 29 else al_l)

        counts["rows_051"] = len(data_rows_1)
        if tot_row_1 and data_rows_1:
            r_first, r_last = data_rows_1[0], data_rows_1[-1]
            for cc in range(24, 29):
                col_l = get_column_letter(cc)
                c = ws1.cell(tot_row_1, cc, f"=SUM({col_l}{r_first}:{col_l}{r_last})")
                c.number_format = NUM
                c.alignment = al_r
                c.font = f_bold
                c.fill = fill_tot
                c.border = b_thin
            c_cnt = ws1.cell(tot_row_1, 30, f'=COUNTIF(AD{r_first}:AD{r_last},"<>Hợp lệ")&" dòng cảnh báo"')
            c_cnt.font = f_bold
            c_cnt.fill = fill_tot
            c_cnt.alignment = al_c
            c_cnt.border = b_thin
            ws1.cell(tot_row_1, 29).fill = fill_tot
            ws1.cell(tot_row_1, 29).border = b_thin

    # ===== 2. SHEET '05-2' (Cột gốc A..L; thêm Test nhanh M..Q = cột 13..17) =====
    if "05-2" in wb.sheetnames:
        ws2 = wb["05-2"]
        if "M1:Q1" not in [str(m) for m in ws2.merged_cells.ranges]:
            ws2.merge_cells("M1:Q1")
        ws2["M1"] = f"TEST NHANH RỦI RO KHẤU TRỪ 10% / 20% (05-2 — NĂM {yr})"
        for cc in range(13, 18):
            c1 = ws2.cell(1, cc)
            c1.font = f_hdr
            c1.fill = fill_hdr
            c1.alignment = al_c
            c1.border = b_thin

        defs_052 = [
            (13, "Thuế suất chuẩn\n(10% / 20%)", "[T1]", 14),
            (14, "Thuế phải KT\n(Test nhanh)", "[T2]", 16),
            (15, "Thuế KT thiếu\n(=MAX(0,Test-K))", "[T3]", 16),
            (16, "Tỷ lệ KT thực tế\n(K / G)", "[T4]", 14),
            (17, "Cảnh báo rủi ro\n(05-2)", "[T5]", 42),
        ]
        for cc, title, code, w in defs_052:
            c2 = ws2.cell(2, cc, title)
            c2.font = f_hdr
            c2.fill = fill_hdr
            c2.alignment = al_c
            c2.border = b_thin
            c3 = ws2.cell(3, cc, code)
            c3.font = f_code
            c3.fill = fill_code
            c3.alignment = al_c
            c3.border = b_thin
            ws2.column_dimensions[get_column_letter(cc)].width = w

        data_rows_2: List[int] = []
        tot_row_2: Optional[int] = None
        for r in range(4, ws2.max_row + 1):
            a_txt = _clean_text(ws2.cell(r, 1).value).upper()
            b_txt = _clean_text(ws2.cell(r, 2).value)
            if "TỔNG CỘNG" in a_txt or "TỔNG CỘNG" in b_txt.upper() or "CÁ NHÂN CƯ TRÚ" in a_txt or "CÁ NHÂN KHÔNG CƯ TRÚ" in a_txt:
                if "TỔNG CỘNG" in a_txt or "TỔNG CỘNG" in b_txt.upper():
                    tot_row_2 = r
                continue
            if not b_txt and not _clean_text(ws2.cell(r, 3).value):
                continue
            data_rows_2.append(r)
            nr_expr = f'OR(UPPER(TRIM(F{r}))="[X]",UPPER(TRIM(F{r}))="X",UPPER(TRIM(F{r}))="TRUE",F{r}=1)'
            ws2.cell(r, 13, f"=IF({nr_expr},0.2,0.1)").number_format = "0%"
            ws2.cell(r, 14, f"=ROUND(MAX(0,G{r}-I{r}-J{r})*M{r},0)").number_format = NUM
            ws2.cell(r, 15, f"=MAX(0,N{r}-K{r})").number_format = NUM
            ws2.cell(r, 16, f"=IF(G{r}>0,K{r}/G{r},0)").number_format = "0.0%"
            ws2.cell(
                r,
                17,
                f'=IF(AND(M{r}=0.2,O{r}>1000),"RỦI RO CAO: KHÔNG CƯ TRÚ KT THIẾU 20%",'
                f'IF(AND(M{r}=0.1,TRIM(C{r})="",G{r}>=2000000,O{r}>1000),"RỦI RO CAO: CHƯA CÓ MST KHÔNG ĐỦ ĐK CAM KẾT 08 (PHẢI KT 10%)",'
                f'IF(AND(M{r}=0.1,G{r}>{self_yr},O{r}>1000),"RỦI RO CAO: TN VƯỢT NGƯỠNG GTBT - KHÔNG ĐƯỢC CAM KẾT 08",'
                f'IF(AND(M{r}=0.1,G{r}>=2000000,O{r}>1000),"CẦN KT CAM KẾT 08/CK-TNCN (KHẤU TRỪ THIẾU 10%)",'
                f'IF(K{r}-N{r}>1000,"LƯU Ý: KHẤU TRỪ VƯỢT MỨC","Hợp lệ")))))',
            )
            for cc in range(13, 18):
                c = ws2.cell(r, cc)
                c.font = f_bold if cc == 17 else f_reg
                c.border = b_thin
                c.alignment = al_c if cc in (13, 16) else (al_r if cc in (14, 15) else al_l)

        counts["rows_052"] = len(data_rows_2)
        if tot_row_2 and data_rows_2:
            r_first, r_last = data_rows_2[0], data_rows_2[-1]
            for cc in (14, 15):
                col_l = get_column_letter(cc)
                c = ws2.cell(tot_row_2, cc, f"=SUM({col_l}{r_first}:{col_l}{r_last})")
                c.number_format = NUM
                c.alignment = al_r
                c.font = f_bold
                c.fill = fill_tot
                c.border = b_thin
            c_cnt = ws2.cell(tot_row_2, 17, f'=COUNTIF(Q{r_first}:Q{r_last},"<>Hợp lệ")&" dòng cảnh báo"')
            c_cnt.font = f_bold
            c_cnt.fill = fill_tot
            c_cnt.alignment = al_c
            c_cnt.border = b_thin
            for cc in (13, 16):
                ws2.cell(tot_row_2, cc).fill = fill_tot
                ws2.cell(tot_row_2, cc).border = b_thin

    return counts


def apply_extra_tms_checks(
    wb: openpyxl.Workbook,
    tax_year: Optional[int] = None,
    company_mst: str = "",
    multi_income_records: Optional[Sequence[MultiIncomeRecord]] = None,
    npt_dir_records: Optional[Sequence[NptDirectoryRecord]] = None,
    npt_income_records: Optional[Sequence[NptIncomeRecord]] = None,
) -> Dict[str, int]:
    """
    Đối chiếu 3 nguồn dữ liệu nâng cao từ TMS trực tiếp lên các sheet:
    - 05-1 (cột AE..AF): Cảnh báo cá nhân KHÔNG đủ điều kiện Ủy quyền QT (có thu nhập nhiều nơi).
    - 05-2 (cột R..S): Cảnh báo cá nhân thuộc diện PHẢI khấu trừ 10% (không đủ ĐK Cam kết 08).
    - 05-3 (cột W..X): Cảnh báo NPT là Hộ kinh doanh đang hoạt động & NPT có thu nhập vượt ngưỡng.
    """
    yr = tax_year or 2024
    policy = get_policy_for_year(yr)
    self_yr = policy.self_deduction_annual

    f_hdr = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_code = Font(name="Calibri", size=9, italic=True, color="1F4E78")
    f_reg = Font(name="Calibri", size=9)
    f_bold = Font(name="Calibri", size=9, bold=True)
    fill_hdr = PatternFill(start_color="833C0C", end_color="833C0C", fill_type="solid")
    fill_code = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_ok = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_err = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    thin = Side(style="thin", color="000000")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")
    NUM = "#,##0"

    summary = {"uq_warnings_051": 0, "kt10_warnings_052": 0, "hkd_warnings_053": 0, "npt_inc_warnings_053": 0}

    # 1. ĐỐI CHIẾU THU NHẬP NHIỀU NƠI LÊN 05-1 (AE..AF) & 05-2 (R..S)
    if multi_income_records:
        if "05-1" in wb.sheetnames:
            ws1 = wb["05-1"]
            if "AE1:AF1" not in [str(m) for m in ws1.merged_cells.ranges]:
                ws1.merge_cells("AE1:AF1")
            ws1["AE1"] = "ĐỐI CHIẾU THU NHẬP NHIỀU NƠI (TMS)"
            for cc in (31, 32):
                c1 = ws1.cell(1, cc)
                c1.font = f_hdr
                c1.fill = fill_hdr
                c1.alignment = al_c
                c1.border = b_thin
            for cc, title, code, w in (
                (31, "TNCT nơi khác\n(TMS)", "[T8]", 18),
                (32, "Cảnh báo Ủy quyền QT\n(Thu nhập nhiều nơi)", "[T9]", 42),
            ):
                c2 = ws1.cell(2, cc, title)
                c2.font = f_hdr
                c2.fill = fill_hdr
                c2.alignment = al_c
                c2.border = b_thin
                c3 = ws1.cell(3, cc, code)
                c3.font = f_code
                c3.fill = fill_code
                c3.alignment = al_c
                c3.border = b_thin
                ws1.column_dimensions[get_column_letter(cc)].width = w

            for r in range(4, ws1.max_row + 1):
                a_txt = _clean_text(ws1.cell(r, 1).value).upper()
                b_txt = _clean_text(ws1.cell(r, 2).value)
                if "TỔNG CỘNG" in a_txt or "TỔNG CỘNG" in b_txt.upper() or (not b_txt and not _clean_text(ws1.cell(r, 3).value)):
                    continue
                mst = clean_mst(ws1.cell(r, 3).value)
                cccd = _clean_id(ws1.cell(r, 5).value)
                f_val = _clean_text(ws1.cell(r, 6).value).upper()
                is_uq = f_val in ("[X]", "X", "TRUE", "1")
                others = evaluate_multi_income_for_person(mst, cccd, company_mst, yr, multi_income_records)
                other_tn = sum(x.tnct for x in others)
                other_cnt = len(others)
                has_other_051 = any(x.loai_bk == "05-1" for x in others)
                has_unwithheld_052 = any(x.loai_bk != "05-1" and x.tnct >= 2_000_000 and x.thue_kt + 1000 < round(x.tnct * 0.1) for x in others)

                c_tn = ws1.cell(r, 31, round(other_tn))
                c_tn.number_format = NUM
                c_tn.alignment = al_r
                c_tn.font = f_reg
                c_tn.border = b_thin

                is_err = False
                if is_uq and other_cnt > 0:
                    if has_other_051:
                        msg = f"KHÔNG ĐỦ ĐK UQ: Có TN 05-1 tại nơi khác ({other_cnt} nơi, {other_tn:,.0f} đ)"
                        is_err = True
                    elif other_tn > 120_000_000:
                        msg = f"KHÔNG ĐỦ ĐK UQ: TN vãng lai nơi khác {other_tn:,.0f} đ > 120tr/năm"
                        is_err = True
                    elif has_unwithheld_052:
                        msg = f"KHÔNG ĐỦ ĐK UQ: TN nơi khác ({other_tn:,.0f} đ) chưa khấu trừ đủ 10%"
                        is_err = True
                    else:
                        msg = f"Đủ ĐK UQ (Vãng lai nơi khác {other_tn:,.0f} đ ≤ 120tr đã KT 10%)"
                elif other_cnt > 0:
                    msg = f"Có TN tại {other_cnt} nơi khác ({other_tn:,.0f} đ - Không UQ)"
                else:
                    msg = "1 nguồn thu nhập"

                c_msg = ws1.cell(r, 32, msg)
                c_msg.alignment = al_l
                c_msg.font = f_bold if is_err else f_reg
                c_msg.border = b_thin
                if is_err:
                    summary["uq_warnings_051"] += 1
                    c_tn.fill = fill_err
                    c_msg.fill = fill_err
                elif other_cnt > 0:
                    c_msg.fill = fill_warn
                else:
                    c_msg.fill = fill_ok

        if "05-2" in wb.sheetnames:
            ws2 = wb["05-2"]
            if "R1:S1" not in [str(m) for m in ws2.merged_cells.ranges]:
                ws2.merge_cells("R1:S1")
            ws2["R1"] = "ĐỐI CHIẾU THU NHẬP NHIỀU NƠI (TMS)"
            for cc in (18, 19):
                c1 = ws2.cell(1, cc)
                c1.font = f_hdr
                c1.fill = fill_hdr
                c1.alignment = al_c
                c1.border = b_thin
            for cc, title, code, w in (
                (18, "TNCT nơi khác\n(TMS)", "[T6]", 18),
                (19, "Cảnh báo Khấu trừ 10%\n(Thu nhập nhiều nơi)", "[T7]", 44),
            ):
                c2 = ws2.cell(2, cc, title)
                c2.font = f_hdr
                c2.fill = fill_hdr
                c2.alignment = al_c
                c2.border = b_thin
                c3 = ws2.cell(3, cc, code)
                c3.font = f_code
                c3.fill = fill_code
                c3.alignment = al_c
                c3.border = b_thin
                ws2.column_dimensions[get_column_letter(cc)].width = w

            for r in range(4, ws2.max_row + 1):
                a_txt = _clean_text(ws2.cell(r, 1).value).upper()
                b_txt = _clean_text(ws2.cell(r, 2).value)
                if any(k in a_txt for k in ("TỔNG CỘNG", "CÁ NHÂN CƯ TRÚ", "CÁ NHÂN KHÔNG CƯ TRÚ")) or "TỔNG CỘNG" in b_txt.upper():
                    continue
                if not b_txt and not _clean_text(ws2.cell(r, 3).value):
                    continue
                mst = clean_mst(ws2.cell(r, 3).value)
                cccd = _clean_id(ws2.cell(r, 5).value)
                is_nonres = _clean_text(ws2.cell(r, 6).value).upper() in ("[X]", "X", "TRUE", "1")
                here_tn = float(ws2.cell(r, 7).value or 0)
                here_kt = float(ws2.cell(r, 11).value or 0)
                others = evaluate_multi_income_for_person(mst, cccd, company_mst, yr, multi_income_records)
                other_tn = sum(x.tnct for x in others)
                other_cnt = len(others)
                tot_all_tn = here_tn + other_tn

                c_tn = ws2.cell(r, 18, round(other_tn))
                c_tn.number_format = NUM
                c_tn.alignment = al_r
                c_tn.font = f_reg
                c_tn.border = b_thin

                is_err = False
                if not is_nonres and here_tn >= 2_000_000 and here_kt + 1000 < round(here_tn * 0.1):
                    if other_cnt > 0 or tot_all_tn > self_yr:
                        msg = f"PHẢI KHẤU TRỪ 10% (Bác CK 08): Có TN tại {other_cnt} nơi khác, tổng TN {tot_all_tn:,.0f} đ"
                        is_err = True
                    else:
                        msg = "Chưa thấy TN nơi khác trên TMS (KT hồ sơ Cam kết 08)"
                elif other_cnt > 0:
                    msg = f"Đã KT đủ (Có TN tại {other_cnt} nơi khác: {other_tn:,.0f} đ)"
                else:
                    msg = "Hợp lệ"

                c_msg = ws2.cell(r, 19, msg)
                c_msg.alignment = al_l
                c_msg.font = f_bold if is_err else f_reg
                c_msg.border = b_thin
                if is_err:
                    summary["kt10_warnings_052"] += 1
                    c_tn.fill = fill_err
                    c_msg.fill = fill_err
                elif other_cnt > 0:
                    c_msg.fill = fill_warn
                else:
                    c_msg.fill = fill_ok

    # 2. ĐỐI CHIẾU DANH BẠ NPT (HKD) & DS NPT CÓ THU NHẬP LÊN 05-3 (W..X = cột 23..24)
    if (npt_dir_records or npt_income_records) and "05-3" in wb.sheetnames:
        ws3 = wb["05-3"]
        if "W1:X1" not in [str(m) for m in ws3.merged_cells.ranges]:
            ws3.merge_cells("W1:X1")
        ws3["W1"] = "KIỂM TRA DANH BẠ HKD & THU NHẬP NPT (TMS)"
        for cc in (23, 24):
            c1 = ws3.cell(1, cc)
            c1.font = f_hdr
            c1.fill = fill_hdr
            c1.alignment = al_c
            c1.border = b_thin
        for cc, title, code, w in (
            (23, "Kiểm tra Danh bạ NPT\n(Hộ kinh doanh)", "[DC12]", 36),
            (24, "Kiểm tra NPT có thu nhập\n(Vượt ngưỡng giảm trừ)", "[DC13]", 40),
        ):
            c2 = ws3.cell(2, cc, title)
            c2.font = f_hdr
            c2.fill = fill_hdr
            c2.alignment = al_c
            c2.border = b_thin
            c3 = ws3.cell(3, cc, code)
            c3.font = f_code
            c3.fill = fill_code
            c3.alignment = al_c
            c3.border = b_thin
            ws3.column_dimensions[get_column_letter(cc)].width = w

        for r in range(4, ws3.max_row + 1):
            ten_nnt = _clean_text(ws3.cell(r, 2).value)
            ten_npt = _clean_text(ws3.cell(r, 4).value)
            mst_npt = clean_mst(ws3.cell(r, 6).value)
            so_gt = _clean_id(ws3.cell(r, 8).value)
            if not mst_npt and not ten_npt:
                continue
            if "TỔNG CỘNG" in (_clean_text(ws3.cell(r, 1).value) + " " + ten_nnt).upper():
                continue
            d_from = parse_ym(_clean_text(ws3.cell(r, 10).value))
            d_to = parse_ym(_clean_text(ws3.cell(r, 11).value))

            # Cột W (23): Danh bạ NPT / HKD
            if npt_dir_records:
                is_hkd_err, hkd_msg = check_npt_hkd_in_period(mst_npt, so_gt, d_from, d_to, npt_dir_records)
            else:
                is_hkd_err, hkd_msg = False, "(Chưa nạp Danh bạ NPT)"
            c_hkd = ws3.cell(r, 23, hkd_msg)
            c_hkd.alignment = al_l
            c_hkd.font = f_bold if is_hkd_err else f_reg
            c_hkd.border = b_thin
            if is_hkd_err:
                summary["hkd_warnings_053"] += 1
                c_hkd.fill = fill_err
            elif npt_dir_records:
                c_hkd.fill = fill_ok

            # Cột X (24): DS NPT có thu nhập
            if npt_income_records:
                is_inc_err, _, inc_msg = check_npt_income_threshold(mst_npt, so_gt, yr, npt_income_records)
            else:
                is_inc_err, inc_msg = False, "(Chưa nạp DS NPT có TN)"
            c_inc = ws3.cell(r, 24, inc_msg)
            c_inc.alignment = al_l
            c_inc.font = f_bold if is_inc_err else f_reg
            c_inc.border = b_thin
            if is_inc_err:
                summary["npt_inc_warnings_053"] += 1
                c_inc.fill = fill_err
            elif npt_income_records:
                c_inc.fill = fill_ok

    return summary


def export_yearly_tncn_with_audit(
    declarations: List[TaxDeclaration],
    out_dir: str,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    tms_npt_path: Optional[str] = None,
    tms_npt_paths: Optional[Sequence[str]] = None,
    ledger_paths: Optional[Sequence[str]] = None,
    multi_income_paths: Optional[Sequence[str]] = None,
    npt_dir_paths: Optional[Sequence[str]] = None,
    npt_income_paths: Optional[Sequence[str]] = None,
) -> Tuple[List[str], Dict[int, Dict[str, Any]]]:
    """
    Lọc tờ khai TNCN theo khoảng [year_from .. year_to], chọn bản bổ sung cuối cùng mỗi năm
    (kế thừa 05-3 từ lần trước), xuất từng file `QTT_TNCN_YYYY.xlsx` chuẩn mẫu `TK, 05-1, 05-2, 05-3`,
    đối chiếu trực tiếp Người phụ thuộc trên sheet `05-3` và tách đúng sổ chi tiết/bảng lương từng năm.
    """
    tncn_valid = []
    for d in declarations:
        if d.error is not None or d.tax_type != "TNCN":
            continue
        if year_from is not None and d.year is not None and d.year < year_from:
            continue
        if year_to is not None and d.year is not None and d.year > year_to:
            continue
        tncn_valid.append(d)

    if not tncn_valid:
        raise ValueError("Không tìm thấy tờ khai Quyết toán TNCN nào hợp lệ trong khoảng năm đã chọn.")

    latest_per_year = filter_latest_per_year(tncn_valid)

    # Gộp nhiều file TMS NPT nếu người dùng chọn nhiều file
    all_tms_paths: List[str] = []
    if tms_npt_path and os.path.isfile(tms_npt_path):
        all_tms_paths.append(tms_npt_path)
    if tms_npt_paths:
        for tp in tms_npt_paths:
            if tp and os.path.isfile(tp) and tp not in all_tms_paths:
                all_tms_paths.append(tp)

    tms_rows: List[TmsNptRow] = []
    for tp in all_tms_paths:
        tms_rows.extend(parse_tms_npt_file(tp))

    multi_inc_records = parse_multi_income_files(multi_income_paths) if multi_income_paths else []
    npt_dir_records = parse_npt_directory_files(npt_dir_paths) if npt_dir_paths else []
    npt_inc_records = parse_npt_income_files(npt_income_paths) if npt_income_paths else []

    # Quét sổ chi tiết & bảng lương tách theo từng năm
    ledgers_by_year = scan_accounting_ledgers_by_year(ledger_paths) if ledger_paths else {}
    undated_ledger = ledgers_by_year.get(None)

    os.makedirs(out_dir, exist_ok=True)
    created_files: List[str] = []
    stats_by_year: Dict[int, Dict[str, Any]] = {}

    for d in latest_per_year:
        yr = d.year or 0
        out_path = os.path.join(out_dir, f"QTT_TNCN_{yr}.xlsx")
        # 1. Xuất đúng 4 sheet chuẩn mẫu tax_xml_to_excel: TK, 05-1, 05-2, 05-3
        write_single_year_tncn_file(out_path, d)

        # 2. Mở lại để đối chiếu trực tiếp NPT trên sheet 05-3 + gắn công thức test nhanh 05-1/05-2 + 3 nguồn TMS mở rộng + quỹ lương
        wb = openpyxl.load_workbook(out_path)
        add_quick_test_formulas_051_052(wb, tax_year=d.year)
        if "05-3" in wb.sheetnames:
            stats_by_year[yr] = reconcile_053_inplace(wb, tms_rows, tax_year=d.year)
        else:
            stats_by_year[yr] = {"total_rows": 0, "matched": 0, "total_diff_months": 0, "total_disallowed_amount": 0}

        if multi_inc_records or npt_dir_records or npt_inc_records:
            extra_st = apply_extra_tms_checks(
                wb,
                tax_year=d.year,
                company_mst=d.mst or "",
                multi_income_records=multi_inc_records,
                npt_dir_records=npt_dir_records,
                npt_income_records=npt_inc_records,
            )
            stats_by_year[yr].update(extra_st)

        yr_ledger = ledgers_by_year.get(yr)
        yr_flagged: List[FlaggedCostEntry] = []
        if yr_ledger:
            _append_payroll_reconcile_sheet(wb, yr, yr_ledger)
            yr_flagged.extend(yr_ledger.flagged_entries)
        if undated_ledger and undated_ledger.flagged_entries:
            yr_flagged.extend(undated_ledger.flagged_entries)
        if yr_flagged:
            _append_cost_risk_sheet(wb, yr_flagged)

        wb.save(out_path)
        created_files.append(out_path)

    return created_files, stats_by_year


def reconcile_existing_tncn_excel(
    excel_path: str,
    tms_npt_path: str,
    tax_year: Optional[int] = None,
    multi_income_paths: Optional[Sequence[str]] = None,
    npt_dir_paths: Optional[Sequence[str]] = None,
    npt_income_paths: Optional[Sequence[str]] = None,
) -> Dict[str, Any]:
    """Đối chiếu trực tiếp NPT trên sheet '05-3' + test nhanh 05-1/05-2 + 3 nguồn TMS mở rộng của file Excel QTT_TNCN_YYYY.xlsx đã xuất sẵn."""
    tms_rows = parse_tms_npt_file(tms_npt_path) if (tms_npt_path and os.path.isfile(tms_npt_path)) else []
    wb = openpyxl.load_workbook(excel_path)
    add_quick_test_formulas_051_052(wb, tax_year=tax_year)
    stats = reconcile_053_inplace(wb, tms_rows, tax_year=tax_year)
    multi_inc = parse_multi_income_files(multi_income_paths) if multi_income_paths else []
    npt_dir = parse_npt_directory_files(npt_dir_paths) if npt_dir_paths else []
    npt_inc = parse_npt_income_files(npt_income_paths) if npt_income_paths else []
    if multi_inc or npt_dir or npt_inc:
        extra_st = apply_extra_tms_checks(
            wb,
            tax_year=tax_year,
            multi_income_records=multi_inc,
            npt_dir_records=npt_dir,
            npt_income_records=npt_inc,
        )
        stats.update(extra_st)
    wb.save(excel_path)
    return stats
