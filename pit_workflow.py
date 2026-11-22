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
from ledger_payroll_scanner import FlaggedCostEntry, scan_accounting_ledgers
from npt_tms_engine import TmsNptRow, audit_dependents, parse_tms_npt_file, parse_ym
from pit_tax_rules import get_policy_for_year
from tax_parser import TaxDeclaration
from xml_05qtt_parser import Row053


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
        mst_nnt = _clean_id(ws_053.cell(r, 3).value)
        ten_npt = _clean_text(ws_053.cell(r, 4).value)
        if not mst_nnt and not ten_npt:
            continue
        if "TỔNG CỘNG" in (str(ws_053.cell(r, 1).value or "") + " " + ten_nnt).upper():
            continue
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
                mst_npt=_clean_id(ws_053.cell(r, 6).value),
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


def export_yearly_tncn_with_audit(
    declarations: List[TaxDeclaration],
    out_dir: str,
    year_from: Optional[int] = None,
    year_to: Optional[int] = None,
    tms_npt_path: Optional[str] = None,
    ledger_paths: Optional[Sequence[str]] = None,
) -> Tuple[List[str], Dict[int, Dict[str, Any]]]:
    """
    Lọc tờ khai TNCN theo khoảng [year_from .. year_to], chọn bản bổ sung cuối cùng mỗi năm
    (kế thừa 05-3 từ lần trước), xuất từng file `QTT_TNCN_YYYY.xlsx` chuẩn mẫu `TK, 05-1, 05-2, 05-3`
    và đối chiếu trực tiếp Người phụ thuộc trên sheet `05-3`.
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
    tms_rows = parse_tms_npt_file(tms_npt_path) if (tms_npt_path and os.path.isfile(tms_npt_path)) else []
    ledger_sum = scan_accounting_ledgers(ledger_paths) if ledger_paths else None

    os.makedirs(out_dir, exist_ok=True)
    created_files: List[str] = []
    stats_by_year: Dict[int, Dict[str, Any]] = {}

    for d in latest_per_year:
        yr = d.year or 0
        out_path = os.path.join(out_dir, f"QTT_TNCN_{yr}.xlsx")
        # 1. Xuất đúng 4 sheet chuẩn mẫu tax_xml_to_excel: TK, 05-1, 05-2, 05-3
        write_single_year_tncn_file(out_path, d)

        # 2. Mở lại để đối chiếu trực tiếp NPT trên sheet 05-3 (luôn chạy kiểm tra độ tuổi & trùng NPT, kể cả khi chưa nạp TMS)
        wb = openpyxl.load_workbook(out_path)
        if "05-3" in wb.sheetnames:
            stats_by_year[yr] = reconcile_053_inplace(wb, tms_rows, tax_year=d.year)
        if ledger_sum and ledger_sum.flagged_entries:
            _append_cost_risk_sheet(wb, ledger_sum.flagged_entries)
        wb.save(out_path)

        created_files.append(out_path)

    return created_files, stats_by_year


def reconcile_existing_tncn_excel(
    excel_path: str,
    tms_npt_path: str,
    tax_year: Optional[int] = None,
) -> Dict[str, Any]:
    """Đối chiếu trực tiếp NPT trên sheet '05-3' của file Excel QTT_TNCN_YYYY.xlsx đã xuất sẵn."""
    tms_rows = parse_tms_npt_file(tms_npt_path)
    wb = openpyxl.load_workbook(excel_path)
    stats = reconcile_053_inplace(wb, tms_rows, tax_year=tax_year)
    wb.save(excel_path)
    return stats
