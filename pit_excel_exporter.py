# -*- coding: utf-8 -*-
"""
Bộ xuất Báo cáo Excel Đối chiếu & Tính lại Thuế TNCN chuyên sâu (7 sheets cài sẵn 100% công thức Excel động):
1. TongHop_KiemTra_TNCN: Khối thông số điều khiển + Bảng tổng hợp kết quả kiểm tra TNCN (thuần đen-trắng để copy vào Biên bản Word)
2. TinhLai_05_1: Chi tiết từng cá nhân trên 05-1, tự động liên kết số tháng NPT bị loại từ DoiChieu_05_3 & tính lại thuế theo biểu lũy tiến 7 bậc
3. TinhLai_05_2: Chi tiết cá nhân trên 05-2 (10% cư trú / 20% không cư trú) + cột điều chỉnh kiểm tra
4. DoiChieu_05_3: Đối chiếu Người phụ thuộc 05-3 vs TMS
5. DS_NPT_TMS: Dữ liệu NPT hệ thống TMS đã làm sạch
6. DoiChieu_QuyLuong_334: Đối chiếu quỹ lương TK 334, BHXH TK 338, Thuế TNCN TK 3335 với 05/QTT-TNCN
7. QuetRuiRo_ChiPhi: Danh sách các bút toán chi phí (641/642/622/627/154/353) có dấu hiệu chịu thuế TNCN bỏ sót
"""

from __future__ import annotations

import os
from typing import Dict, List, Optional, Tuple

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from ledger_payroll_scanner import LedgerAuditSummary
from npt_tms_engine import NptAuditResultRow, TmsNptRow
from pit_tax_rules import excel_formula_progressive_pit_annual, get_policy_for_year
from xml_05qtt_parser import PitDeclarationData


def export_pit_audit_workbook(
    output_path: str,
    decl: PitDeclarationData,
    npt_results: List[NptAuditResultRow],
    npt_per_nnt: Dict[str, Tuple[int, int]],
    tms_rows: List[TmsNptRow],
    ledger_summary: Optional[LedgerAuditSummary] = None,
) -> str:
    policy = get_policy_for_year(decl.year)
    wb = openpyxl.Workbook()
    wb.remove(wb.active)

    # ===== SHARED STYLES =====
    F_NAME = "Times New Roman"
    f_title = Font(name=F_NAME, size=13, bold=True, color="000000")
    f_sub = Font(name=F_NAME, size=10, italic=True, color="000000")
    f_sec = Font(name=F_NAME, size=11, bold=True, color="000000")
    f_hdr_w = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    f_hdr_bw = Font(name=F_NAME, size=10, bold=True, color="000000")
    f_reg = Font(name="Calibri", size=9, color="000000")
    f_bold = Font(name="Calibri", size=9, bold=True, color="000000")
    f_bw_reg = Font(name=F_NAME, size=10, color="000000")
    f_bw_bold = Font(name=F_NAME, size=10, bold=True, color="000000")

    fill_hdr = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_code = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_input = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_ok = PatternFill(start_color="E2EFDA", end_color="E2EFDA", fill_type="solid")
    fill_err = PatternFill(start_color="FCE4D6", end_color="FCE4D6", fill_type="solid")
    fill_warn = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
    fill_none = PatternFill(fill_type=None)

    thin = Side(style="thin", color="000000")
    b_thin = Border(left=thin, right=thin, top=thin, bottom=thin)
    al_c = Alignment(horizontal="center", vertical="center", wrap_text=True)
    al_l = Alignment(horizontal="left", vertical="center", wrap_text=True)
    al_r = Alignment(horizontal="right", vertical="center")
    NUM = "#,##0"

    # Tạo trước các sheet theo thứ tự nghiệp vụ
    ws_sum = wb.create_sheet("TongHop_KiemTra_TNCN")
    ws_051 = wb.create_sheet("TinhLai_05_1")
    ws_052 = wb.create_sheet("TinhLai_05_2")
    ws_053 = wb.create_sheet("DoiChieu_05_3")
    ws_tms = wb.create_sheet("DS_NPT_TMS")
    ws_334 = wb.create_sheet("DoiChieu_QuyLuong_334")
    ws_cost = wb.create_sheet("QuetRuiRo_ChiPhi")

    # =========================================================================
    # SHEET 4: DoiChieu_05_3 (Viết trước để TinhLai_05_1 dùng SUMIF liên kết trực tiếp)
    # =========================================================================
    h_053 = [
        "STT", "Họ tên NNT", "MST NNT", "Họ tên NPT", "Ngày sinh", "MST NPT",
        "Loại giấy tờ", "Số giấy tờ", "Quan hệ", "Từ tháng (05-3)", "Đến tháng (05-3)",
        "Số tháng 05-3", "Từ tháng (TMS)", "Đến tháng (TMS)", "Trạng thái MST NPT (TMS)",
        "Số tháng hợp lệ", "Số tháng bị loại trừ", "Tiền giảm trừ loại trừ", "Kết quả đối chiếu",
    ]
    ws_053.row_dimensions[1].height = 28
    for ci, h in enumerate(h_053, start=1):
        c = ws_053.cell(1, ci, h)
        c.font = f_hdr_w
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    for idx, item in enumerate(npt_results, start=1):
        r = idx + 1
        s = item.src
        vals = [
            idx, s.ho_ten_nnt, s.mst_nnt, s.ho_ten_npt, s.ngay_sinh, s.mst_npt,
            s.loai_gt, s.so_gt, s.quan_he, s.tu_thang, s.den_thang,
            item.decl_months, item.tms_from, item.tms_to, item.tms_status,
            item.valid_months, f"=MAX(0,L{r}-P{r})", f"=Q{r}*TongHop_KiemTra_TNCN!$C$5", item.status,
        ]
        for ci, v in enumerate(vals, start=1):
            c = ws_053.cell(r, ci, v)
            c.font = f_bold if ci == 19 else f_reg
            c.border = b_thin
            if ci in (3, 6, 8):
                c.number_format = "@"
                c.alignment = al_l
            elif ci in (1, 5, 9, 10, 11, 12, 13, 14, 16, 17):
                c.alignment = al_c
            elif ci == 18:
                c.number_format = NUM
                c.alignment = al_r
            else:
                c.alignment = al_l
            if ci in (17, 18, 19):
                c.fill = fill_ok if item.is_ok else (fill_err if item.disallowed_months > 0 else fill_warn)

    last_053 = max(2, len(npt_results) + 1)
    r_tot_053 = last_053 + 1
    ws_053.cell(r_tot_053, 2, "TỔNG CỘNG").font = f_bold
    ws_053.cell(r_tot_053, 12, f"=SUM(L2:L{last_053})" if npt_results else 0).alignment = al_c
    ws_053.cell(r_tot_053, 16, f"=SUM(P2:P{last_053})" if npt_results else 0).alignment = al_c
    ws_053.cell(r_tot_053, 17, f"=SUM(Q2:Q{last_053})" if npt_results else 0).alignment = al_c
    ws_053.cell(r_tot_053, 18, f"=SUM(R2:R{last_053})" if npt_results else 0).number_format = NUM
    for ci in range(1, 20):
        c = ws_053.cell(r_tot_053, ci)
        c.font = f_bold
        c.fill = fill_code
        c.border = b_thin
    for ci, w in enumerate([6, 22, 13, 22, 11, 13, 12, 15, 10, 11, 11, 10, 11, 12, 26, 11, 11, 17, 32], start=1):
        ws_053.column_dimensions[get_column_letter(ci)].width = w
    ws_053.freeze_panes = "D2"

    # =========================================================================
    # SHEET 2: TinhLai_05_1 (Tính lại thuế TNCN từng cá nhân theo biểu lũy tiến 7 bậc)
    # =========================================================================
    h_051 = [
        "STT", "Họ và tên NNT", "MST NNT", "CCCD/CMND", "Ủy quyền QT",
        "Tổng TNCT (05-1)", "TN miễn giảm (05-1)", "Giảm trừ gia cảnh (05-1)",
        "BH bắt buộc & giảm trừ khác", "TN tính thuế (05-1)", "Thuế đã khấu trừ (05-1)",
        "TNCT tăng thêm qua KT\n(Nhập tay)", "Số tháng NPT bị loại\n(Link 05-3)",
        "Tiền giảm trừ NPT bị loại", "Điều chỉnh BH/giảm trừ khác\n(Nhập tay)",
        "TN tính thuế (Kiểm tra)", "Thuế TNCN phải nộp (KT - 7 bậc)",
        "Thuế TNCN truy thu thêm\n(=MAX(0, KT-BC))", "Ghi chú kiểm tra",
    ]
    ws_051.row_dimensions[1].height = 30
    for ci, h in enumerate(h_051, start=1):
        c = ws_051.cell(1, ci, h)
        c.font = f_hdr_w
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    for idx, r1 in enumerate(decl.rows_051, start=1):
        r = idx + 1
        other_ded = r1.tu_thien + r1.bao_hiem + r1.quy_huu_tri
        # Liên kết tự động số tháng NPT bị loại trừ từ DoiChieu_05_3 bằng SUMIF theo MST NNT
        npt_disallowed_formula = (
            f"=IF(C{r}=\"\",0,SUMIF(DoiChieu_05_3!$C$2:$C${last_053},C{r},DoiChieu_05_3!$Q$2:$Q${last_053}))"
            if npt_results else 0
        )
        # TN tính thuế (Kiểm tra) = MAX(0, TNTT_05_1 + TNCT_tăng_thêm + Tiền_NPT_bị_loại - ĐC_BH)
        # Hoặc tính lại đầy đủ từ gốc: MAX(0, (F+L) - G - MAX(0, H - N) - (I + O))
        tntt_kt_formula = f"=MAX(0,(F{r}+L{r})-G{r}-MAX(0,H{r}-N{r})-(I{r}+O{r}))"
        pit_kt_formula = excel_formula_progressive_pit_annual(f"P{r}", year=decl.year)
        truy_thu_formula = f"=MAX(0,Q{r}-K{r})"

        vals = [
            idx, r1.ho_ten, r1.mst, r1.cccd, "X" if r1.uy_quyen else "",
            r1.tong_tnct, r1.tn_mien_giam, r1.giam_tru_gia_canh,
            other_ded, r1.tn_tinh_thue, r1.thue_da_khau_tru,
            0, npt_disallowed_formula, f"=M{r}*TongHop_KiemTra_TNCN!$C$5", 0,
            tntt_kt_formula, pit_kt_formula, truy_thu_formula,
            f'=IF(R{r}>0,"Phát sinh truy thu TNCN",IF(M{r}>0,"Có NPT bị loại (chưa đến ngưỡng nộp thuế)",""))',
        ]
        for ci, v in enumerate(vals, start=1):
            c = ws_051.cell(r, ci, v)
            c.font = f_bold if ci == 18 else f_reg
            c.border = b_thin
            if ci in (3, 4):
                c.number_format = "@"
                c.alignment = al_l
            elif ci in (1, 5, 13):
                c.alignment = al_c
            elif ci in (2, 19):
                c.alignment = al_l
            else:
                c.number_format = NUM
                c.alignment = al_r
            if ci in (12, 15):
                c.fill = fill_input

    last_051 = max(2, len(decl.rows_051) + 1)
    r_tot_051 = last_051 + 1
    ws_051.cell(r_tot_051, 2, "TỔNG CỘNG (05-1)").font = f_bold
    for ci in (6, 7, 8, 9, 10, 11, 12, 13, 14, 15, 16, 17, 18):
        col_l = get_column_letter(ci)
        c = ws_051.cell(r_tot_051, ci, f"=SUM({col_l}2:{col_l}{last_051})" if decl.rows_051 else 0)
        c.number_format = NUM
        c.alignment = al_c if ci == 13 else al_r
    for ci in range(1, 20):
        c = ws_051.cell(r_tot_051, ci)
        c.font = f_bold
        c.fill = fill_code
        c.border = b_thin
    for ci, w in enumerate([6, 24, 14, 14, 9, 16, 14, 16, 15, 16, 16, 16, 12, 16, 14, 16, 17, 17, 28], start=1):
        ws_051.column_dimensions[get_column_letter(ci)].width = w
    ws_051.freeze_panes = "D2"

    # =========================================================================
    # SHEET 3: TinhLai_05_2 (10% cư trú / 20% không cư trú)
    # =========================================================================
    h_052 = [
        "STT", "Họ và tên NNT", "MST NNT", "CCCD/Hộ chiếu", "Không cư trú (20%)",
        "Tổng TNCT (05-2)", "TN miễn giảm (05-2)", "Thuế đã khấu trừ (05-2)",
        "TNCT tăng thêm qua KT\n(Nhập tay)", "Thuế suất áp dụng", "Thuế TNCN phải nộp (KT)",
        "Thuế TNCN truy thu thêm", "Ghi chú",
    ]
    ws_052.row_dimensions[1].height = 28
    for ci, h in enumerate(h_052, start=1):
        c = ws_052.cell(1, ci, h)
        c.font = f_hdr_w
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    for idx, r2 in enumerate(decl.rows_052, start=1):
        r = idx + 1
        rate = 0.20 if r2.khong_cu_tru else 0.10
        vals = [
            idx, r2.ho_ten, r2.mst, r2.cccd, "X" if r2.khong_cu_tru else "",
            r2.tong_tnct, r2.tn_mien_giam, r2.thue_da_khau_tru,
            0, rate, f"=ROUND(MAX(0,F{r}+I{r}-G{r})*J{r},0)", f"=MAX(0,K{r}-H{r})", "",
        ]
        for ci, v in enumerate(vals, start=1):
            c = ws_052.cell(r, ci, v)
            c.font = f_bold if ci == 12 else f_reg
            c.border = b_thin
            if ci in (3, 4):
                c.number_format = "@"
                c.alignment = al_l
            elif ci in (1, 5):
                c.alignment = al_c
            elif ci == 10:
                c.number_format = "0%"
                c.alignment = al_c
            elif ci in (2, 13):
                c.alignment = al_l
            else:
                c.number_format = NUM
                c.alignment = al_r
            if ci == 9:
                c.fill = fill_input

    last_052 = max(2, len(decl.rows_052) + 1)
    r_tot_052 = last_052 + 1
    ws_052.cell(r_tot_052, 2, "TỔNG CỘNG (05-2)").font = f_bold
    for ci in (6, 7, 8, 9, 11, 12):
        col_l = get_column_letter(ci)
        c = ws_052.cell(r_tot_052, ci, f"=SUM({col_l}2:{col_l}{last_052})" if decl.rows_052 else 0)
        c.number_format = NUM
        c.alignment = al_r
    for ci in range(1, 14):
        c = ws_052.cell(r_tot_052, ci)
        c.font = f_bold
        c.fill = fill_code
        c.border = b_thin
    for ci, w in enumerate([6, 24, 14, 15, 11, 16, 15, 16, 16, 11, 17, 17, 24], start=1):
        ws_052.column_dimensions[get_column_letter(ci)].width = w
    ws_052.freeze_panes = "D2"

    # =========================================================================
    # SHEET 5: DS_NPT_TMS
    # =========================================================================
    h_tms = [
        "STT", "Ngày tạo", "Ngày TĐ", "MST NNT", "Tên NNT", "MST NPT", "Tên NPT",
        "Ngày sinh", "CMND/CCCD", "Hộ chiếu", "Quan hệ", "Từ tháng (TMS)", "Đến tháng (TMS)",
        "Trạng thái MST NPT", "Mã TT", "Mẫu BK/TK",
    ]
    ws_tms.row_dimensions[1].height = 24
    for ci, h in enumerate(h_tms, start=1):
        c = ws_tms.cell(1, ci, h)
        c.font = f_hdr_w
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin
    for idx, rec in enumerate(tms_rows, start=1):
        r = idx + 1
        tf = f"{rec.tu_thang[1]:02d}/{rec.tu_thang[0]}" if rec.tu_thang else ""
        tt = f"{rec.den_thang[1]:02d}/{rec.den_thang[0]}" if rec.den_thang else ""
        vals = [
            idx, rec.ngay_tao, rec.ngay_td, rec.mst_nnt, rec.ten_nnt, rec.mst_npt, rec.ten_npt,
            rec.ngay_sinh, rec.cmnd, rec.ho_chieu, rec.quan_he, tf, tt, rec.trang_thai, rec.ma_tt, rec.mau_tk,
        ]
        for ci, v in enumerate(vals, start=1):
            c = ws_tms.cell(r, ci, v)
            c.font = f_reg
            c.border = b_thin
            if ci in (4, 6, 9, 10, 15):
                c.number_format = "@"
                c.alignment = al_l
            elif ci in (1, 2, 3, 8, 11, 12, 13, 16):
                c.alignment = al_c
            else:
                c.alignment = al_l
    for ci, w in enumerate([6, 12, 12, 14, 24, 14, 24, 12, 15, 13, 10, 12, 12, 32, 8, 11], start=1):
        ws_tms.column_dimensions[get_column_letter(ci)].width = w
    ws_tms.freeze_panes = "A2"

    # =========================================================================
    # SHEET 6: DoiChieu_QuyLuong_334
    # =========================================================================
    ws_334.merge_cells("A1:F1")
    ws_334["A1"] = f"ĐỐI CHIẾU QUỸ LƯƠNG (TK 334), BẢO HIỂM (TK 338) & THUẾ TNCN (TK 3335) VỚI 05/QTT-TNCN NĂM {decl.year}"
    ws_334["A1"].font = f_title
    ws_334["A1"].alignment = al_c
    ws_334.row_dimensions[1].height = 24

    h_334 = ["STT", "Chỉ tiêu đối chiếu", "Số liệu trên Sổ kế toán", "Số liệu trên 05/QTT-TNCN", "Chênh lệch (Sổ KT - Tờ khai)", "Ghi chú nghiệp vụ"]
    ws_334.row_dimensions[3].height = 24
    for ci, h in enumerate(h_334, start=1):
        c = ws_334.cell(3, ci, h)
        c.font = f_hdr_w
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    ls = ledger_summary or LedgerAuditSummary()
    rows_334 = [
        (1, "Tổng phát sinh Có TK 334 (Lương, thưởng phải trả) vs Tổng TNCT (05-1 + 05-2)",
         round(ls.tk334_credit_total), f"=TinhLai_05_1!F{r_tot_051}+TinhLai_05_2!F{r_tot_052}",
         "Kiểm tra chênh lệch do thời điểm chi trả (lương T12 năm trước/năm nay) hoặc khoản không chịu thuế"),
        (2, "Tổng phát sinh Nợ TK 334 (Thực chi trả NLĐ) vs Tổng TNCT (05-1 + 05-2)",
         round(ls.tk334_debit_total), f"=TinhLai_05_1!F{r_tot_051}+TinhLai_05_2!F{r_tot_052}",
         "Thuế TNCN tính theo thời điểm thực chi trả (Khoản 2 Điều 8 TT 111/2013)"),
        (3, "Bảo hiểm bắt buộc khấu trừ lương (TK 3383/3384/3386) vs Giảm trừ BH (05-1)",
         round(ls.tk338_insurance_total), f"=TinhLai_05_1!I{r_tot_051}",
         "Đối chiếu số BHXH, BHYT, BHTN phần người lao động đóng được trừ khi tính TNCN"),
        (4, "Thuế TNCN phải nộp phát sinh (Có TK 3335) vs Tổng thuế đã khấu trừ (05-1 + 05-2)",
         round(ls.tk3335_credit_total), f"=TinhLai_05_1!K{r_tot_051}+TinhLai_05_2!H{r_tot_052}",
         "Kiểm tra số thuế TNCN đã khấu trừ trên sổ sách kế toán khớp tờ khai quyết toán"),
        (5, "Tổng chi Quỹ khen thưởng, phúc lợi (Nợ TK 353)",
         round(ls.tk353_debit_total), 0,
         "Rà soát các khoản thưởng/phúc lợi chi đích danh cá nhân có nguy cơ bỏ sót TNCT"),
    ]
    for idx_r, (stt, lbl, acc_v, tk_f, note) in enumerate(rows_334, start=4):
        ws_334.row_dimensions[idx_r].height = 22
        ws_334.cell(idx_r, 1, stt).alignment = al_c
        ws_334.cell(idx_r, 2, lbl).alignment = al_l
        ws_334.cell(idx_r, 3, acc_v).number_format = NUM
        ws_334.cell(idx_r, 3).alignment = al_r
        ws_334.cell(idx_r, 4, tk_f).number_format = NUM
        ws_334.cell(idx_r, 4).alignment = al_r
        ws_334.cell(idx_r, 5, f"=C{idx_r}-D{idx_r}").number_format = NUM
        ws_334.cell(idx_r, 5).alignment = al_r
        ws_334.cell(idx_r, 6, note).alignment = al_l
        for ci in range(1, 7):
            c = ws_334.cell(idx_r, ci)
            c.font = f_bold if ci == 5 else f_reg
            c.border = b_thin
    for ci, w in enumerate([6, 44, 20, 20, 20, 44], start=1):
        ws_334.column_dimensions[get_column_letter(ci)].width = w

    # =========================================================================
    # SHEET 7: QuetRuiRo_ChiPhi (Các bút toán chi phí nghi vấn chịu thuế TNCN bỏ sót)
    # =========================================================================
    h_cost = [
        "STT", "File nguồn", "Ngày CT", "Số CT", "TK Nợ", "TK Có", "Đối tượng",
        "Diễn giải chứng từ", "Số tiền (đồng)", "Nhóm rủi ro TNCN", "Mức độ",
        "Số tiền đưa vào tính truy thu\n(Nhập tay)", "Căn cứ pháp lý / Gợi ý kiểm tra",
    ]
    ws_cost.row_dimensions[1].height = 28
    for ci, h in enumerate(h_cost, start=1):
        c = ws_cost.cell(1, ci, h)
        c.font = f_hdr_w
        c.fill = fill_hdr
        c.alignment = al_c
        c.border = b_thin

    flagged = ls.flagged_entries
    for idx, fe in enumerate(flagged, start=1):
        r = idx + 1
        vals = [
            idx, fe.source_file, fe.ngay_ct, fe.so_ct, fe.tk_no, fe.tk_co, fe.doi_tuong,
            fe.dien_giai, round(fe.so_tien), fe.nhom_rui_ro, fe.muc_do, 0, fe.can_cu_phap_ly,
        ]
        for ci, v in enumerate(vals, start=1):
            c = ws_cost.cell(r, ci, v)
            c.font = f_bold if ci == 11 else f_reg
            c.border = b_thin
            if ci in (1, 3, 4, 5, 6, 11):
                c.alignment = al_c
            elif ci in (9, 12):
                c.number_format = NUM
                c.alignment = al_r
            else:
                c.alignment = al_l
            if ci == 11:
                c.fill = fill_err if fe.muc_do == "CAO" else fill_warn
            if ci == 12:
                c.fill = fill_input

    last_cost = max(2, len(flagged) + 1)
    r_tot_cost = last_cost + 1
    ws_cost.cell(r_tot_cost, 8, "TỔNG CỘNG BÚT TOÁN RỦI RO").font = f_bold
    ws_cost.cell(r_tot_cost, 9, f"=SUM(I2:I{last_cost})" if flagged else 0).number_format = NUM
    ws_cost.cell(r_tot_cost, 12, f"=SUM(L2:L{last_cost})" if flagged else 0).number_format = NUM
    for ci in range(1, 14):
        c = ws_cost.cell(r_tot_cost, ci)
        c.font = f_bold
        c.fill = fill_code
        c.border = b_thin
    for ci, w in enumerate([6, 18, 12, 12, 9, 9, 20, 38, 16, 30, 11, 16, 42], start=1):
        ws_cost.column_dimensions[get_column_letter(ci)].width = w
    ws_cost.freeze_panes = "A2"

    # =========================================================================
    # SHEET 1: TongHop_KiemTra_TNCN (Thuần đen - trắng, sẵn sàng copy vào Biên bản Word)
    # =========================================================================
    ws_sum.merge_cells("A1:E1")
    ws_sum["A1"] = f"BÁO CÁO TỔNG HỢP ĐỐI CHIẾU & TÍNH LẠI THUẾ TNCN NĂM {decl.year}"
    ws_sum["A1"].font = f_title
    ws_sum["A1"].alignment = al_c
    ws_sum.row_dimensions[1].height = 24

    ws_sum.merge_cells("A2:E2")
    ws_sum["A2"] = f"(Căn cứ chính sách: {policy.legal_basis} | Đơn vị tính: Đồng — Định dạng thuần đen-trắng sẵn sàng copy vào Biên bản Word)"
    ws_sum["A2"].font = f_sub
    ws_sum["A2"].alignment = al_c
    ws_sum.row_dimensions[2].height = 18

    # Khối thông số điều khiển (Rows 4-6)
    ctrl_items = [
        (4, "Mức giảm trừ bản thân (đồng/tháng):", policy.self_deduction_monthly),
        (5, "Mức giảm trừ Người phụ thuộc (đồng/tháng — liên kết sang 05-1 & 05-3):", policy.dependent_deduction_monthly),
        (6, "Năm quyết toán thuế TNCN:", decl.year),
    ]
    for r_c, lbl, val in ctrl_items:
        ws_sum.row_dimensions[r_c].height = 20
        ws_sum.merge_cells(start_row=r_c, start_column=1, end_row=r_c, end_column=2)
        ws_sum.cell(r_c, 1, lbl).font = f_bw_bold
        ws_sum.cell(r_c, 1).alignment = al_l
        ws_sum.cell(r_c, 1).border = b_thin
        ws_sum.cell(r_c, 2).border = b_thin
        c_v = ws_sum.cell(r_c, 3, val)
        c_v.font = f_bw_bold
        c_v.number_format = NUM if r_c != 6 else "0"
        c_v.alignment = al_r
        c_v.border = b_thin

    # Bảng tổng hợp đưa vào Biên bản (Thuần đen-trắng, không tô màu nền)
    ws_sum.merge_cells("A8:E8")
    ws_sum["A8"] = "I. BẢNG TỔNG HỢP KẾT QUẢ KIỂM TRA THUẾ TNCN (ĐƯA VÀO BIÊN BẢN)"
    ws_sum["A8"].font = f_sec
    ws_sum.row_dimensions[8].height = 22

    sum_hdrs = ["STT", "Chỉ tiêu", "Số báo cáo của NNT", "Số kiểm tra", "Chênh lệch"]
    ws_sum.row_dimensions[9].height = 24
    for ci, h in enumerate(sum_hdrs, start=1):
        c = ws_sum.cell(9, ci, h)
        c.font = f_hdr_bw
        c.fill = fill_none
        c.alignment = al_c
        c.border = b_thin

    bb_rows = [
        ("1", "Tổng thu nhập chịu thuế (05-1 - Biểu lũy tiến từng phần)",
         f"=TinhLai_05_1!F{r_tot_051}", f"=TinhLai_05_1!F{r_tot_051}+TinhLai_05_1!L{r_tot_051}"),
        ("2", "Tổng thu nhập tính thuế (05-1)",
         f"=TinhLai_05_1!J{r_tot_051}", f"=TinhLai_05_1!P{r_tot_051}"),
        ("3", "Thuế TNCN phải khấu trừ / phải nộp (05-1 - Lũy tiến 7 bậc)",
         f"=TinhLai_05_1!K{r_tot_051}", f"=TinhLai_05_1!Q{r_tot_051}"),
        ("4", "Tổng thu nhập chịu thuế (05-2 - Toàn phần 10% / 20%)",
         f"=TinhLai_05_2!F{r_tot_052}", f"=TinhLai_05_2!F{r_tot_052}+TinhLai_05_2!I{r_tot_052}"),
        ("5", "Thuế TNCN phải khấu trừ / phải nộp (05-2 - 10% / 20%)",
         f"=TinhLai_05_2!H{r_tot_052}", f"=TinhLai_05_2!K{r_tot_052}"),
        ("6", "Tổng cộng Thuế TNCN phải khấu trừ / phải nộp (= 3 + 5)",
         "=C12+C14", "=D12+D14"),
        ("7", "Tổng số thuế TNCN truy thu thêm qua kiểm tra (05-1 + 05-2)",
         0, f"=TinhLai_05_1!R{r_tot_051}+TinhLai_05_2!L{r_tot_052}"),
    ]
    for offset, (stt, lbl, bc_f, kt_f) in enumerate(bb_rows):
        rr = 10 + offset
        is_tot = stt in ("6", "7")
        ws_sum.row_dimensions[rr].height = 20
        ws_sum.cell(rr, 1, stt).alignment = al_c
        ws_sum.cell(rr, 2, lbl).alignment = al_l
        ws_sum.cell(rr, 3, bc_f).number_format = NUM
        ws_sum.cell(rr, 3).alignment = al_r
        ws_sum.cell(rr, 4, kt_f).number_format = NUM
        ws_sum.cell(rr, 4).alignment = al_r
        ws_sum.cell(rr, 5, f"=D{rr}-C{rr}").number_format = NUM
        ws_sum.cell(rr, 5).alignment = al_r
        for ci in range(1, 6):
            c = ws_sum.cell(rr, ci)
            c.font = f_bw_bold if is_tot else f_bw_reg
            c.fill = fill_none
            c.border = b_thin

    # Mục II: Tóm tắt đối chiếu NPT & Quét rủi ro chi phí
    ws_sum.merge_cells("A18:E18")
    ws_sum["A18"] = "II. TÓM TẮT ĐỐI CHIẾU NGƯỜI PHỤ THUỘC (05-3) & RÀ SOÁT SỔ CHI PHÍ"
    ws_sum["A18"].font = f_sec
    ws_sum.row_dimensions[18].height = 22

    stats_rows = [
        ("1", "Tổng số dòng Người phụ thuộc trên 05-3", len(npt_results)),
        ("2", "Tổng số tháng NPT bị loại trừ (Chưa ĐK / Vượt tháng / Trùng NPT)", f"=DoiChieu_05_3!Q{r_tot_053}"),
        ("3", "Tổng số tiền giảm trừ NPT không đủ điều kiện (đồng)", f"=DoiChieu_05_3!R{r_tot_053}"),
        ("4", "Số bút toán chi phí nghi vấn chịu thuế TNCN bỏ sót (TK 641/642/622/627/353)", len(flagged)),
        ("5", "Tổng giá trị bút toán chi phí nghi vấn cần rà soát (đồng)", f"=QuetRuiRo_ChiPhi!I{r_tot_cost}"),
    ]
    for offset, (stt, lbl, val) in enumerate(stats_rows):
        rr = 19 + offset
        ws_sum.row_dimensions[rr].height = 20
        ws_sum.cell(rr, 1, stt).alignment = al_c
        ws_sum.merge_cells(start_row=rr, start_column=2, end_row=rr, end_column=3)
        ws_sum.cell(rr, 2, lbl).alignment = al_l
        ws_sum.merge_cells(start_row=rr, start_column=4, end_row=rr, end_column=5)
        ws_sum.cell(rr, 4, val).number_format = NUM
        ws_sum.cell(rr, 4).alignment = al_r
        for ci in range(1, 6):
            c = ws_sum.cell(rr, ci)
            c.font = f_bw_bold if stt in ("3", "5") else f_bw_reg
            c.fill = fill_none
            c.border = b_thin

    for ci, w in enumerate([8, 48, 22, 22, 22], start=1):
        ws_sum.column_dimensions[get_column_letter(ci)].width = w
    ws_sum.page_setup.orientation = ws_sum.ORIENTATION_LANDSCAPE
    ws_sum.page_setup.paperSize = ws_sum.PAPERSIZE_A4

    os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
    wb.save(output_path)
    return output_path
