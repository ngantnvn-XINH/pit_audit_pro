# -*- coding: utf-8 -*-
"""
Bộ kiểm thử tự động cho phần mềm PIT Audit Pro (pit_audit_pro v1.0).
Toàn bộ dữ liệu kiểm thử là dữ liệu giả lập ẩn danh (CTY A, MST 0100000001...).
"""

from __future__ import annotations

import os
import tempfile
import unittest

import openpyxl

from ledger_payroll_scanner import scan_accounting_ledgers
from npt_tms_engine import audit_dependents, parse_tms_npt_file
from pit_excel_exporter import export_pit_audit_workbook
from pit_tax_rules import (
    calc_housing_taxable_amount,
    calc_progressive_pit_annual,
    calc_progressive_pit_monthly,
    excel_formula_progressive_pit_annual,
    get_policy_for_year,
)
from xml_05qtt_parser import load_pit_declaration


class TestPitAuditPro(unittest.TestCase):
    def test_policy_and_progressive_schedules_across_periods(self) -> None:
        """Kiểm thử chính sách giảm trừ & biểu thuế lũy tiến 7 bậc (<=2025) vs 5 bậc (>=2026)."""
        p19 = get_policy_for_year(2019)
        self.assertEqual((p19.self_deduction_monthly, p19.dependent_deduction_monthly), (9_000_000, 3_600_000))
        p24 = get_policy_for_year(2024)
        self.assertEqual((p24.self_deduction_monthly, p24.dependent_deduction_monthly), (11_000_000, 4_400_000))
        p26 = get_policy_for_year(2026)
        self.assertEqual((p26.self_deduction_monthly, p26.dependent_deduction_monthly), (15_500_000, 6_200_000))

        # Kiểm tra biểu 7 bậc (2024): TNTT 20tr/tháng -> 20tr*20% - 1.65tr = 2.35tr/tháng (28.2tr/năm)
        self.assertEqual(calc_progressive_pit_monthly(20_000_000, year=2024), 2_350_000)
        self.assertEqual(calc_progressive_pit_annual(240_000_000, months=12, year=2024), 28_200_000)

        # Kiểm tra biểu 5 bậc (2026): TNTT 20tr/tháng -> 20tr*10% - 0.5tr = 1.5tr/tháng (18tr/năm)
        self.assertEqual(calc_progressive_pit_monthly(20_000_000, year=2026), 1_500_000)
        self.assertEqual(calc_progressive_pit_annual(240_000_000, months=12, year=2026), 18_000_000)
        # Bậc 5 (2026): TNTT 120tr/tháng -> 120tr*35% - 14.5tr = 27.5tr/tháng
        self.assertEqual(calc_progressive_pit_monthly(120_000_000, year=2026), 27_500_000)

        # Kiểm tra công thức Excel 7 bậc vs 5 bậc
        self.assertIn("60000000", excel_formula_progressive_pit_annual("P2", year=2024))
        self.assertIn("1200000000", excel_formula_progressive_pit_annual("P2", year=2026))
        self.assertNotIn("216000000", excel_formula_progressive_pit_annual("P2", year=2026))

        # Kiểm tra trần tiền thuê nhà 15%
        self.assertEqual(calc_housing_taxable_amount(50_000_000, 200_000_000), 30_000_000)
        self.assertEqual(calc_housing_taxable_amount(20_000_000, 200_000_000), 20_000_000)

    def test_end_to_end_xml_tms_ledger_and_export(self) -> None:
        """Kiểm thử luồng đầy đủ: Đọc XML 05/QTT-TNCN + Đối chiếu NPT TMS + Quét sổ chi phí + Xuất báo cáo 7 sheet."""
        xml_05 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung><TTinTKhaiThue>
      <TKhaiThue><maTKhai>953</maTKhai><tenTKhai>Tờ khai QTT TNCN</tenTKhai>
        <KyKKhaiThue><kyKKhai>2024</kyKKhai><kyKKhaiTuNgay>01/01/2024</kyKKhaiTuNgay><kyKKhaiDenNgay>31/12/2024</kyKKhaiDenNgay></KyKKhaiThue>
        <ngayLapTKhai>2025-03-20</ngayLapTKhai><soLan>0</soLan>
      </TKhaiThue>
      <NNT><mst>0100000001</mst><tenNNT>CTY A</tenNNT></NNT>
    </TTinTKhaiThue></TTinChung>
    <PLuc>
      <PLuc_05_1_BK_QTT>
        <BCaoCTietCNhan>
          <CTietCNhan>
            <ct07>Nguyễn Văn A</ct07><ct08>8000000001</ct08><ct09>079090000001</ct09><ct10>1</ct10>
            <ct11>300000000</ct11><ct15>184800000</ct15><ct17>25200000</ct17><ct19>90000000</ct19><ct20>6000000</ct20>
          </CTietCNhan>
        </BCaoCTietCNhan>
      </PLuc_05_1_BK_QTT>
      <PLuc_05_2_BK_QTT>
        <BCaoCTietCNhan>
          <CTietCNhan>
            <ct07>Trần Văn B</ct07><ct08>8000000002</ct08><ct09>079088000002</ct09><ct10>0</ct10>
            <ct11>30000000</ct11><ct14>3000000</ct14>
          </CTietCNhan>
        </BCaoCTietCNhan>
      </PLuc_05_2_BK_QTT>
      <PLuc_05_3_BK_QTT>
        <BCaoCTietNPT>
          <CTietNPT>
            <ct07>Nguyễn Văn A</ct07><ct08>8000000001</ct08><ct09>Nguyễn Bé Một</ct09>
            <ct10>01/01/2018</ct10><ct11>8900000001</ct11><ct12>CCCD</ct12><ct13>079218000001</ct13>
            <ct14>Con</ct14><ct15>01/2024</ct15><ct16>12/2024</ct16>
          </CTietNPT>
        </BCaoCTietNPT>
      </PLuc_05_3_BK_QTT>
    </PLuc>
  </HSoKhaiThue>
</HSoThueDTu>"""

        tms_csv = (
            "19.11.2026 HIỂN THỊ DANH SÁCH ĐỘNG\n\nDanh sách kết quả\n"
            "Ngày tạo,Ngày TĐ,MST NNT,Tên NNT,MST NPT,Tên NPT,Tên trạng thái NPT,Ngày sinh,CMND,Quan hệ,Tháng,Tháng,Mẫu BK/TK,Mã trạng t\n"
            "10.01.2024,10.07.2024,8000000001,Nguyễn Văn A,8900000001,Nguyễn Bé Một,NNT đang hoạt động,01.01.2018,079218000001,Con,12024,62024,16-TH,00\n"
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            xml_p = os.path.join(tmpdir, "05qtt_2024.xml")
            tms_p = os.path.join(tmpdir, "tms_npt.csv")
            led_p = os.path.join(tmpdir, "so_chi_tiet_334_642.xlsx")
            out_p = os.path.join(tmpdir, "BaoCao_KiemTra_TNCN.xlsx")

            with open(xml_p, "w", encoding="utf-8") as f:
                f.write(xml_05)
            with open(tms_p, "w", encoding="utf-8") as f:
                f.write(tms_csv)

            # Tạo file sổ chi tiết tài khoản giả lập (TK 334, 353, 642)
            wb_l = openpyxl.Workbook()
            ws_l = wb_l.active
            ws_l.title = "SoChiTiet"
            ws_l.append(["Ngày CT", "Số CT", "Diễn giải", "TK Nợ", "TK Có", "Phát sinh Nợ", "Phát sinh Có", "Đối tượng"])
            ws_l.append(["31/12/2024", "PC001", "Trả lương nhân viên năm 2024", "3341", "1121", 300_000_000, 0, "Nhân viên"])
            ws_l.append(["31/12/2024", "PK001", "Hạch toán chi phí lương 2024", "3341", "6421", 0, 330_000_000, "Nhân viên"])
            ws_l.append(["15/06/2024", "PC002", "Chi tiền thuê nhà cho chuyên gia nước ngoài", "6427", "1121", 120_000_000, 0, "Chuyên gia X"])
            ws_l.append(["20/08/2024", "PC003", "Chi hoa hồng môi giới cá nhân ngoài HĐLĐ", "6417", "1111", 15_000_000, 0, "Cá nhân Y"])
            ws_l.append(["25/12/2024", "PC004", "Chi thưởng lễ Tết đích danh từ quỹ phúc lợi", "3531", "1111", 20_000_000, 0, "Nhân viên"])
            wb_l.save(led_p)

            decl = load_pit_declaration(xml_p)
            self.assertEqual(decl.year, 2024)
            self.assertEqual(len(decl.rows_051), 1)
            self.assertEqual(len(decl.rows_052), 1)
            self.assertEqual(len(decl.rows_053), 1)

            tms_rows = parse_tms_npt_file(tms_p)
            npt_res, per_nnt = audit_dependents(decl.rows_053, tms_rows)
            self.assertEqual(npt_res[0].valid_months, 6)
            self.assertEqual(npt_res[0].disallowed_months, 6)
            self.assertEqual(per_nnt["8000000001"], (12, 6))

            led_sum = scan_accounting_ledgers([led_p])
            self.assertEqual(led_sum.tk334_debit_total, 300_000_000)
            self.assertEqual(led_sum.tk334_credit_total, 330_000_000)
            self.assertEqual(led_sum.tk353_debit_total, 20_000_000)
            self.assertEqual(len(led_sum.flagged_entries), 3)

            export_pit_audit_workbook(out_p, decl, npt_res, per_nnt, tms_rows, ledger_summary=led_sum)
            wb_out = openpyxl.load_workbook(out_p, data_only=False)
            self.assertEqual(
                wb_out.sheetnames,
                [
                    "TongHop_KiemTra_TNCN",
                    "TinhLai_05_1",
                    "TinhLai_05_2",
                    "DoiChieu_05_3",
                    "DS_NPT_TMS",
                    "DoiChieu_QuyLuong_334",
                    "QuetRuiRo_ChiPhi",
                ],
            )
            # Kiểm tra sheet Biên bản thuần đen-trắng (không tô màu nền ô)
            ws_sum = wb_out["TongHop_KiemTra_TNCN"]
            for rr in range(9, 17):
                for cc in range(1, 6):
                    self.assertIsNone(ws_sum.cell(rr, cc).fill.fill_type)
            # Kiểm tra TinhLai_05_1 liên kết SUMIF tự động từ DoiChieu_05_3
            self.assertIn("SUMIF(DoiChieu_05_3!", str(wb_out["TinhLai_05_1"]["M2"].value))

    def test_v11_multi_year_supplements_filter_and_direct_053_reconciliation(self) -> None:
        """Kiểm thử v1.1: Nạp nhiều năm + tờ khai bổ sung, lọc Từ năm..Đến năm, kế thừa 05-3, xuất QTT_TNCN_YYYY.xlsx chuẩn mẫu & đối chiếu trực tiếp trên 05-3."""
        from pit_workflow import export_yearly_tncn_with_audit
        from tax_parser import parse_tax_xml

        xml_2022 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung><TTinTKhaiThue>
      <TKhaiThue><maTKhai>953</maTKhai><tenTKhai>Tờ khai QTT TNCN</tenTKhai>
        <KyKKhaiThue><kyKKhai>2022</kyKKhai><kyKKhaiTuNgay>01/01/2022</kyKKhaiTuNgay><kyKKhaiDenNgay>31/12/2022</kyKKhaiDenNgay></KyKKhaiThue>
        <ngayLapTKhai>2023-03-20</ngayLapTKhai><soLan>0</soLan>
      </TKhaiThue>
      <NNT><mst>0100000001</mst><tenNNT>CTY A</tenNNT></NNT>
    </TTinTKhaiThue></TTinChung>
    <CTieuTKhaiChinh><ct21>1</ct21></CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>"""

        xml_2024_ct = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung><TTinTKhaiThue>
      <TKhaiThue><maTKhai>953</maTKhai><tenTKhai>Tờ khai QTT TNCN</tenTKhai>
        <KyKKhaiThue><kyKKhai>2024</kyKKhai><kyKKhaiTuNgay>01/01/2024</kyKKhaiTuNgay><kyKKhaiDenNgay>31/12/2024</kyKKhaiDenNgay></KyKKhaiThue>
        <ngayLapTKhai>2025-03-20</ngayLapTKhai><soLan>0</soLan>
      </TKhaiThue>
      <NNT><mst>0100000001</mst><tenNNT>CTY A</tenNNT></NNT>
    </TTinTKhaiThue></TTinChung>
    <CTieuTKhaiChinh><ct21>1</ct21></CTieuTKhaiChinh>
    <PLuc>
      <PLuc_05_3_BK_QTT>
        <BCaoCTietNPT><CTietNPT>
          <ct07>Nguyễn Văn A</ct07><ct08>8000000001</ct08><ct09>Nguyễn Bé Một</ct09>
          <ct10>01/01/2018</ct10><ct11>8900000001</ct11><ct12>CCCD</ct12><ct13>079218000001</ct13>
          <ct14>Con</ct14><ct15>01/2024</ct15><ct16>12/2024</ct16>
        </CTietNPT></BCaoCTietNPT>
      </PLuc_05_3_BK_QTT>
    </PLuc>
  </HSoKhaiThue>
</HSoThueDTu>"""

        # Tờ khai bổ sung lần 1 năm 2024 (không đính kèm 05-3 -> phải tự kế thừa 05-3 từ bản chính thức!)
        xml_2024_bs1 = """<?xml version="1.0" encoding="UTF-8"?>
<HSoThueDTu xmlns="http://kekhaithue.gdt.gov.vn/TKhaiThue">
  <HSoKhaiThue>
    <TTinChung><TTinTKhaiThue>
      <TKhaiThue><maTKhai>953</maTKhai><tenTKhai>Tờ khai QTT TNCN</tenTKhai>
        <KyKKhaiThue><kyKKhai>2024</kyKKhai><kyKKhaiTuNgay>01/01/2024</kyKKhaiTuNgay><kyKKhaiDenNgay>31/12/2024</kyKKhaiDenNgay></KyKKhaiThue>
        <ngayLapTKhai>2025-05-15</ngayLapTKhai><soLan>1</soLan>
      </TKhaiThue>
      <NNT><mst>0100000001</mst><tenNNT>CTY A</tenNNT></NNT>
    </TTinTKhaiThue></TTinChung>
    <CTieuTKhaiChinh><ct21>2</ct21></CTieuTKhaiChinh>
  </HSoKhaiThue>
</HSoThueDTu>"""

        tms_csv = (
            "19.11.2026 HIỂN THỊ DANH SÁCH ĐỘNG\n\nDanh sách kết quả\n"
            "Ngày tạo,Ngày TĐ,MST NNT,Tên NNT,MST NPT,Tên NPT,Tên trạng thái NPT,Ngày sinh,CMND,Quan hệ,Tháng,Tháng,Mẫu BK/TK,Mã trạng t\n"
            "10.01.2024,10.07.2024,8000000001,Nguyễn Văn A,8900000001,Nguyễn Bé Một,NNT đang hoạt động,01.01.2018,079218000001,Con,12024,62024,16-TH,00\n"
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            f22 = os.path.join(tmpdir, "2022.xml")
            f24_0 = os.path.join(tmpdir, "2024_ct.xml")
            f24_1 = os.path.join(tmpdir, "2024_bs1.xml")
            tms_p = os.path.join(tmpdir, "tms.csv")
            out_dir = os.path.join(tmpdir, "out")
            for p, txt in ((f22, xml_2022), (f24_0, xml_2024_ct), (f24_1, xml_2024_bs1), (tms_p, tms_csv)):
                with open(p, "w", encoding="utf-8") as f:
                    f.write(txt)

            decls = [parse_tax_xml(f22), parse_tax_xml(f24_0), parse_tax_xml(f24_1)]
            # Lọc chỉ lấy từ năm 2023 đến 2025 -> chỉ xuất QTT_TNCN_2024.xlsx (bỏ qua 2022)
            created, stats = export_yearly_tncn_with_audit(
                decls, out_dir, year_from=2023, year_to=2025, tms_npt_path=tms_p
            )
            self.assertEqual(len(created), 1)
            self.assertTrue(created[0].endswith("QTT_TNCN_2024.xlsx"))
            self.assertEqual(stats[2024]["total_rows"], 1)
            self.assertEqual(stats[2024]["total_diff_months"], 6)

            wb = openpyxl.load_workbook(created[0], data_only=False)
            # Đúng 4 sheet chuẩn mẫu TK, 05-1, 05-2, 05-3 + DS_NPT_TMS + DoiChieu_NPT
            for s in ("TK", "05-1", "05-2", "05-3", "DS_NPT_TMS", "DoiChieu_NPT"):
                self.assertIn(s, wb.sheetnames)
            # Đối chiếu trực tiếp trên sheet 05-3 (cột L..V, dòng 4)
            ws_053 = wb["05-3"]
            self.assertEqual(ws_053["L4"].value, 12)
            self.assertEqual(ws_053["M4"].value, 6)  # 2024 - 2018 = 6 tuổi
            self.assertIn("Hợp lệ", str(ws_053["N4"].value))
            self.assertEqual(ws_053["O4"].value, "Không trùng")
            self.assertEqual(ws_053["S4"].value, 6)
            self.assertEqual(ws_053["T4"].value, "=MAX(0,L4-S4)")
            self.assertIn("VƯỢT THỜI GIAN ĐĂNG KÝ", str(ws_053["V4"].value))

    def test_v12_dependent_age_and_duplicate_checks(self) -> None:
        """Kiểm thử v1.2: Kiểm tra độ tuổi NPT (con <18 / 18-22 / >22 / trước tháng sinh; cha mẹ trong/ngoài tuổi LĐ) và trùng NPT (cùng NNT, nhiều NNT trên 05-3, trùng trên TMS)."""
        from npt_tms_engine import TmsNptRow, audit_dependents
        from xml_05qtt_parser import Row053
        import datetime

        rows_053 = [
            # 1. Con 20 tuổi (2024 - 2004 = 20t) -> Cảnh báo KT thẻ SV
            Row053(1, "Nguyễn A", "8000000001", "Nguyễn Con Lớn", "15/05/2004", "8900000001", "CCCD", "079204000001", "Con", "01/2024", "12/2024"),
            # 2. Con sinh 08/2024 nhưng khai từ 01/2024 -> KHAI TRƯỚC THÁNG SINH
            Row053(2, "Nguyễn A", "8000000001", "Nguyễn Bé Mới", "10/08/2024", "8900000002", "GKS", "", "Con", "01/2024", "12/2024"),
            # 3. Cha 50 tuổi (sinh 1974) -> TRONG ĐỘ TUỔI LAO ĐỘNG
            Row053(3, "Trần B", "8000000002", "Trần Văn Cha", "01/01/1974", "8900000003", "CCCD", "079074000003", "Cha/Mẹ", "01/2024", "12/2024"),
            # 4. Mẹ 58 tuổi (sinh 1966, tên đệm Thị) -> Nữ ngoài tuổi LĐ 2024 (>= 56.3t) -> Hợp lệ độ tuổi
            Row053(4, "Trần B", "8000000002", "Lê Thị Mẹ", "01/01/1966", "8900000004", "CCCD", "079166000004", "Cha/Mẹ", "01/2024", "12/2024"),
            # 5 & 6. Trùng NPT giữa 2 NNT khác nhau trên 05-3 (vợ & chồng cùng khai 1 con 8900000005)
            Row053(5, "Lê Chồng", "8000000003", "Lê Bé Chung", "01/03/2019", "8900000005", "CCCD", "079219000005", "Con", "01/2024", "12/2024"),
            Row053(6, "Phạm Vợ", "8000000004", "Lê Bé Chung", "01/03/2019", "8900000005", "CCCD", "079219000005", "Con", "06/2024", "12/2024"),
            # 7. Trùng đăng ký với NNT khác trên TMS (8000000099 đang đăng ký 8900000007 trên TMS)
            Row053(7, "Hoàng E", "8000000005", "Hoàng Bé Bảy", "02/02/2020", "8900000007", "CCCD", "079220000007", "Con", "01/2024", "12/2024"),
        ]

        tms_rows = [
            TmsNptRow("01.01.2023", "01.01.2023", "8000000005", "Hoàng E", "8900000007", "Hoàng Bé Bảy", "Đang hoạt động", "00", "02.02.2020", "079220000007", "", "Con", (2023, 1), None, "16-TH", "", datetime.date(2023, 1, 1)),
            TmsNptRow("01.01.2024", "01.01.2024", "8000000099", "NNT Khác", "8900000007", "Hoàng Bé Bảy", "Đang hoạt động", "00", "02.02.2020", "079220000007", "", "Con", (2024, 1), None, "16-TH", "", datetime.date(2024, 1, 1)),
        ]

        res, _ = audit_dependents(rows_053, tms_rows, tax_year=2024)
        # Kiểm tra độ tuổi
        self.assertFalse(res[0].is_age_ok)
        self.assertIn("18–22 tuổi", res[0].age_status)
        self.assertFalse(res[1].is_age_ok)
        self.assertIn("KHAI TRƯỚC THÁNG SINH", res[1].age_status)
        self.assertFalse(res[2].is_age_ok)
        self.assertIn("TRONG ĐỘ TUỔI LAO ĐỘNG", res[2].age_status)
        self.assertTrue(res[3].is_age_ok)
        self.assertIn("Nữ ngoài độ tuổi lao động", res[3].age_status)

        # Kiểm tra trùng NPT trên 05-3 (rows 5 & 6)
        self.assertTrue(res[4].is_dup)
        self.assertIn("TRÙNG NHIỀU NNT TRÊN 05-3", res[4].dup_status)
        self.assertTrue(res[5].is_dup)
        self.assertIn("8000000003", res[5].dup_status)

        # Kiểm tra trùng NPT với NNT khác trên TMS (row 7)
        self.assertTrue(res[6].is_dup)
        self.assertIn("TRÙNG NNT KHÁC TRÊN TMS", res[6].dup_status)
        self.assertIn("8000000099", res[6].dup_status)

    def test_v13_multi_year_ledger_split(self) -> None:
        """Kiểm thử v1.3: Nạp nhiều file sổ chi tiết/bảng lương nhiều năm -> tự động tách đúng dữ liệu từng năm."""
        from ledger_payroll_scanner import scan_accounting_ledgers_by_year

        with tempfile.TemporaryDirectory() as tmpdir:
            f23 = os.path.join(tmpdir, "SoChiTiet_2023.xlsx")
            f24 = os.path.join(tmpdir, "SoChiTiet_2024.xlsx")

            wb23 = openpyxl.Workbook()
            ws23 = wb23.active
            ws23.append(["Ngày CT", "Số CT", "Diễn giải", "TK Nợ", "TK Có", "Phát sinh Nợ", "Phát sinh Có", "Đối tượng"])
            ws23.append(["20/06/2023", "PC23", "Chi tiền thuê nhà chuyên gia", "6427", "1121", 60_000_000, 0, "Chuyên gia A"])
            ws23.append(["31/12/2023", "PL23", "Chi lương nhân viên 2023", "3341", "1121", 200_000_000, 0, "NV"])
            wb23.save(f23)

            wb24 = openpyxl.Workbook()
            ws24 = wb24.active
            ws24.append(["Ngày CT", "Số CT", "Diễn giải", "TK Nợ", "TK Có", "Phát sinh Nợ", "Phát sinh Có", "Đối tượng"])
            ws24.append(["15/09/2024", "PC24", "Chi hoa hồng môi giới cá nhân", "6417", "1111", 12_000_000, 0, "Cá nhân B"])
            ws24.append(["31/12/2024", "PL24", "Chi lương nhân viên 2024", "3341", "1121", 350_000_000, 0, "NV"])
            wb24.save(f24)

            by_yr = scan_accounting_ledgers_by_year([f23, f24])
            self.assertIn(2023, by_yr)
            self.assertIn(2024, by_yr)
            self.assertEqual(by_yr[2023].tk334_debit_total, 200_000_000)
            self.assertEqual(by_yr[2024].tk334_debit_total, 350_000_000)
            self.assertEqual(len(by_yr[2023].flagged_entries), 1)
            self.assertEqual(by_yr[2023].flagged_entries[0].so_ct, "PC23")
            self.assertEqual(len(by_yr[2024].flagged_entries), 1)
            self.assertEqual(by_yr[2024].flagged_entries[0].so_ct, "PC24")

    def test_v13_legacy_xls_html_and_utf16_tsv_reader(self) -> None:
        """Kiểm thử đọc file .xls giả định dạng HTML và UTF-16 TSV kết xuất từ TMS (không bao giờ lỗi 'File is not a zip file')."""
        from npt_tms_engine import parse_tms_npt_file

        with tempfile.TemporaryDirectory() as tmpdir:
            # 1. File TMS dạng UTF-16LE Tab-delimited mang đuôi .xls
            p_utf16 = os.path.join(tmpdir, "tms_sap.xls")
            tsv_txt = (
                "19.11.2026\tHIỂN THỊ DANH SÁCH ĐỘNG\n\nDanh sách kết quả\n"
                "Ngày tạo\tNgày TĐ\tMST NNT\tTên NNT\tMST NPT\tTên NPT\tTên trạng thái NPT\tNgày sinh\tCMND\tQuan hệ\tTháng\tTháng\tMã trạng t\n"
                "01.01.2024\t01.01.2024\t8000000001\tNguyễn Văn A\t8900000001\tNguyễn Bé Một\tNNT đang hoạt động\t01.01.2018\t079218000001\tCon\t12024\t\t00\n"
            )
            with open(p_utf16, "wb") as f:
                f.write(tsv_txt.encode("utf-16"))
            recs1 = parse_tms_npt_file(p_utf16)
            self.assertEqual(len(recs1), 1)
            self.assertEqual(recs1[0].mst_nnt, "8000000001")
            self.assertEqual(recs1[0].tu_thang, (2024, 1))

            # 2. File TMS dạng HTML table mang đuôi .xls
            p_html = os.path.join(tmpdir, "tms_html.xls")
            html_txt = (
                "<html><body><table>"
                "<tr><th>Ngày tạo</th><th>MST NNT</th><th>MST NPT</th><th>Tên NPT</th><th>Tháng</th><th>Tháng</th><th>Mã trạng t</th></tr>"
                "<tr><td>01.02.2024</td><td>8000000002</td><td>8900000002</td><td>Trần Bé Hai</td><td>22024</td><td>62024</td><td>00</td></tr>"
                "</table></body></html>"
            )
            with open(p_html, "w", encoding="utf-8") as f:
                f.write(html_txt)
            recs2 = parse_tms_npt_file(p_html)
            self.assertEqual(len(recs2), 1)
            self.assertEqual(recs2[0].mst_npt, "8900000002")
            self.assertEqual(recs2[0].den_thang, (2024, 6))

    def test_v14_auto_pad_leading_zero_mst_reconciliation(self) -> None:
        """Kiểm thử tự động bù số 0 ở đầu MST khi đối chiếu NPT:
        - Xử lý triệt để trường hợp Excel làm mất số 0 ở đầu MST NNT hoặc MST NPT (9 số thành 10 số).
        - Đối chiếu khớp 100%, không nói người ta sai khi dữ liệu thực tế khớp.
        - Tự động sửa lại định dạng Text @ và bù số 0 trên sheet 05-3."""
        from npt_tms_engine import audit_dependents, parse_tms_npt_file
        from pit_workflow import reconcile_053_inplace
        from xml_05qtt_parser import Row053, clean_mst

        # 1. Kiểm tra hàm clean_mst trên nhiều định dạng
        self.assertEqual(clean_mst("101234567"), "0101234567")
        self.assertEqual(clean_mst(101234567), "0101234567")
        self.assertEqual(clean_mst("101234567.0"), "0101234567")
        self.assertEqual(clean_mst(101234567.0), "0101234567")
        self.assertEqual(clean_mst("0101234567"), "0101234567")
        self.assertEqual(clean_mst("8000000001"), "8000000001")
        self.assertEqual(clean_mst("890000001"), "0890000001")
        self.assertEqual(clean_mst("101234567-001"), "0101234567-001")
        self.assertEqual(clean_mst("101234567-1"), "0101234567-001")
        self.assertEqual(clean_mst("101234567001"), "0101234567001")
        self.assertEqual(clean_mst(None), "")
        self.assertEqual(clean_mst(""), "")
        self.assertEqual(clean_mst("0"), "")

        with tempfile.TemporaryDirectory() as tmpdir:
            # 2. File TMS bị thiếu số 0 ở đầu MST NNT (100000001 thay vì 0100000001) và MST NPT (890000001 thay vì 0890000001)
            tms_p = os.path.join(tmpdir, "tms_npt_missing_zero.csv")
            tms_csv = (
                "19.11.2026 HIỂN THỊ DANH SÁCH ĐỘNG\n\nDanh sách kết quả\n"
                "Ngày tạo,Ngày TĐ,MST NNT,Tên NNT,MST NPT,Tên NPT,Tên trạng thái NPT,Ngày sinh,CMND,Quan hệ,Tháng,Tháng,Mẫu BK/TK,Mã trạng t\n"
                "10.01.2024,10.07.2024,100000001,Nguyễn Văn A,890000001,Nguyễn Bé Một,NNT đang hoạt động,01.01.2018,079218000001,Con,12024,122024,16-TH,00\n"
            )
            with open(tms_p, "w", encoding="utf-8") as f:
                f.write(tms_csv)

            tms_rows = parse_tms_npt_file(tms_p)
            self.assertEqual(len(tms_rows), 1)
            # Tự động bù số 0 khi parse
            self.assertEqual(tms_rows[0].mst_nnt, "0100000001")
            self.assertEqual(tms_rows[0].mst_npt, "0890000001")

            # 3. Danh sách 05-3: Tờ khai đầy đủ 10 số (0100000001 & 0890000001) đối chiếu với TMS ban đầu bị thiếu số 0
            row_053_a = Row053(
                stt=1,
                ho_ten_nnt="Nguyễn Văn A",
                mst_nnt="0100000001",
                ho_ten_npt="Nguyễn Bé Một",
                ngay_sinh="01/01/2018",
                mst_npt="0890000001",
                loai_gt="CCCD",
                so_gt="079218000001",
                quan_he="Con",
                tu_thang="01/2024",
                den_thang="12/2024",
            )
            res_a, per_a = audit_dependents([row_053_a], tms_rows, tax_year=2024)
            self.assertTrue(res_a[0].is_ok, f"Đối chiếu phải KHỚP nhưng bị báo lỗi: {res_a[0].status}")
            self.assertEqual(res_a[0].status, "KHỚP")
            self.assertEqual(res_a[0].valid_months, 12)
            self.assertEqual(res_a[0].disallowed_months, 0)
            self.assertIn("0100000001", per_a)

            # 4. Ngược lại: Bảng kê 05-3 bị thiếu số 0 (hoặc dạng số nguyên) đối chiếu với TMS có đủ số 0
            row_053_b = Row053(
                stt=1,
                ho_ten_nnt="Nguyễn Văn A",
                mst_nnt="100000001",  # Thiếu số 0
                ho_ten_npt="Nguyễn Bé Một",
                ngay_sinh="01/01/2018",
                mst_npt="890000001",   # Thiếu số 0
                loai_gt="CCCD",
                so_gt="079218000001",
                quan_he="Con",
                tu_thang="01/2024",
                den_thang="12/2024",
            )
            res_b, per_b = audit_dependents([row_053_b], tms_rows, tax_year=2024)
            self.assertTrue(res_b[0].is_ok, f"05-3 thiếu số 0 vẫn phải KHỚP: {res_b[0].status}")
            self.assertEqual(res_b[0].status, "KHỚP")
            self.assertEqual(row_053_b.mst_nnt, "0100000001")
            self.assertEqual(row_053_b.mst_npt, "0890000001")

            # 5. Đối chiếu trực tiếp trên workbook Excel (sheet 05-3 chứa số nguyên mất số 0)
            wb = openpyxl.Workbook()
            ws = wb.active
            ws.title = "05-3"
            ws.append(["STT", "Họ tên NNT", "MST NNT", "Họ tên NPT", "Ngày sinh", "MST NPT", "Loại GT", "Số GT", "Quan hệ", "Từ tháng", "Đến tháng"])
            ws.append(["1", "2", "3", "4", "5", "6", "7", "8", "9", "10", "11"])
            ws.append(["[1]", "[2]", "[3]", "[4]", "[5]", "[6]", "[7]", "[8]", "[9]", "[10]", "[11]"])
            # Dòng 4: MST NNT và MST NPT lưu dạng số (int) mất số 0
            ws.append([1, "Nguyễn Văn A", 100000001, "Nguyễn Bé Một", "01/01/2018", 890000001, "CCCD", "079218000001", "Con", "01/2024", "12/2024"])

            st = reconcile_053_inplace(wb, tms_rows, tax_year=2024)
            self.assertEqual(st["matched"], 1)
            self.assertEqual(st["total_diff_months"], 0)
            # Ô MST trên sheet 05-3 phải được tự động sửa lại đủ 10 số và format Text @
            self.assertEqual(ws.cell(4, 3).value, "0100000001")
            self.assertEqual(ws.cell(4, 3).number_format, "@")
            self.assertEqual(ws.cell(4, 6).value, "0890000001")
            self.assertEqual(ws.cell(4, 6).number_format, "@")
            # Cột V (kết quả đối chiếu tổng hợp) phải là KHỚP
            self.assertEqual(ws.cell(4, 22).value, "KHỚP")

    def test_tms_multi_row_history_merge_and_reversed_dates(self) -> None:
        """
        Kiểm thử hợp nhất lịch sử nhiều dòng trên TMS & xử lý Từ tháng > Đến tháng bị ngược:
        1. Nhiều dòng qua các năm (có dòng trống Đến tháng xen kẽ) kết thúc ở 12/2023 -> hiểu 01/2018..12/2023.
        2. Nhiều dòng qua các năm đều trống Đến tháng, chỉ dòng cuối 01/2023..12/2023 -> hiểu 01/2018..12/2023.
        3. Dòng bị ngược Từ tháng > Đến tháng (01/2025 -> 01/2022) -> tự đảo thành 01/2022..01/2025.
        4. Hai giai đoạn gián đoạn (01/2018..12/2018 và 01/2021..12/2022) -> không tự nối qua 2019-2020.
        """
        import datetime
        from npt_tms_engine import TmsNptRow, audit_dependents
        from xml_05qtt_parser import Row053

        def _mk_tms(mst_nnt: str, mst_npt: str, f, t, sd=(2020, 1, 1)) -> TmsNptRow:
            return TmsNptRow(
                ngay_tao="", ngay_td="", mst_nnt=mst_nnt, ten_nnt="NNT Test",
                mst_npt=mst_npt, ten_npt="NPT Test", trang_thai="Đang hoạt động", ma_tt="00",
                ngay_sinh="01.01.1955", cmnd="111222333", ho_chieu="", quan_he="Cha/Mẹ",
                tu_thang=f, den_thang=t, mau_tk="02TH", gks_so="", sort_date=datetime.date(*sd),
            )

        # Case 1: 01/2018-12/2018, 01/2019-None, 01/2020-12/2020, ..., 01/2023-12/2023
        tms_case1 = [
            _mk_tms("8000000101", "8900000101", (2018, 1), (2018, 12), (2018, 1, 1)),
            _mk_tms("8000000101", "8900000101", (2019, 1), None, (2019, 1, 1)),
            _mk_tms("8000000101", "8900000101", (2020, 1), (2020, 12), (2020, 1, 1)),
            _mk_tms("8000000101", "8900000101", (2021, 1), (2021, 12), (2021, 1, 1)),
            _mk_tms("8000000101", "8900000101", (2022, 1), (2022, 12), (2022, 1, 1)),
            _mk_tms("8000000101", "8900000101", (2023, 1), (2023, 12), (2023, 1, 1)),
        ]
        for yr in (2018, 2019, 2020, 2021, 2022, 2023):
            r053 = Row053(1, "NNT Test", "8000000101", "NPT Test", "01/01/1955", "8900000101", "CMND", "111222333", "Cha/Mẹ", f"01/{yr}", f"12/{yr}")
            res, _ = audit_dependents([r053], tms_case1, tax_year=yr)
            self.assertTrue(res[0].is_ok, f"Year {yr} should be OK, got {res[0].status}")
            self.assertEqual(res[0].tms_from, "01/2018")
            self.assertEqual(res[0].tms_to, "12/2023")
            self.assertEqual(res[0].valid_months, 12)
        # 2024 phải bị cắt (ngoài thời gian đăng ký TMS)
        r053_2024 = Row053(1, "NNT Test", "8000000101", "NPT Test", "01/01/1955", "8900000101", "CMND", "111222333", "Cha/Mẹ", "01/2024", "12/2024")
        res_2024, _ = audit_dependents([r053_2024], tms_case1, tax_year=2024)
        self.assertFalse(res_2024[0].is_ok)
        self.assertEqual(res_2024[0].valid_months, 0)
        self.assertIn("NGOÀI THỜI GIAN ĐĂNG KÝ TMS", res_2024[0].status)

        # Case 2: 01/2018-None .. 01/2022-None, 01/2023-12/2023 -> 01/2018..12/2023
        tms_case2 = [
            _mk_tms("8000000102", "8900000102", (2018, 1), None, (2018, 1, 1)),
            _mk_tms("8000000102", "8900000102", (2019, 1), None, (2019, 1, 1)),
            _mk_tms("8000000102", "8900000102", (2020, 1), None, (2020, 1, 1)),
            _mk_tms("8000000102", "8900000102", (2021, 1), None, (2021, 1, 1)),
            _mk_tms("8000000102", "8900000102", (2022, 1), None, (2022, 1, 1)),
            _mk_tms("8000000102", "8900000102", (2023, 1), (2023, 12), (2023, 1, 1)),
        ]
        r053_c2_2020 = Row053(1, "NNT Test", "8000000102", "NPT Test", "01/01/1955", "8900000102", "CMND", "111222334", "Cha/Mẹ", "01/2020", "12/2020")
        res_c2, _ = audit_dependents([r053_c2_2020], tms_case2, tax_year=2020)
        self.assertTrue(res_c2[0].is_ok)
        self.assertEqual(res_c2[0].tms_from, "01/2018")
        self.assertEqual(res_c2[0].tms_to, "12/2023")
        r053_c2_2024 = Row053(1, "NNT Test", "8000000102", "NPT Test", "01/01/1955", "8900000102", "CMND", "111222334", "Cha/Mẹ", "01/2024", "12/2024")
        res_c2_24, _ = audit_dependents([r053_c2_2024], tms_case2, tax_year=2024)
        self.assertEqual(res_c2_24[0].valid_months, 0)

        # Case 3: Bị ngược Từ tháng > Đến tháng (01/2025 -> 01/2022 => hiểu là 01/2022 -> 01/2025)
        tms_case3 = [_mk_tms("8000000103", "8900000103", (2025, 1), (2022, 1), (2025, 1, 1))]
        r053_c3_2023 = Row053(1, "NNT Test", "8000000103", "NPT Test", "01/01/1955", "8900000103", "CMND", "111222335", "Cha/Mẹ", "01/2023", "12/2023")
        res_c3_23, _ = audit_dependents([r053_c3_2023], tms_case3, tax_year=2023)
        self.assertTrue(res_c3_23[0].is_ok)
        self.assertEqual(res_c3_23[0].tms_from, "01/2022")
        self.assertEqual(res_c3_23[0].tms_to, "01/2025")
        self.assertEqual(res_c3_23[0].valid_months, 12)
        r053_c3_2025 = Row053(1, "NNT Test", "8000000103", "NPT Test", "01/01/1955", "8900000103", "CMND", "111222335", "Cha/Mẹ", "01/2025", "12/2025")
        res_c3_25, _ = audit_dependents([r053_c3_2025], tms_case3, tax_year=2025)
        self.assertEqual(res_c3_25[0].valid_months, 1)
        self.assertEqual(res_c3_25[0].disallowed_months, 11)

        # Case 4 (VD 1): ĐK 01/2023..01/2024 và ĐK 01/2025..(trống) -> tính xuyên suốt từ 01/2023 chưa có thời hạn kết thúc!
        tms_case4 = [
            _mk_tms("8000000104", "8900000104", (2023, 1), (2024, 1), (2023, 1, 1)),
            _mk_tms("8000000104", "8900000104", (2025, 1), None, (2025, 1, 1)),
        ]
        for yr in (2023, 2024, 2025):
            r053_c4 = Row053(1, "NNT Test", "8000000104", "NPT Test", "01/01/1955", "8900000104", "CMND", "111222336", "Cha/Mẹ", f"01/{yr}", f"12/{yr}")
            res_c4, _ = audit_dependents([r053_c4], tms_case4, tax_year=yr)
            self.assertTrue(res_c4[0].is_ok, f"Year {yr} should be OK, got {res_c4[0].status}")
            self.assertEqual(res_c4[0].tms_from, "01/2023")
            self.assertEqual(res_c4[0].tms_to, "Còn hiệu lực")
            self.assertEqual(res_c4[0].valid_months, 12)

        # Case 5 (VD 2): ĐK 01/2022..(trống) và ĐK 01/2023..12/2023 -> tính 01/2022 đến 12/2023 (đã có kỳ kết thúc 12/2023)
        tms_case5 = [
            _mk_tms("8000000105", "8900000105", (2022, 1), None, (2022, 1, 1)),
            _mk_tms("8000000105", "8900000105", (2023, 1), (2023, 12), (2023, 1, 1)),
        ]
        for yr in (2022, 2023):
            r053_c5 = Row053(1, "NNT Test", "8000000105", "NPT Test", "01/01/1955", "8900000105", "CMND", "111222337", "Cha/Mẹ", f"01/{yr}", f"12/{yr}")
            res_c5, _ = audit_dependents([r053_c5], tms_case5, tax_year=yr)
            self.assertTrue(res_c5[0].is_ok)
            self.assertEqual(res_c5[0].tms_from, "01/2022")
            self.assertEqual(res_c5[0].tms_to, "12/2023")
            self.assertEqual(res_c5[0].valid_months, 12)
        r053_c5_24 = Row053(1, "NNT Test", "8000000105", "NPT Test", "01/01/1955", "8900000105", "CMND", "111222337", "Cha/Mẹ", "01/2024", "12/2024")
        res_c5_24, _ = audit_dependents([r053_c5_24], tms_case5, tax_year=2024)
        self.assertEqual(res_c5_24[0].valid_months, 0)

    def test_quick_test_formulas_on_051_and_052(self) -> None:
        """Kiểm thử gắn công thức test nhanh rủi ro khấu trừ / tính sai thuế TNCN trên 05-1 (X..AD) và 05-2 (M..Q)."""
        from pit_workflow import add_quick_test_formulas_051_052

        wb = openpyxl.Workbook()
        ws1 = wb.active
        ws1.title = "05-1"
        ws1.cell(4, 1, 1)
        ws1.cell(4, 2, "Nguyễn Văn A")
        ws1.cell(4, 3, "8000000001")
        ws1.cell(4, 6, "[x]")
        ws1.cell(4, 8, 250_000_000)   # Tổng TNCT (H)
        ws1.cell(4, 12, 1)            # Số NPT (L)
        ws1.cell(4, 13, 184_800_000)  # GTGC (M)
        ws1.cell(4, 15, 20_000_000)   # BH (O)
        ws1.cell(4, 17, 45_200_000)   # TNTT (Q)
        ws1.cell(4, 18, 2_260_000)    # Đã KT (R)
        ws1.cell(4, 20, 2_260_000)    # Phải nộp (T)
        ws1.cell(5, 1, "TỔNG CỘNG")

        ws2 = wb.create_sheet("05-2")
        ws2.cell(4, 1, 1)
        ws2.cell(4, 2, "Trần CTV B")
        ws2.cell(4, 3, "")            # Chưa có MST -> không đủ đk Cam kết 08
        ws2.cell(4, 6, "[]")          # Cư trú 10%
        ws2.cell(4, 7, 15_000_000)    # Tổng TNCT (G)
        ws2.cell(4, 11, 0)            # Đã KT = 0 (K)
        ws2.cell(5, 1, "TỔNG CỘNG")

        cnt = add_quick_test_formulas_051_052(wb, tax_year=2024)
        self.assertEqual(cnt["rows_051"], 1)
        self.assertEqual(cnt["rows_052"], 1)
        # 05-1: X4 (TNTT test), Z4 (7 bậc <=2025), AC4 (GTGC), AD4 (cảnh báo)
        self.assertEqual(ws1["X4"].value, "=MAX(0,H4-J4-K4-M4-N4-O4-P4)")
        self.assertIn("60000000", str(ws1["Z4"].value))
        self.assertIn("132000000", str(ws1["AC4"].value))
        self.assertIn("TÍNH SAI THUẾ", str(ws1["AD4"].value))
        self.assertEqual(ws1["X5"].value, "=SUM(X4:X4)")
        # 05-2: M4 (10%/20%), N4 (Thuế test), O4 (Thiếu), Q4 (Cảnh báo MST/Cam kết 08)
        self.assertEqual(ws2["N4"].value, "=ROUND(MAX(0,G4-I4-J4)*M4,0)")
        self.assertEqual(ws2["O4"].value, "=MAX(0,N4-K4)")
        self.assertIn("CHƯA CÓ MST", str(ws2["Q4"].value))
        self.assertEqual(ws2["O5"].value, "=SUM(O4:O4)")

        # Kiểm tra năm 2026 tự động chuyển sang biểu 5 bậc & GTBT 186tr
        add_quick_test_formulas_051_052(wb, tax_year=2026)
        self.assertIn("120000000", str(ws1["Z4"].value))
        self.assertIn("186000000", str(ws1["AC4"].value))

    def test_v14_extra_tms_folder_checks_multi_income_hkd_and_npt_income(self) -> None:
        """Kiểm thử v1.4: Nạp thư mục Thu nhập nhiều nơi, Danh bạ NPT (HKD), DS NPT có thu nhập -> cảnh báo trực tiếp trên 05-1, 05-2, 05-3."""
        from pit_workflow import apply_extra_tms_checks
        from tms_extra_engine import (
            parse_multi_income_files,
            parse_npt_directory_files,
            parse_npt_income_files,
        )

        with tempfile.TemporaryDirectory() as tmpdir:
            d_mi = os.path.join(tmpdir, "tn_nhieu_noi")
            d_dir = os.path.join(tmpdir, "danh_ba_npt")
            d_inc = os.path.join(tmpdir, "npt_co_tn")
            os.makedirs(d_mi)
            os.makedirs(d_dir)
            os.makedirs(d_inc)

            # 1. File Thu nhập nhiều nơi (Excel)
            wb_mi = openpyxl.Workbook()
            ws_mi = wb_mi.active
            ws_mi.append(["STT", "MST cá nhân", "Họ và tên", "MST tổ chức chi trả", "Tên tổ chức", "Bảng kê", "Năm", "Tổng TNCT", "Thuế đã khấu trừ"])
            ws_mi.append([1, "8000000001", "Nguyễn Văn A", "0109999999", "CTY Khác 1", "05-1", 2024, 150_000_000, 5_000_000])
            ws_mi.append([2, "8000000002", "Trần CTV B", "0108888888", "CTY Khác 2", "05-2", 2024, 30_000_000, 3_000_000])
            wb_mi.save(os.path.join(d_mi, "TN_NhieuNoi_2024.xlsx"))

            # 2. File Danh bạ NPT (HKD đang hoạt động)
            wb_dir = openpyxl.Workbook()
            ws_dir = wb_dir.active
            ws_dir.append(["STT", "MST NPT", "Họ tên NPT", "Loại NNT", "Trạng thái", "Ngày cấp"])
            ws_dir.append([1, "8900000011", "Lê Văn HKD", "Hộ kinh doanh", "Đang hoạt động", "01/01/2020"])
            wb_dir.save(os.path.join(d_dir, "DanhBa_NPT.xlsx"))

            # 3. File DS NPT có thu nhập (> 12tr/năm 2024)
            wb_inc = openpyxl.Workbook()
            ws_inc = wb_inc.active
            ws_inc.append(["STT", "MST NPT", "Họ tên NPT", "Năm", "Tổng thu nhập"])
            ws_inc.append([1, "8900000012", "Phạm Thị Thu Nhập", 2024, 48_000_000])
            wb_inc.save(os.path.join(d_inc, "DS_NPT_CoTN_2024.xlsx"))

            mi_recs = parse_multi_income_files([d_mi])
            dir_recs = parse_npt_directory_files([d_dir])
            inc_recs = parse_npt_income_files([d_inc])
            self.assertEqual(len(mi_recs), 2)
            self.assertEqual(len(dir_recs), 1)
            self.assertEqual(len(inc_recs), 1)

            # Workbook QTT mẫu có 05-1, 05-2, 05-3
            wb = openpyxl.Workbook()
            ws1 = wb.active
            ws1.title = "05-1"
            ws1.cell(4, 1, 1)
            ws1.cell(4, 2, "Nguyễn Văn A")
            ws1.cell(4, 3, "8000000001")
            ws1.cell(4, 6, "[x]")  # Có ủy quyền QT -> phải bị cảnh báo!

            ws2 = wb.create_sheet("05-2")
            ws2.cell(4, 1, 1)
            ws2.cell(4, 2, "Trần CTV B")
            ws2.cell(4, 3, "8000000002")
            ws2.cell(4, 6, "[]")
            ws2.cell(4, 7, 10_000_000)
            ws2.cell(4, 11, 0)  # Chưa khấu trừ 10% nhưng có TN nơi khác -> phải bị cảnh báo!

            ws3 = wb.create_sheet("05-3")
            ws3.cell(4, 1, 1)
            ws3.cell(4, 2, "NNT 1")
            ws3.cell(4, 4, "Lê Văn HKD")
            ws3.cell(4, 6, "8900000011")
            ws3.cell(4, 10, "01/2024")
            ws3.cell(4, 11, "12/2024")

            ws3.cell(5, 1, 2)
            ws3.cell(5, 2, "NNT 2")
            ws3.cell(5, 4, "Phạm Thị Thu Nhập")
            ws3.cell(5, 6, "8900000012")
            ws3.cell(5, 10, "01/2024")
            ws3.cell(5, 11, "12/2024")

            st = apply_extra_tms_checks(
                wb,
                tax_year=2024,
                company_mst="0100000001",
                multi_income_records=mi_recs,
                npt_dir_records=dir_recs,
                npt_income_records=inc_recs,
            )
            self.assertEqual(st["uq_warnings_051"], 1)
            self.assertEqual(st["kt10_warnings_052"], 1)
            self.assertEqual(st["hkd_warnings_053"], 1)
            self.assertEqual(st["npt_inc_warnings_053"], 1)
            self.assertIn("KHÔNG ĐỦ ĐK UQ", str(ws1["AF4"].value))
            self.assertIn("PHẢI KHẤU TRỪ 10%", str(ws2["S4"].value))
            self.assertIn("HỘ KINH DOANH", str(ws3["W4"].value))
            self.assertIn("KHÔNG ĐỦ ĐK GIẢM TRỪ", str(ws3["X5"].value))


if __name__ == "__main__":
    unittest.main()
