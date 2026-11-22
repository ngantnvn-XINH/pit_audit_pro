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
            # Đối chiếu trực tiếp trên sheet 05-3 (cột L..S, dòng 4)
            ws_053 = wb["05-3"]
            self.assertEqual(ws_053["L4"].value, 12)
            self.assertEqual(ws_053["P4"].value, 6)
            self.assertEqual(ws_053["Q4"].value, "=MAX(0,L4-P4)")
            self.assertIn("VƯỢT THỜI GIAN ĐĂNG KÝ", str(ws_053["S4"].value))


if __name__ == "__main__":
    unittest.main()
