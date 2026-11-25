# PIT Audit Pro — Phần Mềm Đối Chiếu & Tính Thuế TNCN Chuyên Sâu (v1.1)

Ứng dụng Desktop Windows độc lập chuyên dụng cho công tác thanh tra, kiểm tra, đối chiếu và tính lại **Thuế Thu nhập cá nhân (TNCN)** theo từng thời kỳ pháp luật Việt Nam (2013 – 2026+).

## Tính năng cốt lõi (v1.1)
- **Nạp hàng loạt tờ khai XML nhiều năm & nhiều lần bổ sung:** Quét cả thư mục con, tự động lọc tờ khai QTT TNCN (`05/QTT-TNCN` cả TT92 & TT80), lọc theo khoảng **`Từ năm ... đến năm ...`**, tự động chốt bản bổ sung cuối cùng mỗi năm và kế thừa phụ lục `05-3` từ lần kê khai trước.
- **Xuất từng năm ra file `QTT_TNCN_YYYY.xlsx` chuẩn 100% mẫu `tax_xml_to_excel` (`TK`, `05-1`, `05-2`, `05-3`)** và **đối chiếu trực tiếp Người phụ thuộc TMS ngay trên các cột `L..S` của sheet `05-3`** (kèm sheet tổng hợp lỗi `DoiChieu_NPT` & dữ liệu sạch `DS_NPT_TMS`).
- **Động cơ chính sách đa thời kỳ (`pit_tax_rules.py`):**
  - Tự động áp dụng mức giảm trừ gia cảnh & biểu thuế lũy tiến theo năm:
    - `2013–2019`: Bản thân `9.000.000 đ/tháng`, NPT `3.600.000 đ/tháng`, biểu lũy tiến **7 bậc**.
    - `2020–2025` (NQ 954/2020/UBTVQH14): Bản thân `11.000.000 đ/tháng`, NPT `4.400.000 đ/tháng`, biểu lũy tiến **7 bậc**.
    - `Từ 2026` (Luật Thuế TNCN 109/2025/QH15 & NQ 110/2025/UBTVQH15): Bản thân `15.500.000 đ/tháng`, NPT `6.200.000 đ/tháng`, biểu lũy tiến **5 bậc** (`5% / 10% / 20% / 30% / 35%`).
  - Hỗ trợ thuế suất toàn phần `10%` (không HĐLĐ ≥ 2 triệu/lần), `20%` (không cư trú), trần tiền thuê nhà `15%` (TT 92/2015).
- **Đối chiếu Người phụ thuộc (`05-3` vs TMS) liên thông tính lại `05-1`:**
  - Đọc file kết xuất TMS (`.xlsx` / `.csv`), hợp nhất lịch sử báo tăng/báo giảm, hỗ trợ cả MST 10 số & Số định danh 12 số (TT 86/2024/TT-BTC).
  - **Tự động bù số 0 đứng đầu MST (v1.1+):** Khắc phục triệt để lỗi người dùng mở file NPT bằng Excel làm mất số 0 ở đầu (9 số thành 10 số, 12 số thành 13 số, MST chi nhánh có gạch nối). Hệ thống tự động bù số 0 trước khi so sánh đa chiều, tự động định dạng chuẩn Text `@` trên sheet 05-3, đảm bảo đối chiếu khớp 100% đúng thực tế, không kết luận sai khi người nộp thuế thực tế đã đăng ký hợp lệ.
  - Tự động đẩy **số tháng NPT bị loại trừ** của từng NNT bằng công thức `SUMIF` sang sheet `TinhLai_05_1` để **tính lại đúng số thuế TNCN phải nộp thêm theo biểu lũy tiến** của chính cá nhân đó.
- **Đối chiếu Quỹ lương (`TK 334`, `338`, `3335`, `353`) & Quét rủi ro Sổ chi phí (`TK 641, 642, 622, 627, 154, 353, 811`):**
  - Tự động phát hiện 6 nhóm khoản chi có nguy cơ chịu thuế TNCN bỏ sót ngoài bảng lương (tiền thuê nhà, thưởng/phúc lợi đích danh, BH nhân thọ/thẻ hội viên, vé máy bay/học phí, phụ cấp vượt định mức, chi cá nhân ngoài HĐLĐ ≥ 2 triệu chưa khấu trừ 10%).
- **Xuất báo cáo Excel 7 sheet cài sẵn 100% công thức động** (kèm sheet `TongHop_KiemTra_TNCN` thuần đen – trắng sẵn sàng copy thẳng vào Biên bản Word).

## Cài đặt & Chạy trên máy mới
```bash
pip install -r requirements.txt
python app.py
```
Hoặc nhấp đúp `run.bat` trên Windows.

## Chạy kiểm thử tự động
```bash
python test_pit_audit.py
```
