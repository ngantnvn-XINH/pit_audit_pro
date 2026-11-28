# -*- coding: utf-8 -*-
"""
Động cơ đọc & đối chiếu 3 nguồn dữ liệu mở rộng từ TMS cho PIT Audit Pro:
1. Thu nhập nhiều nơi -> Đối chiếu trực tiếp 05-1 (cảnh báo không đủ điều kiện Ủy quyền QT)
   & 05-2 (cảnh báo thuộc diện PHẢI khấu trừ 10% / không đủ điều kiện Cam kết 08/CK-TNCN).
2. Danh bạ Người phụ thuộc -> Đối chiếu 05-3 cảnh báo NPT là Hộ kinh doanh đang hoạt động
   trong thời gian giảm trừ (không hợp lệ).
3. Danh sách Người phụ thuộc có thu nhập -> Đối chiếu 05-3 cảnh báo NPT có thu nhập vượt ngưỡng
   (> 1tr/tháng [12tr/năm] cho <= 2025; > 3tr/tháng [36tr/năm] từ 2026) -> không đủ ĐK giảm trừ.

Hỗ trợ đọc mọi định dạng bảng phổ biến từ TMS: .xlsx, .xls (kể cả HTML/UTF-16 TSV), .csv.
"""

from __future__ import annotations

import datetime
import os
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Sequence, Tuple

from npt_tms_engine import parse_ym, ym_idx
from tabular_reader import read_first_sheet_rows
from xml_05qtt_parser import clean_mst


def _clean_text(val: Any) -> str:
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val).replace("\xa0", " ").replace("\u200b", "")).strip()


def _clean_id(val: Any) -> str:
    s = _clean_text(val)
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s.replace(" ", "").upper()


def _to_float(val: Any) -> float:
    if val is None or val == "":
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = _clean_text(val)
    if not s:
        return 0.0
    # Xử lý định dạng số VN/quốc tế: '12,000,000' hoặc '12.000.000'
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if all(len(p) == 3 for p in parts[1:]):
            s = "".join(parts)
        else:
            s = s.replace(",", ".")
    elif "." in s:
        parts = s.split(".")
        if len(parts) > 2 or (len(parts) == 2 and len(parts[1]) == 3 and parts[0].lstrip("-").isdigit()):
            s = "".join(parts)
    try:
        return float(s)
    except ValueError:
        return 0.0


def _extract_year_from_val_or_path(val: Any, file_path: str) -> Optional[int]:
    s = _clean_text(val)
    m = re.search(r"\b(20\d{2})\b", s)
    if m:
        return int(m.group(1))
    m_fn = re.search(r"(20\d{2})", os.path.basename(file_path))
    if m_fn:
        return int(m_fn.group(1))
    return None


def _parse_date_any(val: Any) -> Optional[datetime.date]:
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = _clean_text(val)
    if not s:
        return None
    for fmt in ("%d/%m/%Y", "%d.%m.%Y", "%d-%m-%Y", "%Y-%m-%d", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(s[:10], fmt).date()
        except ValueError:
            pass
    ym = parse_ym(s)
    if ym:
        return datetime.date(ym[0], ym[1], 1)
    return None


def _collect_tabular_paths(paths: Sequence[str]) -> List[str]:
    out: List[str] = []
    seen = set()
    for p in paths:
        if not p:
            continue
        if os.path.isdir(p):
            for r_dir, _, files in os.walk(p):
                for fn in files:
                    if fn.lower().endswith((".xlsx", ".xls", ".csv")) and not fn.startswith("~$"):
                        fp = os.path.normpath(os.path.join(r_dir, fn))
                        if fp not in seen:
                            seen.add(fp)
                            out.append(fp)
        elif os.path.isfile(p) and p.lower().endswith((".xlsx", ".xls", ".csv")) and not os.path.basename(p).startswith("~$"):
            fp = os.path.normpath(p)
            if fp not in seen:
                seen.add(fp)
                out.append(fp)
    return out


# =============================================================================
# 1. THU NHẬP NHIỀU NƠI (ĐỐI CHIẾU 05-1 ỦY QUYỀN QT & 05-2 KHẤU TRỪ 10%)
# =============================================================================
@dataclass
class MultiIncomeRecord:
    mst_nnt: str
    cccd: str
    ho_ten: str
    mst_tcct: str
    ten_tcct: str
    loai_bk: str        # "05-1", "05-2", hoặc ""
    year: Optional[int]
    tnct: float
    thue_kt: float
    source_file: str


def parse_multi_income_files(paths: Sequence[str]) -> List[MultiIncomeRecord]:
    records: List[MultiIncomeRecord] = []
    for fp in _collect_tabular_paths(paths):
        try:
            raw_rows = read_first_sheet_rows(fp)
        except Exception:
            continue
        if not raw_rows:
            continue

        hdr_idx: Optional[int] = None
        cmap: Dict[str, int] = {}
        for idx, row in enumerate(raw_rows[:30]):
            cells = [_clean_text(c).lower() for c in row]
            has_mst = any("mst" in c or "mã số thuế" in c or "số định danh" in c or "cccd" in c for c in cells)
            has_inc = any(
                k in c
                for c in cells
                for k in ("thu nhập", "tnct", "chi trả", "tổ chức", "khấu trừ", "bảng kê", "05-1", "05-2")
            )
            if has_mst and has_inc:
                hdr_idx = idx
                for ci, h in enumerate(cells):
                    if any(k in h for k in ("mst tổ chức", "mst đơn vị", "mst chi trả", "mst cơ quan", "mst tcct", "mst cty", "mã số thuế tổ chức", "mã số thuế đơn vị")):
                        cmap.setdefault("mst_tcct", ci)
                    elif any(k in h for k in ("tên tổ chức", "tên đơn vị", "tên cơ quan", "nơi chi trả", "tên tcct", "đơn vị chi trả")):
                        cmap.setdefault("ten_tcct", ci)
                    elif any(k in h for k in ("mst cá nhân", "mst nnt", "mã số thuế cá nhân", "mã số thuế nnt")) or h in ("mst", "mã số thuế"):
                        cmap.setdefault("mst_nnt", ci)
                    elif any(k in h for k in ("cccd", "cmnd", "định danh", "số giấy tờ")):
                        cmap.setdefault("cccd", ci)
                    elif any(k in h for k in ("họ và tên", "họ tên", "tên cá nhân", "tên nnt")):
                        cmap.setdefault("ho_ten", ci)
                    elif any(k in h for k in ("bảng kê", "mẫu bk", "loại bk", "mẫu tk", "loại thu nhập", "biểu thuế")):
                        cmap.setdefault("loai_bk", ci)
                    elif any(k in h for k in ("năm", "kỳ tính thuế", "kỳ quyết toán", "kỳ kk")):
                        cmap.setdefault("nam", ci)
                    elif any(k in h for k in ("thuế đã khấu trừ", "số thuế đã khấu trừ", "thuế khấu trừ", "thuế tncn")):
                        cmap.setdefault("thue_kt", ci)
                    elif any(k in h for k in ("tổng tnct", "thu nhập chịu thuế", "tổng thu nhập", "tnct", "thu nhập")):
                        cmap.setdefault("tnct", ci)
                break

        if hdr_idx is None:
            continue

        def _g(row: Sequence[Any], k: str) -> Any:
            ci = cmap.get(k)
            return row[ci] if (ci is not None and ci < len(row)) else None

        for row in raw_rows[hdr_idx + 1:]:
            mst_nnt = clean_mst(_g(row, "mst_nnt"))
            cccd = _clean_id(_g(row, "cccd"))
            ho_ten = _clean_text(_g(row, "ho_ten"))
            if not mst_nnt and not cccd and not ho_ten:
                continue
            if "TỔNG CỘNG" in (ho_ten + " " + _clean_text(row[0] if row else "")).upper():
                continue
            bk_raw = _clean_text(_g(row, "loai_bk")).upper()
            if "05-1" in bk_raw or "05_1" in bk_raw or "LŨY TIẾN" in bk_raw or "LUY TIEN" in bk_raw:
                loai_bk = "05-1"
            elif "05-2" in bk_raw or "05_2" in bk_raw or "TOÀN PHẦN" in bk_raw or "TOAN PHAN" in bk_raw or "10%" in bk_raw:
                loai_bk = "05-2"
            else:
                loai_bk = bk_raw
            records.append(
                MultiIncomeRecord(
                    mst_nnt=mst_nnt,
                    cccd=cccd,
                    ho_ten=ho_ten,
                    mst_tcct=clean_mst(_g(row, "mst_tcct")),
                    ten_tcct=_clean_text(_g(row, "ten_tcct")),
                    loai_bk=loai_bk,
                    year=_extract_year_from_val_or_path(_g(row, "nam"), fp),
                    tnct=_to_float(_g(row, "tnct")),
                    thue_kt=_to_float(_g(row, "thue_kt")),
                    source_file=os.path.basename(fp),
                )
            )
    return records


def evaluate_multi_income_for_person(
    mst_nnt: str,
    cccd: str,
    company_mst: str,
    tax_year: Optional[int],
    all_records: Sequence[MultiIncomeRecord],
) -> List[MultiIncomeRecord]:
    """Lấy tất cả dòng thu nhập tại NƠI KHÁC (khác company_mst) của cùng 1 cá nhân trong năm tax_year."""
    c_mst = clean_mst(mst_nnt)
    c_id = _clean_id(cccd)
    comp = clean_mst(company_mst)
    matched: List[MultiIncomeRecord] = []
    for rec in all_records:
        if tax_year and rec.year and rec.year != tax_year:
            continue
        id_match = (bool(c_mst) and rec.mst_nnt == c_mst) or (bool(c_id) and rec.cccd == c_id)
        if not id_match:
            continue
        # Loại trừ chính công ty đang kiểm tra (nếu file có MST tổ chức chi trả)
        if comp and rec.mst_tcct and rec.mst_tcct == comp:
            continue
        matched.append(rec)
    return matched


# =============================================================================
# 2. DANH BẠ NGƯỜI PHỤ THUỘC (CẢNH BÁO NPT LÀ HỘ KINH DOANH HOẠT ĐỘNG)
# =============================================================================
@dataclass
class NptDirectoryRecord:
    mst: str
    cccd: str
    ho_ten: str
    loai_nnt: str
    trang_thai: str
    is_hkd: bool
    is_active: bool
    ngay_bd: Optional[datetime.date]
    ngay_kt: Optional[datetime.date]
    source_file: str


def _looks_like_hkd(loai_nnt: str, trang_thai: str, raw_line: str) -> bool:
    txt = f"{loai_nnt} {trang_thai} {raw_line}".lower()
    hkd_keys = (
        "hộ kinh doanh", "ho kinh doanh", "hkd",
        "cá nhân kinh doanh", "ca nhan kinh doanh", "cnkd",
        "hộ gia đình kinh doanh", "chủ hộ kinh doanh", "kinh doanh cá thể",
        "hộ kd", "cá nhân kd",
    )
    return any(k in txt for k in hkd_keys)


def _looks_like_inactive(trang_thai: str) -> bool:
    tl = trang_thai.lower()
    return any(k in tl for k in ("đã đóng", "ngừng hoạt động", "chấm dứt", "giải thể", "hủy", "mã 01", "01 -"))


def parse_npt_directory_files(paths: Sequence[str]) -> List[NptDirectoryRecord]:
    records: List[NptDirectoryRecord] = []
    for fp in _collect_tabular_paths(paths):
        try:
            raw_rows = read_first_sheet_rows(fp)
        except Exception:
            continue
        if not raw_rows:
            continue

        hdr_idx: Optional[int] = None
        cmap: Dict[str, int] = {}
        for idx, row in enumerate(raw_rows[:30]):
            cells = [_clean_text(c).lower() for c in row]
            has_id = any("mst" in c or "mã số thuế" in c or "cccd" in c or "cmnd" in c or "định danh" in c for c in cells)
            has_dir = any(
                k in c
                for c in cells
                for k in ("trạng thái", "tình trạng", "loại nnt", "loại hình", "hộ kinh doanh", "ngành nghề", "danh bạ", "ngày cấp")
            )
            if has_id and has_dir:
                hdr_idx = idx
                for ci, h in enumerate(cells):
                    if any(k in h for k in ("mst npt", "mst", "mã số thuế")):
                        cmap.setdefault("mst", ci)
                    elif any(k in h for k in ("cccd", "cmnd", "định danh", "số giấy tờ")):
                        cmap.setdefault("cccd", ci)
                    elif any(k in h for k in ("họ tên", "họ và tên", "tên npt", "tên nnt", "tên người")):
                        cmap.setdefault("ho_ten", ci)
                    elif any(k in h for k in ("loại nnt", "loại hình", "đối tượng", "phân loại", "ngành nghề", "hình thức")):
                        cmap.setdefault("loai_nnt", ci)
                    elif any(k in h for k in ("trạng thái", "tình trạng")):
                        cmap.setdefault("trang_thai", ci)
                    elif any(k in h for k in ("ngày cấp", "ngày bắt đầu", "ngày hoạt động", "từ ngày", "ngày tạo")):
                        cmap.setdefault("ngay_bd", ci)
                    elif any(k in h for k in ("ngày đóng", "ngày chấm dứt", "ngày ngừng", "đến ngày")):
                        cmap.setdefault("ngay_kt", ci)
                break

        if hdr_idx is None:
            continue

        def _g(row: Sequence[Any], k: str) -> Any:
            ci = cmap.get(k)
            return row[ci] if (ci is not None and ci < len(row)) else None

        for row in raw_rows[hdr_idx + 1:]:
            mst = clean_mst(_g(row, "mst"))
            cccd = _clean_id(_g(row, "cccd"))
            ho_ten = _clean_text(_g(row, "ho_ten"))
            if not mst and not cccd and not ho_ten:
                continue
            loai = _clean_text(_g(row, "loai_nnt"))
            tt = _clean_text(_g(row, "trang_thai"))
            raw_line = " ".join(_clean_text(x) for x in row if x is not None)
            ngay_bd = _parse_date_any(_g(row, "ngay_bd"))
            ngay_kt = _parse_date_any(_g(row, "ngay_kt"))
            is_hkd = _looks_like_hkd(loai, tt, raw_line)
            is_active = not _looks_like_inactive(tt)
            records.append(
                NptDirectoryRecord(
                    mst=mst,
                    cccd=cccd,
                    ho_ten=ho_ten,
                    loai_nnt=loai,
                    trang_thai=tt,
                    is_hkd=is_hkd,
                    is_active=is_active,
                    ngay_bd=ngay_bd,
                    ngay_kt=ngay_kt,
                    source_file=os.path.basename(fp),
                )
            )
    return records


def check_npt_hkd_in_period(
    mst_npt: str,
    so_gt: str,
    decl_from: Optional[Tuple[int, int]],
    decl_to: Optional[Tuple[int, int]],
    dir_records: Sequence[NptDirectoryRecord],
) -> Tuple[bool, str]:
    """
    Kiểm tra NPT có phải là Hộ kinh doanh hoạt động trong khoảng thời gian giảm trừ [decl_from..decl_to] hay không.
    Trả về (is_violation, message).
    """
    m_npt = clean_mst(mst_npt)
    g_npt = _clean_id(so_gt)
    d_start = ym_idx(decl_from) if decl_from else None
    d_end = ym_idx(decl_to) if decl_to else None

    for rec in dir_records:
        if not rec.is_hkd:
            continue
        if not ((m_npt and rec.mst == m_npt) or (g_npt and rec.cccd == g_npt)):
            continue
        # Kiểm tra giao thoa thời gian hoạt động HKD với khoảng giảm trừ
        hkd_s = ym_idx((rec.ngay_bd.year, rec.ngay_bd.month)) if rec.ngay_bd else None
        hkd_e = ym_idx((rec.ngay_kt.year, rec.ngay_kt.month)) if rec.ngay_kt else None
        if not rec.is_active and hkd_e is None:
            # Đã đóng mà không rõ ngày đóng -> vẫn cảnh báo cần kiểm tra ngày đóng
            return True, f"CẦN KT: NPT từng là HKD ({rec.trang_thai or 'đã đóng'}) - KT ngày ngừng HĐ"
        overlaps = True
        if d_start is not None and d_end is not None:
            if hkd_s is not None and hkd_s > d_end:
                overlaps = False
            if hkd_e is not None and hkd_e < d_start:
                overlaps = False
        if overlaps:
            st_info = rec.trang_thai or "Đang hoạt động"
            return True, f"KHÔNG HỢP LỆ: NPT LÀ HỘ KINH DOANH ({st_info})"
    return False, "Hợp lệ"


# =============================================================================
# 3. DANH SÁCH NGƯỜI PHỤ THUỘC CÓ THU NHẬP (CẢNH BÁO VƯỢT NGƯỠNG GIẢM TRỪ)
# =============================================================================
@dataclass
class NptIncomeRecord:
    mst_npt: str
    cccd: str
    ho_ten: str
    year: Optional[int]
    tong_tn: float
    noi_chi_tra: str
    source_file: str


def parse_npt_income_files(paths: Sequence[str]) -> List[NptIncomeRecord]:
    records: List[NptIncomeRecord] = []
    for fp in _collect_tabular_paths(paths):
        try:
            raw_rows = read_first_sheet_rows(fp)
        except Exception:
            continue
        if not raw_rows:
            continue

        hdr_idx: Optional[int] = None
        cmap: Dict[str, int] = {}
        for idx, row in enumerate(raw_rows[:30]):
            cells = [_clean_text(c).lower() for c in row]
            has_id = any("mst" in c or "mã số thuế" in c or "cccd" in c or "cmnd" in c or "định danh" in c for c in cells)
            has_inc = any(k in c for c in cells for k in ("thu nhập", "tnct", "doanh thu", "số tiền"))
            if has_id and has_inc:
                hdr_idx = idx
                for ci, h in enumerate(cells):
                    if any(k in h for k in ("mst npt", "mst người phụ thuộc")) or (h in ("mst", "mã số thuế") and "mst" not in cmap):
                        cmap["mst_npt"] = ci
                    elif any(k in h for k in ("cccd", "cmnd", "định danh", "số giấy tờ")):
                        cmap.setdefault("cccd", ci)
                    elif any(k in h for k in ("tên npt", "họ tên npt", "họ và tên", "họ tên")):
                        cmap.setdefault("ho_ten", ci)
                    elif any(k in h for k in ("năm", "kỳ tính thuế", "kỳ quyết toán")):
                        cmap.setdefault("nam", ci)
                    elif any(k in h for k in ("tổng tnct", "thu nhập chịu thuế", "tổng thu nhập", "thu nhập", "doanh thu", "số tiền")):
                        cmap.setdefault("tong_tn", ci)
                    elif any(k in h for k in ("đơn vị", "tổ chức", "nơi chi trả")):
                        cmap.setdefault("noi_chi_tra", ci)
                break

        if hdr_idx is None:
            continue

        def _g(row: Sequence[Any], k: str) -> Any:
            ci = cmap.get(k)
            return row[ci] if (ci is not None and ci < len(row)) else None

        for row in raw_rows[hdr_idx + 1:]:
            mst = clean_mst(_g(row, "mst_npt"))
            cccd = _clean_id(_g(row, "cccd"))
            ho_ten = _clean_text(_g(row, "ho_ten"))
            if not mst and not cccd and not ho_ten:
                continue
            if "TỔNG CỘNG" in (ho_ten + " " + _clean_text(row[0] if row else "")).upper():
                continue
            records.append(
                NptIncomeRecord(
                    mst_npt=mst,
                    cccd=cccd,
                    ho_ten=ho_ten,
                    year=_extract_year_from_val_or_path(_g(row, "nam"), fp),
                    tong_tn=_to_float(_g(row, "tong_tn")),
                    noi_chi_tra=_clean_text(_g(row, "noi_chi_tra")),
                    source_file=os.path.basename(fp),
                )
            )
    return records


def check_npt_income_threshold(
    mst_npt: str,
    so_gt: str,
    tax_year: Optional[int],
    income_records: Sequence[NptIncomeRecord],
) -> Tuple[bool, float, str]:
    """
    Kiểm tra tổng thu nhập trong năm của NPT so với ngưỡng quy định:
    - Năm <= 2025: 1.000.000 đ/tháng (12.000.000 đ/năm) — TT 111/2013.
    - Từ năm 2026: 3.000.000 đ/tháng (36.000.000 đ/năm) — Luật 109/2025/QH15.
    Trả về (is_violation, total_income, message).
    """
    m_npt = clean_mst(mst_npt)
    g_npt = _clean_id(so_gt)
    yr = tax_year or 2024
    annual_limit = 36_000_000.0 if yr >= 2026 else 12_000_000.0

    total_inc = 0.0
    for rec in income_records:
        if tax_year and rec.year and rec.year != tax_year:
            continue
        if (m_npt and rec.mst_npt == m_npt) or (g_npt and rec.cccd == g_npt):
            total_inc += rec.tong_tn

    if total_inc <= 0:
        return False, 0.0, "Không phát sinh TN"
    if total_inc > annual_limit:
        return True, total_inc, f"KHÔNG ĐỦ ĐK GIẢM TRỪ: NPT CÓ TN {total_inc:,.0f} đ (> {annual_limit:,.0f} đ/năm)"
    return False, total_inc, f"Có TN {total_inc:,.0f} đ (≤ ngưỡng {annual_limit:,.0f} đ/năm)"
