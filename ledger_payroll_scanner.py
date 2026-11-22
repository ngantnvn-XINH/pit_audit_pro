# -*- coding: utf-8 -*-
"""
Bộ quét & đối chiếu Bảng lương / Sổ chi tiết tài khoản kế toán phục vụ kiểm tra Thuế TNCN:
- Nhóm tài khoản lương & bảo hiểm & thuế: TK 334, TK 3383/3384/3386, TK 3335
- Nhóm tài khoản chi phí & quỹ phúc lợi: TK 353, TK 641, TK 642, TK 622, TK 627, TK 154, TK 811

Tự động nhận diện cột từ sổ chi tiết MISA, FAST, Bravo hoặc Excel tự lập:
  Ngày chứng từ | Số chứng từ | Diễn giải | TK Nợ (hoặc TK) | TK Có (hoặc TK ĐƯ) | Phát sinh Nợ | Phát sinh Có | Đối tượng
"""

from __future__ import annotations

import datetime
import os
import re
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional, Sequence, Tuple

import openpyxl


def _clean_text(val: Any) -> str:
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val).replace("\xa0", " ")).strip()


def _to_float(val: Any) -> float:
    if val is None or isinstance(val, bool):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = _clean_text(val)
    if not s:
        return 0.0
    if "," in s and "." in s:
        if s.rfind(",") > s.rfind("."):
            s = s.replace(".", "").replace(",", ".")
        else:
            s = s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        if len(parts) == 2 and len(parts[1]) <= 2:
            s = s.replace(",", ".")
        else:
            s = s.replace(",", "")
    try:
        return float(s)
    except ValueError:
        return 0.0


def _fmt_date(val: Any) -> str:
    if isinstance(val, (datetime.date, datetime.datetime)):
        return val.strftime("%d/%m/%Y")
    return _clean_text(val)


# 6 nhóm rủi ro chi phí chịu thuế TNCN thường bị bỏ sót ngoài bảng lương
RISK_RULES: List[Tuple[str, str, str, Tuple[str, ...], float]] = [
    (
        "Tiền thuê nhà / điện nước chuyên gia",
        "CAO",
        "Khoản 2 Điều 11 TT 92/2015: Tính vào TNCT nhưng không vượt quá 15% tổng TNCT chưa gồm tiền thuê nhà",
        ("thuê nhà", "thue nha", "nhà ở chuyên gia", "nha o chuyen gia", "tiền nhà", "tien nha", "housing", "rent house", "apartment"),
        0.0,
    ),
    (
        "Thưởng lễ, Tết, nghỉ mát, phúc lợi đích danh",
        "CAO",
        "Điểm đ.3.2 Khoản 2 Điều 2 TT 111/2013: Khoản chi phúc lợi ghi đích danh cá nhân phải tính vào TNCT",
        ("thưởng", "thuong", "nghỉ mát", "nghi mat", "du lịch", "du lich", "phúc lợi", "phuc loi", "quà tặng", "qua tang", "tết", "lễ 30/4", "2/9", "lương tháng 13", "bonus"),
        0.0,
    ),
    (
        "Bảo hiểm nhân thọ / sức khỏe tự nguyện / thẻ hội viên",
        "CAO",
        "Điểm đ.2 Khoản 2 Điều 2 TT 111/2013: BH nhân thọ có tích lũy, thẻ golf/gym ghi tên cá nhân tính vào TNCT",
        ("bảo hiểm nhân thọ", "bao hiem nhan tho", "nhân thọ", "thẻ hội viên", "the hoi vien", "golf", "gym", "club", "bảo hiểm sức khỏe", "health insurance"),
        0.0,
    ),
    (
        "Vé máy bay / học phí / chi trả hộ cá nhân",
        "TRUNG BÌNH",
        "Điểm đ.6, đ.7 Khoản 2 Điều 2 TT 111/2013: Kiểm tra số lần vé máy bay khứ hồi (1 lần/năm) & bậc học miễn thuế",
        ("vé máy bay", "ve may bay", "học phí", "hoc phi", "tuition", "air ticket", "visa", "giấy phép lao động", "work permit"),
        0.0,
    ),
    (
        "Phụ cấp trang phục / ăn ca / điện thoại vượt định mức",
        "TRUNG BÌNH",
        "Điểm đ.4 Khoản 2 Điều 2 TT 111/2013: Trang phục bằng tiền vượt 5tr/năm; ăn ca/điện thoại vượt quy chế DN",
        ("trang phục", "trang phuc", "đồng phục", "dong phuc", "ăn trưa", "an trua", "ăn ca", "an ca", "điện thoại", "dien thoai", "xăng xe"),
        0.0,
    ),
    (
        "Chi trả cá nhân ngoài HĐLĐ >= 2 triệu (nguy cơ sót khấu trừ 10%)",
        "CAO",
        "Điểm i Khoản 1 Điều 25 TT 111/2013: Chi trả cá nhân không ký HĐLĐ từ 2.000.000 đ/lần trở lên phải khấu trừ 10% (05-2)",
        ("môi giới", "moi gioi", "hoa hồng", "hoa hong", "cộng tác viên", "cong tac vien", "ctv", "thù lao", "thu lao", "thuê ngoài", "thue ngoai", "tư vấn cá nhân", "dịch vụ cá nhân", "nhân công thuê ngoài", "tiền công", "bốc xếp", "thợ"),
        2_000_000.0,
    ),
]


@dataclass
class FlaggedCostEntry:
    source_file: str
    ngay_ct: str
    so_ct: str
    tk_no: str
    tk_co: str
    doi_tuong: str
    dien_giai: str
    so_tien: float
    nhom_rui_ro: str
    muc_do: str
    can_cu_phap_ly: str


@dataclass
class LedgerAuditSummary:
    files_scanned: List[str] = field(default_factory=list)
    tk334_credit_total: float = 0.0   # Tổng phát sinh Có TK 334 (Lương/thưởng phải trả phát sinh)
    tk334_debit_total: float = 0.0    # Tổng phát sinh Nợ TK 334 (Lương/thưởng đã chi trả + khấu trừ)
    tk338_insurance_total: float = 0.0  # Tổng BHXH/BHYT/BHTN khấu trừ lương (Nợ 334 / Có 3383,3384,3386)
    tk3335_credit_total: float = 0.0  # Tổng Thuế TNCN phải nộp phát sinh (Có 3335)
    tk3335_debit_total: float = 0.0   # Tổng Thuế TNCN đã nộp NSNN (Nợ 3335)
    tk353_debit_total: float = 0.0    # Tổng chi Quỹ khen thưởng, phúc lợi (Nợ 353)
    flagged_entries: List[FlaggedCostEntry] = field(default_factory=list)


def _detect_columns(rows: List[List[Any]]) -> Tuple[Optional[int], Dict[str, int]]:
    for idx, row in enumerate(rows[:25]):
        norm = [_clean_text(c).lower() for c in row]
        has_desc = any(("diễn giải" in c or "nội dung" in c or "dien giai" in c) for c in norm)
        has_amt = any(("phát sinh" in c or "số tiền" in c or "nợ" == c or "có" == c or "ps nợ" in c) for c in norm)
        if has_desc and has_amt:
            cmap: Dict[str, int] = {}
            for ci, h in enumerate(norm):
                if "ngày" in h and "ngay_ct" not in cmap:
                    cmap["ngay_ct"] = ci
                elif ("số ct" in h or "số chứng từ" in h or "so ct" in h) and "so_ct" not in cmap:
                    cmap["so_ct"] = ci
                elif ("diễn giải" in h or "nội dung" in h or "dien giai" in h) and "dien_giai" not in cmap:
                    cmap["dien_giai"] = ci
                elif ("đối tượng" in h or "khách hàng" in h or "nhân viên" in h or "cá nhân" in h) and "doi_tuong" not in cmap:
                    cmap["doi_tuong"] = ci
                elif h in ("tk nợ", "tài khoản nợ", "tk", "tài khoản") and "tk_no" not in cmap:
                    cmap["tk_no"] = ci
                elif h in ("tk có", "tài khoản có", "tk đối ứng", "tk đư", "đối ứng") and "tk_co" not in cmap:
                    cmap["tk_co"] = ci
                elif ("phát sinh nợ" in h or "ps nợ" in h or h == "nợ" or "số tiền nợ" in h) and "ps_no" not in cmap:
                    cmap["ps_no"] = ci
                elif ("phát sinh có" in h or "ps có" in h or h == "có" or "số tiền có" in h) and "ps_co" not in cmap:
                    cmap["ps_co"] = ci
                elif ("số tiền" in h or "giá trị" in h) and "so_tien" not in cmap:
                    cmap["so_tien"] = ci
            return idx, cmap
    return None, {}


def _classify_risk(dien_giai: str, tk_no: str, tk_co: str, amount: float) -> Optional[Tuple[str, str, str]]:
    desc_low = dien_giai.lower()
    if not desc_low or amount <= 0:
        return None
    # Chỉ quét các tài khoản chi phí / quỹ phúc lợi (hoặc các bút toán không thuộc kết chuyển cuối kỳ)
    if any(k in desc_low for k in ("kết chuyển", "ket chuyen", "khấu hao", "khau hao", "phân bổ ccdc")):
        return None

    is_cost_or_welfare = (
        tk_no.startswith(("641", "642", "622", "627", "154", "353", "811"))
        or tk_co.startswith("353")
        or not tk_no
    )
    if not is_cost_or_welfare:
        return None

    for nhom, muc_do, can_cu, keywords, min_amt in RISK_RULES:
        if amount < min_amt:
            continue
        if any(kw in desc_low for kw in keywords):
            return nhom, muc_do, can_cu

    # Quỹ phúc lợi TK 353 phát sinh Nợ -> luôn đưa vào danh sách cần kiểm tra
    if tk_no.startswith("353") and amount >= 500_000:
        return (
            "Chi từ Quỹ khen thưởng, phúc lợi (TK 353)",
            "TRUNG BÌNH",
            "Điểm đ.3.2 Khoản 2 Điều 2 TT 111/2013: Kiểm tra danh sách cá nhân nhận thưởng/phúc lợi từ TK 353",
        )
    return None


def scan_accounting_ledgers(file_paths: Sequence[str]) -> LedgerAuditSummary:
    """Quét 1 hoặc nhiều file Excel sổ chi tiết tài khoản / bảng lương để tổng hợp TK 334/338/3335/353 và lọc rủi ro chi phí."""
    summary = LedgerAuditSummary()

    for path in file_paths:
        if not path or not os.path.isfile(path):
            continue
        summary.files_scanned.append(os.path.basename(path))
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
        for sname in wb.sheetnames:
            ws = wb[sname]
            rows = [list(r) for r in ws.iter_rows(values_only=True)]
            if not rows:
                continue
            hdr_idx, cmap = _detect_columns(rows)
            if hdr_idx is None:
                continue

            # Dò xem sheet/file này có phải sổ chi tiết riêng của 1 tài khoản (VD: tên sheet '334', '642', '3335')
            m_sheet_tk = re.search(r"\b(334\d*|3335\d*|338[346]\d*|353\d*|641\d*|642\d*|622\d*|627\d*|154\d*|811\d*)\b", sname + " " + os.path.basename(path))
            default_tk = m_sheet_tk.group(1) if m_sheet_tk else ""

            def _get(row: List[Any], key: str) -> Any:
                ci = cmap.get(key)
                return row[ci] if (ci is not None and ci < len(row)) else None

            for row in rows[hdr_idx + 1:]:
                dien_giai = _clean_text(_get(row, "dien_giai"))
                if not dien_giai:
                    continue
                low_d = dien_giai.lower()
                if any(s in low_d for s in ("số dư đầu kỳ", "cộng phát sinh", "số dư cuối kỳ", "tổng cộng", "cộng lũy kế")):
                    continue

                tk_no = _clean_text(_get(row, "tk_no")) or default_tk
                tk_co = _clean_text(_get(row, "tk_co"))
                ps_no = _to_float(_get(row, "ps_no"))
                ps_co = _to_float(_get(row, "ps_co"))
                so_tien = _to_float(_get(row, "so_tien"))
                amt = ps_no if ps_no > 0 else (ps_co if ps_co > 0 else so_tien)
                if amt <= 0:
                    continue

                # 1. Tổng hợp TK 334, 3383/3384/3386, 3335, 353
                if tk_no.startswith("334"):
                    if ps_no > 0:
                        summary.tk334_debit_total += ps_no
                    if ps_co > 0:
                        summary.tk334_credit_total += ps_co
                    if tk_co.startswith(("3383", "3384", "3386")):
                        summary.tk338_insurance_total += (ps_no or amt)
                    if tk_co.startswith("3335"):
                        summary.tk3335_credit_total += (ps_no or amt)
                elif tk_co.startswith("334") and ps_no > 0 and "ps_co" not in cmap:
                    summary.tk334_credit_total += ps_no

                if tk_no.startswith("3335"):
                    if ps_co > 0:
                        summary.tk3335_credit_total += ps_co
                    if ps_no > 0:
                        summary.tk3335_debit_total += ps_no

                if tk_no.startswith("353") and (ps_no > 0 or amt > 0):
                    summary.tk353_debit_total += (ps_no or amt)

                # 2. Quét rủi ro chi phí chịu thuế TNCN bỏ sót
                risk = _classify_risk(dien_giai, tk_no, tk_co, amt)
                if risk is not None:
                    nhom, muc_do, can_cu = risk
                    summary.flagged_entries.append(
                        FlaggedCostEntry(
                            source_file=os.path.basename(path),
                            ngay_ct=_fmt_date(_get(row, "ngay_ct")),
                            so_ct=_clean_text(_get(row, "so_ct")),
                            tk_no=tk_no,
                            tk_co=tk_co,
                            doi_tuong=_clean_text(_get(row, "doi_tuong")),
                            dien_giai=dien_giai,
                            so_tien=amt,
                            nhom_rui_ro=nhom,
                            muc_do=muc_do,
                            can_cu_phap_ly=can_cu,
                        )
                    )
        wb.close()

    return summary
