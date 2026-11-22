# -*- coding: utf-8 -*-
"""
Động cơ đối chiếu Người phụ thuộc (05-3 vs File kết xuất TMS) cho phần mềm pit_audit_pro.
- Đọc file TMS (.xlsx / .csv) định dạng 'HIỂN THỊ DANH SÁCH ĐỘNG'
- Hợp nhất lịch sử báo tăng / báo giảm cho cùng 1 cặp (MST NNT, MST NPT)
- Khớp đa tầng (MST NNT + MST NPT, hoặc MST NNT + CCCD/CMND/Hộ chiếu kể cả mất số 0 đầu)
- Tính chính xác số tháng hợp lệ & số tháng bị loại trừ cho từng dòng 05-3 và gom nhóm theo MST NNT
  để đẩy thẳng vào bảng tính lại thuế TNCN lũy tiến từng phần trên sheet TinhLai_05_1.
"""

from __future__ import annotations

import csv
import datetime
import os
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

import openpyxl

from xml_05qtt_parser import Row053


def _clean_text(val: Any) -> str:
    if val is None:
        return ""
    return re.sub(r"\s+", " ", str(val).replace("\xa0", " ").replace("\u200b", "")).strip()


def _clean_id(val: Any) -> str:
    s = _clean_text(val)
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s.replace(" ", "").upper()


def _norm_no_zero(val: str) -> str:
    c = _clean_id(val)
    return c.lstrip("0") or c


def parse_ym(val: Any) -> Optional[Tuple[int, int]]:
    if val is None:
        return None
    if isinstance(val, (datetime.date, datetime.datetime)):
        return (val.year, val.month)
    s = _clean_text(val)
    if not s:
        return None
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]

    m1 = re.fullmatch(r"(\d{1,2})[/.\-](\d{4})", s)
    if m1:
        mo, yr = int(m1.group(1)), int(m1.group(2))
        if 1 <= mo <= 12 and 1990 <= yr <= 2100:
            return (yr, mo)
    m2 = re.fullmatch(r"(\d{4})[/.\-](\d{1,2})", s)
    if m2:
        yr, mo = int(m2.group(1)), int(m2.group(2))
        if 1 <= mo <= 12 and 1990 <= yr <= 2100:
            return (yr, mo)
    digits = re.sub(r"\D", "", s)
    if len(digits) == 5:
        mo, yr = int(digits[0]), int(digits[1:])
        if 1 <= mo <= 9 and 1990 <= yr <= 2100:
            return (yr, mo)
    elif len(digits) == 6:
        mo, yr = int(digits[:2]), int(digits[2:])
        if 1 <= mo <= 12 and 1990 <= yr <= 2100:
            return (yr, mo)
    return None


def fmt_ym(ym: Optional[Tuple[int, int]]) -> str:
    return f"{ym[1]:02d}/{ym[0]}" if ym else ""


def ym_idx(ym: Tuple[int, int]) -> int:
    return ym[0] * 12 + (ym[1] - 1)


def _parse_dmy(val: Any) -> datetime.date:
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = _clean_text(val)
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return datetime.date.min


@dataclass
class TmsNptRow:
    ngay_tao: str
    ngay_td: str
    mst_nnt: str
    ten_nnt: str
    mst_npt: str
    ten_npt: str
    trang_thai: str
    ma_tt: str
    ngay_sinh: str
    cmnd: str
    ho_chieu: str
    quan_he: str
    tu_thang: Optional[Tuple[int, int]]
    den_thang: Optional[Tuple[int, int]]
    mau_tk: str
    gks_so: str
    sort_date: datetime.date


@dataclass
class NptAuditResultRow:
    src: Row053
    decl_months: int
    tms_from: str
    tms_to: str
    tms_status: str
    valid_months: int
    disallowed_months: int
    status: str
    is_ok: bool


def parse_tms_npt_file(file_path: str) -> List[TmsNptRow]:
    ext = os.path.splitext(file_path)[1].lower()
    raw_rows: List[List[Any]] = []
    if ext in (".csv", ".txt"):
        for enc in ("utf-8-sig", "utf-8", "cp1252", "latin-1"):
            try:
                with open(file_path, "r", encoding=enc, newline="") as f:
                    raw_rows = [r for r in csv.reader(f)]
                break
            except UnicodeDecodeError:
                continue
    else:
        wb = openpyxl.load_workbook(file_path, data_only=True, read_only=True)
        ws = wb[wb.sheetnames[0]]
        raw_rows = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()

    hdr_idx: Optional[int] = None
    cmap: Dict[str, int] = {}
    thang_cols: List[int] = []
    for idx, row in enumerate(raw_rows[:30]):
        cells = [_clean_text(c).lower() for c in row]
        if any("mst nnt" in c for c in cells) and any("mst npt" in c for c in cells):
            hdr_idx = idx
            for c_i, h in enumerate(cells):
                if "ngày tạo" in h or "ngày tạo" in h:
                    cmap.setdefault("ngay_tao", c_i)
                elif h.startswith("ngày tđ") or "ngày thay đổi" in h:
                    cmap.setdefault("ngay_td", c_i)
                elif "mst nnt" in h:
                    cmap.setdefault("mst_nnt", c_i)
                elif "tên nnt" in h:
                    cmap.setdefault("ten_nnt", c_i)
                elif "mst npt" in h:
                    cmap.setdefault("mst_npt", c_i)
                elif "tên npt" in h:
                    cmap.setdefault("ten_npt", c_i)
                elif "trạng thái npt" in h:
                    cmap.setdefault("trang_thai", c_i)
                elif "ngày sinh" in h:
                    cmap.setdefault("ngay_sinh", c_i)
                elif h in ("cmnd", "cccd", "cmnd/cccd", "số giấy tờ"):
                    cmap.setdefault("cmnd", c_i)
                elif "quan hệ" in h:
                    cmap.setdefault("quan_he", c_i)
                elif h in ("tháng", "tháng", "từ tháng", "đến tháng"):
                    thang_cols.append(c_i)
                elif "mẫu bk" in h or "mẫu tk" in h:
                    cmap.setdefault("mau_tk", c_i)
                elif "hộ chiếu" in h:
                    cmap.setdefault("ho_chieu", c_i)
                elif "gks/số" in h or "số gks" in h:
                    cmap.setdefault("gks_so", c_i)
                elif h.startswith("mã trạng t"):
                    cmap.setdefault("ma_tt", c_i)
            break

    if hdr_idx is None:
        raise ValueError("Không tìm thấy tiêu đề 'MST NNT' & 'MST NPT' trong file NPT hệ thống.")
    if len(thang_cols) >= 1:
        cmap["tu_thang"] = thang_cols[0]
    if len(thang_cols) >= 2:
        cmap["den_thang"] = thang_cols[1]

    def _g(row: Sequence[Any], k: str) -> Any:
        i = cmap.get(k)
        return row[i] if (i is not None and i < len(row)) else None

    out: List[TmsNptRow] = []
    for row in raw_rows[hdr_idx + 1:]:
        mst_nnt = _clean_id(_g(row, "mst_nnt"))
        mst_npt = _clean_id(_g(row, "mst_npt"))
        ten_npt = _clean_text(_g(row, "ten_npt"))
        if not mst_nnt and not mst_npt and not ten_npt:
            continue
        nt = _clean_text(_g(row, "ngay_tao"))
        ntd = _clean_text(_g(row, "ngay_td"))
        out.append(
            TmsNptRow(
                ngay_tao=nt,
                ngay_td=ntd,
                mst_nnt=mst_nnt,
                ten_nnt=_clean_text(_g(row, "ten_nnt")),
                mst_npt=mst_npt,
                ten_npt=ten_npt,
                trang_thai=_clean_text(_g(row, "trang_thai")),
                ma_tt=_clean_text(_g(row, "ma_tt")),
                ngay_sinh=_clean_text(_g(row, "ngay_sinh")),
                cmnd=_clean_id(_g(row, "cmnd")),
                ho_chieu=_clean_id(_g(row, "ho_chieu")),
                quan_he=_clean_text(_g(row, "quan_he")),
                tu_thang=parse_ym(_g(row, "tu_thang")),
                den_thang=parse_ym(_g(row, "den_thang")),
                mau_tk=_clean_text(_g(row, "mau_tk")),
                gks_so=_clean_text(_g(row, "gks_so")),
                sort_date=max(_parse_dmy(nt), _parse_dmy(ntd)),
            )
        )
    return out


def _resolve_latest(recs: List[TmsNptRow]) -> TmsNptRow:
    pool = [r for r in recs if r.tu_thang is not None] or recs
    pool_s = sorted(
        pool,
        key=lambda r: (r.sort_date, 1 if r.den_thang is not None else 0, ym_idx(r.tu_thang) if r.tu_thang else -1),
    )
    return pool_s[-1]


def audit_dependents(
    rows_053: List[Row053],
    tms_rows: List[TmsNptRow],
) -> Tuple[List[NptAuditResultRow], Dict[str, Tuple[int, int]]]:
    """
    Đối chiếu danh sách 05-3 với danh sách TMS.
    Trả về:
    - Danh sách kết quả chi tiết từng dòng 05-3
    - Bản đồ theo MST NNT: mst_nnt -> (tổng_tháng_05_3, tổng_tháng_bị_loại_trừ)
    """
    by_mst: Dict[Tuple[str, str], List[TmsNptRow]] = {}
    by_doc: Dict[Tuple[str, str], List[TmsNptRow]] = {}
    by_npt_only: Dict[str, List[TmsNptRow]] = {}

    for rec in tms_rows:
        if rec.mst_nnt and rec.mst_npt:
            by_mst.setdefault((rec.mst_nnt, rec.mst_npt), []).append(rec)
        if rec.mst_nnt:
            for d in (rec.cmnd, rec.ho_chieu):
                if d:
                    by_doc.setdefault((rec.mst_nnt, _norm_no_zero(d)), []).append(rec)
        if rec.mst_npt:
            by_npt_only.setdefault(rec.mst_npt, []).append(rec)

    res_mst = {k: _resolve_latest(v) for k, v in by_mst.items()}
    res_doc = {k: _resolve_latest(v) for k, v in by_doc.items()}

    # Phát hiện trùng NPT giữa nhiều NNT trên 05-3
    npt_nnts: Dict[str, Set[str]] = {}
    for r in rows_053:
        key = r.mst_npt or _norm_no_zero(r.so_gt)
        if key and r.mst_nnt:
            npt_nnts.setdefault(key, set()).add(r.mst_nnt)

    results: List[NptAuditResultRow] = []
    per_nnt: Dict[str, List[int]] = {}  # mst_nnt -> [decl_m, disallowed_m]

    for r in rows_053:
        ym_f = parse_ym(r.tu_thang)
        ym_t = parse_ym(r.den_thang)
        decl_m = (ym_idx(ym_t) - ym_idx(ym_f) + 1) if (ym_f and ym_t and ym_idx(ym_t) >= ym_idx(ym_f)) else 0

        matched: Optional[TmsNptRow] = None
        if r.mst_nnt and r.mst_npt and (r.mst_nnt, r.mst_npt) in res_mst:
            matched = res_mst[(r.mst_nnt, r.mst_npt)]
        elif r.mst_nnt and r.so_gt and (r.mst_nnt, _norm_no_zero(r.so_gt)) in res_doc:
            matched = res_doc[(r.mst_nnt, _norm_no_zero(r.so_gt))]

        npt_key = r.mst_npt or _norm_no_zero(r.so_gt)
        is_dup = bool(npt_key and len(npt_nnts.get(npt_key, set())) > 1)

        reasons: List[str] = []
        tms_f_s = ""
        tms_t_s = ""
        tms_st_s = ""
        valid_m = 0

        if not tms_rows:
            # Chưa nạp file TMS -> giữ nguyên số tháng kê khai
            valid_m = decl_m
            status = "CHƯA NẠP FILE TMS"
            is_ok = True
        else:
            if matched is None:
                if r.mst_npt and r.mst_npt in by_npt_only:
                    others = sorted({x.mst_nnt for x in by_npt_only[r.mst_npt] if x.mst_nnt != r.mst_nnt})
                    reasons.append(f"ĐĂNG KÝ CHO NNT KHÁC ({', '.join(others[:2])})")
                else:
                    reasons.append("CHƯA ĐĂNG KÝ TRÊN HỆ THỐNG")
                valid_m = 0
            else:
                tms_f_s = fmt_ym(matched.tu_thang)
                tms_t_s = fmt_ym(matched.den_thang) if matched.den_thang else "Còn hiệu lực"
                tms_st_s = f"{matched.ma_tt} - {matched.trang_thai}".strip(" -")
                if matched.ma_tt and matched.ma_tt != "00":
                    reasons.append(f"MST NPT NGỪNG HĐ (Mã {matched.ma_tt})")
                if matched.tu_thang is None:
                    reasons.append("CẦN KIỂM TRA (TMS THIẾU TỪ THÁNG)")
                    valid_m = 0
                elif ym_f and ym_t:
                    t1 = ym_idx(matched.tu_thang)
                    t2 = ym_idx(matched.den_thang) if matched.den_thang else ym_idx((9999, 12))
                    lo = max(ym_idx(ym_f), t1)
                    hi = min(ym_idx(ym_t), t2)
                    valid_m = max(0, hi - lo + 1)
                    if valid_m < decl_m:
                        reasons.append(
                            "NGOÀI THỜI GIAN ĐĂNG KÝ TMS" if valid_m == 0
                            else f"VƯỢT THỜI GIAN ĐĂNG KÝ ({decl_m - valid_m} tháng)"
                        )
            if is_dup:
                reasons.append("TRÙNG NPT GIỮA NHIỀU NNT TRÊN 05-3")

            is_ok = len(reasons) == 0
            status = "KHỚP" if is_ok else "; ".join(reasons)

        disallowed_m = max(0, decl_m - valid_m)
        results.append(
            NptAuditResultRow(
                src=r,
                decl_months=decl_m,
                tms_from=tms_f_s,
                tms_to=tms_t_s,
                tms_status=tms_st_s,
                valid_months=valid_m,
                disallowed_months=disallowed_m,
                status=status,
                is_ok=is_ok,
            )
        )
        if r.mst_nnt:
            acc = per_nnt.setdefault(r.mst_nnt, [0, 0])
            acc[0] += decl_m
            acc[1] += disallowed_m

    return results, {k: (v[0], v[1]) for k, v in per_nnt.items()}
