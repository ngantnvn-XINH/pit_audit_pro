# -*- coding: utf-8 -*-
"""
Động cơ đối chiếu Người phụ thuộc (05-3 vs File kết xuất TMS) cho phần mềm pit_audit_pro:
1. Đọc file TMS (.xlsx / .csv) định dạng 'HIỂN THỊ DANH SÁCH ĐỘNG' (xử lý \xa0, 2 cột 'Tháng', hợp nhất lịch sử báo tăng/giảm).
2. Kiểm tra ĐỘ TUỔI Người phụ thuộc theo Điểm d Khoản 1 Điều 9 TT 111/2013/TT-BTC & Bộ luật Lao động 2019 (NĐ 135/2020/NĐ-CP):
   - Con: <18t (hợp lệ), 18-22t (KT thẻ SV/đang học), >22t (rủi ro cao), khai trước tháng sinh.
   - Cha/Mẹ, Vợ/Chồng, Khác: đối chiếu lộ trình tuổi nghỉ hưu theo từng năm (trước 2021: Nam 60/Nữ 55; từ 2021 tăng dần).
3. Kiểm tra TRÙNG ĐĂNG KÝ NPT đa chiều:
   - Trùng trong cùng 1 NNT trên 05-3 (kê khai lặp có tháng giao nhau).
   - Trùng giữa >= 2 NNT khác nhau trên 05-3 (theo MST NPT, CCCD/CMND, hoặc Họ tên + Ngày sinh).
   - Trùng đăng ký với NNT khác có thời gian giao nhau trên hệ thống TMS.
"""

from __future__ import annotations

import csv
import datetime
import os
import re
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Sequence, Set, Tuple

import openpyxl

from tabular_reader import read_first_sheet_rows
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


def _norm_name(val: str) -> str:
    return re.sub(r"\s+", " ", _clean_text(val).lower()).strip()


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


def parse_date_flexible(val: Any) -> Optional[datetime.date]:
    if isinstance(val, datetime.datetime):
        return val.date()
    if isinstance(val, datetime.date):
        return val
    s = _clean_text(val)
    if not s:
        return None
    for fmt in ("%d.%m.%Y", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"):
        try:
            return datetime.datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _parse_dmy(val: Any) -> datetime.date:
    return parse_date_flexible(val) or datetime.date.min


def retirement_age_thresholds(tax_year: int) -> Tuple[float, float]:
    """
    Trả về (tuổi_nghỉ_hưu_Nữ, tuổi_nghỉ_hưu_Nam) theo Bộ luật Lao động 2019 (NĐ 135/2020/NĐ-CP):
    - Trước 2021: Nữ 55.0, Nam 60.0
    - Từ 2021: Nữ tăng +4 tháng/năm (đến 60 tuổi năm 2035), Nam tăng +3 tháng/năm (đến 62 tuổi năm 2028)
    """
    if tax_year <= 2020:
        return (55.0, 60.0)
    delta = tax_year - 2020
    female = min(60.0, 55.0 + (delta * 4) / 12.0)
    male = min(62.0, 60.0 + (delta * 3) / 12.0)
    return (female, male)


def evaluate_dependent_age(
    dob: Optional[datetime.date],
    quan_he: str,
    ten_npt: str,
    tax_year: int,
    decl_from: Optional[Tuple[int, int]],
) -> Tuple[Optional[int], str, bool]:
    """
    Kiểm tra độ tuổi người phụ thuộc trong năm tính thuế `tax_year`.
    Trả về: (tuổi_tính_theo_năm, thông_báo_độ_tuổi, is_age_ok)
    """
    if dob is None:
        return (None, "Thiếu ngày sinh (cần kiểm tra)", False)

    age = tax_year - dob.year
    qh_low = _clean_text(quan_he).lower()
    is_child = ("con" in qh_low) or (qh_low == "01")

    # Kiểm tra khai trước tháng sinh
    if decl_from is not None:
        if ym_idx(decl_from) < ym_idx((dob.year, dob.month)):
            return (age, f"KHAI TRƯỚC THÁNG SINH (Sinh {dob.month:02d}/{dob.year})", False)

    if is_child:
        if age < 18:
            return (age, "Hợp lệ (Con < 18 tuổi)", True)
        if age <= 22:
            return (age, "CẦN KIỂM TRA (Con 18–22 tuổi: KT thẻ SV / hồ sơ đang học)", False)
        return (age, f"RỦI RO ĐỘ TUỔI (Con {age} tuổi > 22t: KT hồ sơ học / khuyết tật)", False)

    # Cha/Mẹ, Vợ/Chồng, Khác -> kiểm tra ngưỡng tuổi nghỉ hưu theo năm
    female_ret, male_ret = retirement_age_thresholds(tax_year)
    if age >= male_ret:
        return (age, "Hợp lệ (Ngoài độ tuổi lao động)", True)
    if age < female_ret:
        return (
            age,
            f"TRONG ĐỘ TUỔI LAO ĐỘNG ({age}t < {female_ret:.1f}t: KT hồ sơ khuyết tật / mất khả năng LĐ)",
            False,
        )
    # Vùng giữa tuổi nghỉ hưu Nữ và Nam (VD 56–60t): nhận diện nhanh tên đệm 'Thị' là Nữ
    name_tokens = _norm_name(ten_npt).split()
    if "thị" in name_tokens or "thi" in name_tokens:
        return (age, "Hợp lệ (Nữ ngoài độ tuổi lao động)", True)
    return (
        age,
        f"CẦN KIỂM TRA GIỚI TÍNH ({age}t: Nữ hợp lệ, Nam còn trong tuổi LĐ < {male_ret:.1f}t)",
        False,
    )


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
    age: Optional[int]
    age_status: str
    is_age_ok: bool
    dup_status: str
    is_dup: bool
    tms_from: str
    tms_to: str
    tms_status: str
    valid_months: int
    disallowed_months: int
    status: str
    is_ok: bool


def parse_tms_npt_file(file_path: str) -> List[TmsNptRow]:
    raw_rows = read_first_sheet_rows(file_path)

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


def _intervals_overlap(
    a1: Optional[Tuple[int, int]],
    a2: Optional[Tuple[int, int]],
    b1: Optional[Tuple[int, int]],
    b2: Optional[Tuple[int, int]],
) -> int:
    if a1 is None or a2 is None or b1 is None:
        return 0
    s1, e1 = ym_idx(a1), ym_idx(a2)
    s2 = ym_idx(b1)
    e2 = ym_idx(b2) if b2 is not None else ym_idx((9999, 12))
    if e1 < s1 or e2 < s2:
        return 0
    return max(0, min(e1, e2) - max(s1, s2) + 1)


def audit_dependents(
    rows_053: List[Row053],
    tms_rows: List[TmsNptRow],
    tax_year: Optional[int] = None,
) -> Tuple[List[NptAuditResultRow], Dict[str, Tuple[int, int]]]:
    """
    Đối chiếu toàn diện danh sách 05-3 với hệ thống TMS + Độ tuổi NPT + Trùng đăng ký NPT:
    - Chạy kiểm tra độ tuổi & trùng NPT ngay cả khi chưa nạp file TMS.
    - Trả về (danh sách kết quả chi tiết từng dòng, bản đồ per_nnt: mst_nnt -> (decl_m, disallowed_m)).
    """
    by_mst: Dict[Tuple[str, str], List[TmsNptRow]] = {}
    by_doc: Dict[Tuple[str, str], List[TmsNptRow]] = {}
    by_npt_only: Dict[str, List[TmsNptRow]] = {}
    by_doc_only: Dict[str, List[TmsNptRow]] = {}

    for rec in tms_rows:
        if rec.mst_nnt and rec.mst_npt:
            by_mst.setdefault((rec.mst_nnt, rec.mst_npt), []).append(rec)
        if rec.mst_nnt:
            for d in (rec.cmnd, rec.ho_chieu):
                if d:
                    by_doc.setdefault((rec.mst_nnt, _norm_no_zero(d)), []).append(rec)
        if rec.mst_npt:
            by_npt_only.setdefault(rec.mst_npt, []).append(rec)
        for d in (rec.cmnd, rec.ho_chieu):
            if d:
                by_doc_only.setdefault(_norm_no_zero(d), []).append(rec)

    res_mst = {k: _resolve_latest(v) for k, v in by_mst.items()}
    res_doc = {k: _resolve_latest(v) for k, v in by_doc.items()}

    # Chuẩn bị bản đồ phát hiện trùng trên 05-3 theo 3 khóa: MST NPT, Số giấy tờ, và (Họ tên + Ngày sinh)
    # Mỗi entry lưu (row_idx_in_list, mst_nnt, ym_from, ym_to)
    decl_entries_by_key: Dict[str, List[Tuple[int, str, Optional[Tuple[int, int]], Optional[Tuple[int, int]]]]] = {}
    row_keys: List[List[str]] = []

    for idx, r in enumerate(rows_053):
        ym_f = parse_ym(r.tu_thang)
        ym_t = parse_ym(r.den_thang)
        dob = parse_date_flexible(r.ngay_sinh)
        keys: List[str] = []
        if r.mst_npt:
            keys.append(f"MST:{r.mst_npt}")
        if r.so_gt:
            keys.append(f"DOC:{_norm_no_zero(r.so_gt)}")
        if r.ho_ten_npt and dob is not None:
            keys.append(f"NAME:{_norm_name(r.ho_ten_npt)}|{dob.isoformat()}")
        row_keys.append(keys)
        for k in keys:
            decl_entries_by_key.setdefault(k, []).append((idx, r.mst_nnt, ym_f, ym_t))

    results: List[NptAuditResultRow] = []
    per_nnt: Dict[str, List[int]] = {}

    for idx, r in enumerate(rows_053):
        ym_f = parse_ym(r.tu_thang)
        ym_t = parse_ym(r.den_thang)
        decl_m = (ym_idx(ym_t) - ym_idx(ym_f) + 1) if (ym_f and ym_t and ym_idx(ym_t) >= ym_idx(ym_f)) else 0
        eff_year = tax_year or (ym_f[0] if ym_f else (ym_t[0] if ym_t else datetime.date.today().year))

        # 1. KIỂM TRA ĐỘ TUỔI NPT
        dob = parse_date_flexible(r.ngay_sinh)
        age_val, age_status, is_age_ok = evaluate_dependent_age(dob, r.quan_he, r.ho_ten_npt, eff_year, ym_f)

        # 2. KIỂM TRA TRÙNG NPT (TRÊN 05-3 & TRÊN HỆ THỐNG TMS)
        dup_reasons: List[str] = []
        other_nnts_053: Set[str] = set()
        same_nnt_overlap = 0

        for k in row_keys[idx]:
            for o_idx, o_mst_nnt, o_f, o_t in decl_entries_by_key.get(k, []):
                if o_idx == idx:
                    continue
                ov = _intervals_overlap(ym_f, ym_t, o_f, o_t)
                if ov > 0:
                    if o_mst_nnt and r.mst_nnt and o_mst_nnt != r.mst_nnt:
                        other_nnts_053.add(o_mst_nnt)
                    elif o_mst_nnt == r.mst_nnt:
                        same_nnt_overlap = max(same_nnt_overlap, ov)

        if same_nnt_overlap > 0:
            dup_reasons.append(f"TRÙNG CÙNG NNT TRÊN 05-3 (giao nhau {same_nnt_overlap} tháng)")
        if other_nnts_053:
            dup_reasons.append(f"TRÙNG NHIỀU NNT TRÊN 05-3 (với MST {', '.join(sorted(other_nnts_053)[:2])})")

        # Kiểm tra trùng với NNT khác có thời gian giao nhau trên TMS
        if tms_rows and ym_f and ym_t:
            tms_other_nnts: Set[str] = set()
            candidates: List[TmsNptRow] = []
            if r.mst_npt and r.mst_npt in by_npt_only:
                candidates.extend(by_npt_only[r.mst_npt])
            if r.so_gt and _norm_no_zero(r.so_gt) in by_doc_only:
                candidates.extend(by_doc_only[_norm_no_zero(r.so_gt)])
            # Gom theo từng MST NNT khác và lấy bản hiệu lực mới nhất của NNT đó
            other_grouped: Dict[str, List[TmsNptRow]] = {}
            for cand in candidates:
                if cand.mst_nnt and cand.mst_nnt != r.mst_nnt:
                    other_grouped.setdefault(cand.mst_nnt, []).append(cand)
            for o_mst, o_recs in other_grouped.items():
                o_eff = _resolve_latest(o_recs)
                if _intervals_overlap(ym_f, ym_t, o_eff.tu_thang, o_eff.den_thang) > 0:
                    tms_other_nnts.add(o_mst)
            if tms_other_nnts:
                dup_reasons.append(f"TRÙNG NNT KHÁC TRÊN TMS (MST {', '.join(sorted(tms_other_nnts)[:2])})")

        is_dup = len(dup_reasons) > 0
        dup_status = "; ".join(dup_reasons) if is_dup else "Không trùng"

        # 3. ĐỐI CHIẾU THỜI GIAN & TRẠNG THÁI VỚI TMS
        matched: Optional[TmsNptRow] = None
        if r.mst_nnt and r.mst_npt and (r.mst_nnt, r.mst_npt) in res_mst:
            matched = res_mst[(r.mst_nnt, r.mst_npt)]
        elif r.mst_nnt and r.so_gt and (r.mst_nnt, _norm_no_zero(r.so_gt)) in res_doc:
            matched = res_doc[(r.mst_nnt, _norm_no_zero(r.so_gt))]

        reasons: List[str] = []
        tms_f_s = ""
        tms_t_s = ""
        tms_st_s = ""
        valid_m = decl_m

        if tms_rows:
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
                    valid_m = _intervals_overlap(ym_f, ym_t, matched.tu_thang, matched.den_thang)
                    if valid_m < decl_m:
                        reasons.append(
                            "NGOÀI THỜI GIAN ĐĂNG KÝ TMS" if valid_m == 0
                            else f"VƯỢT THỜI GIAN ĐĂNG KÝ ({decl_m - valid_m} tháng)"
                        )

        if not is_age_ok:
            reasons.append(age_status)
        if is_dup:
            reasons.extend(dup_reasons)

        is_ok = len(reasons) == 0
        status = "KHỚP" if is_ok else "; ".join(reasons)
        disallowed_m = max(0, decl_m - valid_m)

        results.append(
            NptAuditResultRow(
                src=r,
                decl_months=decl_m,
                age=age_val,
                age_status=age_status,
                is_age_ok=is_age_ok,
                dup_status=dup_status,
                is_dup=is_dup,
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
