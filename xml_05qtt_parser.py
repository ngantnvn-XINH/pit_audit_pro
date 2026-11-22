# -*- coding: utf-8 -*-
"""
Module đọc dữ liệu Quyết toán thuế TNCN (05/QTT-TNCN) từ:
1. File XML HTKK / eTax (hỗ trợ cả TT92/156 maTKhai=395 và TT80 maTKhai=953).
2. File Excel đã xuất chuẩn 4 sheet ('TK', '05-1', '05-2', '05-3') như QTT_TNCN_YYYY.xlsx.
"""

from __future__ import annotations

import datetime
import os
import re
import xml.etree.ElementTree as ET
from dataclasses import dataclass, field
from typing import Any, Dict, List, Optional

import openpyxl


def _strip_ns(tag: str) -> str:
    return tag.split("}", 1)[-1] if "}" in tag else tag


def _clean_str(val: Any) -> str:
    if val is None:
        return ""
    s = str(val).replace("\xa0", " ").replace("\u200b", "")
    return re.sub(r"\s+", " ", s).strip()


def _clean_id(val: Any) -> str:
    s = _clean_str(val)
    if s.endswith(".0") and s[:-2].isdigit():
        s = s[:-2]
    return s.replace(" ", "").upper()


def _to_float(val: Any) -> float:
    if val is None or isinstance(val, bool):
        return 0.0
    if isinstance(val, (int, float)):
        return float(val)
    s = _clean_str(val)
    if not s:
        return 0.0
    # Xử lý số có dấu phân cách
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


@dataclass
class Row051:
    stt: int
    ho_ten: str
    mst: str
    cccd: str
    uy_quyen: bool
    tong_tnct: float
    tn_mien_giam: float
    giam_tru_gia_canh: float
    tu_thien: float
    bao_hiem: float
    quy_huu_tri: float
    tn_tinh_thue: float
    thue_da_khau_tru: float
    thue_phai_nop: float
    thue_nop_them: float
    thue_nop_thua: float


@dataclass
class Row052:
    stt: int
    ho_ten: str
    mst: str
    cccd: str
    khong_cu_tru: bool
    tong_tnct: float
    tn_mien_giam: float
    thue_da_khau_tru: float


@dataclass
class Row053:
    stt: int
    ho_ten_nnt: str
    mst_nnt: str
    ho_ten_npt: str
    ngay_sinh: str
    mst_npt: str
    loai_gt: str
    so_gt: str
    quan_he: str
    tu_thang: str
    den_thang: str


@dataclass
class PitDeclarationData:
    source_path: str
    year: int
    mst_cqct: str = ""
    ten_cqct: str = ""
    so_lan: int = 0
    ngay_lap: str = ""
    is_tt156: bool = False
    rows_051: List[Row051] = field(default_factory=list)
    rows_052: List[Row052] = field(default_factory=list)
    rows_053: List[Row053] = field(default_factory=list)


def _child_text(elem: ET.Element, tag_names: List[str]) -> str:
    wanted = {t.lower() for t in tag_names}
    for ch in elem:
        if _strip_ns(ch.tag).lower() in wanted:
            return _clean_str(ch.text)
    return ""


def _find_first_text(root: ET.Element, tag_names: List[str]) -> str:
    wanted = {t.lower() for t in tag_names}
    for el in root.iter():
        if _strip_ns(el.tag).lower() in wanted and el.text and el.text.strip():
            return _clean_str(el.text)
    return ""


def parse_05qtt_xml(xml_path: str) -> PitDeclarationData:
    """Đọc trực tiếp file XML tờ khai 05/QTT-TNCN (cả TT92/156 và TT80)."""
    tree = ET.parse(xml_path)
    root = tree.getroot()

    ma_tk = _find_first_text(root, ["maTKhai"])
    is_tt156 = ma_tk in ("395", "05/KK-TNCN", "05/QTT-TNCN_156")
    ky_kk = _find_first_text(root, ["kyKKhai"])
    m_yr = re.search(r"(20\d{2})", ky_kk)
    if not m_yr:
        den_ngay = _find_first_text(root, ["kyKKhaiDenNgay"])
        m_yr = re.search(r"(20\d{2})", den_ngay)
    yr = int(m_yr.group(1)) if m_yr else datetime.date.today().year
    if yr <= 2020 and ma_tk == "395":
        is_tt156 = True

    so_lan_s = _find_first_text(root, ["soLan", "soLanBoSung", "lanBoSung"])
    so_lan = int(so_lan_s) if so_lan_s.isdigit() else 0

    data = PitDeclarationData(
        source_path=xml_path,
        year=yr,
        mst_cqct=_clean_id(_find_first_text(root, ["mst"])),
        ten_cqct=_find_first_text(root, ["tenNNT"]),
        so_lan=so_lan,
        ngay_lap=_find_first_text(root, ["ngayLapTKhai"]),
        is_tt156=is_tt156,
    )

    for el in root.iter():
        tname = _strip_ns(el.tag).lower()
        if "05_1" in tname or "pluc_05_1" in tname:
            idx = 0
            for item in el.iter():
                if _strip_ns(item.tag).lower() in ("ctietcnhan", "cnhan"):
                    ho_ten = _child_text(item, ["ct07", "hoVaTen"])
                    mst = _clean_id(_child_text(item, ["ct08", "mst"]))
                    if not ho_ten and not mst:
                        continue
                    idx += 1
                    cccd = _clean_id(_child_text(item, ["ct09", "soCMND"]))
                    if is_tt156:
                        uy_q = _child_text(item, ["ct10"]).lower() in ("1", "true", "x", "có")
                        tong_tn = _to_float(_child_text(item, ["ct11"]))
                        tn_mg = _to_float(_child_text(item, ["ct12"]))
                        gt_gc = _to_float(_child_text(item, ["ct14"]))
                        tu_thien = _to_float(_child_text(item, ["ct15"]))
                        bh = _to_float(_child_text(item, ["ct16"]))
                        huu_tri = _to_float(_child_text(item, ["ct17"]))
                        tntt = _to_float(_child_text(item, ["ct18"]))
                        thue_kt = _to_float(_child_text(item, ["ct19"]))
                        thue_pn = _to_float(_child_text(item, ["ct21"]))
                        thue_them = _to_float(_child_text(item, ["ct22"]))
                        thue_thua = _to_float(_child_text(item, ["ct23"]))
                    else:
                        uy_q = _child_text(item, ["ct10"]).lower() in ("1", "true", "x", "có")
                        tong_tn = _to_float(_child_text(item, ["ct11"]))
                        tn_mg = _to_float(_child_text(item, ["ct12"])) + _to_float(_child_text(item, ["ct13"]))
                        gt_gc = _to_float(_child_text(item, ["ct15"]))
                        tu_thien = _to_float(_child_text(item, ["ct16"]))
                        bh = _to_float(_child_text(item, ["ct17"]))
                        huu_tri = _to_float(_child_text(item, ["ct18"]))
                        tntt = _to_float(_child_text(item, ["ct19"]))
                        thue_kt = _to_float(_child_text(item, ["ct20"]))
                        thue_pn = _to_float(_child_text(item, ["ct23"]))
                        thue_them = _to_float(_child_text(item, ["ct24"]))
                        thue_thua = _to_float(_child_text(item, ["ct25"]))
                    data.rows_051.append(
                        Row051(
                            stt=idx, ho_ten=ho_ten, mst=mst, cccd=cccd, uy_quyen=uy_q,
                            tong_tnct=tong_tn, tn_mien_giam=tn_mg, giam_tru_gia_canh=gt_gc,
                            tu_thien=tu_thien, bao_hiem=bh, quy_huu_tri=huu_tri,
                            tn_tinh_thue=tntt, thue_da_khau_tru=thue_kt,
                            thue_phai_nop=thue_pn, thue_nop_them=thue_them, thue_nop_thua=thue_thua,
                        )
                    )
        elif "05_2" in tname or "pluc_05_2" in tname:
            idx = 0
            for item in el.iter():
                if _strip_ns(item.tag).lower() in ("ctietcnhan", "cnhan"):
                    ho_ten = _child_text(item, ["ct07", "hoVaTen"])
                    mst = _clean_id(_child_text(item, ["ct08", "mst"]))
                    if not ho_ten and not mst:
                        continue
                    idx += 1
                    cccd = _clean_id(_child_text(item, ["ct09", "soCMND"]))
                    kct = _child_text(item, ["ct10"]).lower() in ("1", "true", "x", "có")
                    tong_tn = _to_float(_child_text(item, ["ct11"]))
                    tn_mg = _to_float(_child_text(item, ["ct12"]))
                    thue_kt = _to_float(_child_text(item, ["ct14", "ct15", "ct13"]))
                    data.rows_052.append(
                        Row052(
                            stt=idx, ho_ten=ho_ten, mst=mst, cccd=cccd, khong_cu_tru=kct,
                            tong_tnct=tong_tn, tn_mien_giam=tn_mg, thue_da_khau_tru=thue_kt,
                        )
                    )
        elif "05_3" in tname or "pluc_05_3" in tname:
            idx = 0
            for item in el.iter():
                if _strip_ns(item.tag).lower() in ("ctietnpt", "ctietcnhan", "npt"):
                    ten_nnt = _child_text(item, ["ct07", "hoVaTenNNT"])
                    mst_nnt = _clean_id(_child_text(item, ["ct08", "mstNNT"]))
                    ten_npt = _child_text(item, ["ct09", "hoVaTenNPT"])
                    if not ten_npt and not mst_nnt:
                        continue
                    idx += 1
                    data.rows_053.append(
                        Row053(
                            stt=idx,
                            ho_ten_nnt=ten_nnt,
                            mst_nnt=mst_nnt,
                            ho_ten_npt=ten_npt,
                            ngay_sinh=_child_text(item, ["ct10", "ngaySinh"]),
                            mst_npt=_clean_id(_child_text(item, ["ct11", "mstNPT"])),
                            loai_gt=_child_text(item, ["ct12", "loaiGiayTo"]),
                            so_gt=_clean_id(_child_text(item, ["ct13", "soGiayTo"])),
                            quan_he=_child_text(item, ["ct14", "quanHe"]),
                            tu_thang=_child_text(item, ["ct15", "tuThang"]),
                            den_thang=_child_text(item, ["ct16", "denThang"]),
                        )
                    )

    return data


def parse_05qtt_excel(excel_path: str) -> PitDeclarationData:
    """Đọc file Excel QTT TNCN (có các sheet '05-1', '05-2', '05-3')."""
    wb = openpyxl.load_workbook(excel_path, data_only=True)
    m_yr = re.search(r"(20\d{2})", os.path.basename(excel_path))
    yr = int(m_yr.group(1)) if m_yr else datetime.date.today().year

    data = PitDeclarationData(source_path=excel_path, year=yr)

    if "05-1" in wb.sheetnames:
        ws = wb["05-1"]
        idx = 0
        for r in range(4, ws.max_row + 1):
            ho_ten = _clean_str(ws.cell(r, 2).value)
            mst = _clean_id(ws.cell(r, 3).value)
            if not ho_ten and not mst:
                continue
            if "TỔNG CỘNG" in (str(ws.cell(r, 1).value or "") + " " + ho_ten).upper():
                continue
            idx += 1
            uy_q = _clean_str(ws.cell(r, 5).value).lower() in ("1", "true", "x", "có")
            data.rows_051.append(
                Row051(
                    stt=idx,
                    ho_ten=ho_ten,
                    mst=mst,
                    cccd=_clean_id(ws.cell(r, 4).value),
                    uy_quyen=uy_q,
                    tong_tnct=_to_float(ws.cell(r, 8).value),
                    tn_mien_giam=_to_float(ws.cell(r, 9).value) + _to_float(ws.cell(r, 10).value),
                    giam_tru_gia_canh=_to_float(ws.cell(r, 12).value),
                    tu_thien=_to_float(ws.cell(r, 13).value),
                    bao_hiem=_to_float(ws.cell(r, 14).value),
                    quy_huu_tri=_to_float(ws.cell(r, 15).value),
                    tn_tinh_thue=_to_float(ws.cell(r, 16).value),
                    thue_da_khau_tru=_to_float(ws.cell(r, 18).value),
                    thue_phai_nop=_to_float(ws.cell(r, 20).value),
                    thue_nop_them=_to_float(ws.cell(r, 21).value),
                    thue_nop_thua=_to_float(ws.cell(r, 22).value),
                )
            )

    if "05-2" in wb.sheetnames:
        ws = wb["05-2"]
        idx = 0
        for r in range(4, ws.max_row + 1):
            ho_ten = _clean_str(ws.cell(r, 2).value)
            mst = _clean_id(ws.cell(r, 3).value)
            if not ho_ten and not mst:
                continue
            if "TỔNG CỘNG" in (str(ws.cell(r, 1).value or "") + " " + ho_ten).upper():
                continue
            idx += 1
            kct = _clean_str(ws.cell(r, 5).value).lower() in ("1", "true", "x", "có")
            data.rows_052.append(
                Row052(
                    stt=idx,
                    ho_ten=ho_ten,
                    mst=mst,
                    cccd=_clean_id(ws.cell(r, 4).value),
                    khong_cu_tru=kct,
                    tong_tnct=_to_float(ws.cell(r, 7).value),
                    tn_mien_giam=_to_float(ws.cell(r, 8).value),
                    thue_da_khau_tru=_to_float(ws.cell(r, 11).value),
                )
            )

    if "05-3" in wb.sheetnames:
        ws = wb["05-3"]
        idx = 0
        for r in range(4, ws.max_row + 1):
            ten_nnt = _clean_str(ws.cell(r, 2).value)
            mst_nnt = _clean_id(ws.cell(r, 3).value)
            ten_npt = _clean_str(ws.cell(r, 4).value)
            if not ten_npt and not mst_nnt:
                continue
            if "TỔNG CỘNG" in (str(ws.cell(r, 1).value or "") + " " + ten_nnt).upper():
                continue
            idx += 1
            tu_t = _clean_str(ws.cell(r, 10).value)
            den_t = _clean_str(ws.cell(r, 11).value)
            m_y = re.search(r"(20\d{2})", tu_t or den_t)
            if m_y and not m_yr:
                data.year = int(m_y.group(1))
            data.rows_053.append(
                Row053(
                    stt=idx,
                    ho_ten_nnt=ten_nnt,
                    mst_nnt=mst_nnt,
                    ho_ten_npt=ten_npt,
                    ngay_sinh=_clean_str(ws.cell(r, 5).value),
                    mst_npt=_clean_id(ws.cell(r, 6).value),
                    loai_gt=_clean_str(ws.cell(r, 7).value),
                    so_gt=_clean_id(ws.cell(r, 8).value),
                    quan_he=_clean_str(ws.cell(r, 9).value),
                    tu_thang=tu_t,
                    den_thang=den_t,
                )
            )

    wb.close()
    return data


def load_pit_declaration(path: str) -> PitDeclarationData:
    ext = os.path.splitext(path)[1].lower()
    if ext == ".xml":
        return parse_05qtt_xml(path)
    return parse_05qtt_excel(path)
