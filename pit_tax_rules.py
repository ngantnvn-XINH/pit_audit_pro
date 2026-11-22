# -*- coding: utf-8 -*-
"""
Động cơ chính sách & tính thuế Thu nhập cá nhân (TNCN) Việt Nam qua các thời kỳ (2015 - 2026+).
Căn cứ pháp lý:
- Luật Thuế TNCN 04/2007/QH12 & Luật sửa đổi 26/2012/QH13
- Thông tư 111/2013/TT-BTC & Thông tư 92/2015/TT-BTC
- Nghị quyết 954/2020/UBTVQH14 (hiệu lực từ kỳ tính thuế năm 2020: Bản thân 11tr/tháng, NPT 4,4tr/tháng)
- Nghị quyết 110/2025/UBTVQH15 (hiệu lực từ kỳ tính thuế năm 2026: Bản thân 15,5tr/tháng, NPT 6,2tr/tháng)
"""

from __future__ import annotations

from dataclasses import dataclass
from typing import Optional, Tuple


@dataclass(frozen=True)
class PitPolicyParams:
    year: int
    self_deduction_monthly: int
    dependent_deduction_monthly: int
    self_deduction_annual: int
    non_contract_threshold: int = 2_000_000
    flat_rate_resident: float = 0.10
    flat_rate_non_resident: float = 0.20
    housing_cap_ratio: float = 0.15
    uniform_cash_cap_annual: int = 5_000_000
    legal_basis: str = ""


def get_policy_for_year(year: Optional[int]) -> PitPolicyParams:
    """Trả về bộ tham số chính sách thuế TNCN chuẩn theo năm tính thuế."""
    y = year or 2024
    if y <= 2019:
        return PitPolicyParams(
            year=y,
            self_deduction_monthly=9_000_000,
            dependent_deduction_monthly=3_600_000,
            self_deduction_annual=108_000_000,
            legal_basis="Luật 26/2012/QH13 & TT 111/2013/TT-BTC",
        )
    if y <= 2025:
        return PitPolicyParams(
            year=y,
            self_deduction_monthly=11_000_000,
            dependent_deduction_monthly=4_400_000,
            self_deduction_annual=132_000_000,
            legal_basis="Nghị quyết 954/2020/UBTVQH14 & TT 111/2013/TT-BTC",
        )
    return PitPolicyParams(
        year=y,
        self_deduction_monthly=15_500_000,
        dependent_deduction_monthly=6_200_000,
        self_deduction_annual=186_000_000,
        legal_basis="Nghị quyết 110/2025/UBTVQH15",
    )


# Biểu thuế lũy tiến từng phần 7 bậc (áp dụng đến hết kỳ tính thuế 2025) - Phụ lục 01/PL-TNCN TT 111/2013/TT-BTC
# Mỗi phần tử: (Ngưỡng trên tháng, Thuế suất, Số trừ nhanh theo tháng)
PROGRESSIVE_7_BRACKETS_MONTHLY: Tuple[Tuple[float, float, int], ...] = (
    (5_000_000, 0.05, 0),
    (10_000_000, 0.10, 250_000),
    (18_000_000, 0.15, 750_000),
    (32_000_000, 0.20, 1_650_000),
    (52_000_000, 0.25, 3_250_000),
    (80_000_000, 0.30, 5_850_000),
    (float("inf"), 0.35, 9_850_000),
)

# Biểu thuế lũy tiến từng phần 5 bậc (áp dụng từ kỳ tính thuế năm 2026) - Luật Thuế TNCN 109/2025/QH15
PROGRESSIVE_5_BRACKETS_MONTHLY: Tuple[Tuple[float, float, int], ...] = (
    (10_000_000, 0.05, 0),
    (30_000_000, 0.10, 500_000),
    (60_000_000, 0.20, 3_500_000),
    (100_000_000, 0.30, 9_500_000),
    (float("inf"), 0.35, 14_500_000),
)

# Giữ alias tương thích ngược
PROGRESSIVE_MONTHLY_BRACKETS = PROGRESSIVE_7_BRACKETS_MONTHLY


def calc_progressive_pit_monthly(monthly_taxable_income: float, year: int = 2024) -> int:
    """Tính thuế TNCN theo biểu lũy tiến từng phần cho 1 tháng (7 bậc <=2025; 5 bậc >=2026)."""
    if monthly_taxable_income <= 0:
        return 0
    brackets = PROGRESSIVE_5_BRACKETS_MONTHLY if year >= 2026 else PROGRESSIVE_7_BRACKETS_MONTHLY
    for upper, rate, quick_sub in brackets:
        if monthly_taxable_income <= upper:
            return max(0, round(monthly_taxable_income * rate - quick_sub))
    return 0


def calc_progressive_pit_annual(annual_taxable_income: float, months: int = 12, year: int = 2024) -> int:
    """
    Tính thuế TNCN quyết toán năm theo biểu lũy tiến từng phần (7 bậc <=2025; 5 bậc >=2026):
    - Mặc định chia bình quân 12 tháng (hoặc `months` tháng nếu quyết toán theo số tháng cư trú thực tế).
    """
    if annual_taxable_income <= 0 or months <= 0:
        return 0
    avg_month = annual_taxable_income / months
    brackets = PROGRESSIVE_5_BRACKETS_MONTHLY if year >= 2026 else PROGRESSIVE_7_BRACKETS_MONTHLY
    for upper, rate, quick_sub in brackets:
        if avg_month <= upper:
            return max(0, round((avg_month * rate - quick_sub) * months))
    return 0


def excel_formula_progressive_pit_annual(annual_taxable_cell: str, year: int = 2024) -> str:
    """
    Sinh công thức Excel chuẩn tính thuế TNCN quyết toán năm (7 bậc <=2025; 5 bậc >=2026).
    Tương thích mọi phiên bản Excel (không cần XLOOKUP/LET/IFS).
    """
    c = annual_taxable_cell
    if year >= 2026:
        return (
            f"=IF({c}<=0,0,ROUND("
            f"IF({c}<=120000000,{c}*0.05,"
            f"IF({c}<=360000000,{c}*0.1-6000000,"
            f"IF({c}<=720000000,{c}*0.2-42000000,"
            f"IF({c}<=1200000000,{c}*0.3-114000000,"
            f"{c}*0.35-174000000)))),0))"
        )
    return (
        f"=IF({c}<=0,0,ROUND("
        f"IF({c}<=60000000,{c}*0.05,"
        f"IF({c}<=120000000,{c}*0.1-3000000,"
        f"IF({c}<=216000000,{c}*0.15-9000000,"
        f"IF({c}<=384000000,{c}*0.2-19800000,"
        f"IF({c}<=624000000,{c}*0.25-39000000,"
        f"IF({c}<=960000000,{c}*0.3-70200000,"
        f"{c}*0.35-118200000)))))),0))"
    )


def calc_housing_taxable_amount(actual_housing: float, other_taxable_income: float) -> int:
    """
    Tính tiền thuê nhà tính vào TNCT theo Khoản 2 Điều 11 Thông tư 92/2015/TT-BTC:
    Tính theo số thực tế chi trả hộ nhưng không vượt quá 15% tổng TNCT phát sinh (chưa bao gồm tiền thuê nhà).
    """
    if actual_housing <= 0:
        return 0
    cap = max(0.0, other_taxable_income * 0.15)
    return round(min(actual_housing, cap))
