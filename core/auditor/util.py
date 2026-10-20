"""Shared helpers for the office-audit engine (no third-party deps beyond openpyxl)."""
from __future__ import annotations

import math
import re
from datetime import date, datetime
from decimal import Decimal

# Excel error values that can surface as cached results.
ERROR_VALUES = {
    "#REF!",
    "#DIV/0!",
    "#VALUE!",
    "#NAME?",
    "#N/A",
    "#NUM!",
    "#NULL!",
    "#SPILL!",
    "#CALC!",
    "#GETTING_DATA",
    "#FIELD!",
    "#BLOCKED!",
    "#CONNECT!",
    "#UNKNOWN!",
}

# Placeholder text that finance workbooks use for "zero / not applicable".
PLACEHOLDER_TEXTS = {"-", "—", "–", "－", "N/A", "n/a", "NA", "不适用", "—", "无"}

# Ratio-like row labels (percent-typed rows).
RATIO_WORDS = (
    "毛利率",
    "净利率",
    "收益率",
    "报酬率",
    "周转率",
    "负债率",
    "现金比率",
    "流动比率",
    "速动比率",
    "保障倍数",
    "费率",
    "增长率",
    "税率",
    "利率",
    "占比",
    "比例",
    "比重",
    "毛利率",
    "净利率",
    "收益率",
    "报酬率",
    "周转率",
    "周转天数",
    "负债率",
    "保障倍数",
    "流动比率",
    "速动比率",
    "现金比率",
    "每股",
    "EBITDA率",
    "ROE",
    "ROA",
    "ROIC",
)
# Rows expressed in days (must NOT be formatted as percent).
DAY_WORDS = ("周转天数", "天数", "账龄", "周期", "天数）", "（天）", "(天)")
# Rows expressed as multiples.
MULTIPLE_WORDS = ("倍数", "倍", "周转率")
# Rows expressed as money.
MONEY_WORDS = (
    "资产",
    "负债",
    "权益",
    "收入",
    "营收",
    "成本",
    "费用",
    "利润",
    "现金",
    "存款",
    "借款",
    "应付",
    "应收",
    "存货",
    "固定资产",
    "无形资产",
    "商誉",
    "投资",
    "总额",
    "合计",
    "余额",
    "金额",
    "资本",
    "支出",
    "现金流",
    "折旧",
    "摊销",
    "税",
    "薪酬",
    "工资",
    "租金",
    "利息",
    "分红",
    "股利",
)

# Unambiguous money words: a percent format on these rows is a real defect.
STRONG_MONEY_WORDS = (
    "合计", "总额", "余额", "资产", "负债", "应收", "应付", "存货",
    "货币资金", "借款", "营业收入", "净利润", "现金流", "费用", "成本合计",
)

# Accounting labels used by the three-statement cross-checks.
LABELS = {
    "assets_total": ("资产总计", "资产合计", "总资产", "资产总额"),
    "liabilities_total": ("负债合计", "负债总计", "负债总额"),
    "equity_total": ("所有者权益合计", "股东权益合计", "所有者权益总计", "股东权益总计", "权益合计"),
    "liab_equity_total": ("负债和所有者权益总计", "负债及所有者权益总计", "负债和股东权益合计", "负债及股东权益总计"),
    "retained_earnings": ("留存收益", "未分配利润", "留存利润"),
    "net_income": ("净利润", "归属于母公司股东的净利润", "归母净利润", "税后利润"),
    "dividends": ("分配股利", "股利分配", "利润分配", "分红", "分配利润"),
    "cash_bs": ("货币资金", "现金及现金等价物", "现金及银行存款", "银行存款"),
    "cash_cf_end": ("期末现金及现金等价物余额", "期末现金余额", "期末现金"),
    "cash_cf_begin": ("期初现金及现金等价物余额", "期初现金余额", "期初现金"),
    "cash_cf_net": ("现金及现金等价物净增加额", "现金净增加额", "现金及现金等价物净变动额"),
    "cf_operating": ("经营活动产生的现金流量净额", "经营活动现金流量净额", "经营活动净现金流"),
    "cf_investing": ("投资活动产生的现金流量净额", "投资活动现金流量净额", "投资活动净现金流"),
    "cf_financing": ("筹资活动产生的现金流量净额", "筹资活动现金流量净额", "筹资活动净现金流"),
    "revenue": ("营业收入", "营收", "销售收入", "营业总收入"),
    "total_costs": ("营业成本", "营业总成本"),
}

# Derived / analysis sheets are not raw statements: their rows are shares or
# ratios, so amount-format, growth and three-statement checks do not apply.
DERIVED_SHEET_WORDS = (
    "common-size", "commonsize", "common size", "结构", "占比", "垂直",
    "同比", "勾稽校验", "模型审计", "敏感性", "sensitivity",
)


def is_derived_sheet_name(name) -> bool:
    norm = str(name).lower().replace(" ", "")
    return any(word in norm for word in DERIVED_SHEET_WORDS)


CJK_RE = re.compile(r"[\u4e00-\u9fff]")
YEAR_RE = re.compile(r"(?<!\d)(20\d{2})(?!\d)")
QUARTER_RE = re.compile(r"(?:20\d{2}\s*)?(?:Q|q|第)?\s*([1-4])\s*(?:季|季度)?")
PERCENT_HINT_RE = re.compile(r"%|％|率|占比|比重")
DAYS_HINT_RE = re.compile(r"天|days?|账龄|周期", re.IGNORECASE)
MULTIPLE_HINT_RE = re.compile(r"倍|times", re.IGNORECASE)
CURRENCY_HINT_RE = re.compile(r"[¥$€£￥]|元|万元|亿元|千元")


def is_formula(value) -> bool:
    return isinstance(value, str) and value.startswith("=")


def is_number(value) -> bool:
    return isinstance(value, (int, float, Decimal)) and not isinstance(value, bool)


def to_number(value):
    """Best-effort numeric conversion; returns None when not numeric."""
    if is_number(value):
        try:
            f = float(value)
        except (TypeError, ValueError):
            return None
        if math.isnan(f) or math.isinf(f):
            return None
        return f
    if isinstance(value, str):
        text = value.strip().replace(",", "").replace("，", "")
        if text in ("", "-", "—", "–", "N/A", "n/a", "NA"):
            return None
        pct = text.endswith("%") or text.endswith("％")
        if pct:
            text = text[:-1]
        try:
            f = float(text)
        except ValueError:
            return None
        return f / 100.0 if pct else f
    return None


def is_error_value(value) -> bool:
    return isinstance(value, str) and value.strip().upper() in ERROR_VALUES


def normalize_label(text) -> str:
    if not isinstance(text, str):
        return ""
    text = text.replace("\u3000", " ").replace("\xa0", " ")
    text = re.sub(r"\s+", "", text)
    text = re.sub(r"[（(].*?[）)]", "", text)
    text = text.rstrip(":：")
    return text


def label_has(label: str, words) -> bool:
    norm = normalize_label(label)
    return any(word in norm for word in words)


def find_label(labels, words, *, exclude=()) -> tuple[int, str] | None:
    """Return (row_index_in_list, label) for the first label matching any word."""
    for idx, label in labels:
        norm = normalize_label(label)
        if any(word in norm for word in words) and not any(bad in norm for bad in exclude):
            return idx, label
    return None


def period_of_header(text):
    """Extract a (kind, key) period tuple from a header cell, or None."""
    if text is None:
        return None
    s = str(text).strip()
    if s == "":
        return None
    year_match = YEAR_RE.search(s)
    year = int(year_match.group(1)) if year_match else None
    if year is None:
        return None
    # Quarter or month suffix, e.g. 2026Q1 / 2026年一季度 / 2026-03 / 2026年3月
    q = re.search(r"(?:Q|q)\s*([1-4])", s)
    if q:
        return ("quarter", year, int(q.group(1)))
    m = re.search(r"(?:第?\s*([1-4])\s*季)", s)
    if m:
        return ("quarter", year, int(m.group(1)))
    month = re.search(r"(\d{1,2})\s*月", s)
    if month and 1 <= int(month.group(1)) <= 12 and year_match:
        return ("month", year, int(month.group(1)))
    return ("year", year)


def is_percent_format(fmt) -> bool:
    return isinstance(fmt, str) and ("%" in fmt or "0.0%" in fmt or "#,##0.0%" in fmt)


def format_compact(value) -> str:
    num = to_number(value)
    if num is None:
        return "" if value is None else str(value)
    if abs(num) >= 1e8:
        return f"{num / 1e8:,.2f}亿"
    if abs(num) >= 1e4:
        return f"{num / 1e4:,.2f}万"
    if abs(num) >= 1000:
        return f"{num:,.2f}"
    if num == int(num):
        return str(int(num))
    return f"{num:,.4g}"


def describe_cell(value) -> str:
    if value is None:
        return "空"
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (datetime, date)):
        return value.isoformat()
    if isinstance(value, str) and len(value) > 60:
        return value[:57] + "…"
    return format_compact(value) if is_number(value) else str(value)
