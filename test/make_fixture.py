"""Generate a deliberately broken financial-model fixture for the audit engine.

Run: python test/make_fixture.py test/fixtures/broken-model.xlsx
The fixture covers: formula errors, broken refs, formula inconsistency,
hardcoded constants, percent-format misuse, BS imbalance, RE roll-forward
break, cash tie-out break, CF subtotal break, ratio conventions (hardcoded 365,
mixed periods, pre-revenue), data cleaning, and an outlier growth rate.
"""
from __future__ import annotations

import os
import sys

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

BLUE = Font(color="0000FF")
YELLOW = PatternFill("solid", fgColor="FFFF00")


def main(path: str):
    wb = Workbook()
    ws = wb.active
    ws.title = "三表"
    headers = ["项目", "2022", "2023", "2024", "2026Q1"]
    for col, text in enumerate(headers, start=1):
        ws.cell(row=1, column=col, value=text).font = Font(bold=True)

    def put(row, label, values, *, formula=None, fmt=None):
        ws.cell(row=row, column=1, value=label)
        for idx, col in enumerate(range(2, 6)):
            cell = ws.cell(row=row, column=col)
            if formula is not None:
                cell.value = formula(idx, get_column_letter(col))
            else:
                cell.value = values[idx]
            if fmt:
                cell.number_format = fmt

    # Balance sheet (rows 3-15): 2024 is deliberately out of balance by 100k.
    put(3, "货币资金", [1_000_000, 2_000_000, 3_000_000, 3_500_000], fmt="#,##0.00")
    put(4, "应收账款", [500_000, 600_000, 700_000, 800_000], fmt="#,##0.00")
    put(5, "存货", [300_000, 350_000, 400_000, 450_000], fmt="#,##0.00")
    put(6, "流动资产合计", None, formula=lambda i, c: f"=SUM({c}3:{c}5)", fmt="#,##0.00")
    put(7, "固定资产", [2_000_000, 2_100_000, 2_200_000, 2_300_000], fmt="#,##0.00")
    put(8, "资产总计", None, formula=lambda i, c: f"={c}6+{c}7", fmt="#,##0.00")
    put(9, "短期借款", [800_000, 700_000, 600_000, 500_000], fmt="#,##0.00")
    put(10, "应付账款", [400_000, 450_000, 500_000, 550_000], fmt="#,##0.00")
    put(11, "负债合计", None, formula=lambda i, c: f"={c}9+{c}10", fmt="#,##0.00")
    put(12, "实收资本", [2_000_000, 2_000_000, 2_000_000, 2_000_000], fmt="#,##0.00")
    put(13, "未分配利润", [600_000, 1_900_000, 3_100_000, 4_000_000], fmt="#,##0.00")
    put(14, "所有者权益合计", None, formula=lambda i, c: f"={c}12+{c}13", fmt="#,##0.00")
    # 2024 gets an extra +100000 -> BS imbalance in 2024 only.
    ws["D15"] = "=D11+D14+100000"
    ws["B15"] = "=B11+B14"
    ws["C15"] = "=C11+C14"
    ws["E15"] = "=E11+E14"
    ws["A15"] = "负债和所有者权益总计"
    for col in "BCDE":
        ws[f"{col}15"].number_format = "#,##0.00"

    # Income statement (rows 17-21)
    ws["A17"] = "营业收入"
    for col, value in zip("BCDE", [90_000, 120_000, 6_000_000, 1_500_000]):
        ws[f"{col}17"] = value
        ws[f"{col}17"].number_format = "#,##0.00"
    ws["A18"] = "营业成本"
    for col, value in zip("BCDE", [60_000, 80_000, 4_000_000, 1_000_000]):
        ws[f"{col}18"] = value
        ws[f"{col}18"].number_format = "#,##0.00"
    ws["A19"] = "净利润"
    for col in "BCDE":
        ws[f"{col}19"] = f"={col}17-{col}18-20000"
        ws[f"{col}19"].number_format = "#,##0.00"
    ws["A20"] = "分配股利、利润或偿付利息所支付的现金"
    for col, value in zip("BCDE", [0, 30_000, 0, 0]):
        ws[f"{col}20"] = value
        ws[f"{col}20"].number_format = "#,##0.00"

    # Cash flow (rows 23-29)
    ws["A23"] = "经营活动产生的现金流量净额"
    for col, value in zip("BCDE", [50_000, 60_000, 700_000, 350_000]):
        ws[f"{col}23"] = value
        ws[f"{col}23"].number_format = "#,##0.00"
    ws["A24"] = "投资活动产生的现金流量净额"
    for col, value in zip("BCDE", [-20_000, -30_000, -400_000, -200_000]):
        ws[f"{col}24"] = value
        ws[f"{col}24"].number_format = "#,##0.00"
    ws["A25"] = "筹资活动产生的现金流量净额"
    for col in "BCDE":
        ws[f"{col}25"] = 0
        ws[f"{col}25"].number_format = "#,##0.00"
    ws["A26"] = "现金及现金等价物净增加额"
    for idx, col in enumerate("BCDE"):
        if col == "D":
            ws[f"{col}26"] = 200_000  # hardcoded, should be the sum (300,000)
        else:
            ws[f"{col}26"] = f"={col}23+{col}24+{col}25"
        ws[f"{col}26"].number_format = "#,##0.00"
    ws["A27"] = "期初现金及现金等价物余额"
    ws["B27"] = 900_000
    for prev, col in zip("BCDE", "CDE"):
        ws[f"{col}27"] = f"={prev}28"
    for col in "BCDE":
        ws[f"{col}27"].number_format = "#,##0.00"
    ws["A28"] = "期末现金及现金等价物余额"
    for col in "BCDE":
        ws[f"{col}28"] = f"={col}27+{col}26"
        ws[f"{col}28"].number_format = "#,##0.00"

    # Broken reference + formula error.
    ws["H3"] = "='已删除的工作表'!A1"
    ws["H4"] = "=1/0"

    # Formula inconsistency: row 31 should be 4 identical formulas; one is offset.
    ws["A31"] = "测试公式一致性"
    for col in "BCDE":
        ws[f"{col}31"] = f"=SUM({col}3:{col}5)"
    ws["E31"] = "=SUM(E4:E5)"  # wrong reference

    # Hardcoded constants mixed with formulas (row 32: B:D formulas, E constant).
    ws["A32"] = "测试硬编码"
    for col in "BCD":
        ws[f"{col}32"] = f"={col}8*0.5"
    ws["E32"] = 1234567

    # Text-stored number + placeholder + blank in a numeric column.
    ws["A34"] = "测试数据清洗"
    ws["B34"] = "1234.5"  # text number
    ws["C34"] = "-"
    # D34 left blank

    # Ratio sheet -------------------------------------------------------
    rs = wb.create_sheet("比率表")
    for col, text in enumerate(["指标", "2022", "2023", "2024", "2026Q1"], start=1):
        rs.cell(row=1, column=col, value=text).font = Font(bold=True)
    # AR turnover days: hardcoded 365, tiny 2022 revenue -> absurd days.
    rs["A2"] = "应收账款周转天数"
    for idx, col in enumerate("BCDE"):
        rs[f"{col}2"] = f"=365*'三表'!{col}4/'三表'!{col}17"
    # AP turnover days
    rs["A3"] = "应付账款周转天数"
    for col in "BCDE":
        rs[f"{col}3"] = f"=365*'三表'!{col}10/'三表'!{col}18"
    # Net margin (percent expected)
    rs["A4"] = "净利率"
    for col in "BCDE":
        rs[f"{col}4"] = f"='三表'!{col}19/'三表'!{col}17"
    # Money row wrongly formatted as percent (the user's real bug).
    rs["A5"] = "有息负债合计"
    for col in "BCDE":
        rs[f"{col}5"] = f"='三表'!{col}9"
        rs[f"{col}5"].number_format = "0.00%"
    # Days row wrongly formatted as percent.
    rs["A6"] = "现金转换周期"
    for col in "BCDE":
        rs[f"{col}6"] = f"={col}2+30-{col}3"
        rs[f"{col}6"].number_format = "0.00%"
    # Ratio row with a division that can be #DIV/0! in the first period.
    rs["A7"] = "流动比率"
    for col in "BCDE":
        rs[f"{col}7"] = f"='三表'!{col}6/'三表'!{col}9"
    rs["B7"] = "='三表'!B6/0"

    # Mixed periods are already present (2022..2024 + 2026Q1).

    # Data-cleaning sheet ----------------------------------------------
    ds = wb.create_sheet("数据")
    ds["A1"] = "科目"
    ds["B1"] = "金额"
    ds["A2"] = "营业收入"
    ds["B2"] = "1234.5"  # text-stored number
    ds["A3"] = "营业成本"
    ds["B3"] = "-"       # placeholder
    ds["A4"] = "销售费用"  # blank amount
    ds["A5"] = "管理费用"
    ds["B5"] = 100
    ds["A6"] = "管理费用"
    ds["B6"] = 100       # duplicate row
    ds["A7"] = "财务费用"
    ds["B7"] = 50

    os.makedirs(os.path.dirname(os.path.abspath(path)), exist_ok=True)
    wb.save(path)
    print(f"wrote {path}")


if __name__ == "__main__":
    main(sys.argv[1] if len(sys.argv) > 1 else "test/fixtures/broken-model.xlsx")
