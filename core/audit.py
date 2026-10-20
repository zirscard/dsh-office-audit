#!/usr/bin/env python3
"""office_audit engine CLI.

Usage:
  python core/audit.py --input model.xlsx [--recalc recalculated.xlsx] \
      --out-json report.json --out-md report.md [--profile auto] \
      [--checks a,b] [--write-sheet copy.xlsx]

Prints one JSON object to stdout (the same payload written to --out-json).
Progress and errors go to stderr.
"""
from __future__ import annotations

import argparse
import json
import os
import sys
import time
import traceback

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from auditor import checks_finance as finance  # noqa: E402
from auditor import checks_generic as generic  # noqa: E402
from auditor.findings import AuditResult  # noqa: E402
from auditor.model import WorkbookModel  # noqa: E402
from auditor.report import render_markdown, write_audit_sheet  # noqa: E402
from auditor.util import normalize_label  # noqa: E402

ENGINE_VERSION = "0.2.0"

GENERIC_CHECKS = [
    "formula-errors",
    "broken-references",
    "formula-consistency",
    "hardcoded-values",
    "number-format",
    "data-cleaning",
    "outliers",
    "workbook-structure",
]
FINANCE_CHECKS = [
    "balance-sheet",
    "retained-earnings",
    "cash-tieout",
    "cash-flow-subtotals",
]
RATIO_CHECKS = ["ratio-conventions"]

CHECK_FUNCS = {
    "formula-errors": generic.check_formula_errors,
    "broken-references": generic.check_broken_references,
    "formula-consistency": generic.check_formula_consistency,
    "hardcoded-values": generic.check_hardcoded_values,
    "number-format": generic.check_number_format,
    "data-cleaning": generic.check_data_cleaning,
    "outliers": generic.check_outliers,
    "workbook-structure": generic.check_workbook_structure,
    "balance-sheet": finance.check_balance_sheet,
    "retained-earnings": finance.check_retained_earnings,
    "cash-tieout": finance.check_cash_tieout,
    "cash-flow-subtotals": finance.check_cash_flow_subtotals,
    "ratio-conventions": finance.check_ratio_conventions,
}


def _period_label(period) -> str:
    if period[0] == "quarter" and len(period) > 2:
        return f"{period[1]}Q{period[2]}"
    if period[0] == "month" and len(period) > 2:
        return f"{period[1]}-{period[2]:02d}"
    return str(period[1])


def detect_profile(model: WorkbookModel) -> str:
    labels = set()
    ratio_words = ("周转率", "周转天数", "毛利率", "净利率", "资产负债率", "流动比率", "速动比率", "ROE", "ROA")
    for sheet in model.iter_sheets():
        for _row, label in sheet.label_rows():
            norm = normalize_label(label)
            if norm:
                labels.add(norm)
    three_statement_words = ("资产总计", "资产合计", "负债合计", "所有者权益合计", "留存收益", "未分配利润",
                             "经营活动产生的现金流量净额", "现金及现金等价物")
    if any(word in label for label in labels for word in three_statement_words):
        return "financial-model"
    if any(word in label for label in labels for word in ratio_words):
        return "ratio-table"
    return "generic"


def check_list_for(profile: str, requested, verbose=False):
    if requested:
        unknown = [c for c in requested if c not in CHECK_FUNCS]
        if unknown:
            raise ValueError(
                f"unknown check(s): {', '.join(unknown)}; valid: {', '.join(sorted(CHECK_FUNCS))}"
            )
        return list(requested)
    if profile == "financial-model":
        return GENERIC_CHECKS + FINANCE_CHECKS + RATIO_CHECKS
    if profile == "ratio-table":
        return GENERIC_CHECKS + RATIO_CHECKS
    return GENERIC_CHECKS


def sheet_stats(model: WorkbookModel):
    stats = []
    for sheet in model.iter_sheets():
        formulas = 0
        constants = 0
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                cell = sheet.formula_cell(row, col)
                if cell.value is None:
                    continue
                if cell.data_type == "f" or (isinstance(cell.value, str) and cell.value.startswith("=")):
                    formulas += 1
                elif isinstance(cell.value, (int, float)) and not isinstance(cell.value, bool):
                    constants += 1
        period_cols = [
            {"column": col, "period": period[0], "label": _period_label(period)}
            for col, period in sorted(sheet.period_columns.items())
        ]
        stats.append({
            "name": sheet.name,
            "rows": sheet.max_row,
            "columns": sheet.max_column,
            "formulaCells": formulas,
            "constantCells": constants,
            "headerRow": sheet.header_row,
            "periodColumns": period_cols,
            "hidden": sheet.ws.sheet_state != "visible",
        })
    return stats


def main(argv=None) -> int:
    parser = argparse.ArgumentParser(description="Audit an Excel workbook")
    parser.add_argument("--input", required=True)
    parser.add_argument("--recalc", default=None)
    parser.add_argument("--out-json", required=True)
    parser.add_argument("--out-md", required=True)
    parser.add_argument("--profile", default="auto", choices=["auto", "financial-model", "ratio-table", "generic"])
    parser.add_argument("--checks", default="")
    parser.add_argument("--write-sheet", default=None)
    parser.add_argument("--verbose", action="store_true")
    args = parser.parse_args(argv)

    started = time.time()

    def log(message):
        if args.verbose:
            print(message, file=sys.stderr)

    try:
        log(f"loading {args.input}")
        model = WorkbookModel(args.input, recalc_path=args.recalc)
        profile = args.profile
        if profile == "auto":
            profile = detect_profile(model)
        log(f"profile={profile}")

        result = AuditResult(workbook=os.path.abspath(args.input), profile=profile)
        opts = {"max_examples": 8, "growth_threshold": 20.0, "revenue_threshold": 1_000_000}
        try:
            checks = check_list_for(profile, [c.strip() for c in args.checks.split(",") if c.strip()], args.verbose)
        except ValueError as exc:
            print(f"error: {exc}", file=sys.stderr)
            return 2

        for name in checks:
            log(f"check: {name}")
            CHECK_FUNCS[name](model, result, opts)

        result.sheets = sheet_stats(model)
        result.stats = {
            "durationMs": int((time.time() - started) * 1000),
            "formulaCells": sum(s["formulaCells"] for s in result.sheets),
            "constantCells": sum(s["constantCells"] for s in result.sheets),
            "checks": checks,
        }
        result.meta.update({
            "recalculated": args.recalc is not None,
            "engineVersion": ENGINE_VERSION,
            "generatedAt": time.strftime("%Y-%m-%dT%H:%M:%S"),
        })

        os.makedirs(os.path.dirname(os.path.abspath(args.out_json)), exist_ok=True)
        os.makedirs(os.path.dirname(os.path.abspath(args.out_md)), exist_ok=True)
        markdown = render_markdown(result, engine=f"dsh-office-audit {ENGINE_VERSION} / openpyxl")
        with open(args.out_md, "w", encoding="utf-8") as handle:
            handle.write(markdown)
        payload = result.to_dict()
        with open(args.out_json, "w", encoding="utf-8") as handle:
            json.dump(payload, handle, ensure_ascii=False, indent=2)

        if args.write_sheet:
            log(f"writing audit sheet to {args.write_sheet}")
            os.makedirs(os.path.dirname(os.path.abspath(args.write_sheet)), exist_ok=True)
            write_audit_sheet(result, args.input, args.write_sheet)

        print(json.dumps(payload, ensure_ascii=False))
        log("done")
        return 0
    except Exception:  # noqa: BLE001 - CLI boundary: report, never swallow
        traceback.print_exc()
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
