"""Generic workbook checks: formulas, references, formatting, data hygiene."""
from __future__ import annotations

import re

from openpyxl.formula import Tokenizer
from openpyxl.utils import column_index_from_string, get_column_letter

from .util import (
    CURRENCY_HINT_RE,
    ERROR_VALUES,
    STRONG_MONEY_WORDS,
    is_derived_sheet_name,
    PLACEHOLDER_TEXTS,
    describe_cell,
    format_compact,
    is_error_value,
    is_formula,
    is_number,
    label_has,
    normalize_label,
    to_number,
)

CELL_REF_RE = re.compile(r"^(\$?)([A-Za-z]{1,3})(\$?)(\d+)$")
MAX_FINDINGS_PER_RULE = 20


def cell_ref(row: int, col: int) -> str:
    return f"{get_column_letter(col)}{row}"


# ---------------------------------------------------------------------------
# reference parsing
# ---------------------------------------------------------------------------
def split_sheet_ref(ref: str):
    """Split an A1-style reference into (sheet_prefix or None, rest).

    Handles quoted sheet names ('My Sheet'!A1), 3-D prefixes (Sheet1:Sheet3!A1)
    and external links ([1]Sheet1!A1 -> prefix '[1]Sheet1').
    """
    s = ref.strip()
    if "!" not in s:
        return None, s
    in_quote = False
    last_bang = -1
    for i, ch in enumerate(s):
        if ch == "'":
            in_quote = not in_quote
        elif ch == "!" and not in_quote:
            last_bang = i
    if last_bang < 0:
        return None, s
    prefix = s[:last_bang]
    rest = s[last_bang + 1:]
    if prefix.startswith("["):
        return prefix, rest
    prefix = prefix.strip("'").replace("''", "'")
    return prefix, rest


def normalize_ref_part(part: str, row: int, col: int) -> str:
    m = CELL_REF_RE.match(part)
    if not m:
        return part
    col_abs = m.group(1) == "$"
    row_abs = m.group(3) == "$"
    ref_col = column_index_from_string(m.group(2).upper())
    ref_row = int(m.group(4))
    col_part = f"C{ref_col}$" if col_abs else f"C[{ref_col - col}]"
    row_part = f"R{ref_row}$" if row_abs else f"R[{ref_row - row}]"
    return row_part + col_part


def normalize_formula(formula: str, row: int, col: int) -> str:
    try:
        tokens = Tokenizer(formula).items
    except Exception:  # noqa: BLE001 - unparsable formula, compare raw
        return formula
    out = []
    for token in tokens:
        if token.type == "OPERAND" and token.subtype == "RANGE":
            prefix, rest = split_sheet_ref(token.value)
            if prefix is not None:
                head = f"{prefix}!"
            else:
                head = ""
            pieces = rest.split(":")
            out.append(head + ":".join(normalize_ref_part(p, row, col) for p in pieces))
        elif token.type == "FUNC":
            out.append(token.value.rstrip("(") + "(")
        else:
            out.append(token.value)
    return "".join(out)


def sheet_exists(model, name: str) -> bool:
    folded = name.casefold()
    return any(existing.casefold() == folded for existing in model.sheet_names)


# ---------------------------------------------------------------------------
# checks
# ---------------------------------------------------------------------------
def check_formula_errors(model, result, opts):
    if model.wbv is None:
        result.add(
            "info",
            "engine",
            "未重算，无法检查公式错误值与缓存结果",
            detail="LibreOffice Kit 未参与本次审计，公式单元格只有公式文本、没有缓存值。",
            suggestion="确认 LibreOffice Kit 可用后重跑；静态检查仍会报告断链、硬编码与格式问题。",
        )
        return
    errors = {}
    for sheet in model.iter_sheets():
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                if not sheet.is_formula(row, col):
                    continue
                value = sheet.value(row, col)
                if not is_error_value(value):
                    continue
                key = value.strip().upper()
                errors.setdefault(key, []).append(
                    (sheet.name, cell_ref(row, col), sheet.formula_text(row, col) or "")
                )
    for error_value, hits in sorted(errors.items()):
        shown = hits[: opts["max_examples"]]
        evidence = "；".join(f"{sheet}!{cell} {formula}" for sheet, cell, formula in shown)
        if len(hits) > len(shown):
            evidence += f"；…共 {len(hits)} 处"
        result.add(
            "critical",
            "formula-errors",
            f"公式错误值 {error_value} × {len(hits)}",
            sheet=hits[0][0],
            cell=hits[0][1],
            detail=f"重算后仍有 {len(hits)} 个单元格返回 {error_value}，下游引用会全部污染。",
            evidence=evidence,
            suggestion="逐处修复：检查被引用单元格是否为空/为 0、名称是否拼错、跨表引用是否断链，修完重算复验。",
        )


def check_broken_references(model, result, opts):
    missing = {}
    external = {}
    scanned = 0
    for sheet, row, col, cell in model.all_formula_cells():
        scanned += 1
        formula = cell.value if is_formula(cell.value) else None
        if formula is None:
            continue
        try:
            tokens = Tokenizer(formula).items
        except Exception:  # noqa: BLE001
            continue
        for token in tokens:
            if token.type != "OPERAND" or token.subtype != "RANGE":
                continue
            prefix, _rest = split_sheet_ref(token.value)
            if prefix is None:
                continue
            if prefix.startswith("["):
                external.setdefault(prefix, (sheet.name, cell_ref(row, col), formula))
                continue
            names = [part for part in prefix.split(":") if part]
            for name in names:
                if not sheet_exists(model, name):
                    missing.setdefault(name, (sheet.name, cell_ref(row, col), formula))
    for name, (sheet_name, ref, formula) in sorted(missing.items()):
        result.add(
            "critical",
            "broken-references",
            f"公式引用了不存在的工作表「{name}」",
            sheet=sheet_name,
            cell=ref,
            detail="公式中的跨表引用指向的工作表已被删除或改名，结果必然为 #REF!（或引用错误值）。",
            evidence=f"首个命中：{ref} {formula}",
            suggestion="将引用改为现存工作表，或恢复被删/被改名的工作表；随后用公式错误值检查复验。",
        )
    for prefix, (sheet_name, ref, formula) in sorted(external.items()):
        result.add(
            "warning",
            "broken-references",
            f"存在外部工作簿链接 {prefix}",
            sheet=sheet_name,
            cell=ref,
            detail="公式引用了另一个工作簿，交付/移动文件后链接会断，值不再更新。",
            evidence=f"{ref} {formula}",
            suggestion="把外部数据复制为值或迁移到本工作簿的输入表，删除外部链接。",
        )
    if not missing and not external:
        result.add(
            "info",
            "broken-references",
            "未发现断链或外部链接",
            detail=f"已解析 {scanned} 个公式单元格的引用。",
        )


def check_formula_consistency(model, result, opts):
    reported = 0

    def scan_line(sheet, cells, axis):
        nonlocal reported
        if len(cells) < 4 or reported >= MAX_FINDINGS_PER_RULE:
            return
        normalized = []
        for row, col in cells:
            text = sheet.formula_text(row, col)
            if text is None:
                return
            normalized.append((row, col, text, normalize_formula(text, row, col)))
        patterns = {}
        for row, col, text, norm in normalized:
            patterns.setdefault(norm, []).append((row, col, text))
        if len(patterns) < 2:
            return
        # Only flag when a clear majority pattern exists (minority ≤ 34%).
        ranked = sorted(patterns.items(), key=lambda kv: len(kv[1]), reverse=True)
        majority_norm, majority_cells = ranked[0]
        first_cell = normalized[0]
        last_cell = normalized[-1]
        minority = [
            item for _, group in ranked[1:] for item in group
            if not (item[0] == first_cell[0] and item[1] == first_cell[1])
            and not (item[0] == last_cell[0] and item[1] == last_cell[1])
        ]
        if len(minority) == 0 or len(minority) / len(normalized) > 0.34:
            return
        axis_name = "行" if axis == "row" else "列"
        evidence = "；".join(
            f"{cell_ref(r, c)} {t}" for r, c, t in minority[:3]
        ) + f"；主流模式示例 {cell_ref(majority_cells[0][0], majority_cells[0][1])} {majority_cells[0][2]}"
        result.add(
            "info",
            "formula-consistency",
            f"同一{axis_name}内公式模式不一致（{len(minority)} 个异常）",
            sheet=sheet.name,
            cell=cell_ref(minority[0][0], minority[0][1]),
            detail="相邻单元格的公式模式不同。若为首期/末期或开关分支的有意特殊逻辑可忽略；否则可能是手工改错或漏拖。",
            evidence=evidence,
            suggestion="确认差异是否符合业务口径；若为无意改错，统一为多数模式后重新拖拽填充。",
        )
        reported += 1

    for sheet in model.iter_sheets():
        # rows
        for row in range(1, sheet.max_row + 1):
            segment = []
            for col in range(1, sheet.max_column + 1):
                if sheet.is_formula(row, col):
                    segment.append((row, col))
                else:
                    if segment:
                        scan_line(sheet, segment, "row")
                        segment = []
            if segment:
                scan_line(sheet, segment, "row")



def is_input_style(cell) -> bool:
    """Blue font or yellow fill = a modeler-flagged hard-coded input."""
    try:
        font = cell.font
        if font is not None and font.color is not None:
            rgb = font.color.rgb
            if isinstance(rgb, str):
                rgb = rgb.upper()
                # ARGB or RGB; blue channel FF with R=G=0.
                if (rgb.endswith("0000FF") or rgb.endswith("00007F")) and rgb[:2] in ("FF", "00"):
                    return True
        fill = cell.fill
        if fill is not None and fill.fgColor is not None:
            rgb = fill.fgColor.rgb
            if isinstance(rgb, str):
                rgb = rgb.upper()
                if rgb.startswith("FF"):
                    r, g, b = int(rgb[2:4], 16), int(rgb[4:6], 16), int(rgb[6:8], 16)
                    if r >= 240 and g >= 210 and b <= 160:
                        return True
    except Exception:  # noqa: BLE001 - style objects can be odd in some files
        return False
    return False


def check_hardcoded_values(model, result, opts):
    reported = 0
    param_sheets = [
        name for name in model.sheet_names
        if any(word in normalize_label(name) for word in ("参数", "假设", "输入", "参数表", "Assumptions", "Input"))
    ]
    for sheet in model.iter_sheets():
        if sheet.name in param_sheets:
            continue
        if sheet.header_row is None:
            continue
        # Record-style tables (a repeated label per row, e.g. one company per
        # row) are not model rows: a constant next to a formula there is normal.
        labels = [normalize_label(label) for _row, label in sheet.label_rows()]
        unique_ratio = len(set(labels)) / max(1, len(labels))
        if not sheet.period_columns and unique_ratio < 0.8:
            continue
        start = sheet.header_row + 1
        for row in range(start, sheet.max_row + 1):
            if reported >= MAX_FINDINGS_PER_RULE:
                return
            formula_cols = []
            hard_cols = []
            for col in range(1, sheet.max_column + 1):
                cell = sheet.formula_cell(row, col)
                if cell.data_type == "f" or is_formula(cell.value):
                    formula_cols.append(col)
                    continue
                value = cell.value
                if is_number(value) and not is_formula(value):
                    if is_input_style(cell):
                        continue
                    hard_cols.append(col)
            if len(formula_cols) >= 2 and hard_cols:
                label = sheet.label(row)
                evidence = "；".join(
                    f"{cell_ref(row, col)}={format_compact(sheet.value(row, col))}" for col in hard_cols[:5]
                )
                if len(hard_cols) > 5:
                    evidence += f"；…共 {len(hard_cols)} 个常数"
                result.add(
                    "warning",
                    "hardcoded-values",
                    f"公式与硬编码常数混排：{label or '（无标签行）'}",
                    sheet=sheet.name,
                    cell=cell_ref(row, hard_cols[0]),
                    detail=f"该行有 {len(formula_cols)} 个公式，同时有 {len(hard_cols)} 个未标输入色的常数；口径改动时公式与常数会脱节。",
                    evidence=evidence,
                    suggestion="把常数集中到参数/输入表并标蓝或黄底，公式改为引用；或确认这些是历史实际数后加输入色标记。",
                )
                reported += 1


def check_number_format(model, result, opts):
    reported = 0
    for sheet in model.iter_sheets():
        if sheet.header_row is None or is_derived_sheet_name(sheet.name):
            continue
        for row in range(sheet.header_row + 1, sheet.max_row + 1):
            if reported >= MAX_FINDINGS_PER_RULE:
                return
            label = sheet.label(row)
            if not label:
                continue
            kind = sheet.row_kind(row)
            if kind == "other":
                continue
            percent_cells = []
            currency_cells = []
            for col in range(2, sheet.max_column + 1):
                value = sheet.value(row, col)
                num = to_number(value)
                if num is None:
                    continue
                fmt = sheet.number_format(row, col) or ""
                is_pct = "%" in fmt
                if is_pct:
                    percent_cells.append((col, num, fmt))
                elif CURRENCY_HINT_RE.search(fmt):
                    currency_cells.append((col, num, fmt))
            if kind in ("money", "days", "multiple") and percent_cells:
                kind_label = {"money": "金额", "days": "天数", "multiple": "倍数"}[kind]
                strong = kind != "money" or any(word in normalize_label(label) for word in STRONG_MONEY_WORDS)
                shown = "；".join(
                    f"{cell_ref(row, col)}={format_compact(num)} 格式={fmt}" for col, num, fmt in percent_cells[:4]
                )
                result.add(
                    "warning" if strong else "info",
                    "number-format",
                    f"{kind_label}行被套用百分比格式：{label}",
                    sheet=sheet.name,
                    cell=cell_ref(row, percent_cells[0][0]),
                    detail="数字格式与行语义不符，表格显示会失真（金额显示成 %、天数显示成 %），下游取数也容易误读。",
                    evidence=shown,
                    suggestion="按行设置正确格式：金额用 #,##0.00 / 会计格式，天数用 0.0 或 0，倍数用 0.00\"倍\"，比率才用 0.00%。",
                )
                reported += 1
            elif kind == "ratio" and not percent_cells and currency_cells:
                shown = "；".join(
                    f"{cell_ref(row, col)}={format_compact(num)} 格式={fmt}" for col, num, fmt in currency_cells[:4]
                )
                result.add(
                    "warning",
                    "number-format",
                    f"比率行使用金额格式：{label}",
                    sheet=sheet.name,
                    cell=cell_ref(row, currency_cells[0][0]),
                    detail="比率行套用了货币/千分位格式，数值会被读成金额。",
                    evidence=shown,
                    suggestion="改为 0.00% 格式；若该行实际是金额，请把行名改回金额口径。",
                )
                reported += 1
            elif kind == "ratio":
                # ratio row with no percent format and abnormally large values (>100)
                big = [(col, num) for col in range(2, sheet.max_column + 1)
                       if (num := to_number(sheet.value(row, col))) is not None and abs(num) > 100]
                if big and not percent_cells:
                    result.add(
                        "info",
                        "number-format",
                        f"比率行未使用百分比格式：{label}",
                        sheet=sheet.name,
                        cell=cell_ref(row, big[0][0]),
                        detail=f"该行最大值 {format_compact(big[0][1])}，但格式不是百分比；若是比率请统一为 0.00%。",
                        evidence="；".join(f"{cell_ref(row, c)}={format_compact(v)}" for c, v in big[:4]),
                        suggestion="统一比率行的数字格式，避免同一列混用小数与百分数。",
                    )
                    reported += 1


def check_data_cleaning(model, result, opts):
    reported = 0
    for sheet in model.iter_sheets():
        if sheet.header_row is None or not sheet.has_data():
            continue
        start = sheet.header_row + 1
        numeric_cols = sheet.numeric_columns()
        if not numeric_cols:
            continue
        for col in sorted(numeric_cols):
            if reported >= MAX_FINDINGS_PER_RULE:
                return
            blanks, text_cells, numeric_text, placeholders = [], [], [], []
            for row in range(start, sheet.max_row + 1):
                if not sheet.label(row):
                    continue
                cell = sheet.formula_cell(row, col)
                if is_formula(cell.value):
                    continue
                value = cell.value
                if value is None:
                    blanks.append(row)
                elif isinstance(value, str):
                    text = value.strip()
                    if text in PLACEHOLDER_TEXTS:
                        placeholders.append(row)
                    else:
                        num = to_number(text)
                        if num is not None and not text.endswith(("%", "％")):
                            numeric_text.append((row, text))
                        else:
                            text_cells.append((row, text))
            if blank_count := len(blanks):
                if sheet.period_columns and blank_count >= 3 and blank_count / max(1, (sheet.max_row - start + 1)) > 0.1:
                    result.add(
                        "info",
                        "data-cleaning",
                        f"数据列存在空白单元格 × {blank_count}",
                        sheet=sheet.name,
                        cell=cell_ref(blanks[0], col),
                        detail="数值列中的空白会让加总/比率静默跳过这些期间，需确认是漏填还是真为零。",
                        evidence="空白行：" + "、".join(str(r) for r in blanks[:8]) + ("…" if blank_count > 8 else ""),
                        suggestion="补齐数据；若确为零，填 0 并备注口径，不要留空。",
                    )
                    reported += 1
            if numeric_text:
                result.add(
                    "warning",
                    "data-cleaning",
                    f"数字以文本形式存储 × {len(numeric_text)}",
                    sheet=sheet.name,
                    cell=cell_ref(numeric_text[0][0], col),
                    detail="文本型数字不参与 SUM/AVERAGE，且排序、透视会错位。",
                    evidence="；".join(f"{cell_ref(r, col)}={v!r}" for r, v in numeric_text[:5]),
                    suggestion="用「分列」或 VALUE() 转为数值，并统一该列单元格格式。",
                )
                reported += 1
            if placeholders:
                result.add(
                    "info",
                    "data-cleaning",
                    f"数值列出现占位符（{len(placeholders)} 处）",
                    sheet=sheet.name,
                    cell=cell_ref(placeholders[0], col),
                    detail="源材料用「-」等占位符表示无数据；公式会把它当 0 或报 #VALUE!，口径需登记。",
                    evidence="、".join(cell_ref(r, col) for r in placeholders[:8]),
                    suggestion="在「说明」表登记占位符口径，或统一替换为 0 / 空值（二选一并保持全表一致）。",
                )
                reported += 1
            if text_cells and len(text_cells) <= 3:
                # Mixed text/number inside a numeric column.
                result.add(
                    "info",
                    "data-cleaning",
                    f"数值列混入文本 × {len(text_cells)}",
                    sheet=sheet.name,
                    cell=cell_ref(text_cells[0][0], col),
                    detail="同列既有数字又有文本，可能是单位/备注混入数据区。",
                    evidence="；".join(f"{cell_ref(r, col)}={v!r}" for r, v in text_cells[:3]),
                    suggestion="把备注/单位移出数据列，保持一列一类型。",
                )
                reported += 1

        # Duplicate rows inside the data block.
        seen = {}
        for row in range(start, sheet.max_row + 1):
            values = tuple(
                sheet.value(row, col) for col in sorted(numeric_cols)
                if sheet.value(row, col) is not None and not is_formula(sheet.formula_cell(row, col).value)
            )
            if len(values) < 2:
                continue
            key = tuple(round(v, 6) if isinstance(v, float) else v for v in values)
            seen.setdefault(key, []).append(row)
        for rows in seen.values():
            if len(rows) > 1 and reported < MAX_FINDINGS_PER_RULE:
                result.add(
                    "info",
                    "data-cleaning",
                    f"疑似重复行 × {len(rows)}",
                    sheet=sheet.name,
                    cell=cell_ref(rows[0], 1),
                    detail="数据区存在数值完全相同的行，可能是重复录入/重复导入。",
                    evidence="行号：" + "、".join(str(r) for r in rows[:8]),
                    suggestion="核对是否为重复记录；若是，删除重复行并重算。",
                )
                reported += 1


def check_outliers(model, result, opts):
    threshold = opts.get("growth_threshold", 20.0)
    reported = 0
    for sheet in model.iter_sheets():
        if is_derived_sheet_name(sheet.name):
            continue
        periods = sheet.period_columns
        if len(periods) < 2:
            continue
        by_kind = {}
        for col, period in periods.items():
            by_kind.setdefault(period[0], []).append((col, period))
        for kind, cols in by_kind.items():
            cols.sort(key=lambda item: item[1])
            if len(cols) < 2:
                continue
            for row in range(1, sheet.max_row + 1):
                if reported >= MAX_FINDINGS_PER_RULE:
                    return
                kind_row = sheet.row_kind(row)
                if kind_row in ("ratio", "days", "multiple"):
                    continue
                label = sheet.label(row)
                if not label:
                    continue
                for idx in range(1, len(cols)):
                    prev_col, prev_period = cols[idx - 1]
                    cur_col, cur_period = cols[idx]
                    prev = to_number(sheet.value(row, prev_col))
                    cur = to_number(sheet.value(row, cur_col))
                    if prev is None or cur is None or prev <= 0:
                        continue
                    growth = (cur - prev) / prev
                    if abs(growth) <= threshold:
                        continue
                    result.add(
                        "warning",
                        "outliers",
                        f"增速异常 {growth * 100:,.0f}%：{label}",
                        sheet=sheet.name,
                        cell=cell_ref(row, cur_col),
                        detail=f"{prev_period[1]} 年 → {cur_period[1]} 年，从 {format_compact(prev)} 变为 {format_compact(cur)}，超出 {threshold * 100:,.0f}% 阈值。",
                        evidence=f"{cell_ref(row, prev_col)}={format_compact(prev)} → {cell_ref(row, cur_col)}={format_compact(cur)}",
                        suggestion="确认是业务爆发（如新产线投产、并购）还是单位/口径错误；在说明表解释后可将该期标记为已复核。",
                    )
                    reported += 1


def check_workbook_structure(model, result, opts):
    hidden = [sheet.name for sheet in model.iter_sheets() if sheet.ws.sheet_state != "visible"]
    if hidden:
        result.add(
            "info",
            "workbook-structure",
            f"存在隐藏工作表 × {len(hidden)}",
            detail="隐藏表可能是历史版本或中间计算，交付前确认是否需要保留。",
            evidence="、".join(hidden[:10]),
            suggestion="删除或显式标注为「底稿-勿删」；隐藏不是保护。",
        )
    if model._defined_names:
        result.add(
            "info",
            "workbook-structure",
            f"定义名称 × {len(model._defined_names)}",
            detail="定义名称被公式引用，改名/删除会引发 #NAME?。",
            evidence="、".join(list(model._defined_names)[:10]),
            suggestion="保留命名区域的说明；删除前先全文搜索引用。",
        )
    # Formula ratio per sheet.
    for sheet in model.iter_sheets():
        if not sheet.has_data():
            continue
        formulas = 0
        constants = 0
        for row in range(1, sheet.max_row + 1):
            for col in range(1, sheet.max_column + 1):
                cell = sheet.formula_cell(row, col)
                if cell.value is None:
                    continue
                if cell.data_type == "f" or is_formula(cell.value):
                    formulas += 1
                elif is_number(cell.value):
                    constants += 1
        if formulas == 0 and constants > 20:
            result.add(
                "info",
                "workbook-structure",
                f"工作表「{sheet.name}」没有公式（纯数值）",
                detail=f"该表有 {constants} 个常数、0 个公式，可能是导出的结果表而非模型表。",
                suggestion="若为模型表，确认公式是否被「复制为值」覆盖。",
            )
