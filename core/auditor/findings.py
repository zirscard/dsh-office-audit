"""Finding model and report collector."""
from __future__ import annotations

from dataclasses import asdict, dataclass, field

SEVERITIES = ("critical", "warning", "info")
_SEVERITY_LABEL = {"critical": "严重", "warning": "警告", "info": "提示"}

CATEGORY_LABELS = {
    "formula-errors": "公式错误值",
    "broken-references": "断链引用",
    "formula-consistency": "公式一致性",
    "hardcoded-values": "硬编码",
    "number-format": "格式与单位",
    "balance-sheet": "资产负债表平衡",
    "retained-earnings": "留存收益滚动",
    "cash-tieout": "现金勾稽",
    "cash-flow-subtotals": "现金流量表加总",
    "ratio-conventions": "比率表口径",
    "data-cleaning": "数据清洗",
    "outliers": "异常值",
    "workbook-structure": "工作簿结构",
    "engine": "审计引擎",
}


@dataclass
class Finding:
    id: str
    severity: str
    category: str
    title: str
    sheet: str | None = None
    cell: str | None = None
    detail: str = ""
    evidence: str = ""
    suggestion: str = ""

    def as_dict(self):
        return asdict(self)


@dataclass
class AuditResult:
    workbook: str = ""
    profile: str = "auto"
    findings: list = field(default_factory=list)
    stats: dict = field(default_factory=dict)
    sheets: list = field(default_factory=list)
    meta: dict = field(default_factory=dict)

    def add(self, severity, category, title, *, sheet=None, cell=None, detail="", evidence="", suggestion=""):
        if severity not in SEVERITIES:
            raise ValueError(f"bad severity {severity!r}")
        finding = Finding(
            id=f"F-{len(self.findings) + 1:03d}",
            severity=severity,
            category=category,
            title=title,
            sheet=sheet,
            cell=cell,
            detail=detail,
            evidence=evidence,
            suggestion=suggestion,
        )
        self.findings.append(finding)
        return finding

    def summary(self):
        counts = {sev: 0 for sev in SEVERITIES}
        for f in self.findings:
            counts[f.severity] += 1
        counts["total"] = len(self.findings)
        return counts

    def sorted_findings(self):
        order = {sev: i for i, sev in enumerate(SEVERITIES)}
        return sorted(self.findings, key=lambda f: (order.get(f.severity, 9), f.category, f.sheet or "", f.cell or ""))

    def to_dict(self):
        return {
            "workbook": self.workbook,
            "profile": self.profile,
            "summary": self.summary(),
            "findings": [f.as_dict() for f in self.sorted_findings()],
            "sheets": self.sheets,
            "stats": self.stats,
            "meta": self.meta,
        }
