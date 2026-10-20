# dsh-office-audit · 财务模型 / Excel 审计器

把「审计器」做成 DSH 插件：agent 直接调用 **`office_audit`** 工具，对任意 `.xlsx/.xlsm` 做一次结构化审计，输出 `critical / warning / info` 分级的问题清单 + markdown/JSON 报告，可选在工作簿副本里生成「模型审计」表。

审计引擎移植自 Anthropic `financial-services` 仓库的 `audit-xls` / `clean-data-xls` 清单（Apache-2.0），并按中文财务模型的实际口径改写；不依赖任何外部技能包。

```
模型 → office_audit → 重算（LibreOffice Kit）→ 13 类检查 → 报告 + 审计表
```

## 一、装了之后发生什么

- 注册一个 agent 工具 `office_audit`（所有会话可用）。
- 注入一段系统提示：当用户要求审计 Excel / 交付含公式的模型时，agent 会先调用它。
- 纯 host 插件：不新增 GUI、不改 DSH 源码、没有 npm 运行时依赖。

## 二、调用方式

```
office_audit(
  path: "/abs/or/relative/model.xlsx",   # 必填；.xlsx/.xlsm
  profile: "auto",                       # auto | financial-model | ratio-table | generic
  recalculate: true,                     # 默认 true：先用 LibreOffice Kit 重算公式缓存
  writeAuditSheet: false,                # true：在工作簿副本里写「模型审计」表
  checks: [],                            # 可选：只跑指定类别
  outputDir: "..."                       # 可选：报告目录（默认 <工作簿目录>/.audit）
)
```

返回：`summary`（三级计数）、`findings`（含 sheet/cell/证据/建议）、`report.markdownPath`、`report.jsonPath`、可选 `auditWorkbook`。

**审计从不修改原文件。** 写审计表时生成 `<stem>-audit-<时间戳>.xlsx` 副本。

## 三、检查清单（13 类）

| 类别 | 级别 | 内容 |
| --- | --- | --- |
| `formula-errors` | critical | 重算后仍为 `#REF! / #DIV/0! / #VALUE! / #NAME? / #N/A / #NUM! / #SPILL! / #CALC!` 的单元格 |
| `broken-references` | critical / warning | 公式引用不存在的工作表（用 tokenizer 解析，不用正则）；外部工作簿链接 |
| `formula-consistency` | info | 同一行内公式模式不一致（归一化相对引用后比较；首尾特殊逻辑自动排除） |
| `hardcoded-values` | warning | 公式与未标输入色的常数混排（蓝字/黄底视为合法输入） |
| `number-format` | warning / info | 金额行/天数行套用 `%` 格式、比率行用金额格式（按行标签语义判断） |
| `balance-sheet` | critical | 资产总计 ≠ 负债合计 + 所有者权益合计（0.1% 容差，逐期） |
| `retained-earnings` | warning | 期末留存收益 ≠ 期初 + 净利润 − 分红（识别法定口径 vs 模拟口径差异） |
| `cash-tieout` | critical | BS 货币资金 vs CF 期末现金及现金等价物（跨表、按期间匹配） |
| `cash-flow-subtotals` | critical | 净增加额 ≠ 经营+投资+筹资；期末现金 ≠ 期初 + 净增加额 |
| `ratio-conventions` | warning / info | 内联硬编码 365/360/366；季度/年度混用；未投产期分母过小；平均余额 vs 期末口径混用；首期 `/0` |
| `data-cleaning` | warning / info | 空白（仅时间序列表）、数字存文本、占位符「-」、混合类型、重复行 |
| `outliers` | warning | 同比增速超过 2000% 阈值（可按期间类型对齐，季度只比季度） |
| `workbook-structure` | info | 隐藏工作表、定义名称、纯数值表 |

自动匹配三表行标签（资产总计/负债合计/所有者权益合计/留存收益/净利润/分红/货币资金/经营活动现金流净额……），中文表头可用；找不到对应行时静默跳过，不误报。

## 四、安装（desktop profile）

```bash
CLI="/Applications/DeepSeek Harness.app/Contents/Resources/runtime/cli/bin/dsh"
"$CLI" plugin --profile desktop add link:/path/to/plugins/dsh-office-audit
```

然后确认 `~/.dsh/profiles/desktop/package.json`：

```json
{
  "dependencies": { "dsh-office-audit": "link:/…/plugins/dsh-office-audit" },
  "dsh": { "profile": { "bundles": [ "…", "dsh-office-audit" ] } }
}
```

`bundles` 里没有这一行时插件不会加载。改完确认 GUI 工具列表里出现 `office_audit`（DSH 会自动重载 profile；若没出现，重启一次 DSH）。

归档包安装（离线/换机）：

```bash
"$CLI" plugin --profile desktop add file:/path/to/deploy/office-audit/vendor/dsh-office-audit-0.2.0.tgz
```

卸载：`"$CLI" plugin --profile desktop remove dsh-office-audit`（不删报告）。

## 五、运行环境（插件自动探测，无需配置）

| 依赖 | 来源 |
| --- | --- |
| Python 3 + openpyxl | `~/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python` |
| LibreOffice Kit CLI | DSH app 包内 `app.asar.unpacked/dsh/node_modules/@deepseek-ai/libreoffice-kit` |
| Node | DSH runtime 的 node（用于跑 kit CLI） |

可用环境变量覆盖：`DSH_AUDIT_PYTHON`、`DSH_AUDIT_NODE`、`DSH_AUDIT_OFFICE_CLI`。

**降级行为**：LibreOffice Kit 缺失时自动跳过重算，只做静态检查（断链、硬编码、格式、数据清洗），并在报告里标注「未重算」；Python 缺失时工具报错并提示用 `load_workspace_dependencies` 检查运行时。

## 六、目录

```
plugins/dsh-office-audit/
├── index.js              # host 插件：注册 office_audit + 系统提示
├── runner.js             # 探测 python / node / LibreOffice Kit，spawn 审计
├── core/
│   ├── audit.py          # 引擎 CLI（--input/--recalc/--out-json/--out-md/--write-sheet）
│   └── auditor/
│       ├── model.py      # 工作簿模型、表头/期间/标签识别
│       ├── checks_generic.py   # 通用检查
│       ├── checks_finance.py   # 三表 + 比率表专项
│       ├── findings.py   # 分级结果模型
│       ├── report.py     # markdown + 「模型审计」表
│       └── util.py       # 数值/标签/期间解析
└── test/
    └── make_fixture.py   # 生成覆盖 13 类问题的坏模型
```

## 七、自测

```bash
# 回归：fixture + 本机三张真实表（存在才跑）
node test/regression.mjs

# 只跑引擎（不重算）
PY="$HOME/.dsh/dsh-runtimes/dsh-primary-runtime/dependencies/python/bin/python3"
PY core/audit.py --input test/fixtures/broken-model.xlsx --out-json /tmp/a.json --out-md /tmp/a.md
```

`test/fixtures/broken-model.xlsx` 的期望结果是 **6 critical / 12 warning / 5 info**：critical 覆盖断链、公式错误、BS 不平衡、CF 加总、现金 tie-out；warning 覆盖格式误用、硬编码、增速、比率口径、数据清洗。回归结果写入 `test-output/regression-<时间戳>.json`。

## 八、已知限制

- 只支持 `.xlsx/.xlsm`；`.xls` 需先转换。
- 派生表（`common-size` / 结构 / 同比 / 敏感性）自动跳过金额格式、增速与三表检查——它们的行是占比，按金额口径检查会误报。
- 公式一致性降为 `info`：真实模型的首期/末期特殊逻辑很常见，报告里只作提示。
- 三表/比率检查依赖中文行标签（资产总计、负债合计、留存收益……）；非中文模板会退化为通用检查。
- 公式一致性、硬编码、格式判断是启发式，可能有个别误报；报告里给了证据和单元格位置，便于人工复核。
- 重算用 LibreOffice，个别 Excel 专有函数（如 `XLOOKUP` 旧版本、部分动态数组）可能不支持。
- 体量：默认最多扫描 12 万个公式、40 万个单元格；超大模型会截断并在报告标注。

## 九、来源与许可

- 检查清单参考 [anthropics/financial-services](https://github.com/anthropics/financial-services)（Apache-2.0）的 `audit-xls` / `clean-data-xls` 技能，按中文财务模型口径重写，未复制其代码。
- 本插件：MIT。
