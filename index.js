/**
 * Host half of dsh-office-audit.
 *
 * Registers one agent tool, `office_audit`, plus a short system-prompt section
 * that tells the agent when to use it. All heavy lifting lives in the bundled
 * Python audit engine (`core/audit.py`, openpyxl) and the DSH-bundled
 * LibreOffice Kit CLI (formula recalculation); `runner.js` locates those
 * runtimes, spawns them, and normalizes their JSON.
 *
 * The plugin keeps ZERO runtime npm dependencies, mirroring dsh-taskboard: it
 * compiles its parameter spec to the raw JSON-Schema subset the tool registry
 * expects and never imports @deepseek-ai packages at runtime.
 *
 * @module dsh-office-audit
 */
import { fileURLToPath } from 'node:url'
import { dirname, isAbsolute, resolve } from 'node:path'
import { runAudit } from './runner.js'

/** Cordis plugin name. */
export const name = 'dsh-office-audit'

/** Required host services: tool registry + prompt assembly. */
export const inject = ['tools', 'systemPrompt']

const HERE = dirname(fileURLToPath(import.meta.url))

/* ------------------------------------------------------------------ */
/* defineTool (structure-compatible reimplementation of the SDK's)     */
/* ------------------------------------------------------------------ */

/** Compile one value spec to the raw JSON-Schema subset. */
function compileValue(spec) {
  const node = {}
  if (typeof spec.description === 'string' && spec.description.length > 0) node.description = spec.description
  const type = spec.type
  if (type === undefined || type === 'json') return node
  if (type === 'object') {
    node.type = 'object'
    node.additionalProperties = spec.additionalProperties
    if (spec.properties !== undefined) node.properties = compilePropertyMap(spec.properties).properties
    return node
  }
  if (type === 'array') {
    node.type = 'array'
    if (spec.items !== undefined) node.items = compileValue(spec.items)
    return node
  }
  node.type = type
  if (spec.enum !== undefined) node.enum = [...spec.enum]
  if (spec.const !== undefined) node.const = spec.const
  return node
}

function compilePropertyMap(spec) {
  const properties = {}
  const required = []
  for (const [key, entry] of Object.entries(spec)) {
    const { required: isRequired, ...valueSpec } = entry
    properties[key] = compileValue(valueSpec)
    if (isRequired === true) required.push(key)
  }
  return required.length > 0 ? { properties, required } : { properties }
}

function matchesScalarType(value, type) {
  switch (type) {
    case 'string': return typeof value === 'string'
    case 'number': return typeof value === 'number'
    case 'integer': return typeof value === 'number' && Number.isInteger(value)
    case 'boolean': return typeof value === 'boolean'
    case 'null': return value === null
    default: return true
  }
}

function validateValue(schema, value, path) {
  if (typeof schema.type !== 'string' || schema.type.length === 0) return []
  if (schema.type === 'object') {
    if (typeof value !== 'object' || value === null || Array.isArray(value)) return [`${path} must be an object`]
    const violations = []
    const present = value
    for (const key of schema.required ?? []) if (!(key in present)) violations.push(`${path}.${key} is required`)
    if (schema.additionalProperties === false) {
      const known = new Set(Object.keys(schema.properties ?? {}))
      for (const key of Object.keys(present)) if (!known.has(key)) violations.push(`${path}.${key} is not a declared property`)
    }
    for (const [key, child] of Object.entries(schema.properties ?? {})) {
      if (key in present) violations.push(...validateValue(child, present[key], `${path}.${key}`))
    }
    return violations
  }
  if (schema.type === 'array') {
    if (!Array.isArray(value)) return [`${path} must be an array`]
    const violations = []
    if (schema.items !== undefined) value.forEach((item, i) => violations.push(...validateValue(schema.items, item, `${path}[${i}]`)))
    return violations
  }
  if (!matchesScalarType(value, schema.type)) return [`${path} must be ${schema.type}`]
  if (schema.enum !== undefined && !schema.enum.some((v) => v === value)) return [`${path} must be one of ${schema.enum.map(String).join(', ')}`]
  if (schema.const !== undefined && schema.const !== value) return [`${path} must be ${String(schema.const)}`]
  return []
}

function defineTool(options) {
  const compiled = compilePropertyMap(options.parameters)
  const parameters = { type: 'object', properties: compiled.properties }
  if (compiled.required !== undefined) parameters.required = compiled.required
  return {
    name: options.name,
    description: options.description,
    parameters,
    output: {
      schema: {},
      render(args, value) {
        return options.output.render(args, value)
      },
    },
    async execute(args, exec) {
      const violations = validateValue(parameters, args, 'arguments')
      if (violations.length > 0) throw new Error(`Error: invalid arguments: ${violations.join('; ')}`)
      return options.execute(args, exec)
    },
  }
}

/* ------------------------------------------------------------------ */
/* rendering                                                           */
/* ------------------------------------------------------------------ */

const SEVERITY_LABEL = { critical: '严重', warning: '警告', info: '提示' }
const SEVERITY_ORDER = { critical: 0, warning: 1, info: 2 }

function findingLine(f) {
  const where = [f.sheet, f.cell].filter(Boolean).join('!')
  const sev = SEVERITY_LABEL[f.severity] ?? f.severity
  const tag = where.length > 0 ? `[${sev}] ${where}：` : `[${sev}] `
  return `- ${tag}${f.title}${f.detail ? ` — ${f.detail}` : ''}`
}

function renderAudit(_args, value) {
  if (value === null || typeof value !== 'object') return [{ type: 'text', text: String(value) }]
  if (value.error !== undefined) return [{ type: 'text', text: `审计失败：${value.error}` }]
  const v = value
  const s = v.summary ?? {}
  const findings = Array.isArray(v.findings) ? v.findings : []
  const lines = []
  lines.push(`审计完成：${v.workbook ?? ''}`)
  const meta = []
  if (v.profile) meta.push(`模式 ${v.profile}`)
  if (v.recalculated) meta.push('已用 LibreOffice 重算')
  else meta.push('未重算（静态检查）')
  if (typeof v.formulaCells === 'number') meta.push(`公式 ${v.formulaCells} 个`)
  if (meta.length > 0) lines.push(meta.join(' · '))
  lines.push(`结果：严重 ${s.critical ?? 0} / 警告 ${s.warning ?? 0} / 提示 ${s.info ?? 0}`)
  const top = [...findings]
    .sort((a, b) => (SEVERITY_ORDER[a.severity] ?? 9) - (SEVERITY_ORDER[b.severity] ?? 9))
    .slice(0, 24)
  if (top.length > 0) {
    lines.push('', '主要问题：')
    for (const f of top) lines.push(findingLine(f))
    if (findings.length > top.length) lines.push(`… 另有 ${findings.length - top.length} 条，见完整报告`)
  } else {
    lines.push('未发现需要报告的问题。')
  }
  if (v.report?.markdownPath) lines.push('', `完整报告：${v.report.markdownPath}`)
  if (v.auditWorkbook) lines.push(`审计表副本：${v.auditWorkbook}`)
  return [{ type: 'text', text: lines.join('\n') }]
}

/* ------------------------------------------------------------------ */
/* plugin apply                                                        */
/* ------------------------------------------------------------------ */

const PROMPT_SECTION = `## 办公审计器（office_audit）

当用户要求检查/审计 Excel 工作簿（尤其是财务模型、三表联动、财务比率表），或你即将交付一个含公式的 .xlsx 时，先调用 office_audit 工具，再向用户汇报。

- 参数 path：工作簿路径（绝对路径，或相对当前会话工作目录）。支持 .xlsx / .xlsm，不支持旧版 .xls。
- 工具自动用 LibreOffice 重算公式缓存，然后检查：公式错误值（#REF! 等）、断链引用、公式一致性、硬编码、数字格式与单位（% / 金额 / 天）、三表勾稽（BS 平衡、留存收益滚动、现金 tie-out、现金流量表加总）、比率表口径（年天数、季度/年度混用、未投产期分母过小）、数据清洗（空白、混合类型、数字存文本、符号语义）、异常增速。
- 输出按 critical / warning / info 分级，附完整 markdown + JSON 报告路径；传入 writeAuditSheet=true 会在工作簿副本里生成「模型审计」表。
- 审计不修改原文件。修复前先把 critical 与 warning 完整汇报给用户（附工作表和单元格位置）；修完必须重跑一次审计复验。`

function resolveWorkbookPath(args, exec) {
  let p = args.path
  if (typeof p !== 'string' || p.trim().length === 0) throw new Error('Error: invalid_input: path is required')
  if (!isAbsolute(p)) {
    const cwd = exec?.agent?.session?.header?.cwd
    if (typeof cwd !== 'string' || cwd.length === 0) throw new Error('Error: invalid_input: relative path given but the session has no working directory')
    p = resolve(cwd, p)
  }
  return p
}

export function apply(ctx) {
  const disposers = []

  disposers.push(ctx.tools.register(defineTool({
    name: 'office_audit',
    description: 'Audit an Excel workbook (.xlsx/.xlsm) as a financial model: recalculate formulas with LibreOffice, then check formula errors, broken sheet references, formula-consistency, hardcoded constants, number-format/unit mismatches, balance-sheet balance, retained-earnings roll-forward, cash tie-out, cash-flow subtotals, ratio-table conventions (day counts, quarterly/annual mixing, pre-revenue denominators), data-cleaning issues, and outlier growth rates. Returns severity-ranked findings plus markdown/JSON report paths. Does not modify the original workbook.',
    parameters: {
      path: {
        type: 'string',
        required: true,
        description: 'Workbook path (.xlsx/.xlsm). Absolute, or relative to the session working directory.',
      },
      profile: {
        type: 'string',
        enum: ['auto', 'financial-model', 'ratio-table', 'generic'],
        description: 'Check profile. auto (default) detects three-statement / ratio tables and applies those checks when detected.',
      },
      recalculate: {
        type: 'boolean',
        description: 'Recalculate formulas with the bundled LibreOffice Kit before auditing (default true). Falls back to static checks when the kit is unavailable.',
      },
      writeAuditSheet: {
        type: 'boolean',
        description: 'Write a 「模型审计」 findings sheet into a copy of the workbook (default false; the original is never touched).',
      },
      checks: {
        type: 'array',
        items: { type: 'string' },
        description: 'Optional category ids to restrict the audit, e.g. ["formula-errors","balance-sheet"]. Empty = all.',
      },
      outputDir: {
        type: 'string',
        description: 'Directory for the report files (default: <workbook dir>/.audit).',
      },
    },
    output: { render: renderAudit },
    async execute(args, exec) {
      const workbook = resolveWorkbookPath(args, exec)
      const result = await runAudit({
        workbook,
        profile: args.profile ?? 'auto',
        recalculate: args.recalculate !== false,
        writeAuditSheet: args.writeAuditSheet === true,
        checks: Array.isArray(args.checks) ? args.checks : [],
        outputDir: typeof args.outputDir === 'string' && args.outputDir.length > 0
          ? (isAbsolute(args.outputDir) ? args.outputDir : resolve(dirname(workbook), args.outputDir))
          : undefined,
        engineRoot: HERE,
      })
      return result
    },
  })))

  const disposeSection = ctx.systemPrompt.section({
    name: 'office-audit',
    order: 185,
    text: PROMPT_SECTION,
  })
  ctx.effect(() => disposeSection, 'dsh-office-audit: prompt section')

  ctx.effect(() => () => {
    for (const dispose of disposers.splice(0)) dispose()
  }, 'dsh-office-audit: tools')
}
