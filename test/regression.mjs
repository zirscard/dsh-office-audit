/**
 * Regression harness for dsh-office-audit.
 *
 * Runs the audit engine against the bundled broken fixture plus any real
 * workbooks that are present on this machine, writes a JSON evidence file to
 * test-output/, and exits non-zero when a fixture expectation regresses.
 *
 * Usage: node test/regression.mjs [--no-recalc]
 */
import { mkdirSync, readdirSync, writeFileSync } from 'node:fs'
import { dirname, join, resolve } from 'node:path'
import { fileURLToPath } from 'node:url'
import { runAudit } from '../runner.js'

const HERE = dirname(fileURLToPath(import.meta.url))
const ROOT = resolve(HERE, '..')
const RECALC = !process.argv.includes('--no-recalc')
const STAMP = new Date().toISOString().replace(/[:.]/g, '-')

/** Pick the first workbook in a project folder that matches a pattern. */
function discover(dir, pattern) {
  try {
    const hit = readdirSync(dir).find((name) => pattern.test(name) && !name.startsWith('~$'))
    return hit === undefined ? undefined : join(dir, hit)
  } catch {
    return undefined
  }
}

const CASES = [
  {
    name: 'fixture-broken-model',
    path: join(ROOT, 'test', 'fixtures', 'broken-model.xlsx'),
    expect: { critical: 6, minWarning: 12 },
  },
  {
    name: 'sample-ratio-table',
    path: discover(process.env.AUDIT_SAMPLE_RATIO_DIR || '', /_01_历年财务数据\.xlsx$/),
    optional: true,
    expect: { critical: 0, maxWarning: 0 },
  },
  {
    name: 'sample-dcf',
    path: process.env.AUDIT_SAMPLE_DCF || '',
    optional: true,
    expect: { critical: 0 },
  },
  {
    name: 'sample-financial-data',
    path: discover(process.env.AUDIT_SAMPLE_FIN_DIR || '', /_01_历年财务数据\.xlsx$/),
    optional: true,
    expect: { critical: 0 },
  },
]

const cases = CASES.filter((item) => item.path !== undefined || !item.optional)

const evidence = {
  generatedAt: new Date().toISOString(),
  recalculate: RECALC,
  engine: 'dsh-office-audit',
  results: [],
}
let failures = 0

for (const testCase of cases) {
  const outputDir = join(ROOT, 'test-output', 'regression', testCase.name)
  let result
  try {
    result = await runAudit({
      workbook: testCase.path,
      recalculate: RECALC,
      profile: 'auto',
      engineRoot: ROOT,
      outputDir,
    })
  } catch (error) {
    result = { error: String(error) }
  }

  if (result.error !== undefined) {
    if (testCase.optional) {
      evidence.results.push({ name: testCase.name, skipped: true, reason: result.error })
      console.log(`SKIP ${testCase.name}: ${result.error}`)
      continue
    }
    failures += 1
    evidence.results.push({ name: testCase.name, error: result.error })
    console.log(`FAIL ${testCase.name}: ${result.error}`)
    continue
  }

  const summary = result.summary
  const problems = []
  if (testCase.expect.critical !== undefined && summary.critical !== testCase.expect.critical) {
    problems.push(`critical=${summary.critical}, expected ${testCase.expect.critical}`)
  }
  if (testCase.expect.minWarning !== undefined && summary.warning < testCase.expect.minWarning) {
    problems.push(`warning=${summary.warning}, expected ≥ ${testCase.expect.minWarning}`)
  }
  if (testCase.expect.maxWarning !== undefined && summary.warning > testCase.expect.maxWarning) {
    problems.push(`warning=${summary.warning}, expected ≤ ${testCase.expect.maxWarning}`)
  }
  const passed = problems.length === 0
  if (!passed) failures += 1

  evidence.results.push({
    name: testCase.name,
    path: testCase.path,
    recalculated: result.recalculated,
    summary,
    problems,
    passed,
  })
  const mark = passed ? 'PASS' : 'FAIL'
  console.log(
    `${mark} ${testCase.name}: critical=${summary.critical} warning=${summary.warning} info=${summary.info}` +
      (problems.length > 0 ? ` (${problems.join('; ')})` : ''),
  )
}

const outDir = join(ROOT, 'test-output')
mkdirSync(outDir, { recursive: true })
const outFile = join(outDir, `regression-${STAMP}.json`)
writeFileSync(outFile, JSON.stringify(evidence, null, 2))
console.log(`\nEvidence: ${outFile}`)
process.exit(failures > 0 ? 1 : 0)
