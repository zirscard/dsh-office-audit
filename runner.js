/**
 * Runtime bridge for dsh-office-audit.
 *
 * Locates the DSH-bundled Python runtime and LibreOffice Kit CLI, recalculates
 * the workbook into a temp copy, then runs `core/audit.py` and returns its
 * normalized JSON. Every external dependency is optional in the sense that
 * failure modes are reported, not crashed: a missing kit downgrades to static
 * checks, a missing Python is a hard error with a remediation hint.
 *
 * @module dsh-office-audit/runner
 */
import { spawn } from 'node:child_process'
import { existsSync, mkdtempSync, rmSync, statSync } from 'node:fs'
import { homedir, tmpdir } from 'node:os'
import { basename, dirname, extname, join, resolve } from 'node:path'

/** Resolve $DSH_HOME like the host does. */
function dshHome() {
  const override = process.env.DSH_HOME
  return (override !== undefined && override.length > 0 ? override : join(homedir(), '.dsh'))
}

function firstExisting(candidates) {
  for (const candidate of candidates) {
    if (candidate === undefined || candidate.length === 0) continue
    try {
      if (existsSync(candidate) && statSync(candidate).isFile()) return candidate
    } catch {
      /* ignore unreadable candidates */
    }
  }
  return undefined
}

/** Locate the bundled Python (or a PATH python3 as last resort). */
export function findPython() {
  const runtime = join(dshHome(), 'dsh-runtimes', 'dsh-primary-runtime', 'dependencies', 'python')
  const fromEnv = process.env.DSH_AUDIT_PYTHON
  const found = firstExisting([
    fromEnv,
    join(runtime, 'bin', 'python3'),
    join(runtime, 'bin', 'python'),
    join(runtime, 'python.exe'),
    join(runtime, 'bin', 'python.exe'),
  ])
  if (found !== undefined) return found
  return 'python3'
}

/** Locate the bundled Node used to run the LibreOffice Kit CLI. */
export function findNode() {
  const runtime = join(dshHome(), 'dsh-runtimes', 'dsh-primary-runtime', 'dependencies', 'node')
  const found = firstExisting([
    process.env.DSH_AUDIT_NODE,
    join(runtime, 'bin', 'node'),
    join(runtime, 'node.exe'),
    join(runtime, 'bin', 'node.exe'),
  ])
  if (found !== undefined) return found
  return process.execPath
}

/** Locate the LibreOffice Kit CLI (bundled with the DSH app). */
export function findOfficeCli() {
  const rel = ['node_modules', '@deepseek-ai', 'libreoffice-kit', 'lib', 'cli.js']
  const roots = []
  if (typeof process.resourcesPath === 'string' && process.resourcesPath.length > 0) {
    roots.push(join(process.resourcesPath, 'app.asar.unpacked', 'dsh'))
    roots.push(join(process.resourcesPath, 'app.asar.unpacked'))
  }
  roots.push('/Applications/DeepSeek Harness.app/Contents/Resources/app.asar.unpacked/dsh')
  roots.push(join(dirname(process.execPath), '..', 'Resources', 'app.asar.unpacked', 'dsh'))
  const candidates = [process.env.DSH_AUDIT_OFFICE_CLI]
  for (const root of roots) candidates.push(join(root, ...rel))
  return firstExisting(candidates)
}

/** Run a child process, capture stdout/stderr, enforce a timeout. */
function exec(command, args, options = {}) {
  return new Promise((resolvePromise) => {
    const child = spawn(command, args, {
      cwd: options.cwd,
      env: { ...process.env, ...(options.env ?? {}) },
      stdio: ['ignore', 'pipe', 'pipe'],
    })
    let stdout = ''
    let stderr = ''
    let settled = false
    const timeoutMs = options.timeoutMs ?? 300_000
    const timer = setTimeout(() => {
      if (settled) return
      settled = true
      child.kill('SIGKILL')
      resolvePromise({ code: -1, stdout, stderr: `${stderr}\n[timeout after ${timeoutMs}ms]`, timedOut: true })
    }, timeoutMs)
    child.stdout.on('data', (chunk) => { stdout += String(chunk) })
    child.stderr.on('data', (chunk) => { stderr += String(chunk) })
    child.on('error', (error) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolvePromise({ code: -1, stdout, stderr: `${stderr}\n${String(error?.message ?? error)}`, spawnError: error })
    })
    child.on('close', (code) => {
      if (settled) return
      settled = true
      clearTimeout(timer)
      resolvePromise({ code: code ?? -1, stdout, stderr })
    })
  })
}

/** Recalculate formulas via the bundled LibreOffice Kit CLI. */
async function recalculate(workbook, output, engineRoot) {
  const cli = findOfficeCli()
  if (cli === undefined) return { ok: false, error: 'LibreOffice Kit CLI not found (DSH app bundle is missing app.asar.unpacked/dsh/node_modules/@deepseek-ai/libreoffice-kit)' }
  const node = findNode()
  const result = await exec(node, [cli, 'recalculate', '--input', workbook, '--output', output], {
    cwd: engineRoot,
    timeoutMs: 240_000,
  })
  if (result.code !== 0) {
    const detail = (result.stderr || result.stdout || '').trim().split('\n').slice(-6).join('\n')
    return { ok: false, error: `LibreOffice recalculation failed (exit ${result.code})${detail.length > 0 ? `: ${detail}` : ''}` }
  }
  if (!existsSync(output)) return { ok: false, error: 'LibreOffice recalculation produced no output file' }
  return { ok: true }
}

function timestamp() {
  const d = new Date()
  const pad = (n) => String(n).padStart(2, '0')
  return `${d.getFullYear()}${pad(d.getMonth() + 1)}${pad(d.getDate())}-${pad(d.getHours())}${pad(d.getMinutes())}${pad(d.getSeconds())}`
}

/**
 * Run a full audit.
 *
 * @param {object} options
 * @param {string} options.workbook - absolute workbook path
 * @param {string} [options.profile] - auto | financial-model | ratio-table | generic
 * @param {boolean} [options.recalculate] - recalculate with LibreOffice first
 * @param {boolean} [options.writeAuditSheet] - write a 模型审计 sheet into a copy
 * @param {string[]} [options.checks] - category filter
 * @param {string} [options.outputDir] - report directory
 * @param {string} options.engineRoot - plugin root (contains core/audit.py)
 */
export async function runAudit(options) {
  const workbook = resolve(options.workbook)
  if (!existsSync(workbook)) return { error: `workbook not found: ${workbook}` }
  const ext = extname(workbook).toLowerCase()
  if (ext !== '.xlsx' && ext !== '.xlsm') return { error: `unsupported workbook type "${ext}"; only .xlsx/.xlsm are supported (convert .xls first)` }

  const python = findPython()
  const auditScript = join(options.engineRoot, 'core', 'audit.py')
  if (!existsSync(auditScript)) return { error: `audit engine missing: ${auditScript}` }

  const reportDir = options.outputDir ?? join(dirname(workbook), '.audit')
  const stem = basename(workbook, ext)
  const stamp = timestamp()
  const markdownPath = join(reportDir, `${stem}-audit-${stamp}.md`)
  const jsonPath = join(reportDir, `${stem}-audit-${stamp}.json`)

  const tempDir = mkdtempSync(join(tmpdir(), 'dsh-office-audit-'))
  try {
    let recalculated = false
    let recalcNote
    let recalcPath
    if (options.recalculate !== false) {
      recalcPath = join(tempDir, 'recalculated.xlsx')
      const recalc = await recalculate(workbook, recalcPath, options.engineRoot)
      if (recalc.ok) recalculated = true
      else {
        recalcNote = recalc.error
        recalcPath = undefined
      }
    }

    const args = [
      auditScript,
      '--input', workbook,
      '--out-json', jsonPath,
      '--out-md', markdownPath,
      '--profile', options.profile ?? 'auto',
    ]
    if (recalcPath !== undefined) args.push('--recalc', recalcPath)
    if (Array.isArray(options.checks) && options.checks.length > 0) args.push('--checks', options.checks.join(','))
    let auditWorkbook
    if (options.writeAuditSheet === true) {
      auditWorkbook = join(reportDir, `${stem}-audit-${stamp}.xlsx`)
      args.push('--write-sheet', auditWorkbook)
    }

    const result = await exec(python, args, { cwd: dirname(auditScript), timeoutMs: 600_000 })
    if (result.code !== 0) {
      const detail = (result.stderr || result.stdout || '').trim().split('\n').slice(-8).join('\n')
      return {
        error: `audit engine failed (exit ${result.code})${detail.length > 0 ? `: ${detail}` : ''}`,
        hint: 'check that the DSH Python runtime has openpyxl (load_workspace_dependencies reports it)',
      }
    }
    let payload
    try {
      payload = JSON.parse(result.stdout)
    } catch {
      return { error: 'audit engine returned invalid JSON', stdoutTail: result.stdout.slice(-2000), stderrTail: result.stderr.slice(-2000) }
    }
    if (recalcNote !== undefined) payload.recalcNote = recalcNote
    payload.workbook = workbook
    payload.recalculated = recalculated
    payload.report = { markdownPath, jsonPath }
    if (auditWorkbook !== undefined && existsSync(auditWorkbook)) payload.auditWorkbook = auditWorkbook
    // Ensure the on-disk JSON matches what the agent sees.
    return payload
  } finally {
    try {
      rmSync(tempDir, { recursive: true, force: true })
    } catch {
      /* temp cleanup is best effort */
    }
  }
}
