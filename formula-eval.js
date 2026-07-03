// Lightweight evaluator for the handful of aggregate formulas that spreadsheet
// apps leave uncached (no stored result) so we can still show a value instead of
// a blank cell. Supports SUM/MIN/MAX/AVERAGE/COUNT over ranges, AVERAGEIF, and
// the MIN/MAX/AVERAGE(IF(cond,range)) array pattern. Dates are handled because
// callers feed date cells in as Excel serial numbers (see toExcelSerial).

// ExcelJS returns dates as UTC midnight; converting to an Excel serial keeps the
// calendar day stable regardless of the host timezone (otherwise users west of
// UTC would see dates shifted one day earlier).
const EXCEL_EPOCH_UTC_MS = Date.UTC(1899, 11, 30)
function toExcelSerial(d) {
  return (d.getTime() - EXCEL_EPOCH_UTC_MS) / 86400000
}

// Column letters (A, B, ..., AA) to a 0-based column index.
function colFromLetter(s) {
  let n = 0
  for (const ch of s.toUpperCase()) n = n * 26 + (ch.charCodeAt(0) - 64)
  return n - 1
}

// Parse "A1" or "A1:B2" (with optional $ anchors) into a 0-based range. Returns
// null for anything else (named ranges, cross-sheet refs, etc.).
function parseRef(s) {
  s = s.trim()
  let m = s.match(/^\$?([A-Z]+)\$?(\d+):\$?([A-Z]+)\$?(\d+)$/i)
  if (m) {
    return { r1: +m[2] - 1, c1: colFromLetter(m[1]), r2: +m[4] - 1, c2: colFromLetter(m[3]) }
  }
  m = s.match(/^\$?([A-Z]+)\$?(\d+)$/i)
  if (m) {
    const r = +m[2] - 1, c = colFromLetter(m[1])
    return { r1: r, c1: c, r2: r, c2: c }
  }
  return null
}

// Expand a range into its list of {r,c} cells, top-to-bottom, left-to-right.
function rangeCoords(rng) {
  const r1 = Math.min(rng.r1, rng.r2), r2 = Math.max(rng.r1, rng.r2)
  const c1 = Math.min(rng.c1, rng.c2), c2 = Math.max(rng.c1, rng.c2)
  const out = []
  for (let r = r1; r <= r2; r++) {
    for (let c = c1; c <= c2; c++) out.push({ r, c })
  }
  return out
}

// Split a comma-separated argument list, respecting nested parens and quotes so
// that e.g. AVERAGEIF(A:A,"<"&DATE(2100,1,1),B:B) splits into three args.
function splitArgs(s) {
  const args = []
  let depth = 0, cur = '', inQuote = false
  for (const ch of s) {
    if (ch === '"') { inQuote = !inQuote; cur += ch; continue }
    if (inQuote) { cur += ch; continue }
    if (ch === '(') depth++
    else if (ch === ')') depth--
    else if (ch === ',' && depth === 0) { args.push(cur.trim()); cur = ''; continue }
    cur += ch
  }
  args.push(cur.trim())
  return args
}

// Evaluate a scalar sub-expression to a number: a numeric literal or DATE(y,m,d).
// Returns null if it isn't one of those.
function evalNum(expr) {
  expr = expr.trim()
  const dm = expr.match(/^DATE\(\s*(\d+)\s*,\s*(\d+)\s*,\s*(\d+)\s*\)$/i)
  if (dm) return toExcelSerial(new Date(Date.UTC(+dm[1], +dm[2] - 1, +dm[3])))
  if (/^-?\d+(\.\d+)?$/.test(expr)) return Number(expr)
  return null
}

// Parse an AVERAGEIF-style criteria argument into {op, num}. Handles the two
// common shapes: a concatenated operator ("<"&DATE(2100,1,1)) and an all-in-one
// string literal (">5", "<=100"). A bare value means equality.
function parseCriteria(arg) {
  arg = arg.trim()
  let op = '='
  let m = arg.match(/^"(<=|>=|<>|<|>|=)"\s*&\s*(.+)$/)
  if (m) { op = m[1]; arg = m[2].trim() }
  else {
    m = arg.match(/^"(<=|>=|<>|<|>|=)?(.*)"$/)
    if (m) { op = m[1] || '='; arg = m[2].trim() }
  }
  const num = evalNum(arg)
  if (num === null) return null
  return { op, num }
}

// Parse an IF condition like "$C$2:$C$89<DATE(2100,1,1)" into {rng, op, num}.
function parseCondition(expr) {
  const m = expr.trim().match(/^(.+?)(<=|>=|<>|<|>|=)(.+)$/)
  if (!m) return null
  const rng = parseRef(m[1].trim())
  const num = evalNum(m[3].trim())
  if (!rng || num === null) return null
  return { rng, op: m[2], num }
}

function testCriteria(n, { op, num }) {
  switch (op) {
    case '<': return n < num
    case '>': return n > num
    case '<=': return n <= num
    case '>=': return n >= num
    case '<>': return n !== num
    default: return n === num
  }
}

function applyAgg(fn, nums) {
  if (fn === 'COUNT') return nums.length
  if (nums.length === 0) return null
  switch (fn) {
    case 'SUM': return nums.reduce((a, b) => a + b, 0)
    case 'MIN': return Math.min(...nums)
    case 'MAX': return Math.max(...nums)
    case 'AVERAGE': return nums.reduce((a, b) => a + b, 0) / nums.length
    default: return null
  }
}

// Pair two ranges cell-by-cell in expansion order (like AVERAGEIF pairing its
// criteria range with its average range). Truncates to the shorter length.
function pairedCoords(aRng, bRng) {
  const a = rangeCoords(aRng), b = rangeCoords(bRng)
  const n = Math.min(a.length, b.length)
  const pairs = []
  for (let i = 0; i < n; i++) pairs.push([a[i], b[i]])
  return pairs
}

// ctx: { rows, cols, get(r,c) -> number|undefined, pending(r,c) -> bool }
// Returns null (unsupported), { pending: true } (inputs not ready yet), or
// { value } (value may be null for an empty result → caller renders blank).
function evalFormula(formula, ctx) {
  if (!formula) return null
  const f = formula.trim()

  let m = f.match(/^(SUM|MIN|MAX|AVERAGE|COUNT)\((.*)\)$/i)
  if (m) {
    const fn = m[1].toUpperCase()
    const inner = m[2].trim()

    // Array pattern: MAX(IF(cond, range)) and friends.
    const ifm = inner.match(/^IF\((.*)\)$/i)
    if (ifm) {
      const ifargs = splitArgs(ifm[1])
      if (ifargs.length < 2) return null
      const cond = parseCondition(ifargs[0])
      const valRng = parseRef(ifargs[1])
      if (!cond || !valRng) return null
      const pairs = pairedCoords(cond.rng, valRng)
      const nums = []
      for (const [cc, vc] of pairs) {
        if (ctx.pending(cc.r, cc.c) || ctx.pending(vc.r, vc.c)) return { pending: true }
        const cv = ctx.get(cc.r, cc.c)
        if (typeof cv !== 'number' || !testCriteria(cv, cond)) continue
        const vv = ctx.get(vc.r, vc.c)
        if (typeof vv === 'number') nums.push(vv)
      }
      return { value: applyAgg(fn, nums) }
    }

    // Plain aggregate over one or more ranges/cells.
    const coords = []
    for (const a of splitArgs(inner)) {
      const rng = parseRef(a)
      if (!rng) return null
      coords.push(...rangeCoords(rng))
    }
    const nums = []
    for (const { r, c } of coords) {
      if (r < 0 || r >= ctx.rows || c < 0 || c >= ctx.cols) continue
      if (ctx.pending(r, c)) return { pending: true }
      const n = ctx.get(r, c)
      if (typeof n === 'number') nums.push(n)
    }
    return { value: applyAgg(fn, nums) }
  }

  m = f.match(/^AVERAGEIF\((.*)\)$/i)
  if (m) {
    const args = splitArgs(m[1])
    if (args.length < 2) return null
    const critRng = parseRef(args[0])
    const crit = parseCriteria(args[1])
    // A third arg is the average range; if omitted, average the criteria range.
    const avgRng = args[2] ? parseRef(args[2]) : critRng
    if (!critRng || !crit || !avgRng) return null
    const pairs = pairedCoords(critRng, avgRng)
    let sum = 0, cnt = 0
    for (const [cc, ac] of pairs) {
      if (ctx.pending(cc.r, cc.c) || ctx.pending(ac.r, ac.c)) return { pending: true }
      const cv = ctx.get(cc.r, cc.c)
      if (typeof cv !== 'number' || !testCriteria(cv, crit)) continue
      const av = ctx.get(ac.r, ac.c)
      if (typeof av === 'number') { sum += av; cnt++ }
    }
    return { value: cnt === 0 ? null : sum / cnt }
  }

  return null
}

module.exports = { evalFormula, toExcelSerial, EXCEL_EPOCH_UTC_MS }
