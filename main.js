const { app, BrowserWindow, ipcMain, dialog, shell, Menu } = require('electron')
const path = require('path')
const fs = require('fs')
const ExcelJS = require('exceljs')
const JSZip = require('jszip')
const SSF = require('ssf')
const XLSX = require('xlsx')
const { evalFormula, toExcelSerial } = require('./formula-eval')

let pendingFilePaths = [] // files received before app is ready

const MAX_RECENT_FILES = 10

// Recent files are persisted as a JSON array of absolute paths in userData.
function recentFilesPath() {
  return path.join(app.getPath('userData'), 'recent-files.json')
}

// Read the recent list, dropping any entries whose file no longer exists.
function getRecentFiles() {
  try {
    const list = JSON.parse(fs.readFileSync(recentFilesPath(), 'utf-8'))
    if (!Array.isArray(list)) return []
    return list.filter((p) => typeof p === 'string' && fs.existsSync(p))
  } catch {
    return []
  }
}

// Move filePath to the front of the recent list (most recent first), de-duped.
function addRecentFile(filePath) {
  try {
    const abs = path.resolve(filePath)
    const list = getRecentFiles().filter((p) => p !== abs)
    list.unshift(abs)
    fs.writeFileSync(recentFilesPath(), JSON.stringify(list.slice(0, MAX_RECENT_FILES), null, 2))
    app.addRecentDocument(abs) // macOS Dock / Windows JumpList "Open Recent"
    buildMenu() // refresh the File > Open Recent submenu
  } catch { /* ignore — recents are best-effort */ }
}

function clearRecentFiles() {
  try {
    fs.writeFileSync(recentFilesPath(), JSON.stringify([]))
  } catch { /* ignore */ }
  app.clearRecentDocuments()
  buildMenu()
}

// Drop a single path from the recent list (e.g. it was deleted) and refresh menu.
function removeRecentFile(filePath) {
  try {
    const abs = path.resolve(filePath)
    const list = getRecentFiles().filter((p) => p !== abs)
    fs.writeFileSync(recentFilesPath(), JSON.stringify(list, null, 2))
  } catch { /* ignore */ }
  buildMenu()
}

// Tell the user a recent file is gone, then prune it from the list.
function reportMissingFile(filePath, win) {
  removeRecentFile(filePath)
  const opts = {
    type: 'warning',
    buttons: ['OK'],
    message: 'File not found',
    detail: `"${path.basename(filePath)}" can't be opened — it may have been moved, ` +
      `renamed, or deleted.\n\nIt has been removed from your recent files.`,
  }
  if (win && !win.isDestroyed()) dialog.showMessageBox(win, opts)
  else dialog.showMessageBox(opts)
}

// Tell the user we opened the file but couldn't read it as a spreadsheet.
function reportUnsupportedFile(filePath, detail, win) {
  removeRecentFile(filePath)
  const opts = {
    type: 'warning',
    buttons: ['OK'],
    message: "Can't open this file",
    detail,
  }
  if (win && !win.isDestroyed()) dialog.showMessageBox(win, opts)
  else dialog.showMessageBox(opts)
}

// Surface any parse failure to the user as a dialog. Every path through here
// ends with the renderer getting null, so a failed open never blanks a window.
function reportParseError(err, filePath, win) {
  if (err && err.code === 'ENOENT') {
    reportMissingFile(filePath, win)
  } else if (err && err.code === 'EUNSUPPORTED') {
    reportUnsupportedFile(filePath, err.message, win)
  } else {
    reportUnsupportedFile(
      filePath,
      `"${path.basename(filePath)}" couldn't be opened.\n\n` +
        `${(err && err.message) || 'Unknown error.'}`,
      win
    )
  }
}

// Open a file path, reusing the focused window if it's empty, else a new one.
function openPathInWindow(filePath) {
  if (!fs.existsSync(filePath)) {
    reportMissingFile(filePath, BrowserWindow.getFocusedWindow())
    return
  }
  const focused = BrowserWindow.getFocusedWindow()
  if (focused && !focused._hasFile) {
    focused._hasFile = true
    focused.webContents.send('open-file', filePath)
  } else {
    createWindow(filePath)._hasFile = true
  }
}

// Trigger the same Open dialog flow as the in-app button.
function triggerOpenDialog() {
  const focused = BrowserWindow.getFocusedWindow()
  if (focused) focused.webContents.send('menu-open')
  else createWindow()
}

// Build the application menu, including the dynamic Open Recent submenu.
function buildMenu() {
  const isMac = process.platform === 'darwin'
  const recent = getRecentFiles()

  const recentItems = recent.length
    ? recent.map((p) => ({ label: path.basename(p), click: () => openPathInWindow(p) }))
    : [{ label: 'No Recent Files', enabled: false }]
  recentItems.push(
    { type: 'separator' },
    { label: 'Clear Recent', enabled: recent.length > 0, click: clearRecentFiles },
  )

  const template = [
    ...(isMac ? [{ role: 'appMenu' }] : []),
    {
      label: 'File',
      submenu: [
        { label: 'Open…', accelerator: 'CmdOrCtrl+O', click: triggerOpenDialog },
        { label: 'Open Recent', submenu: recentItems },
        { type: 'separator' },
        isMac ? { role: 'close' } : { role: 'quit' },
      ],
    },
    { role: 'editMenu' },
    { role: 'viewMenu' },
    { role: 'windowMenu' },
  ]

  Menu.setApplicationMenu(Menu.buildFromTemplate(template))
}

function createWindow(filePath) {
  const win = new BrowserWindow({
    width: 1200,
    height: 800,
    minWidth: 800,
    minHeight: 600,
    titleBarStyle: 'hiddenInset',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
    },
  })

  win.loadFile('index.html')

  // Open links in external browser instead of navigating the app
  win.webContents.setWindowOpenHandler(({ url }) => {
    shell.openExternal(url)
    return { action: 'deny' }
  })
  win.webContents.on('will-navigate', (event, url) => {
    if (!url.startsWith('file://')) {
      event.preventDefault()
      shell.openExternal(url)
    }
  })

  if (filePath) {
    win.webContents.on('did-finish-load', () => {
      win.webContents.send('open-file', filePath)
    })
  }

  return win
}

// macOS: handle file opened via Finder (right-click > Open With, drag to dock, etc.)
app.on('open-file', (event, filePath) => {
  event.preventDefault()
  if (app.isReady()) {
    if (fs.existsSync(filePath)) createWindow(filePath)
    else reportMissingFile(filePath, BrowserWindow.getFocusedWindow())
  } else {
    pendingFilePaths.push(filePath)
  }
})

app.whenReady().then(() => {
  buildMenu()
  if (pendingFilePaths.length > 0) {
    pendingFilePaths.forEach((fp) => createWindow(fp))
    pendingFilePaths = []
  } else {
    createWindow()
  }
  app.on('activate', () => {
    if (BrowserWindow.getAllWindows().length === 0) createWindow()
  })
})

app.on('window-all-closed', () => {
  if (process.platform !== 'darwin') app.quit()
})

// Default Excel indexed color palette (indices 0-63, plus 64/65 for system fg/bg)
const DEFAULT_INDEXED_COLORS = [
  '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF', // 0-7
  '000000','FFFFFF','FF0000','00FF00','0000FF','FFFF00','FF00FF','00FFFF', // 8-15
  '800000','008000','000080','808000','800080','008080','C0C0C0','808080', // 16-23
  '9999FF','993366','FFFFCC','CCFFFF','660066','FF8080','0066CC','CCCCFF', // 24-31
  '000080','FF00FF','FFFF00','00FFFF','800080','800000','008080','0000FF', // 32-39
  '00CCFF','CCFFFF','CCFFCC','FFFF99','99CCFF','FF99CC','CC99FF','FFCC99', // 40-47
  '3366FF','33CCCC','99CC00','FFCC00','FF9900','FF6600','666699','969696', // 48-55
  '003366','339966','003300','333300','993300','993366','333399','333333', // 56-63
  '000000','FFFFFF', // 64-65: system foreground / background
]

// Parse xlsx zip for custom palette and HYPERLINK formula display text
async function parseXlsxMeta(buffer) {
  const result = { palette: null, hyperlinkMap: {} }
  try {
    const zip = await JSZip.loadAsync(buffer)

    // Extract custom indexed colors
    const stylesXml = await zip.file('xl/styles.xml')?.async('text')
    if (stylesXml) {
      const match = stylesXml.match(/<indexedColors>([\s\S]*?)<\/indexedColors>/)
      if (match) {
        const colors = []
        const regex = /rgb="([A-Fa-f0-9]+)"/g
        let m
        while ((m = regex.exec(match[1])) !== null) {
          const argb = m[1]
          colors.push(argb.length === 8 ? argb.slice(2) : argb)
        }
        if (colors.length > 0) result.palette = colors
      }
    }

    // Extract HYPERLINK formulas from each sheet XML
    // Build map: sheetIndex (1-based) -> { cellAddr -> { url, text } }
    const sheetFiles = zip.file(/^xl\/worksheets\/sheet\d+\.xml$/)
    for (const file of sheetFiles) {
      const sheetIdx = parseInt(file.name.match(/sheet(\d+)\.xml$/)[1])
      const xml = await file.async('text')
      const map = {}
      // Split on </c> so each chunk contains at most one cell
      const cells = xml.split('</c>')
      for (const chunk of cells) {
        const addrMatch = chunk.match(/<c\s+r="([A-Z]+\d+)"/)
        // Try both escaped (&quot;) and unescaped (") quote styles
        const hlMatch = chunk.match(/HYPERLINK\("([^"]*)"\s*,\s*"([^"]*)"\)/) ||
                        chunk.match(/HYPERLINK\(&quot;([^&]*)&quot;\s*,\s*&quot;([^&]*)&quot;\)/)
        if (addrMatch && hlMatch) {
          map[addrMatch[1]] = { url: hlMatch[1], text: hlMatch[2] }
        }
      }
      if (Object.keys(map).length > 0) result.hyperlinkMap[sheetIdx] = map
    }
  } catch { /* ignore */ }
  return result
}

// Resolve an ExcelJS color object to a 6-char hex string using the palette
function resolveColor(color, palette) {
  if (!color) return null
  if (color.argb) {
    const hex = color.argb.length === 8 ? color.argb.slice(2) : color.argb
    return hex
  }
  if (color.indexed !== undefined && color.indexed < palette.length) {
    return palette[color.indexed]
  }
  return null
}

// Build an inline CSS string from an ExcelJS cell's style properties
function getCellCSS(cell, palette) {
  if (!cell) return ''
  const parts = []

  const font = cell.font || {}
  if (font.bold)      parts.push('font-weight:bold')
  if (font.italic)    parts.push('font-style:italic')
  if (font.underline) parts.push('text-decoration:underline')
  if (font.strike)    parts.push('text-decoration:line-through')

  if (font.size) parts.push(`font-size:${font.size}pt`)

  const fColor = resolveColor(font.color, palette)
  if (fColor && !/^000000$/i.test(fColor)) parts.push(`color:#${fColor}`)

  const fill = cell.fill || {}
  if (fill.type === 'pattern' && fill.pattern === 'solid') {
    const bg = resolveColor(fill.fgColor, palette)
    if (bg && !/^FFFFFF$/i.test(bg)) parts.push(`background-color:#${bg}`)
  }

  const align = cell.alignment || {}
  if (align.horizontal === 'center')      parts.push('text-align:center')
  else if (align.horizontal === 'right')  parts.push('text-align:right')
  else if (align.horizontal === 'left')   parts.push('text-align:left')

  return parts.join(';')
}

// Parse HYPERLINK("url", "text") formula → { url, text } or null
function parseHyperlink(formula) {
  if (!formula) return null
  const match = formula.match(/^HYPERLINK\(\s*"([^"]+)"\s*,\s*"([^"]+)"\s*\)$/i)
  if (match) return { url: match[1], text: match[2] }
  // Single-arg: HYPERLINK("url")
  const match2 = formula.match(/^HYPERLINK\(\s*"([^"]+)"\s*\)$/i)
  if (match2) return { url: match2[1], text: match2[1] }
  return null
}

// Decode XML/HTML numeric and named character references (e.g. &#225; → á)
function decodeEntities(s) {
  if (typeof s !== 'string' || s.indexOf('&') === -1) return s
  return s
    .replace(/&#x([0-9a-fA-F]+);/g, (_, h) => String.fromCodePoint(parseInt(h, 16)))
    .replace(/&#(\d+);/g, (_, d) => String.fromCodePoint(parseInt(d, 10)))
    .replace(/&amp;/g, '&')
    .replace(/&lt;/g, '<')
    .replace(/&gt;/g, '>')
    .replace(/&quot;/g, '"')
    .replace(/&apos;/g, "'")
}

// Format an ExcelJS cell value into a display string, returning { v, num, link }
function formatCell(cell) {
  const val = cell.value
  if (val === null || val === undefined) return { v: '', num: null }

  // Native hyperlink: { text: "..." or {richText: [...]}, hyperlink: "http://..." }
  if (typeof val === 'object' && val.hyperlink) {
    let text = val.hyperlink
    if (typeof val.text === 'string') {
      text = val.text
    } else if (val.text && val.text.richText) {
      text = val.text.richText.map((rt) => rt.text).join('')
    }
    return { v: text, num: null, link: val.hyperlink }
  }

  // Formula cell — use cached result
  if (typeof val === 'object' && (val.formula !== undefined || val.sharedFormula !== undefined)) {
    const formula = val.formula || val.sharedFormula
    // Check for HYPERLINK formula
    const hl = parseHyperlink(formula)
    if (hl) return { v: hl.text, num: null, link: hl.url }

    const result = val.result
    if (result === null || result === undefined) return { v: null, num: null, formula } // uncached
    if (typeof result === 'number') {
      return { v: formatNumber(result, cell.numFmt), num: result }
    }
    if (result instanceof Date) {
      return { v: formatDate(result, cell.numFmt), num: toExcelSerial(result) }
    }
    return { v: String(result), num: null }
  }

  // Rich text
  if (typeof val === 'object' && val.richText) {
    return { v: val.richText.map((rt) => rt.text).join(''), num: null }
  }

  // Error
  if (typeof val === 'object' && val.error) {
    return { v: val.error, num: null }
  }

  if (typeof val === 'number') {
    return { v: formatNumber(val, cell.numFmt), num: val }
  }

  if (val instanceof Date) {
    // Expose the Excel serial as num so aggregate formulas (MIN/MAX/AVERAGE over
    // date columns) can operate on date cells.
    return { v: formatDate(val, cell.numFmt), num: toExcelSerial(val) }
  }

  if (typeof val === 'boolean') {
    return { v: val ? 'TRUE' : 'FALSE', num: null }
  }

  return { v: String(val), num: null }
}

// Format a number using SSF (same engine SheetJS uses)
function formatNumber(n, numFmt) {
  if (!numFmt || numFmt === 'General') return String(n)
  try {
    return SSF.format(numFmt, n)
  } catch {
    return String(n)
  }
}

// Format a date using SSF or fallback to locale string.
// ExcelJS returns dates as UTC midnight; converting to an Excel serial keeps
// the calendar day stable regardless of the host timezone (otherwise users
// west of UTC would see dates shifted one day earlier).
function formatDate(d, numFmt) {
  const utcMidnight = new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate())
  if (!numFmt || numFmt === 'General') return utcMidnight.toLocaleDateString()
  try {
    // SSF returns '' (rather than throwing) for serials outside Excel's
    // valid date range, e.g. a year past 9999 — treat that as a failure too.
    const s = SSF.format(numFmt, toExcelSerial(d))
    if (s !== '') return s
  } catch {}
  return utcMidnight.toLocaleDateString()
}

// Convert column number (1-based) to Excel letter (1=A, 2=B, ..., 27=AA)
function colToLetter(c) {
  let name = ''
  while (c > 0) {
    c--
    name = String.fromCharCode(65 + (c % 26)) + name
    c = Math.floor(c / 26)
  }
  return name
}


// Parse an Excel range like "B1:D1" into {s:{r,c}, e:{r,c}} (0-based)
function decodeRange(ref) {
  const m = ref.match(/^([A-Z]+)(\d+):([A-Z]+)(\d+)$/)
  if (!m) return null
  const colFromLetter = (s) => {
    let n = 0
    for (const ch of s) n = n * 26 + (ch.charCodeAt(0) - 64)
    return n - 1
  }
  return {
    s: { r: parseInt(m[2]) - 1, c: colFromLetter(m[1]) },
    e: { r: parseInt(m[4]) - 1, c: colFromLetter(m[3]) },
  }
}

// Parse a single worksheet (ExcelJS only)
function parseSheet(exWs, palette, hlMap) {
  if (!exWs || exWs.rowCount === 0) return { rows: [], colWidths: [] }

  // exWs.rowCount / columnCount can be inflated when a sheet has formatting on
  // empty trailing rows/columns. Use the largest populated row/column instead,
  // otherwise we'd allocate millions of empty cells and OOM the renderer.
  let maxRow = 0
  let maxCol = 0
  exWs.eachRow({ includeEmpty: false }, (row, rn) => {
    if (rn > maxRow) maxRow = rn
    row.eachCell({ includeEmpty: false }, (_cell, cn) => {
      if (cn > maxCol) maxCol = cn
    })
  })
  if (maxRow === 0 || maxCol === 0) return { rows: [], colWidths: [] }

  // Build merge map from worksheet model: "r,c" (1-based) -> {rowspan,colspan} or {skip:true}
  const mergeMap = {}
  const mergeRefs = exWs.model?.merges || []
  for (const ref of mergeRefs) {
    const range = decodeRange(ref)
    if (!range) continue
    const { s, e } = range
    mergeMap[`${s.r + 1},${s.c + 1}`] = { rowspan: e.r - s.r + 1, colspan: e.c - s.c + 1 }
    for (let rr = s.r; rr <= e.r; rr++) {
      for (let cc = s.c; cc <= e.c; cc++) {
        if (rr === s.r && cc === s.c) continue
        mergeMap[`${rr + 1},${cc + 1}`] = { skip: true }
      }
    }
  }

  const rows = []
  const rawNums = []
  const formulas = []   // formula string for uncached cells, else null
  const numFmts  = []   // numFmt for uncached formula cells

  for (let r = 1; r <= maxRow; r++) {
    const row = []
    const numRow = []
    const fRow = []
    const fmtRow = []
    const exRow = exWs.getRow(r)

    for (let c = 1; c <= maxCol; c++) {
      const cell = exRow.getCell(c)
      const css  = getCellCSS(cell, palette)
      let { v, num, link, formula } = formatCell(cell)

      // Fallback: if we have a hyperlink without display text, check raw XML map
      if (link && (v === link || !v) && hlMap) {
        const addr = colToLetter(c) + r
        const hlData = hlMap[addr]
        if (hlData) { v = hlData.text; link = hlData.url }
      }

      const cellData = { v, css }
      if (link) cellData.link = link
      const m = mergeMap[`${r},${c}`]
      if (m) {
        if (m.skip) cellData.skip = true
        else { cellData.rowspan = m.rowspan; cellData.colspan = m.colspan }
      }
      row.push(cellData)
      numRow.push(num)
      fRow.push(formula || null)
      fmtRow.push(formula ? cell.numFmt : null)
    }
    rows.push(row)
    rawNums.push(numRow)
    formulas.push(fRow)
    numFmts.push(fmtRow)
  }

  // Evaluate uncached aggregate formulas (SUM/MIN/MAX/AVERAGE/AVERAGEIF/...) from
  // rawNums so cells whose stored result was left blank still show a value.
  // Iterate because a column-total may depend on row-totals that are themselves
  // uncached; each pass resolves any cell whose inputs are now known.
  const evalCtx = {
    rows: rows.length,
    cols: maxCol,
    get: (r, c) => rawNums[r][c],
    // A cell is pending if it still holds an unresolved formula.
    pending: (r, c) =>
      r >= 0 && r < rows.length && c >= 0 && c < maxCol &&
      rows[r][c].v === null && !!formulas[r][c],
  }
  for (let pass = 0; pass < rows.length + 1; pass++) {
    let changed = false
    for (let r = 0; r < rows.length; r++) {
      for (let c = 0; c < maxCol; c++) {
        if (rows[r][c].v !== null || !formulas[r][c]) continue
        const res = evalFormula(formulas[r][c], evalCtx)
        if (!res || res.pending) continue // unsupported, or deps not ready yet
        if (res.value === null || res.value === undefined) {
          // Computed but empty (e.g. no rows matched a criteria) — render blank.
          rows[r][c].v = ''
        } else {
          rows[r][c].v = formatNumber(res.value, numFmts[r][c])
          rawNums[r][c] = res.value
        }
        changed = true
      }
    }
    if (!changed) break
  }

  // Any remaining null cells (unparseable formulas) become blank
  for (let r = 0; r < rows.length; r++) {
    for (let c = 0; c < maxCol; c++) {
      if (rows[r][c].v === null) rows[r][c].v = ''
    }
  }

  // Extract column widths (Excel units → pixels: ~7px per unit)
  const colWidths = []
  for (let c = 1; c <= maxCol; c++) {
    const w = exWs.getColumn(c)?.width
    colWidths.push(w ? Math.round(w * 7) : null)
  }

  // Frozen panes (xSplit = frozen cols, ySplit = frozen rows)
  const view = (exWs.views || [])[0]
  const freezeRows = view && view.state === 'frozen' ? (view.ySplit || 0) : 0
  const freezeCols = view && view.state === 'frozen' ? (view.xSplit || 0) : 0

  return {
    rows: rows.map((row) => row.map((cell) => {
      const rawV = cell.v ?? ''
      const out = { v: typeof rawV === 'string' ? decodeEntities(rawV) : rawV, css: cell.css }
      if (cell.link) out.link = decodeEntities(cell.link)
      if (cell.skip) out.skip = true
      if (cell.rowspan) out.rowspan = cell.rowspan
      if (cell.colspan) out.colspan = cell.colspan
      return out
    })),
    colWidths,
    freezeRows,
    freezeCols,
  }
}

// Build inline CSS for a SheetJS cell's style block
function getXlsCellCSS(cell) {
  const s = cell.s
  if (!s) return ''
  const parts = []

  const font = s.font || {}
  if (font.bold)      parts.push('font-weight:bold')
  if (font.italic)    parts.push('font-style:italic')
  if (font.underline) parts.push('text-decoration:underline')
  if (font.strike)    parts.push('text-decoration:line-through')
  if (font.sz)        parts.push(`font-size:${font.sz}pt`)

  if (font.color && font.color.rgb) {
    const hex = font.color.rgb.length === 8 ? font.color.rgb.slice(2) : font.color.rgb
    if (!/^000000$/i.test(hex)) parts.push(`color:#${hex}`)
  }

  const fill = s.fill || {}
  const fillRgb = fill.fgColor?.rgb || fill.bgColor?.rgb
  if (fillRgb) {
    const hex = fillRgb.length === 8 ? fillRgb.slice(2) : fillRgb
    if (!/^FFFFFF$/i.test(hex)) parts.push(`background-color:#${hex}`)
  }

  const align = s.alignment || {}
  if (align.horizontal === 'center')      parts.push('text-align:center')
  else if (align.horizontal === 'right')  parts.push('text-align:right')
  else if (align.horizontal === 'left')   parts.push('text-align:left')

  return parts.join(';')
}

// Convert a SheetJS cell to the renderer's { v, css, link? } shape
function formatXlsCell(cell) {
  if (!cell) return { v: '', css: '' }

  let v = ''
  if (cell.w !== undefined && cell.w !== null) {
    v = cell.w
  } else if (cell.v !== undefined && cell.v !== null) {
    if (cell.v instanceof Date) v = cell.v.toLocaleDateString()
    else if (typeof cell.v === 'boolean') v = cell.v ? 'TRUE' : 'FALSE'
    else v = String(cell.v)
  }

  const out = { v: decodeEntities(v), css: getXlsCellCSS(cell) }
  if (cell.l && cell.l.Target) out.link = decodeEntities(cell.l.Target)
  return out
}

// Parse a single SheetJS worksheet into { rows, colWidths }
function parseXlsSheet(ws) {
  if (!ws || !ws['!ref']) return { rows: [], colWidths: [] }

  const range = XLSX.utils.decode_range(ws['!ref'])
  const minRow = range.s.r
  const maxRow = range.e.r
  const minCol = range.s.c
  const maxCol = range.e.c

  // Build merge map from SheetJS !merges: "r,c" (0-based) -> {rowspan,colspan} or {skip:true}
  const mergeMap = {}
  const wsMerges = ws['!merges'] || []
  for (const { s, e } of wsMerges) {
    mergeMap[`${s.r},${s.c}`] = { rowspan: e.r - s.r + 1, colspan: e.c - s.c + 1 }
    for (let rr = s.r; rr <= e.r; rr++) {
      for (let cc = s.c; cc <= e.c; cc++) {
        if (rr === s.r && cc === s.c) continue
        mergeMap[`${rr},${cc}`] = { skip: true }
      }
    }
  }

  const rows = []
  for (let r = minRow; r <= maxRow; r++) {
    const row = []
    for (let c = minCol; c <= maxCol; c++) {
      const addr = XLSX.utils.encode_cell({ r, c })
      const cellData = formatXlsCell(ws[addr])
      const m = mergeMap[`${r},${c}`]
      if (m) {
        if (m.skip) cellData.skip = true
        else { cellData.rowspan = m.rowspan; cellData.colspan = m.colspan }
      }
      row.push(cellData)
    }
    rows.push(row)
  }

  // Column widths: SheetJS exposes wpx (pixels) or wch (chars); fall back to ~7px per char
  const cols = ws['!cols'] || []
  const colWidths = []
  for (let c = minCol; c <= maxCol; c++) {
    const col = cols[c]
    if (col?.wpx) colWidths.push(Math.round(col.wpx))
    else if (col?.wch) colWidths.push(Math.round(col.wch * 7))
    else colWidths.push(null)
  }

  const freeze = ws['!freeze']
  const freezeRows = freeze?.ySplit || freeze?.r || 0
  const freezeCols = freeze?.xSplit || freeze?.c || 0

  return { rows, colWidths, freezeRows, freezeCols }
}

// Parse a legacy .xls (BIFF) file using SheetJS — ExcelJS doesn't support this format
function parseXlsBuffer(buffer, fileName) {
  const wb = XLSX.read(buffer, {
    type: 'buffer',
    cellStyles: true,
    cellNF: true,
    cellDates: true,
    cellFormula: true,
    cellHTML: false,
  })

  const sheetNames = wb.SheetNames
  const sheets = {}
  for (const name of sheetNames) {
    sheets[name] = parseXlsSheet(wb.Sheets[name])
  }
  return { fileName, sheetNames, sheets }
}

// Parse a CSV file into the same format as parseSheet output
function parseCsvContent(text) {
  // Parse CSV handling quoted fields with commas/newlines
  const rows = []
  let current = ''
  let inQuotes = false
  let row = []

  for (let i = 0; i < text.length; i++) {
    const ch = text[i]
    if (inQuotes) {
      if (ch === '"' && text[i + 1] === '"') {
        current += '"'
        i++
      } else if (ch === '"') {
        inQuotes = false
      } else {
        current += ch
      }
    } else {
      if (ch === '"') {
        inQuotes = true
      } else if (ch === ',') {
        row.push(current)
        current = ''
      } else if (ch === '\n' || (ch === '\r' && text[i + 1] === '\n')) {
        row.push(current)
        current = ''
        rows.push(row)
        row = []
        if (ch === '\r') i++
      } else if (ch === '\r') {
        row.push(current)
        current = ''
        rows.push(row)
        row = []
      } else {
        current += ch
      }
    }
  }
  // Last field/row. Skip it only when it's the empty artifact of a trailing
  // newline (no leftover field text and no fields already collected this row);
  // genuine blank lines between data are pushed in the loop above and kept.
  if (current !== '' || row.length > 0) {
    row.push(current)
    rows.push(row)
  }

  // Convert to cell format
  const maxCols = rows.reduce((max, r) => Math.max(max, r.length), 0)
  const cellRows = rows.map((r) => {
    const cells = []
    for (let c = 0; c < maxCols; c++) {
      cells.push({ v: r[c] || '', css: '' })
    }
    return cells
  })

  return { rows: cellRows, colWidths: [] }
}

// An openable file whose contents aren't a workbook we can read. Carries a
// human-readable reason so the UI can say *why* rather than just failing.
function unsupportedFileError(detail) {
  const err = new Error(detail)
  err.code = 'EUNSUPPORTED'
  return err
}

// Best-effort identification of what a file actually is, by content rather than
// by extension — a file renamed to .xlsx is the common case (e.g. an ODS export
// saved with an .xlsx name). Returns { name, ext, convertible }, or null if
// unrecognized. `convertible` marks formats a spreadsheet app can re-save as
// .xlsx, so we only suggest that when it's actually useful advice.
async function identifyFormat(buffer) {
  if (buffer.length >= 4 && buffer.readUInt32BE(0) === 0xd0cf11e0) {
    return { name: 'a legacy Microsoft Office file', ext: '.xls / .doc / .ppt', convertible: true }
  }
  if (buffer.slice(0, 4).toString('latin1') === '%PDF') {
    return { name: 'a PDF', ext: '.pdf', convertible: false }
  }
  if (buffer.slice(0, 2).toString('latin1') !== 'PK') return null

  try {
    const zip = await JSZip.loadAsync(buffer)
    const mimetype = (await zip.file('mimetype')?.async('text'))?.trim() || ''
    const ODF = 'application/vnd.oasis.opendocument.'
    if (mimetype.startsWith(ODF)) {
      const kind = mimetype.slice(ODF.length)
      const names = {
        spreadsheet: ['an OpenDocument Spreadsheet', '.ods', true],
        text: ['an OpenDocument Text document', '.odt', false],
        presentation: ['an OpenDocument Presentation', '.odp', false],
      }
      const [name, ext, convertible] = names[kind] || ['an OpenDocument file', '.' + kind, false]
      return { name, ext, convertible }
    }
    if (zip.file('word/document.xml')) {
      return { name: 'a Word document', ext: '.docx', convertible: false }
    }
    if (zip.file(/^ppt\/slides\//).length) {
      return { name: 'a PowerPoint presentation', ext: '.pptx', convertible: false }
    }
    return { name: 'a Zip archive, not a spreadsheet', ext: '.zip', convertible: false }
  } catch {
    return null
  }
}

async function parseFile(filePath) {
  const ext = path.extname(filePath).toLowerCase()
  const buffer = fs.readFileSync(filePath)
  const data = await parseFileBuffer(buffer, ext, path.basename(filePath))
  addRecentFile(filePath) // only remember files we could actually read
  return data
}

async function parseFileBuffer(buffer, ext, fileName) {
  // CSV/TSV: plain text parsing
  if (ext === '.csv' || ext === '.tsv') {
    const text = buffer.toString('utf-8')
    const sheetName = 'Sheet1'
    return {
      fileName,
      sheetNames: [sheetName],
      sheets: { [sheetName]: parseCsvContent(text) },
    }
  }

  // Legacy .xls (BIFF) — ExcelJS only handles OOXML, so use SheetJS
  if (ext === '.xls') {
    let xlsData
    try {
      xlsData = parseXlsBuffer(buffer, fileName)
    } catch {
      throw unsupportedFileError(await unsupportedDetail(buffer, fileName, ext))
    }
    if (xlsData.sheetNames.length === 0) {
      throw unsupportedFileError(await unsupportedDetail(buffer, fileName, ext))
    }
    return xlsData
  }

  // Modern Excel formats (.xlsx, .xlsm)
  const meta = await parseXlsxMeta(buffer)
  const palette = meta.palette || DEFAULT_INDEXED_COLORS

  const exWb = new ExcelJS.Workbook()
  try {
    await exWb.xlsx.load(buffer)
  } catch {
    throw unsupportedFileError(await unsupportedDetail(buffer, fileName, ext))
  }

  // ExcelJS resolves with an empty workbook rather than throwing when the zip
  // isn't OOXML at all, so an empty sheet list means "couldn't read it".
  if (exWb.worksheets.length === 0) {
    throw unsupportedFileError(await unsupportedDetail(buffer, fileName, ext))
  }

  const sheetNames = exWb.worksheets.map((ws) => ws.name)
  const sheets = {}
  for (const ws of exWb.worksheets) {
    const hlMap = meta.hyperlinkMap[ws.id] || null
    sheets[ws.name] = parseSheet(ws, palette, hlMap)
  }

  return { fileName, sheetNames, sheets }
}

const SUPPORTED_LIST = 'Excel Reader can open .xlsx, .xlsm, .xls, .csv, and .tsv files.'
const CONVERT_HINT =
  'Open it in Excel, Numbers, or LibreOffice and save it as .xlsx to view it here.'

// Build the explanation shown to the user for a file we can't read.
async function unsupportedDetail(buffer, fileName, ext) {
  const actual = await identifyFormat(buffer)
  if (!actual) {
    return `"${fileName}" isn't a spreadsheet Excel Reader can read — it may be ` +
      `damaged, or saved in an unsupported format.\n\n${SUPPORTED_LIST}`
  }
  // Only call out the mismatch when the name actually disagrees with the bytes.
  const lead = actual.ext === ext
    ? `"${fileName}" is ${actual.name}, which Excel Reader can't open.`
    : `"${fileName}" is ${actual.name} (${actual.ext}), despite its name.`
  return `${lead} ${SUPPORTED_LIST}` + (actual.convertible ? `\n\n${CONVERT_HINT}` : '')
}

ipcMain.handle('open-and-parse-file', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender)
  const { canceled, filePaths } = await dialog.showOpenDialog(win, {
    properties: ['openFile'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls', 'xlsm', 'csv', 'tsv'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  })
  if (canceled || filePaths.length === 0) return null

  const filePath = filePaths[0]
  let data
  try {
    data = await parseFile(filePath)
  } catch (err) {
    reportParseError(err, filePath, win)
    return null // renderer skips loading on null
  }

  // Open in a new window if the current window already has a file loaded
  const isEmptyWindow = !win._hasFile
  if (isEmptyWindow) {
    win._hasFile = true
    return data
  } else {
    const newWin = createWindow()
    newWin._hasFile = true
    newWin.webContents.on('did-finish-load', () => {
      newWin.webContents.send('open-file', filePath)
    })
    return null // signal to caller: file opened in new window
  }
})

ipcMain.handle('parse-file', async (event, filePath) => {
  try {
    return await parseFile(filePath)
  } catch (err) {
    reportParseError(err, filePath, BrowserWindow.fromWebContents(event.sender))
    return null // renderer skips loading on null
  }
})

// Return recent files as { path, name, dir } for the empty-state list.
const HOME_DIR = require('os').homedir()
ipcMain.handle('get-recent-files', async () => {
  return getRecentFiles().map((p) => ({
    path: p,
    name: path.basename(p),
    dir: path.dirname(p).replace(HOME_DIR, '~'),
  }))
})
