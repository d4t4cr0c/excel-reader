const { app, BrowserWindow, ipcMain, dialog, shell } = require('electron')
const path = require('path')
const fs = require('fs')
const ExcelJS = require('exceljs')
const JSZip = require('jszip')
const SSF = require('ssf')

let pendingFilePaths = [] // files received before app is ready

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
    createWindow(filePath)
  } else {
    pendingFilePaths.push(filePath)
  }
})

app.whenReady().then(() => {
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

// Format an ExcelJS cell value into a display string, returning { v, num, link }
function formatCell(cell) {
  const val = cell.value
  if (val === null || val === undefined) return { v: '', num: null }

  // Native hyperlink: { text: "...", hyperlink: "http://..." }
  if (typeof val === 'object' && val.hyperlink) {
    const text = val.text || val.hyperlink
    return { v: String(text), num: null, link: val.hyperlink }
  }

  // Formula cell — use cached result
  if (typeof val === 'object' && (val.formula !== undefined || val.sharedFormula !== undefined)) {
    // Check for HYPERLINK formula
    const hl = parseHyperlink(val.formula || val.sharedFormula)
    if (hl) return { v: hl.text, num: null, link: hl.url }

    const result = val.result
    if (result === null || result === undefined) return { v: null, num: null } // uncached
    if (typeof result === 'number') {
      return { v: formatNumber(result, cell.numFmt), num: result }
    }
    if (result instanceof Date) {
      return { v: formatDate(result, cell.numFmt), num: null }
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
    return { v: formatDate(val, cell.numFmt), num: null }
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

// Format a date using SSF or fallback to locale string
function formatDate(d, numFmt) {
  if (!numFmt || numFmt === 'General') return d.toLocaleDateString()
  try {
    return SSF.format(numFmt, d)
  } catch {
    return d.toLocaleDateString()
  }
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

// Parse a single worksheet (ExcelJS only)
function parseSheet(exWs, palette, hlMap) {
  if (!exWs || exWs.rowCount === 0) return { rows: [], colWidths: [] }

  const maxRow = exWs.rowCount
  const maxCol = exWs.columnCount

  const rows = []
  const rawNums = []

  for (let r = 1; r <= maxRow; r++) {
    const row = []
    const numRow = []
    const exRow = exWs.getRow(r)

    for (let c = 1; c <= maxCol; c++) {
      const cell = exRow.getCell(c)
      const css  = getCellCSS(cell, palette)
      let { v, num, link } = formatCell(cell)

      // Fallback: if we have a hyperlink without display text, check raw XML map
      if (link && (v === link || !v) && hlMap) {
        const addr = colToLetter(c) + r
        const hlData = hlMap[addr]
        if (hlData) { v = hlData.text; link = hlData.url }
      }

      const cellData = { v, css }
      if (link) cellData.link = link
      row.push(cellData)
      numRow.push(num)
    }
    rows.push(row)
    rawNums.push(numRow)
  }

  // Compute missing values for rows with mixed null/non-null cells (uncached SUM rows)
  for (let r = 1; r < rows.length; r++) {
    const row = rows[r]
    const hasNull  = row.some((c) => c.v === null)
    const hasValue = row.some((c) => c.v !== null && c.v !== '')
    if (!hasNull || !hasValue) continue

    for (let c = 0; c < maxCol; c++) {
      if (row[c].v !== null) continue
      let sum = 0, hasNums = false
      for (let pr = 0; pr < r; pr++) {
        const n = rawNums[pr][c]
        if (n !== null) { sum += n; hasNums = true }
      }
      row[c].v      = hasNums ? sum.toLocaleString('en-US', { minimumFractionDigits: 2, maximumFractionDigits: 2 }) : ''
      rawNums[r][c] = hasNums ? sum : null
    }
  }

  // Extract column widths (Excel units → pixels: ~7px per unit)
  const colWidths = []
  for (let c = 1; c <= maxCol; c++) {
    const w = exWs.getColumn(c)?.width
    colWidths.push(w ? Math.round(w * 7) : null)
  }

  return {
    rows: rows.map((row) => row.map((cell) => {
      const out = { v: cell.v ?? '', css: cell.css }
      if (cell.link) out.link = cell.link
      return out
    })),
    colWidths,
  }
}

async function parseFile(filePath) {
  const buffer = fs.readFileSync(filePath)

  const meta = await parseXlsxMeta(buffer)
  const palette = meta.palette || DEFAULT_INDEXED_COLORS

  const exWb = new ExcelJS.Workbook()
  await exWb.xlsx.load(buffer)

  const sheetNames = exWb.worksheets.map((ws) => ws.name)
  const sheets = {}
  for (const ws of exWb.worksheets) {
    const hlMap = meta.hyperlinkMap[ws.id] || null
    sheets[ws.name] = parseSheet(ws, palette, hlMap)
  }

  return { fileName: path.basename(filePath), sheetNames, sheets }
}

ipcMain.handle('open-and-parse-file', async (event) => {
  const win = BrowserWindow.fromWebContents(event.sender)
  const { canceled, filePaths } = await dialog.showOpenDialog(win, {
    properties: ['openFile'],
    filters: [
      { name: 'Excel Files', extensions: ['xlsx', 'xls', 'xlsm', 'csv'] },
      { name: 'All Files', extensions: ['*'] },
    ],
  })
  if (canceled || filePaths.length === 0) return null

  const filePath = filePaths[0]
  const data = await parseFile(filePath)

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

ipcMain.handle('parse-file', async (_event, filePath) => {
  return parseFile(filePath)
})
