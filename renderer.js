let workbookData = null
let currentSheetRows = null
let currentNumRows = 0
let currentNumCols = 0

const openBtn = document.getElementById('open-btn')
const openBtnCenter = document.getElementById('open-btn-center')
const dropZone = document.getElementById('drop-zone')
const tableContainer = document.getElementById('table-container')
const sheetTabs = document.getElementById('sheet-tabs')
const searchInput = document.getElementById('search-input')
const searchCount = document.getElementById('search-count')

let searchMatches = []
let currentMatchIndex = -1
let lastSearchQuery = ''

let selectionAnchor = null
let selectionFocus = null
let isSelecting = false

async function openFile() {
  const result = await window.api.openAndParseFile()
  if (!result) return // null means file was opened in a new window
  loadWorkbookData(result)
}

function loadWorkbookData(data) {
  workbookData = data
  document.title = data.fileName + ' — Excel Reader'
  document.getElementById('app-title').textContent = data.fileName
  renderTabs()
  selectSheet(data.sheetNames[0])
}

function renderTabs() {
  sheetTabs.innerHTML = ''
  sheetTabs.classList.remove('hidden')
  workbookData.sheetNames.forEach((name) => {
    const tab = document.createElement('div')
    tab.className = 'tab'
    tab.textContent = name
    tab.addEventListener('click', () => selectSheet(name))
    sheetTabs.appendChild(tab)
  })
}

function selectSheet(name) {
  document.querySelectorAll('#sheet-tabs .tab').forEach((t) => {
    t.classList.toggle('active', t.textContent === name)
  })
  if (workbookData) renderSheet(workbookData.sheets[name])
  clearSearch()
  clearSelection()
}

function renderSheet(sheetData) {
  dropZone.classList.add('hidden')
  tableContainer.classList.remove('hidden')

  const rows = sheetData.rows || sheetData
  const colWidths = sheetData.colWidths || []
  const freezeRows = sheetData.freezeRows || 0
  const freezeCols = sheetData.freezeCols || 0

  if (!rows || rows.length === 0) {
    tableContainer.innerHTML = '<p style="padding:20px;color:#6e6e73">This sheet is empty.</p>'
    return
  }

  const numCols = rows.reduce((max, row) => Math.max(max, row.length), 0)
  currentSheetRows = rows
  currentNumRows = rows.length
  currentNumCols = numCols

  const colHeaders = []
  for (let c = 0; c < numCols; c++) {
    colHeaders.push(colLetter(c))
  }

  let html = '<table>'

  html += '<colgroup><col style="width:40px">'
  for (let c = 0; c < numCols; c++) {
    const w = colWidths[c]
    html += w ? `<col style="width:${w}px">` : '<col>'
  }
  html += '</colgroup>'

  html += '<thead><tr><th class="row-header"></th>'
  colHeaders.forEach((h, c) => { html += `<th data-col="${c}">${h}</th>` })
  html += '</tr></thead><tbody>'

  rows.forEach((row, r) => {
    html += `<tr><td class="row-num" data-rownum="${r}">${r + 1}</td>`
    for (let c = 0; c < numCols; c++) {
      const cell = row[c] || { v: '', css: '' }
      if (cell.skip) continue
      const value = cell.v !== undefined ? String(cell.v) : ''
      const isNum = value !== '' && !isNaN(value) && !value.includes('/')
      const styleAttr = cell.css ? ` style="${escapeAttr(cell.css)}"` : ''
      const rowspan = cell.rowspan > 1 ? cell.rowspan : 1
      const colspan = cell.colspan > 1 ? cell.colspan : 1
      const spanAttr =
        (rowspan > 1 ? ` rowspan="${rowspan}"` : '') +
        (colspan > 1 ? ` colspan="${colspan}"` : '')
      const content = cell.link
        ? `<a href="${escapeAttr(cell.link)}" target="_blank" rel="noopener">${escapeHtml(value)}</a>`
        : escapeHtml(value)
      html += `<td class="${isNum ? 'numeric' : ''}"${styleAttr}${spanAttr} data-row="${r}" data-col="${c}" data-rowspan="${rowspan}" data-colspan="${colspan}" title="${escapeAttr(value)}">${content}</td>`
    }
    html += '</tr>'
  })

  html += '</tbody></table>'
  tableContainer.innerHTML = html

  applyFreeze(freezeRows, freezeCols)
}

function applyFreeze(freezeRows, freezeCols) {
  const table = tableContainer.querySelector('table')
  if (!table) return
  const thead = table.querySelector('thead')
  const tbody = table.querySelector('tbody')
  if (!thead || !tbody) return

  const headerHeight = thead.offsetHeight || 0
  const bodyRows = Array.from(tbody.querySelectorAll('tr'))

  // Frozen rows: first N tbody rows stick below the header
  let topAcc = headerHeight
  for (let i = 0; i < freezeRows && i < bodyRows.length; i++) {
    const tr = bodyRows[i]
    Array.from(tr.children).forEach((cell) => {
      cell.classList.add('frozen-row')
      cell.style.top = `${topAcc}px`
    })
    topAcc += tr.offsetHeight
  }

  if (freezeCols <= 0) return

  // Frozen columns: row-num + first N data cols stick to the left
  const headerTr = thead.querySelector('tr')
  const allRows = headerTr ? [headerTr, ...bodyRows] : bodyRows
  if (allRows.length === 0) return

  const firstRow = bodyRows[0] || headerTr
  const widths = Array.from(firstRow.children).map((c) => c.offsetWidth)

  const stickyLefts = [0]
  for (let i = 0; i <= freezeCols; i++) stickyLefts.push((stickyLefts[i] || 0) + (widths[i] || 0))

  for (const tr of allRows) {
    const cells = Array.from(tr.children)
    for (let i = 0; i <= freezeCols && i < cells.length; i++) {
      cells[i].classList.add('frozen-col')
      cells[i].style.left = `${stickyLefts[i]}px`
    }
  }
}

function colLetter(index) {
  let name = ''
  index++
  while (index > 0) {
    index--
    name = String.fromCharCode(65 + (index % 26)) + name
    index = Math.floor(index / 26)
  }
  return name
}

function escapeHtml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
}

function escapeAttr(str) {
  return str.replace(/"/g, '&quot;')
}

openBtn.addEventListener('click', openFile)
openBtnCenter.addEventListener('click', openFile)

tableContainer.addEventListener('dblclick', (e) => {
  const td = e.target.closest('td')
  if (!td || td.classList.contains('row-num')) return
  if (e.target.closest('a')) return
  td.classList.toggle('cell-expanded')
})

function clearSelection() {
  selectionAnchor = null
  selectionFocus = null
  isSelecting = false
  tableContainer.querySelectorAll('td.cell-selected').forEach((td) => {
    td.classList.remove('cell-selected')
  })
}

function selectionRect() {
  if (!selectionAnchor || !selectionFocus) return null
  return {
    r1: Math.min(selectionAnchor.r, selectionFocus.r),
    r2: Math.max(selectionAnchor.r, selectionFocus.r),
    c1: Math.min(selectionAnchor.c, selectionFocus.c),
    c2: Math.max(selectionAnchor.c, selectionFocus.c),
  }
}

function applySelection() {
  tableContainer.querySelectorAll('td.cell-selected').forEach((td) => {
    td.classList.remove('cell-selected')
  })
  const rect = selectionRect()
  if (!rect) return
  const cells = tableContainer.querySelectorAll('td[data-row]')
  cells.forEach((td) => {
    const r = parseInt(td.dataset.row, 10)
    const c = parseInt(td.dataset.col, 10)
    const rs = parseInt(td.dataset.rowspan || '1', 10)
    const cs = parseInt(td.dataset.colspan || '1', 10)
    if (r > rect.r2 || r + rs - 1 < rect.r1) return
    if (c > rect.c2 || c + cs - 1 < rect.c1) return
    td.classList.add('cell-selected')
  })
}

function copySelection() {
  const rect = selectionRect()
  if (!rect || !currentSheetRows) return
  const lines = []
  for (let r = rect.r1; r <= rect.r2; r++) {
    const row = currentSheetRows[r] || []
    const cells = []
    for (let c = rect.c1; c <= rect.c2; c++) {
      const cell = row[c]
      if (!cell || cell.skip || cell.v === undefined || cell.v === null) {
        cells.push('')
        continue
      }
      const s = String(cell.v)
      cells.push(/[\t\n\r"]/.test(s) ? '"' + s.replace(/"/g, '""') + '"' : s)
    }
    lines.push(cells.join('\t'))
  }
  const text = lines.join('\n')
  navigator.clipboard.writeText(text).catch(() => {
    // Fallback for older contexts
    const ta = document.createElement('textarea')
    ta.value = text
    ta.style.position = 'fixed'
    ta.style.opacity = '0'
    document.body.appendChild(ta)
    ta.select()
    try { document.execCommand('copy') } catch (_) {}
    document.body.removeChild(ta)
  })
}

tableContainer.addEventListener('mousedown', (e) => {
  if (e.button !== 0) return
  if (e.target.closest('a')) return

  const rowNum = e.target.closest('td.row-num')
  if (rowNum) {
    e.preventDefault()
    const r = parseInt(rowNum.dataset.rownum, 10)
    if (Number.isNaN(r)) return
    if (e.shiftKey && selectionAnchor) {
      selectionFocus = { r, c: currentNumCols - 1 }
    } else {
      selectionAnchor = { r, c: 0 }
      selectionFocus = { r, c: currentNumCols - 1 }
    }
    isSelecting = false
    applySelection()
    return
  }

  const th = e.target.closest('th[data-col]')
  if (th) {
    e.preventDefault()
    const c = parseInt(th.dataset.col, 10)
    if (Number.isNaN(c)) return
    if (e.shiftKey && selectionAnchor) {
      selectionFocus = { r: currentNumRows - 1, c }
    } else {
      selectionAnchor = { r: 0, c }
      selectionFocus = { r: currentNumRows - 1, c }
    }
    isSelecting = false
    applySelection()
    return
  }

  const td = e.target.closest('td[data-row]')
  if (!td) return
  const r = parseInt(td.dataset.row, 10)
  const c = parseInt(td.dataset.col, 10)
  if (Number.isNaN(r) || Number.isNaN(c)) return

  if (e.shiftKey && selectionAnchor) {
    selectionFocus = { r, c }
  } else {
    selectionAnchor = { r, c }
    selectionFocus = { r, c }
  }
  isSelecting = true
  applySelection()
})

tableContainer.addEventListener('mousemove', (e) => {
  if (!isSelecting) return
  lastPointer = { x: e.clientX, y: e.clientY }
  updateFocusFromPointer(e.clientX, e.clientY)
  updateAutoScroll(e.clientX, e.clientY)
})

document.addEventListener('mouseup', () => {
  isSelecting = false
  stopAutoScroll()
})

let lastPointer = null
let autoScrollRAF = null
let autoScrollDX = 0
let autoScrollDY = 0

function updateFocusFromPointer(x, y) {
  let el = document.elementFromPoint(x, y)
  let td = el && el.closest ? el.closest('td[data-row]') : null
  if (!td) {
    const rect = tableContainer.getBoundingClientRect()
    const cx = Math.min(Math.max(x, rect.left + 1), rect.right - 1)
    const cy = Math.min(Math.max(y, rect.top + 1), rect.bottom - 1)
    el = document.elementFromPoint(cx, cy)
    td = el && el.closest ? el.closest('td[data-row]') : null
  }
  if (!td) return
  const r = parseInt(td.dataset.row, 10)
  const c = parseInt(td.dataset.col, 10)
  if (Number.isNaN(r) || Number.isNaN(c)) return
  if (selectionFocus && selectionFocus.r === r && selectionFocus.c === c) return
  selectionFocus = { r, c }
  applySelection()
}

function updateAutoScroll(x, y) {
  const rect = tableContainer.getBoundingClientRect()
  const margin = 40
  const maxSpeed = 24

  let dx = 0
  let dy = 0
  if (y < rect.top + margin) dy = -ramp(rect.top + margin - y, margin, maxSpeed)
  else if (y > rect.bottom - margin) dy = ramp(y - (rect.bottom - margin), margin, maxSpeed)
  if (x < rect.left + margin) dx = -ramp(rect.left + margin - x, margin, maxSpeed)
  else if (x > rect.right - margin) dx = ramp(x - (rect.right - margin), margin, maxSpeed)

  autoScrollDX = dx
  autoScrollDY = dy

  if (dx === 0 && dy === 0) {
    stopAutoScroll()
  } else if (autoScrollRAF === null) {
    autoScrollRAF = requestAnimationFrame(autoScrollTick)
  }
}

function ramp(distance, margin, maxSpeed) {
  const t = Math.min(distance / margin, 1)
  return Math.max(1, Math.round(t * maxSpeed))
}

function autoScrollTick() {
  autoScrollRAF = null
  if (!isSelecting) return
  if (autoScrollDX === 0 && autoScrollDY === 0) return
  const beforeLeft = tableContainer.scrollLeft
  const beforeTop = tableContainer.scrollTop
  tableContainer.scrollLeft = beforeLeft + autoScrollDX
  tableContainer.scrollTop = beforeTop + autoScrollDY
  if (lastPointer) updateFocusFromPointer(lastPointer.x, lastPointer.y)
  autoScrollRAF = requestAnimationFrame(autoScrollTick)
}

function stopAutoScroll() {
  autoScrollDX = 0
  autoScrollDY = 0
  if (autoScrollRAF !== null) {
    cancelAnimationFrame(autoScrollRAF)
    autoScrollRAF = null
  }
}

function clearSearch() {
  searchMatches = []
  currentMatchIndex = -1
  lastSearchQuery = ''
  if (searchCount) searchCount.textContent = ''
  tableContainer.querySelectorAll('td.search-current-cell').forEach((td) => {
    td.classList.remove('search-current-cell')
  })
  tableContainer.querySelectorAll('mark.search-hit').forEach((m) => {
    const parent = m.parentNode
    parent.replaceChild(document.createTextNode(m.textContent), m)
    parent.normalize()
  })
}

function highlightMatches(query) {
  const q = query.toLowerCase()
  const matches = []
  const cells = tableContainer.querySelectorAll('tbody td:not(.row-num)')
  cells.forEach((td) => {
    const text = td.textContent
    if (!text.toLowerCase().includes(q)) return
    highlightInNode(td, q)
    td.querySelectorAll('mark.search-hit').forEach((m) => matches.push(m))
  })
  return matches
}

function highlightInNode(node, qLower) {
  const walker = document.createTreeWalker(node, NodeFilter.SHOW_TEXT, null)
  const textNodes = []
  let n
  while ((n = walker.nextNode())) textNodes.push(n)
  textNodes.forEach((tn) => {
    const text = tn.nodeValue
    const lower = text.toLowerCase()
    let idx = lower.indexOf(qLower)
    if (idx === -1) return
    const frag = document.createDocumentFragment()
    let cursor = 0
    while (idx !== -1) {
      if (idx > cursor) frag.appendChild(document.createTextNode(text.slice(cursor, idx)))
      const mark = document.createElement('mark')
      mark.className = 'search-hit'
      mark.textContent = text.slice(idx, idx + qLower.length)
      frag.appendChild(mark)
      cursor = idx + qLower.length
      idx = lower.indexOf(qLower, cursor)
    }
    if (cursor < text.length) frag.appendChild(document.createTextNode(text.slice(cursor)))
    tn.parentNode.replaceChild(frag, tn)
  })
}

function focusMatch(i) {
  if (searchMatches.length === 0) return
  searchMatches.forEach((m) => m.classList.remove('current'))
  tableContainer.querySelectorAll('td.search-current-cell').forEach((td) => {
    td.classList.remove('search-current-cell')
  })
  currentMatchIndex = ((i % searchMatches.length) + searchMatches.length) % searchMatches.length
  const mark = searchMatches[currentMatchIndex]
  mark.classList.add('current')
  const td = mark.closest('td')
  if (td) td.classList.add('search-current-cell')
  mark.scrollIntoView({ behavior: 'smooth', block: 'center', inline: 'center' })
  searchCount.textContent = `${currentMatchIndex + 1}/${searchMatches.length}`
}

searchInput.addEventListener('keydown', (e) => {
  if (e.key === 'Escape') {
    searchInput.value = ''
    clearSearch()
    searchInput.blur()
    return
  }
  if ((e.key === 'Tab' || e.key === 'Enter') && searchMatches.length > 1) {
    e.preventDefault()
    focusMatch(currentMatchIndex + (e.shiftKey ? -1 : 1))
  }
})

searchInput.addEventListener('input', () => {
  const query = searchInput.value.trim()
  if (query === lastSearchQuery) return
  clearSearch()
  if (!query) return
  lastSearchQuery = query
  searchMatches = highlightMatches(query)
  if (searchMatches.length === 0) {
    searchCount.textContent = '0/0'
    return
  }
  focusMatch(0)
})

document.addEventListener('keydown', (e) => {
  if ((e.metaKey || e.ctrlKey) && e.key === 'f') {
    e.preventDefault()
    searchInput.focus()
    searchInput.select()
    return
  }

  const inInput = document.activeElement && (
    document.activeElement.tagName === 'INPUT' ||
    document.activeElement.tagName === 'TEXTAREA' ||
    document.activeElement.isContentEditable
  )

  if ((e.metaKey || e.ctrlKey) && (e.key === 'c' || e.key === 'C')) {
    if (inInput) return
    if (!selectionAnchor || !selectionFocus) return
    e.preventDefault()
    copySelection()
    return
  }

  if ((e.metaKey || e.ctrlKey) && (e.key === 'a' || e.key === 'A')) {
    if (inInput) return
    if (currentNumRows === 0 || currentNumCols === 0) return
    e.preventDefault()
    selectionAnchor = { r: 0, c: 0 }
    selectionFocus = { r: currentNumRows - 1, c: currentNumCols - 1 }
    applySelection()
    return
  }

  if (e.key === 'Escape' && !inInput) {
    clearSelection()
  }
})

// Handle files opened via Finder (right-click > Open With, etc.)
window.api.onOpenFile(async (filePath) => {
  const result = await window.api.parseFile(filePath)
  if (result) loadWorkbookData(result)
})

document.addEventListener('dragover', (e) => e.preventDefault())
document.addEventListener('drop', (e) => {
  e.preventDefault()
})
