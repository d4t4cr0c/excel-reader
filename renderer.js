let workbookData = null

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
  colHeaders.forEach((h) => { html += `<th>${h}</th>` })
  html += '</tr></thead><tbody>'

  rows.forEach((row, r) => {
    html += `<tr><td class="row-num">${r + 1}</td>`
    for (let c = 0; c < numCols; c++) {
      const cell = row[c] || { v: '', css: '' }
      if (cell.skip) continue
      const value = cell.v !== undefined ? String(cell.v) : ''
      const isNum = value !== '' && !isNaN(value) && !value.includes('/')
      const styleAttr = cell.css ? ` style="${escapeAttr(cell.css)}"` : ''
      const spanAttr =
        (cell.rowspan > 1 ? ` rowspan="${cell.rowspan}"` : '') +
        (cell.colspan > 1 ? ` colspan="${cell.colspan}"` : '')
      const content = cell.link
        ? `<a href="${escapeAttr(cell.link)}" target="_blank" rel="noopener">${escapeHtml(value)}</a>`
        : escapeHtml(value)
      html += `<td class="${isNum ? 'numeric' : ''}"${styleAttr}${spanAttr} title="${escapeAttr(value)}">${content}</td>`
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

tableContainer.addEventListener('click', (e) => {
  const td = e.target.closest('td')
  if (!td || td.classList.contains('row-num')) return
  if (e.target.closest('a')) return
  td.classList.toggle('cell-expanded')
})

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
