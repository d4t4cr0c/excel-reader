let workbookData = null

const openBtn = document.getElementById('open-btn')
const openBtnCenter = document.getElementById('open-btn-center')
const dropZone = document.getElementById('drop-zone')
const tableContainer = document.getElementById('table-container')
const sheetTabs = document.getElementById('sheet-tabs')

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
}

function renderSheet(sheetData) {
  dropZone.classList.add('hidden')
  tableContainer.classList.remove('hidden')

  const rows = sheetData.rows || sheetData
  const colWidths = sheetData.colWidths || []

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

// Handle files opened via Finder (right-click > Open With, etc.)
window.api.onOpenFile(async (filePath) => {
  const result = await window.api.parseFile(filePath)
  if (result) loadWorkbookData(result)
})

document.addEventListener('dragover', (e) => e.preventDefault())
document.addEventListener('drop', (e) => {
  e.preventDefault()
})
