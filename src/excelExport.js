import ExcelJS from 'exceljs'

// ---------- Theme ----------
const NAVY = 'FF1E3A5F'
const BLUE = 'FF2563EB'
const LIGHT = 'FFEFF6FF'
const ZEBRA = 'FFF8FAFC'
const TOTAL_BG = 'FFDBEAFE'
const BORDER_COLOR = 'FF94A3B8'
const FONT = 'Calibri'

const thin = { style: 'thin', color: { argb: BORDER_COLOR } }
const BORDER = { top: thin, left: thin, bottom: thin, right: thin }
const medium = { style: 'medium', color: { argb: NAVY } }

function styleTitle(ws, title, subtitle, lastCol) {
  ws.mergeCells(1, 1, 1, lastCol)
  const t = ws.getCell(1, 1)
  t.value = title
  t.font = { name: FONT, size: 16, bold: true, color: { argb: 'FFFFFFFF' } }
  t.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
  t.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  ws.getRow(1).height = 30

  ws.mergeCells(2, 1, 2, lastCol)
  const s = ws.getCell(2, 1)
  s.value = subtitle
  s.font = { name: FONT, size: 10, italic: true, color: { argb: 'FF475569' } }
  s.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: LIGHT } }
  s.alignment = { vertical: 'middle', horizontal: 'left', indent: 1 }
  ws.getRow(2).height = 20
}

function styleHeaderCell(cell) {
  cell.font = { name: FONT, size: 11, bold: true, color: { argb: 'FFFFFFFF' } }
  cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: BLUE } }
  cell.alignment = { vertical: 'middle', horizontal: 'center', wrapText: true }
  cell.border = BORDER
}

function styleBodyRow(row, firstCol, lastCol, index, leftCols = [1], numFmts = {}) {
  for (let c = firstCol; c <= lastCol; c++) {
    const cell = row.getCell(c)
    cell.font = { name: FONT, size: 10 }
    cell.border = BORDER
    cell.alignment = {
      vertical: 'middle',
      horizontal: leftCols.includes(c) ? 'left' : 'center',
      wrapText: leftCols.includes(c)
    }
    if (index % 2 === 1) cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: ZEBRA } }
    if (numFmts[c]) cell.numFmt = numFmts[c]
  }
  row.height = 18
}

function styleTotalRow(row, firstCol, lastCol, leftCols = [1]) {
  for (let c = firstCol; c <= lastCol; c++) {
    const cell = row.getCell(c)
    cell.font = { name: FONT, size: 11, bold: true, color: { argb: NAVY } }
    cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: TOTAL_BG } }
    cell.border = { top: medium, left: thin, bottom: medium, right: thin }
    cell.alignment = { vertical: 'middle', horizontal: leftCols.includes(c) ? 'left' : 'center' }
  }
  row.height = 22
}

function setupPage(ws, lastCol, headerRows, landscape = true) {
  ws.pageSetup = {
    orientation: landscape ? 'landscape' : 'portrait',
    paperSize: 9,
    fitToPage: true,
    fitToWidth: 1,
    fitToHeight: 0,
    margins: { left: 0.4, right: 0.4, top: 0.6, bottom: 0.6, header: 0.3, footer: 0.3 }
  }
  ws.headerFooter = {
    oddFooter: '&LDisposal Report&CPage &P of &N&R&D'
  }
  ws.views = [{ state: 'frozen', xSplit: 1, ySplit: headerRows, showGridLines: false }]
  ws.pageSetup.printTitlesRow = `1:${headerRows}`
}

function colLetter(n) {
  let s = ''
  while (n > 0) {
    const m = (n - 1) % 26
    s = String.fromCharCode(65 + m) + s
    n = Math.floor((n - 1) / 26)
  }
  return s
}

// Generic "judge x columns + total column + grand total row" sheet
function addSimpleReport(wb, name, title, subtitle, columns, rows, widths) {
  const ws = wb.addWorksheet(name, { properties: { tabColor: { argb: BLUE } } })
  const lastCol = columns.length
  styleTitle(ws, title, subtitle, lastCol)

  const headerRow = ws.getRow(4)
  columns.forEach((label, i) => {
    headerRow.getCell(i + 1).value = label
    styleHeaderCell(headerRow.getCell(i + 1))
  })
  headerRow.height = 30
  ws.getRow(3).height = 6

  const dataStart = 5
  rows.forEach((values, i) => {
    const row = ws.getRow(dataStart + i)
    values.forEach((v, c) => { row.getCell(c + 1).value = v })
    styleBodyRow(row, 1, lastCol, i)
    for (let c = 2; c <= lastCol; c++) row.getCell(c).numFmt = '#,##0'
  })

  // Grand total row using live SUM formulas (with cached result)
  const totalRowNum = dataStart + rows.length
  const totalRow = ws.getRow(totalRowNum)
  totalRow.getCell(1).value = 'GRAND TOTAL'
  for (let c = 2; c <= lastCol; c++) {
    const L = colLetter(c)
    const result = rows.reduce((s, r) => s + (Number(r[c - 1]) || 0), 0)
    totalRow.getCell(c).value = rows.length
      ? { formula: `SUM(${L}${dataStart}:${L}${totalRowNum - 1})`, result }
      : 0
    totalRow.getCell(c).numFmt = '#,##0'
  }
  styleTotalRow(totalRow, 1, lastCol)

  ws.columns = widths.map(w => ({ width: w }))
  setupPage(ws, lastCol, 4)
  ws.autoFilter = undefined
  return ws
}

// Judge x source PDF → From Date / To Date (earliest / latest disposal date), no totals
function addDateRangeSheet(wb, name, title, subtitle, files, rows) {
  const ws = wb.addWorksheet(name, { properties: { tabColor: { argb: 'FF7C3AED' } } })
  const lastCol = 1 + files.length * 2
  styleTitle(ws, title, subtitle, lastCol)
  ws.getRow(3).height = 6

  ws.mergeCells(4, 1, 5, 1)
  ws.getCell(4, 1).value = 'Judge Name'
  styleHeaderCell(ws.getCell(4, 1))
  styleHeaderCell(ws.getCell(5, 1))
  files.forEach((file, i) => {
    const start = 2 + i * 2
    ws.mergeCells(4, start, 4, start + 1)
    ws.getCell(4, start).value = file
    for (let c = start; c <= start + 1; c++) styleHeaderCell(ws.getCell(4, c))
    ;['From Date', 'To Date'].forEach((label, k) => {
      const cell = ws.getCell(5, start + k)
      cell.value = label
      styleHeaderCell(cell)
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
    })
  })
  ws.getRow(4).height = 30
  ws.getRow(5).height = 20

  const toDate = s => {
    const m = String(s || '').match(/^(\d{2})-(\d{2})-(\d{4})$/)
    return m ? new Date(Date.UTC(+m[3], +m[2] - 1, +m[1])) : null
  }
  rows.forEach(({ judge, ranges }, i) => {
    const row = ws.getRow(6 + i)
    row.getCell(1).value = judge
    files.forEach((file, fi) => {
      row.getCell(2 + fi * 2).value = toDate(ranges[file]?.from)
      row.getCell(3 + fi * 2).value = toDate(ranges[file]?.to)
    })
    styleBodyRow(row, 1, lastCol, i)
    for (let c = 2; c <= lastCol; c++) row.getCell(c).numFmt = 'dd-mm-yyyy'
  })

  ws.columns = [{ width: 28 }, ...Array(lastCol - 1).fill({ width: 13 })]
  setupPage(ws, lastCol, 5)
  return ws
}

// Two-level header sheet (Judge x age category x Contested/Uncontested/Total)
function addAgeNatureSheet(wb, name, title, subtitle, groups, rows, tabColor) {
  const ws = wb.addWorksheet(name.slice(0, 31), { properties: { tabColor: { argb: tabColor } } })
  const lastCol = 1 + groups.length * 3
  styleTitle(ws, title, subtitle, lastCol)
  ws.getRow(3).height = 6

  // Row 4 group header, row 5 sub header
  ws.mergeCells(4, 1, 5, 1)
  ws.getCell(4, 1).value = 'Judge Name'
  styleHeaderCell(ws.getCell(4, 1))
  styleHeaderCell(ws.getCell(5, 1))

  groups.forEach((group, i) => {
    const start = 2 + i * 3
    ws.mergeCells(4, start, 4, start + 2)
    const g = ws.getCell(4, start)
    g.value = group
    for (let c = start; c <= start + 2; c++) styleHeaderCell(ws.getCell(4, c))
    ;['Contested', 'Uncontested', 'TOTAL'].forEach((label, k) => {
      const cell = ws.getCell(5, start + k)
      cell.value = label
      styleHeaderCell(cell)
      cell.fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: NAVY } }
    })
    // Slightly different shade for the TOTAL group
    if (group === 'TOTAL') {
      for (let c = start; c <= start + 2; c++) {
        ws.getCell(4, c).fill = { type: 'pattern', pattern: 'solid', fgColor: { argb: 'FF0F766E' } }
      }
    }
  })
  ws.getRow(4).height = 22
  ws.getRow(5).height = 22

  const dataStart = 6
  const body = rows.filter(r => r.judge !== 'GRAND TOTAL')
  body.forEach(({ judge, counts }, i) => {
    const row = ws.getRow(dataStart + i)
    row.getCell(1).value = judge
    groups.forEach((group, gi) => {
      const start = 2 + gi * 3
      row.getCell(start).value = counts[group].contested
      row.getCell(start + 1).value = counts[group].uncontested
      row.getCell(start + 2).value = counts[group].total
    })
    styleBodyRow(row, 1, lastCol, i)
    for (let c = 2; c <= lastCol; c++) row.getCell(c).numFmt = '#,##0'
    // Emphasise total columns of each group
    for (let gi = 0; gi < groups.length; gi++) {
      row.getCell(4 + gi * 3).font = { name: FONT, size: 10, bold: true }
    }
  })

  const totalRowNum = dataStart + body.length
  const totalRow = ws.getRow(totalRowNum)
  totalRow.getCell(1).value = 'GRAND TOTAL'
  const grand = rows.find(r => r.judge === 'GRAND TOTAL')
  groups.forEach((group, gi) => {
    ;['contested', 'uncontested', 'total'].forEach((key, k) => {
      const c = 2 + gi * 3 + k
      const L = colLetter(c)
      totalRow.getCell(c).value = body.length
        ? { formula: `SUM(${L}${dataStart}:${L}${totalRowNum - 1})`, result: grand ? grand.counts[group][key] : 0 }
        : 0
      totalRow.getCell(c).numFmt = '#,##0'
    })
  })
  styleTotalRow(totalRow, 1, lastCol)

  ws.columns = [{ width: 28 }, ...Array(lastCol - 1).fill({ width: 13 })]
  setupPage(ws, lastCol, 5)
  return ws
}

// ---------- Main entry ----------
export async function buildWorkbook(records, helpers) {
  const { ageInDays, ageInYears, ageCategory, AGE_CATEGORIES, buildWebReports } = helpers
  // Same numbers (and row order) as the web dashboard
  const reports = buildWebReports(records)
  const withoutGrandTotal = rows => rows.filter(r => r.judge !== 'GRAND TOTAL')
  const wb = new ExcelJS.Workbook()
  wb.creator = 'PDF TO EXCEL PRO'
  wb.created = new Date()

  const generated = `Generated on ${new Date().toLocaleString('en-IN')}  |  Total records: ${records.length.toLocaleString('en-IN')}`

  // ----- 1. Disposal Register -----
  const headers = [
    'S.No.', 'Source File Name', 'Court Name', 'Court Of', 'Judge (Header)', 'Disposal Nature',
    'Case Type', 'SIDE', 'Reg. No./Year', 'CASE NO', 'Date of Filing', 'Disposal Date',
    'AGE(D)', 'AGE(Y)', 'AGE(CAT)', 'Disposal Type', 'Duration (Y)', 'Duration (M)',
    'Duration (D)', 'Judge Name'
  ]
  const widths = [7, 24, 24, 48, 14, 13, 10, 11, 13, 18, 12, 12, 9, 9, 15, 36, 9, 9, 9, 16]
  const leftCols = [2, 3, 4, 5, 16, 20]

  const ws = wb.addWorksheet('Disposal Register', { properties: { tabColor: { argb: NAVY } } })
  const lastCol = headers.length
  styleTitle(ws, 'DISPOSAL REGISTER', generated, lastCol)
  ws.getRow(3).height = 6
  const hr = ws.getRow(4)
  headers.forEach((h, i) => { hr.getCell(i + 1).value = h; styleHeaderCell(hr.getCell(i + 1)) })
  hr.height = 32

  const dataStart = 5
  const toDate = (s) => {
    const m = String(s || '').match(/^(\d{2})-(\d{2})-(\d{4})$/)
    return m ? new Date(Date.UTC(+m[3], +m[2] - 1, +m[1])) : s
  }
  records.forEach((r, i) => {
    const row = ws.getRow(dataStart + i)
    const days = ageInDays(r.filing_date, r.disposal_date)
    const yrs = ageInYears(r.filing_date, r.disposal_date)
    const vals = [
      r.s_no, r.filename, r.court_name, r.court_of, r.judge_in_header, r.disposal_nature,
      r.case_type, r.side, r.reg_no, r.case_no, toDate(r.filing_date), toDate(r.disposal_date),
      days, yrs, ageCategory(r.filing_date, r.disposal_date),
      r.disposal_type, r.duration_y, r.duration_m, r.duration_d, r.judge_name
    ]
    vals.forEach((v, c) => { row.getCell(c + 1).value = v })
    styleBodyRow(row, 1, lastCol, i, leftCols, { 11: 'dd-mm-yyyy', 12: 'dd-mm-yyyy', 13: '#,##0', 14: '0.00' })
    // Single-line rows: wrapped text would be clipped by the fixed row height
    leftCols.forEach(c => { row.getCell(c).alignment = { vertical: 'middle', horizontal: 'left' } })
  })

  // Total row: case count + average age (a sum of ages is meaningless)
  const totalRowNum = dataStart + records.length
  const tr = ws.getRow(totalRowNum)
  const last = totalRowNum - 1
  const ages = records.map(r => ageInDays(r.filing_date, r.disposal_date)).filter(v => v !== '')
  const avgDays = ages.length ? ages.reduce((s, v) => s + v, 0) / ages.length : 0
  tr.getCell(1).value = 'TOTAL'
  tr.getCell(2).value = { formula: `COUNTA(J${dataStart}:J${last})&" cases"`, result: `${records.length} cases` }
  tr.getCell(12).value = 'Average age'
  tr.getCell(13).value = { formula: `IFERROR(AVERAGE(M${dataStart}:M${last}),0)`, result: avgDays }
  tr.getCell(14).value = { formula: `IFERROR(AVERAGE(N${dataStart}:N${last}),0)`, result: Math.round((avgDays / 365.2425) * 100) / 100 }
  styleTotalRow(tr, 1, lastCol, [1, 2])
  tr.getCell(12).alignment = { vertical: 'middle', horizontal: 'right' }
  tr.getCell(13).numFmt = '#,##0'
  tr.getCell(14).numFmt = '0.00'

  ws.columns = widths.map(w => ({ width: w }))
  ws.autoFilter = { from: { row: 4, column: 1 }, to: { row: totalRowNum - 1, column: lastCol } }
  setupPage(ws, lastCol, 4)
  // 20 columns on one A4 page are unreadable: print 2 pages wide and repeat S.No. + file on each
  ws.pageSetup.fitToWidth = 2
  ws.pageSetup.printTitlesColumn = 'A:B'
  ws.pageSetup.pageOrder = 'overThenDown' // print both halves of each row block together
  ws.views = [{ state: 'frozen', xSplit: 2, ySplit: 4, showGridLines: false }]

  // ----- 2. Judge Summary -----
  addSimpleReport(wb, 'Judge Summary', 'JUDGE SUMMARY', generated,
    ['Judge Name', 'CIVIL', 'CRIMINAL', 'TOTAL DISPOSALS'],
    withoutGrandTotal(reports.judgeSummary).map(r => [r.judge, r.civil, r.criminal, r.total]), [30, 14, 14, 20])

  // ----- 3. Judge Age Wise -----
  addSimpleReport(wb, 'Judge Age Wise', 'JUDGE AGE WISE DISPOSALS', generated,
    ['Judge Name', ...AGE_CATEGORIES, 'Total'],
    withoutGrandTotal(reports.ageWise).map(r => [r.judge, ...AGE_CATEGORIES.map(k => r[k]), r.total]),
    [30, 12, 12, 12, 12, 18, 12])

  // ----- 4-6. Age category nature -----
  const groups = [...AGE_CATEGORIES, 'TOTAL']
  addAgeNatureSheet(wb, 'Age Category Nature', 'AGE CATEGORY NATURE – ALL CASES', generated,
    groups, reports.ageNature, 'FF2563EB')
  addAgeNatureSheet(wb, 'AGE CATEGORY NATURE(CIVIL)', 'AGE CATEGORY NATURE – CIVIL', generated,
    groups, reports.ageNatureCivil, 'FF16A34A')
  addAgeNatureSheet(wb, 'AGE CATEGORY NATURE(CRIMINAL)', 'AGE CATEGORY NATURE – CRIMINAL', generated,
    groups, reports.ageNatureCriminal, 'FFDC2626')

  // ----- 7. Judge Source Pivot -----
  addSimpleReport(wb, 'Judge Source Pivot', 'JUDGE vs SOURCE PDF', generated,
    ['Judge Name', ...reports.sourceFiles, 'Grand Total'],
    withoutGrandTotal(reports.sourcePivot).map(r => [r.judge, ...reports.sourceFiles.map(f => r[f]), r.total]),
    [30, ...reports.sourceFiles.map(() => 26), 14])

  // ----- 8. Judge Source Date Range -----
  addDateRangeSheet(wb, 'Judge Source Date Range', 'JUDGE vs SOURCE PDF – DISPOSAL DATE RANGE', generated,
    reports.sourceFiles, reports.sourceDateRange)

  return wb
}

export async function downloadStyledExcel(records, originalName, helpers) {
  if (!records.length) return
  const wb = await buildWorkbook(records, helpers)
  const buffer = await wb.xlsx.writeBuffer()
  const blob = new Blob([buffer], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
  })
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = originalName.replace(/\.pdf$/i, '') + '_Disposal_Report.xlsx'
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1000)
}
