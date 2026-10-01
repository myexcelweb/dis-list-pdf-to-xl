import { useState, useRef, useCallback, useMemo, useEffect } from 'react'
import { downloadStyledExcel } from './excelExport.js'
import sideClassificationCsv from './side_classification_mappings.csv?raw'
import {
  parseSideCsv, extractTextFromPdfDocument, parseDisposalRegister, ageInDays, ageInYears,
  ageCategory, AGE_CATEGORIES, buildWebReports, removeDuplicateEntries
} from './parser.js'
import './App.css'

const SIDE_CLASSIFICATION = parseSideCsv(sideClassificationCsv)

// ---------- PDF text extraction using pdf.js ----------
async function extractTextFromPdf(file) {
  const pdfjs = await import('pdfjs-dist')
  // Use the worker from the same package (Vite handles it)
  pdfjs.GlobalWorkerOptions.workerSrc = new URL(
    'pdfjs-dist/build/pdf.worker.min.mjs',
    import.meta.url
  ).toString()

  const arrayBuffer = await file.arrayBuffer()
  const pdf = await pdfjs.getDocument({ data: arrayBuffer }).promise
  return extractTextFromPdfDocument(pdf)
}

// ---------- Excel generation (styled, see excelExport.js) ----------
function downloadExcel(records, originalName) {
  return downloadStyledExcel(records, originalName, {
    ageInDays, ageInYears, ageCategory, AGE_CATEGORIES, buildWebReports
  })
}

// SHA-256 of the file bytes, used to skip the same PDF uploaded twice (even under another name)
async function fileFingerprint(file) {
  // crypto.subtle only exists on https / localhost; fall back to size + name elsewhere
  if (!globalThis.crypto?.subtle) return `${file.size}:${file.name}`
  const digest = await crypto.subtle.digest('SHA-256', await file.arrayBuffer())
  return Array.from(new Uint8Array(digest), b => b.toString(16).padStart(2, '0')).join('')
}

// ---------- React Component ----------
export default function App() {
  const [status, setStatus] = useState('idle') // idle | parsing | done | error
  const [message, setMessage] = useState('')
  const [notice, setNotice] = useState('') // non-fatal warnings (duplicates, unreadable files)
  const [records, setRecords] = useState([])
  const [fileName, setFileName] = useState('')
  const [dragOver, setDragOver] = useState(false)
  const [activePage, setActivePage] = useState('upload')
  const [activeView, setActiveView] = useState('register')
  const [searchText, setSearchText] = useState('')
  const [judgeFilter, setJudgeFilter] = useState('')
  const [sourceFilter, setSourceFilter] = useState('')
  const [ageFilter, setAgeFilter] = useState('')
  const [natureFilter, setNatureFilter] = useState('')
  const [sideFilter, setSideFilter] = useState('')
  const [page, setPage] = useState(1)
  const [pageSize, setPageSize] = useState(50)
  const tableWrapRef = useRef(null)
  const [fullscreen, setFullscreen] = useState(false)
  const panelRef = useRef(null)
  const inputRef = useRef(null)

  const webReports = useMemo(() => buildWebReports(records), [records])
  const judgeOptions = useMemo(() => [...new Set(records.map(r => r.judge_name || 'Unknown'))].sort((a, b) => a.localeCompare(b)), [records])
  const sourceOptions = useMemo(() => [...new Set(records.map(r => r.filename || 'Unknown'))].sort((a, b) => a.localeCompare(b)), [records])
  const natureOptions = useMemo(() => [...new Set(records.map(r => r.disposal_nature || 'Unknown'))].sort((a, b) => a.localeCompare(b)), [records])
  const sideOptions = useMemo(() => [...new Set(records.map(r => r.side || 'UNMAPPED'))].sort((a, b) => a.localeCompare(b)), [records])
  const filteredRecords = useMemo(() => {
    const query = searchText.trim().toLowerCase()
    return records.filter(r => {
      if (judgeFilter && (r.judge_name || 'Unknown') !== judgeFilter) return false
      if (sourceFilter && (r.filename || 'Unknown') !== sourceFilter) return false
      if (ageFilter && ageCategory(r.filing_date, r.disposal_date) !== ageFilter) return false
      if (natureFilter && (r.disposal_nature || 'Unknown') !== natureFilter) return false
      if (sideFilter && (r.side || 'UNMAPPED') !== sideFilter) return false
      if (query && ![
        r.s_no, r.filename, r.court_name, r.court_of, r.judge_in_header, r.disposal_nature,
        r.case_type, r.side, r.reg_no, r.case_no, r.filing_date, r.disposal_date, r.disposal_type,
        ageInDays(r.filing_date, r.disposal_date), ageInYears(r.filing_date, r.disposal_date),
        ageCategory(r.filing_date, r.disposal_date), r.duration_y, r.duration_m,
        r.duration_d, r.judge_name
      ].some(value => String(value ?? '').toLowerCase().includes(query))) return false
      return true
    })
  }, [records, searchText, judgeFilter, sourceFilter, ageFilter, natureFilter, sideFilter])

  const processFiles = useCallback(async (fileList) => {
    const files = Array.from(fileList || [])
    if (!files.length) return
    if (files.some(file => !file.name.toLowerCase().endsWith('.pdf'))) {
      setStatus('error')
      setMessage('Please select PDF files only.')
      return
    }

    setStatus('parsing')
    setMessage(`Extracting text from ${files.length} PDF${files.length === 1 ? '' : 's'}…`)
    setRecords([])
    setNotice('')
    setFileName(files.length === 1 ? files[0].name : `${files.length} PDF files`)
    setActiveView('register')
    setSearchText('')
    setJudgeFilter('')
    setSourceFilter('')
    setAgeFilter('')
    setNatureFilter('')
    setSideFilter('')
    setPage(1)

    try {
      const allRecords = []
      const failedFiles = []
      const duplicateFiles = []
      const seen = new Map() // fingerprint → first file name
      for (const file of files) {
        setMessage(`Extracting and parsing ${file.name}…`)
        try {
          const fingerprint = await fileFingerprint(file)
          if (seen.has(fingerprint)) {
            duplicateFiles.push(`${file.name} (same as ${seen.get(fingerprint)})`)
            continue
          }
          seen.set(fingerprint, file.name)
          const text = await extractTextFromPdf(file)
          const parsed = parseDisposalRegister(text, file.name, SIDE_CLASSIFICATION)
          if (!parsed.length) failedFiles.push(file.name)
          allRecords.push(...parsed)
        } catch (err) {
          console.error(`Failed to process ${file.name}`, err)
          failedFiles.push(file.name)
        }
      }

      if (allRecords.length === 0) {
        setStatus('error')
        setMessage(`No records could be extracted${failedFiles.length ? ` from: ${failedFiles.join(', ')}` : ''}. The PDFs may use a different Court Disposal Register format.`)
        return
      }

      // Same CASE NO + Judge Name → keep the first entry only
      const { unique, duplicates } = removeDuplicateEntries(allRecords)
      if (duplicates.length) console.info('Skipped duplicate entries (same CASE NO + Judge Name):', duplicates)
      const examples = duplicates.slice(0, 3).map(r => `${r.case_no} – ${r.judge_name} (${r.filename} S.No. ${r.s_no})`).join('; ')

      setRecords(unique)
      setStatus('done')
      const parsedCount = files.length - failedFiles.length - duplicateFiles.length
      setMessage(`Successfully extracted ${unique.length} records from ${parsedCount} of ${files.length} PDF${files.length === 1 ? '' : 's'}.`)
      setNotice([
        duplicateFiles.length ? `Skipped duplicate file${duplicateFiles.length === 1 ? '' : 's'}: ${duplicateFiles.join(', ')}.` : '',
        duplicates.length ? `Skipped ${duplicates.length} duplicate entr${duplicates.length === 1 ? 'y' : 'ies'} with the same CASE NO and Judge Name: ${examples}${duplicates.length > 3 ? ` and ${duplicates.length - 3} more` : ''}.` : '',
        failedFiles.length ? `No records found in: ${failedFiles.join(', ')}.` : ''
      ].filter(Boolean).join(' '))
      setActivePage('dashboard')
    } catch (err) {
      console.error(err)
      setStatus('error')
      setMessage('Failed to process PDFs: ' + (err.message || 'Unknown error'))
    }
  }, [])

  const onDrop = (e) => {
    e.preventDefault()
    setDragOver(false)
    if (e.dataTransfer.files?.length) processFiles(e.dataTransfer.files)
  }

  const onFileChange = (e) => {
    if (e.target.files?.length) processFiles(e.target.files)
    e.target.value = ''
  }

  const clearFilesAndReports = () => {
    setStatus('idle')
    setMessage('')
    setNotice('')
    setRecords([])
    setFileName('')
    setActiveView('register')
    setActivePage('upload')
    setSearchText('')
    setJudgeFilter('')
    setSourceFilter('')
    setAgeFilter('')
    setNatureFilter('')
    setSideFilter('')
    setPage(1)
    if (inputRef.current) inputRef.current.value = ''
  }

  const fmt = (v) => (typeof v === 'number' ? v.toLocaleString('en-IN') : (v ?? 0))

  const renderReportTable = (columns, rows) => (
    <div className="table-wrap report-table-wrap">
      <table className="report-table">
        <thead>
          <tr>{columns.map((column, i) => (
            <th key={column.key} className={i === 0 ? 'col-name' : 'col-num'} title={column.label}>{column.label}</th>
          ))}</tr>
        </thead>
        <tbody>
          {rows.map((row, index) => (
            <tr key={`${row.judge}-${index}`} className={row.judge === 'GRAND TOTAL' ? 'total-row' : ''}>
              {columns.map((column, i) => (
                <td key={column.key} className={i === 0 ? 'col-name' : 'col-num'}>{fmt(row[column.key])}</td>
              ))}
            </tr>
          ))}
        </tbody>
      </table>
    </div>
  )

  // Judge × source PDF → earliest / latest disposal date (no totals)
  const renderDateRangeTable = (files, rows) => (
    <div className="table-wrap report-table-wrap">
      <table className="report-table grouped-report-table date-range-table">
        <colgroup>
          <col className="col-judge" />
          {files.flatMap(file => [<col key={`${file}-from`} />, <col key={`${file}-to`} />])}
        </colgroup>
        <thead>
          <tr>
            <th rowSpan="2" className="col-name">Judge Name</th>
            {files.map(file => <th key={file} colSpan="2" className="group-th" title={file}>{file}</th>)}
          </tr>
          <tr>
            {files.flatMap(file => [
              <th key={`${file}-from`} className="sub-th col-num group-start">From Date</th>,
              <th key={`${file}-to`} className="sub-th col-num">To Date</th>
            ])}
          </tr>
        </thead>
        <tbody>{rows.map(row => (
          <tr key={row.judge}>
            <td className="col-name">{row.judge}</td>
            {files.flatMap(file => [
              <td key={`${file}-from`} className="col-num group-start">{row.ranges[file]?.from || ''}</td>,
              <td key={`${file}-to`} className="col-num">{row.ranges[file]?.to || ''}</td>
            ])}
          </tr>
        ))}</tbody>
      </table>
    </div>
  )

  const GROUPS = [...AGE_CATEGORIES, 'TOTAL']
  const renderAgeNatureTable = (rows) => (
    <div className="table-wrap report-table-wrap">
      <table className="report-table grouped-report-table">
        <colgroup>
          <col className="col-judge" />
          {GROUPS.flatMap(group => [0, 1, 2].map(k => <col key={`${group}-${k}`} />))}
        </colgroup>
        <thead>
          <tr>
            <th rowSpan="2" className="col-name">Judge Name</th>
            {GROUPS.map(group => (
              <th key={group} colSpan="3" className={`group-th ${group === 'TOTAL' ? 'group-total' : ''}`}>{group}</th>
            ))}
          </tr>
          <tr>
            {/* Soft hyphens (­) let long labels wrap cleanly in narrow columns */}
            {GROUPS.flatMap(group => [['Contested', 'Con­tested'], ['Uncontested', 'Uncon­tested'], ['TOTAL', 'TOTAL']].map(([key, label], k) => (
              <th key={`${group}-${key}`} className={`sub-th col-num ${k === 0 ? 'group-start' : ''}`} title={`${group} – ${key}`}>{label}</th>
            )))}
          </tr>
        </thead>
        <tbody>{rows.map((row, index) => (
          <tr key={`${row.judge}-${index}`} className={row.judge === 'GRAND TOTAL' ? 'total-row' : ''}>
            <td className="col-name">{row.judge}</td>
            {GROUPS.flatMap(group => [
              <td key={`${group}-contested`} className="col-num group-start">{fmt(row.counts[group].contested)}</td>,
              <td key={`${group}-uncontested`} className="col-num">{fmt(row.counts[group].uncontested)}</td>,
              <td key={`${group}-total`} className="col-num sum-col">{fmt(row.counts[group].total)}</td>
            ])}
          </tr>
        ))}</tbody>
      </table>
    </div>
  )

  const REGISTER_COLUMNS = [
    ['S.No.', 'num'], ['Source File Name', 'text'], ['Court Name', 'text'], ['Court Of', 'text'], ['Judge (Header)', 'text'],
    ['Disposal Nature', 'text'], ['Case Type', 'text'], ['SIDE', 'badge'], ['Reg. No./Year', 'num'],
    ['CASE NO', 'text'], ['Date of Filing', 'num'], ['Disposal Date', 'num'], ['AGE(D)', 'num'],
    ['AGE(Y)', 'num'], ['AGE(CAT)', 'text'], ['Disposal Type', 'text'], ['Duration (Y)', 'num'],
    ['Duration (M)', 'num'], ['Duration (D)', 'num'], ['Judge Name', 'text']
  ]

  const VIEWS = [
    ['register', 'Disposal Register', '📋'],
    ['judge', 'Judge Summary', '👨‍⚖️'],
    ['age', 'Judge Age Wise', '⏳'],
    ['nature', 'Age Category Nature', '📊'],
    ['natureCivil', 'Age Category Nature (Civil)', '🏛️'],
    ['natureCriminal', 'Age Category Nature (Criminal)', '🔒'],
    ['source', 'Judge Source Pivot', '📁'],
    ['sourceDates', 'Judge Source Date Range', '📅']
  ]
  const activeViewLabel = VIEWS.find(v => v[0] === activeView)?.[1] || ''

  const stats = useMemo(() => {
    const civil = records.filter(r => r.side === 'CIVIL').length
    const criminal = records.filter(r => r.side === 'CRIMINAL').length
    const judges = new Set(records.map(r => r.judge_name || 'Unknown')).size
    const files = new Set(records.map(r => r.filename || 'Unknown')).size
    return { civil, criminal, judges, files }
  }, [records])

  const resetFilters = () => { setSearchText(''); setJudgeFilter(''); setSourceFilter(''); setAgeFilter(''); setNatureFilter(''); setSideFilter('') }
  const hasFilters = searchText || judgeFilter || sourceFilter || ageFilter || natureFilter || sideFilter

  const renderCell = (r, label) => {
    switch (label) {
      case 'S.No.': return r.s_no
      case 'Source File Name': return r.filename
      case 'Court Name': return r.court_name
      case 'Court Of': return r.court_of
      case 'Judge (Header)': return r.judge_in_header
      case 'Disposal Nature': return <span className={`pill ${/^un/i.test(r.disposal_nature) ? 'pill-amber' : 'pill-blue'}`}>{r.disposal_nature}</span>
      case 'Case Type': return r.case_type
      case 'SIDE': return <span className={`pill ${r.side === 'CIVIL' ? 'pill-green' : r.side === 'CRIMINAL' ? 'pill-red' : 'pill-gray'}`}>{r.side}</span>
      case 'Reg. No./Year': return r.reg_no
      case 'CASE NO': return r.case_no
      case 'Date of Filing': return r.filing_date
      case 'Disposal Date': return r.disposal_date
      case 'AGE(D)': return ageInDays(r.filing_date, r.disposal_date)
      case 'AGE(Y)': { const y = ageInYears(r.filing_date, r.disposal_date); return y === '' ? '' : y.toFixed(2) }
      case 'AGE(CAT)': return ageCategory(r.filing_date, r.disposal_date)
      case 'Disposal Type': return r.disposal_type
      case 'Duration (Y)': return r.duration_y
      case 'Duration (M)': return r.duration_m
      case 'Duration (D)': return r.duration_d
      case 'Judge Name': return r.judge_name
      default: return ''
    }
  }

  // pageSize 0 = show all rows
  const effectivePageSize = pageSize || Math.max(1, filteredRecords.length)
  const totalPages = Math.max(1, Math.ceil(filteredRecords.length / effectivePageSize))
  const safePage = Math.min(page, totalPages)
  const pageStart = (safePage - 1) * effectivePageSize
  const pageRecords = filteredRecords.slice(pageStart, pageStart + effectivePageSize)
  const pageEnd = pageStart + pageRecords.length
  const goToPage = p => {
    setPage(Math.min(Math.max(1, p), totalPages))
    if (tableWrapRef.current) tableWrapRef.current.scrollTop = 0
  }
  // Compact list of page numbers around the current page, e.g. 1 … 4 5 [6] 7 8 … 20
  const pageNumbers = (() => {
    const pages = new Set([1, totalPages])
    for (let p = safePage - 2; p <= safePage + 2; p++) if (p >= 1 && p <= totalPages) pages.add(p)
    const sorted = [...pages].sort((a, b) => a - b)
    const out = []
    sorted.forEach((p, i) => { if (i > 0 && p - sorted[i - 1] > 1) out.push('…' + p); out.push(p) })
    return out
  })()

  // Full screen mode for the report panel (with Exit). Uses the browser Fullscreen API,
  // with a CSS fixed-overlay fallback. Esc also exits.
  const enterFullscreen = () => {
    setFullscreen(true)
    const el = panelRef.current
    if (el && el.requestFullscreen) el.requestFullscreen().catch(() => {})
  }
  const exitFullscreen = () => {
    setFullscreen(false)
    if (document.fullscreenElement && document.exitFullscreen) document.exitFullscreen().catch(() => {})
  }
  useEffect(() => {
    const onFs = () => { if (!document.fullscreenElement) setFullscreen(false) }
    const onKey = e => { if (e.key === 'Escape') setFullscreen(false) }
    document.addEventListener('fullscreenchange', onFs)
    document.addEventListener('keydown', onKey)
    return () => { document.removeEventListener('fullscreenchange', onFs); document.removeEventListener('keydown', onKey) }
  }, [])

  const isDash = activePage === 'dashboard' && status === 'done' && records.length > 0
  return (
    <div className={`app ${isDash ? 'app-dash' : ''}`}>
      <header className={`header ${isDash ? 'header-compact' : ''}`}>
        <div className="container header-inner">
          <div className="logo">
            <span className="logo-icon" aria-hidden="true">⚖️</span>
            <div>
              <h1>PDF TO EXCEL PRO</h1>
              <p className="tagline">Court Disposal Register → Professional Excel Reports</p>
            </div>
          </div>
          <nav className="page-tabs" aria-label="App pages">
            <button className={`page-tab ${activePage === 'upload' ? 'active' : ''}`} onClick={() => setActivePage('upload')} aria-pressed={activePage === 'upload'}>
              Upload
            </button>
            <button className={`page-tab ${activePage === 'dashboard' ? 'active' : ''}`} onClick={() => records.length > 0 && setActivePage('dashboard')} disabled={records.length === 0} aria-pressed={activePage === 'dashboard'}>
              Dashboard
            </button>
            <button className={`page-tab ${activePage === 'about' ? 'active' : ''}`} onClick={() => setActivePage('about')} aria-pressed={activePage === 'about'}>
              About
            </button>
          </nav>
        </div>
      </header>

      <main className={`main ${isDash ? 'main-dash' : ''}`}>
        <div className={`container ${activePage === 'dashboard' ? 'container-wide' : ''}`}>

          {activePage === 'about' && (
            <section className="about-card">
              <div className="about-icon" aria-hidden="true">⚖️</div>
              <h2>PDF TO EXCEL PRO</h2>
              <p className="about-tagline">Court Disposal Register → Professional Excel Reports</p>
              <p className="about-text">
                Converts Court Disposal Register PDFs into a formatted, multi-sheet Excel workbook with
                judge-wise, age-wise and source-wise reports. Everything runs in your browser — no data leaves your computer.
              </p>
              <div className="about-dev">
                <span className="about-label">Designed &amp; Developed by</span>
                <strong className="about-name">Parimal Hodar</strong>
                <a className="about-email" href="mailto:parimalhodar.dev@gmail.com">✉ parimalhodar.dev@gmail.com</a>
              </div>
            </section>
          )}

          {activePage === 'upload' && <>
            <section className="hero">
              <h2>Convert Disposal Registers to Excel</h2>
              <p>Upload one or more Court Disposal Register PDFs and get a fully formatted, multi-sheet Excel workbook.</p>
            </section>

            <section
              className={`upload-zone ${dragOver ? 'drag-over' : ''} ${status === 'parsing' ? 'busy' : ''}`}
              onDragOver={(e) => { e.preventDefault(); setDragOver(true) }}
              onDragLeave={() => setDragOver(false)}
              onDrop={onDrop}
              onClick={() => status !== 'parsing' && inputRef.current?.click()}
              role="button"
              tabIndex={0}
              onKeyDown={(e) => { if ((e.key === 'Enter' || e.key === ' ') && status !== 'parsing') inputRef.current?.click() }}
            >
              <input ref={inputRef} type="file" accept=".pdf,application/pdf" multiple onChange={onFileChange} hidden />
              {status === 'parsing' ? (
                <div className="spinner-wrap">
                  <div className="spinner" />
                  <p>{message}</p>
                </div>
              ) : (
                <>
                  <div className="upload-icon" aria-hidden="true">📄</div>
                  <h2>Drop your PDF files here</h2>
                  <p>or click to browse — multiple files supported</p>
                  <span className="btn-primary upload-btn">Choose PDF files</span>
                  <p className="hint">🔒 100% in your browser — no data ever leaves your computer</p>
                </>
              )}
            </section>

            <div className="upload-actions">
              <button className="btn-secondary" onClick={clearFilesAndReports} disabled={status === 'parsing'}>Clear files and reports</button>
            </div>

            {status === 'error' && <div className="alert error" role="alert">{message}</div>}

            <section className="info-card">
              <h3>How it works</h3>
              <div className="steps">
                <div className="step"><span className="step-no">1</span><div><strong>Upload</strong><p>Select Court Disposal Register PDFs (Taluka Court format).</p></div></div>
                <div className="step"><span className="step-no">2</span><div><strong>Review</strong><p>Browse 7 reports with search, filters and totals in the dashboard.</p></div></div>
                <div className="step"><span className="step-no">3</span><div><strong>Download</strong><p>Get a formatted Excel with headers, borders and total rows.</p></div></div>
              </div>
              <p className="note">
                SIDE is mapped from CASE TYPE using <code>side_classification_mappings.csv</code> (unlisted types show as <code>UNMAPPED</code>).
                CASE NO combines CASE TYPE and Reg. No./Year, e.g. <code>CC/3122/2025</code>.
              </p>
            </section>
          </>}

          {activePage === 'dashboard' && status === 'done' && records.length > 0 && (
            <>
              <section className="dash-bar">
                <div className="dash-bar-title">
                  <h2 className="dash-title">Reports Dashboard</h2>
                  <p className="file-name" title={`${fileName} · ${message}`}>{fileName} · {message}</p>
                </div>
                <div className="kpis" aria-label="Summary">
                  <div className="kpi"><span className="kpi-label">Total</span><span className="kpi-value">{records.length.toLocaleString('en-IN')}</span></div>
                  <div className="kpi kpi-green"><span className="kpi-label">Civil</span><span className="kpi-value">{stats.civil.toLocaleString('en-IN')}</span></div>
                  <div className="kpi kpi-red"><span className="kpi-label">Criminal</span><span className="kpi-value">{stats.criminal.toLocaleString('en-IN')}</span></div>
                  <div className="kpi kpi-purple"><span className="kpi-label">Judges</span><span className="kpi-value">{stats.judges}</span></div>
                  <div className="kpi kpi-amber"><span className="kpi-label">PDFs</span><span className="kpi-value">{stats.files}</span></div>
                </div>
                <div className="dash-actions">
                  <button className="btn-secondary btn-sm" onClick={clearFilesAndReports}>Clear</button>
                  <button className="btn-primary btn-sm" onClick={() => downloadExcel(records, fileName.endsWith(' PDF files') ? 'Combined' : fileName)}>
                    ⬇ Download Excel
                  </button>
                </div>
              </section>

              {notice && <div className="alert warning" role="status">⚠ {notice}<button className="alert-close" onClick={() => setNotice('')} aria-label="Dismiss">✕</button></div>}

              <div className="dashboard-layout">
                <section ref={panelRef} className={`dashboard-report-content ${fullscreen ? 'is-fullscreen' : ''}`}>
                  <div className="report-heading">
                    <nav className="view-tabs" aria-label="Report views">
                      {VIEWS.map(([view, label, icon]) => (
                        <button key={view} className={`view-tab ${activeView === view ? 'active' : ''}`} onClick={() => setActiveView(view)} aria-pressed={activeView === view}>
                          <span className="view-icon" aria-hidden="true">{icon}</span><span>{label}</span>
                        </button>
                      ))}
                    </nav>
                    <div className="report-heading-right">
                      {activeView === 'register' && <span className="record-count">{filteredRecords.length.toLocaleString('en-IN')} of {records.length.toLocaleString('en-IN')}</span>}
                      {fullscreen
                        ? <button className="btn-exit" onClick={exitFullscreen} title="Exit full screen (Esc)">✕ Exit full screen</button>
                        : <button className="btn-secondary btn-sm" onClick={enterFullscreen} title="View report in full screen">⛶ Full screen</button>}
                    </div>
                  </div>

                  {activeView === 'register' && (
                    <>
                      <div className="filter-toolbar">
                        <label className="filter-control search-control">
                          <span>Search</span>
                          <input type="search" value={searchText} onChange={e => { setSearchText(e.target.value); setPage(1) }} placeholder="Case number, date, judge…" />
                        </label>
                        <label className="filter-control">
                          <span>Judge</span>
                          <select value={judgeFilter} onChange={e => { setJudgeFilter(e.target.value); setPage(1) }}>
                            <option value="">All judges</option>
                            {judgeOptions.map(value => <option key={value} value={value}>{value}</option>)}
                          </select>
                        </label>
                        <label className="filter-control">
                          <span>Source PDF</span>
                          <select value={sourceFilter} onChange={e => { setSourceFilter(e.target.value); setPage(1) }}>
                            <option value="">All files</option>
                            {sourceOptions.map(value => <option key={value} value={value}>{value}</option>)}
                          </select>
                        </label>
                        <label className="filter-control">
                          <span>Age</span>
                          <select value={ageFilter} onChange={e => { setAgeFilter(e.target.value); setPage(1) }}>
                            <option value="">All ages</option>
                            {AGE_CATEGORIES.map(value => <option key={value} value={value}>{value}</option>)}
                          </select>
                        </label>
                        <label className="filter-control">
                          <span>Nature</span>
                          <select value={natureFilter} onChange={e => { setNatureFilter(e.target.value); setPage(1) }}>
                            <option value="">All natures</option>
                            {natureOptions.map(value => <option key={value} value={value}>{value}</option>)}
                          </select>
                        </label>
                        <label className="filter-control">
                          <span>Side</span>
                          <select value={sideFilter} onChange={e => { setSideFilter(e.target.value); setPage(1) }}>
                            <option value="">All sides</option>
                            {sideOptions.map(value => <option key={value} value={value}>{value}</option>)}
                          </select>
                        </label>
                        <button className="btn-secondary" disabled={!hasFilters} onClick={() => { resetFilters(); setPage(1) }}>Clear filters</button>
                      </div>

                      <div className="table-wrap full-data-table" ref={tableWrapRef}>
                        <table className="register-table">
                          <thead><tr>
                            {REGISTER_COLUMNS.map(([label, kind], i) => (
                              <th key={label} className={`${kind === 'num' ? 'col-num' : ''} ${i === 0 ? 'sticky-col' : ''}`}>{label}</th>
                            ))}
                          </tr></thead>
                          <tbody>
                            {pageRecords.map((r, idx) => (
                              <tr key={`${r.filename}-${r.s_no}-${idx}`}>
                                {REGISTER_COLUMNS.map(([label, kind], i) => (
                                  <td key={label} className={`${kind === 'num' ? 'col-num' : ''} ${kind === 'badge' ? 'col-center' : ''} ${label === 'Disposal Type' ? 'disposal-type' : ''} ${i === 0 ? 'sticky-col' : ''}`}>
                                    {renderCell(r, label)}
                                  </td>
                                ))}
                              </tr>
                            ))}
                          </tbody>
                        </table>
                        {filteredRecords.length === 0 && <p className="more-note">No records match these filters.</p>}
                      </div>

                      <div className="pager">
                        <span className="pager-info">
                          {filteredRecords.length === 0
                            ? 'No records'
                            : `Showing ${(pageStart + 1).toLocaleString('en-IN')}–${pageEnd.toLocaleString('en-IN')} of ${filteredRecords.length.toLocaleString('en-IN')}`}
                        </span>
                        <div className="pager-nav">
                          <button className="btn-secondary" disabled={safePage <= 1} onClick={() => goToPage(1)} title="First page">«</button>
                          <button className="btn-secondary" disabled={safePage <= 1} onClick={() => goToPage(safePage - 1)}>← Prev</button>
                          {pageNumbers.map(p => typeof p === 'string'
                            ? <span key={p} className="pager-gap">…</span>
                            : <button key={p} className={`btn-secondary pager-num ${p === safePage ? 'active' : ''}`} onClick={() => goToPage(p)} aria-current={p === safePage ? 'page' : undefined}>{p}</button>)}
                          <button className="btn-secondary" disabled={safePage >= totalPages} onClick={() => goToPage(safePage + 1)}>Next →</button>
                          <button className="btn-secondary" disabled={safePage >= totalPages} onClick={() => goToPage(totalPages)} title="Last page">»</button>
                        </div>
                        <label className="pager-size">
                          Rows per page
                          <select value={pageSize} onChange={e => { setPageSize(Number(e.target.value)); goToPage(1) }}>
                            {[25, 50, 100, 250].map(n => <option key={n} value={n}>{n}</option>)}
                            <option value={0}>All</option>
                          </select>
                        </label>
                      </div>
                    </>
                  )}

                  {activeView === 'judge' && renderReportTable(
                    [
                      { key: 'judge', label: 'Judge Name' },
                      { key: 'civil', label: 'CIVIL' },
                      { key: 'criminal', label: 'CRIMINAL' },
                      { key: 'total', label: 'TOTAL DISPOSALS' }
                    ],
                    webReports.judgeSummary
                  )}
                  {activeView === 'age' && renderReportTable(
                    [{ key: 'judge', label: 'Judge Name' }, ...AGE_CATEGORIES.map(category => ({ key: category, label: category })), { key: 'total', label: 'Total' }],
                    webReports.ageWise
                  )}
                  {activeView === 'nature' && renderAgeNatureTable(webReports.ageNature)}
                  {activeView === 'natureCivil' && renderAgeNatureTable(webReports.ageNatureCivil)}
                  {activeView === 'natureCriminal' && renderAgeNatureTable(webReports.ageNatureCriminal)}
                  {activeView === 'sourceDates' && renderDateRangeTable(webReports.sourceFiles, webReports.sourceDateRange)}
                  {activeView === 'source' && renderReportTable(
                    [{ key: 'judge', label: 'Judge Name' }, ...webReports.sourceFiles.map(sourceFile => ({ key: sourceFile, label: sourceFile })), { key: 'total', label: 'Grand Total' }],
                    webReports.sourcePivot
                  )}
                </section>
              </div>
            </>
          )}
        </div>
      </main>

      {!isDash && <footer className="footer">
        <div className="container">PDF TO EXCEL PRO · Runs entirely in your browser</div>
      </footer>}
    </div>
  )
}
