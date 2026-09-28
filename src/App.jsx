import { useState, useRef, useCallback } from 'react'
import * as XLSX from 'xlsx'
import './App.css'

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

  let fullText = ''
  for (let i = 1; i <= pdf.numPages; i++) {
    const page = await pdf.getPage(i)
    const content = await page.getTextContent()
    // Reconstruct lines roughly by y-position
    const items = content.items
    let lastY = null
    let line = ''
    for (const item of items) {
      if (lastY !== null && Math.abs(item.transform[5] - lastY) > 5) {
        fullText += line.trim() + '\n'
        line = ''
      }
      line += item.str + ' '
      lastY = item.transform[5]
    }
    if (line.trim()) fullText += line.trim() + '\n'
  }
  return fullText
}

// ---------- Parser (robust – correctly applies Contested / Uncontested) ----------
function parseDisposalRegister(fullText, filename = 'uploaded.pdf') {
  const lines = fullText.split('\n').map(l => l.trim()).filter(Boolean)

  // ---- Header info ----
  let courtName = 'Unknown'
  let judgeInHeader = 'Unknown'
  for (let i = 0; i < Math.min(20, lines.length); i++) {
    const upper = lines[i].toUpperCase()
    if (upper.includes('COURT,') || (upper.includes('COURT') && upper.includes(','))) {
      courtName = lines[i]
    } else if (upper.includes('IN THE COURT OF')) {
      for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
        if (/^[A-Z][A-Z.\s]+$/.test(lines[j]) && !lines[j].toUpperCase().includes('DISPOSAL')) {
          judgeInHeader = lines[j]
          break
        }
      }
    }
  }

  // ---- Helpers ----
  // Lines that belong to page headers / footers – never part of a record
  const isPageHeader = (line) => {
    const u = line.toUpperCase()
    return (
      u.includes('DISPOSAL REGISTER') ||
      u.includes('S.NO.') ||
      u.includes('CASE TYPE') ||
      u === 'Y M D' ||
      u === 'DURATION' ||
      u.includes('TALUKA COURT') ||
      u.includes('CIVIL COURT') ||
      u.includes('IN THE COURT OF') ||
      u.includes('JUDGMENT JUDGE') ||
      u.includes('JUDGEMENT JUDGE') ||
      /^\d+\s*\/\s*\d+$/.test(line) ||          // page numbers 1/22
      /^PRINCIPAL\s+(SENIOR\s+)?CIVIL/i.test(line)
    )
  }

  // Pure noise that should never be joined
  const isNoise = (line) => isPageHeader(line)

  // Words that legitimately continue a multi-line field
  const CONTINUATION_WORDS = new Set([
    'ALLOWED', 'ACQUITTAL', 'CONVICTION', 'DEFAULT', 'GUILTY', 'PARTLY',
    'PRELIMINARY', 'DECREE', 'DISMISSED', 'FOR', 'DISPOSED', 'OF',
    'PLEAD', 'WITHDRAWN', 'ABATED', 'COMPROMISED', 'COMMITAL',
    'CONVERTED', 'TO', 'LOK', 'ADALAT', 'REJECTED'
  ])

  // Known short case-type suffixes that appear on their own line
  const CASE_TYPE_SUFFIXES = new Set(['J', 'S', 'EN', 'SC', 'R'])

  // ---- Build record blocks, tracking Disposal Nature ----
  const recordBlocks = []   // { nature, text }
  let currentBlockLines = []
  let currentDisposalNature = 'Unknown'
  let justSawPageHeader = false   // skip pure-name lines that follow a page header

  const flushBlock = () => {
    if (currentBlockLines.length === 0) return
    const text = currentBlockLines.join(' ').replace(/\s+/g, ' ').trim()
    if (text) recordBlocks.push({ nature: currentDisposalNature, text })
    currentBlockLines = []
  }

  for (const line of lines) {
    // Page header → force close any open record so header text never joins it
    if (isPageHeader(line)) {
      flushBlock()
      justSawPageHeader = true
      continue
    }

    // After a page header, the next pure name line is the judge name in the header – skip it
    if (justSawPageHeader && /^[A-Z][A-Z.\s]{2,}$/.test(line) && !CASE_TYPE_SUFFIXES.has(line.toUpperCase())) {
      // still keep the flag for one more possible name line, then clear
      continue
    }
    justSawPageHeader = false

    // Standalone "Disposal Nature:Contested" / "Uncontested"
    if (/^Disposal Nature\s*:/i.test(line)) {
      flushBlock()
      const parts = line.split(':')
      currentDisposalNature = (parts[1] || '').trim() || 'Unknown'
      continue
    }

    // New numbered record
    if (/^\d{1,4}\s+[A-Z]/.test(line)) {
      flushBlock()
      currentBlockLines = [line]
      continue
    }

    // Continuation of current record
    if (currentBlockLines.length > 0) {
      const up = line.toUpperCase()
      // Keep short case-type suffixes and known multi-line field words
      if (
        !isNoise(line) ||
        CASE_TYPE_SUFFIXES.has(up) ||
        CONTINUATION_WORDS.has(up)
      ) {
        // Never append a pure long uppercase name (these are header leftovers)
        const isPureName = /^[A-Z][A-Z.\s]{3,}$/.test(line) && !CASE_TYPE_SUFFIXES.has(up) && !CONTINUATION_WORDS.has(up)
        if (!isPureName) {
          currentBlockLines.push(line)
        }
      }
    }
  }
  flushBlock()

  // ---- Parse each block with robust field splitting ----
  const results = []

  // Clean garbage that sometimes sticks to the end of judge name after page breaks
  const cleanJudgeName = (raw) => {
    if (!raw) return ''
    let j = raw
      .replace(/\bJudgment\s+Judge\s+Name\b/gi, '')
      .replace(/\bJudgement\s+Judge\s+Name\b/gi, '')
      .replace(/\bJudge\s+Name\b/gi, '')
      .replace(/\bDuration\b/gi, '')
      .replace(/\bY\s+M\s+D\b/gi, '')
      .replace(/\s+/g, ' ')
      .trim()

    // If the same name is repeated (e.g. "K.R.RATHOD K.R.RATHOD"), keep only the first
    const parts = j.split(/\s+/)
    if (parts.length >= 2) {
      // Detect exact consecutive duplicates of a multi-token name
      const mid = Math.floor(parts.length / 2)
      const firstHalf = parts.slice(0, mid).join(' ')
      const secondHalf = parts.slice(mid).join(' ')
      if (firstHalf === secondHalf) {
        j = firstHalf
      }
    }
    // Also handle "NAME NAME" where NAME is a single token
    if (parts.length === 2 && parts[0] === parts[1]) {
      j = parts[0]
    }
    return j.trim()
  }

  for (const { nature, text } of recordBlocks) {
    // 1) Extract fixed prefix: SNo  CaseType  RegNo  FilingDate  DisposalDate  REST
    const prefix = text.match(
      /^(\d+)\s+(.+?)\s+(\d+\/\d{4,})\s+(\d{2}-\d{2}-\d{4})\s+(\d{2}-\d{2}-\d{4})\s+(.+)$/
    )
    if (!prefix) continue

    let [, sNo, caseType, regNo, filingDate, disposalDate, rest] = prefix
    caseType = caseType.trim()
    rest = rest.trim()

    // 2) Find the duration (last three numbers) and everything after them
    const durMatch = rest.match(/(\d{1,2})\s+(\d{1,2})\s+(\d{1,2})\s+([A-Z].*)$/)
    if (!durMatch) continue

    const y = parseInt(durMatch[1], 10)
    const m = parseInt(durMatch[2], 10)
    const d = parseInt(durMatch[3], 10)
    let after = durMatch[4].trim()
    let disposalType = rest.slice(0, durMatch.index).trim()

    // 3) Move stray continuation / case-type words out of the judge area
    const tokens = after.split(/\s+/)
    const cleanTokens = []
    for (const t of tokens) {
      const up = t.toUpperCase()
      if (CASE_TYPE_SUFFIXES.has(up) && up.length <= 2) {
        caseType = (caseType + ' ' + t).trim()
      } else if (CONTINUATION_WORDS.has(up)) {
        disposalType = (disposalType + ' ' + t).trim()
      } else {
        cleanTokens.push(t)
      }
    }
    let judgeName = cleanJudgeName(cleanTokens.join(' '))

    // Final tidy
    disposalType = disposalType.replace(/\s+/g, ' ').trim()
    caseType = caseType.replace(/\s+/g, ' ').trim()

    // Skip obviously broken rows (no judge or disposal type empty + weird)
    if (!judgeName || judgeName.length < 2) {
      // fallback: use the header judge if available
      judgeName = judgeInHeader !== 'Unknown' ? judgeInHeader : judgeName
    }

    results.push({
      s_no: parseInt(sNo, 10),
      case_type: caseType,
      reg_no: regNo,
      filing_date: filingDate,
      disposal_date: disposalDate,
      disposal_type: disposalType,
      duration_y: y,
      duration_m: m,
      duration_d: d,
      judge_name: judgeName,
      filename,
      court_name: courtName,
      judge_in_header: judgeInHeader,
      disposal_nature: nature
    })
  }

  results.sort((a, b) => a.s_no - b.s_no)
  return results
}

// ---------- Excel generation ----------
function downloadExcel(records, originalName) {
  if (!records.length) return

  const headers = [
    'S.No.', 'Filename', 'Court Name', 'Judge (Header)', 'Disposal Nature',
    'Case Type', 'Reg. No./Year', 'Date of Filing', 'Disposal Date',
    'Disposal Type', 'Duration (Y)', 'Duration (M)', 'Duration (D)', 'Judge Name'
  ]

  const rows = records.map(r => [
    r.s_no, r.filename, r.court_name, r.judge_in_header, r.disposal_nature,
    r.case_type, r.reg_no, r.filing_date, r.disposal_date,
    r.disposal_type, r.duration_y, r.duration_m, r.duration_d, r.judge_name
  ])

  const ws = XLSX.utils.aoa_to_sheet([headers, ...rows])

  // Set column widths
  ws['!cols'] = [
    { wch: 8 }, { wch: 18 }, { wch: 28 }, { wch: 16 }, { wch: 16 },
    { wch: 12 }, { wch: 14 }, { wch: 14 }, { wch: 14 },
    { wch: 22 }, { wch: 10 }, { wch: 10 }, { wch: 10 }, { wch: 16 }
  ]

  const wb = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(wb, ws, 'Disposal Register')

  // Add a simple summary sheet
  const judges = {}
  records.forEach(r => {
    judges[r.judge_name] = (judges[r.judge_name] || 0) + 1
  })
  const summaryData = [
    ['Judge Name', 'Total Disposals'],
    ...Object.entries(judges).sort((a, b) => b[1] - a[1]),
    ['GRAND TOTAL', records.length]
  ]
  const ws2 = XLSX.utils.aoa_to_sheet(summaryData)
  ws2['!cols'] = [{ wch: 20 }, { wch: 16 }]
  XLSX.utils.book_append_sheet(wb, ws2, 'Judge Summary')

  const outName = originalName.replace(/\.pdf$/i, '') + '_Disposal_Report.xlsx'
  XLSX.writeFile(wb, outName)
}

// ---------- React Component ----------
export default function App() {
  const [status, setStatus] = useState('idle') // idle | parsing | done | error
  const [message, setMessage] = useState('')
  const [records, setRecords] = useState([])
  const [fileName, setFileName] = useState('')
  const [dragOver, setDragOver] = useState(false)
  const inputRef = useRef(null)

  const processFile = useCallback(async (file) => {
    if (!file || !file.name.toLowerCase().endsWith('.pdf')) {
      setStatus('error')
      setMessage('Please upload a PDF file.')
      return
    }

    setStatus('parsing')
    setMessage('Extracting text from PDF…')
    setRecords([])
    setFileName(file.name)

    try {
      const text = await extractTextFromPdf(file)
      setMessage('Parsing disposal records…')
      const parsed = parseDisposalRegister(text, file.name)

      if (parsed.length === 0) {
        setStatus('error')
        setMessage('No records could be extracted. The PDF layout may be different from the expected Court Disposal Register format.')
        return
      }

      setRecords(parsed)
      setStatus('done')
      setMessage(`Successfully extracted ${parsed.length} records.`)
    } catch (err) {
      console.error(err)
      setStatus('error')
      setMessage('Failed to process PDF: ' + (err.message || 'Unknown error'))
    }
  }, [])

  const onDrop = (e) => {
    e.preventDefault()
    setDragOver(false)
    const file = e.dataTransfer.files?.[0]
    if (file) processFile(file)
  }

  const onFileChange = (e) => {
    const file = e.target.files?.[0]
    if (file) processFile(file)
  }

  return (
    <div className="app">
      <header className="header">
        <div className="container">
          <div className="logo">
            <span className="logo-icon">⚖️</span>
            <div>
              <h1>PDF TO EXCEL PRO</h1>
              <p className="tagline">Court Disposal Register → Excel (Client-side only)</p>
            </div>
          </div>
        </div>
      </header>

      <main className="main">
        <div className="container">
          {/* Upload Zone */}
          <section
            className={`upload-zone ${dragOver ? 'drag-over' : ''} ${status === 'parsing' ? 'busy' : ''}`}
            onDragOver={(e) => { e.preventDefault(); setDragOver(true) }}
            onDragLeave={() => setDragOver(false)}
            onDrop={onDrop}
            onClick={() => status !== 'parsing' && inputRef.current?.click()}
          >
            <input
              ref={inputRef}
              type="file"
              accept=".pdf,application/pdf"
              onChange={onFileChange}
              hidden
            />
            {status === 'parsing' ? (
              <div className="spinner-wrap">
                <div className="spinner" />
                <p>{message}</p>
              </div>
            ) : (
              <>
                <div className="upload-icon">📄</div>
                <h2>Upload Court Disposal PDF</h2>
                <p>Drag & drop or click to select a PDF file</p>
                <p className="hint">Works entirely in your browser — no server, no data leaves your computer</p>
              </>
            )}
          </section>

          {/* Status / Result */}
          {status === 'error' && (
            <div className="alert error">{message}</div>
          )}

          {status === 'done' && records.length > 0 && (
            <>
              <div className="alert success">
                {message}
                <button
                  className="btn-primary"
                  onClick={() => downloadExcel(records, fileName)}
                >
                  ⬇ Download Excel Report
                </button>
              </div>

              {/* Preview Table */}
              <section className="preview">
                <div className="preview-header">
                  <h3>Preview ({records.length} records)</h3>
                  <span className="file-name">{fileName}</span>
                </div>
                <div className="table-wrap">
                  <table>
                    <thead>
                      <tr>
                        <th>#</th>
                        <th>Case Type</th>
                        <th>Reg. No.</th>
                        <th>Filing</th>
                        <th>Disposal</th>
                        <th>Type</th>
                        <th>Y</th>
                        <th>M</th>
                        <th>D</th>
                        <th>Judge</th>
                      </tr>
                    </thead>
                    <tbody>
                      {records.slice(0, 50).map((r, idx) => (
                        <tr key={idx}>
                          <td>{r.s_no}</td>
                          <td>{r.case_type}</td>
                          <td>{r.reg_no}</td>
                          <td>{r.filing_date}</td>
                          <td>{r.disposal_date}</td>
                          <td className="disposal-type">{r.disposal_type}</td>
                          <td>{r.duration_y}</td>
                          <td>{r.duration_m}</td>
                          <td>{r.duration_d}</td>
                          <td>{r.judge_name}</td>
                        </tr>
                      ))}
                    </tbody>
                  </table>
                  {records.length > 50 && (
                    <p className="more-note">Showing first 50 of {records.length} records. Full data is in the Excel download.</p>
                  )}
                </div>
              </section>
            </>
          )}

          {/* Info */}
          <section className="info-card">
            <h3>How it works</h3>
            <ul>
              <li>Upload a Court Disposal Register PDF (same format as Taluka Court registers).</li>
              <li>Text is extracted and parsed <strong>entirely in your browser</strong>.</li>
              <li>Excel is generated with the same columns as the original Python STEP1 tool.</li>
              <li>Two sheets: <code>Disposal Register</code> (all rows) + <code>Judge Summary</code>.</li>
              <li>No backend, no server, no data is uploaded anywhere — perfect for Netlify static hosting.</li>
            </ul>
            <p className="note">
              For maximum accuracy on complex multi-line fields, the original desktop Python scripts (STEP1–STEP4) remain the gold standard.
              This web version is a convenient client-side alternative.
            </p>
          </section>
        </div>
      </main>

      <footer className="footer">
        <div className="container">
          PDF TO EXCEL PRO · Pure client-side React · Deploy on Netlify via Git
        </div>
      </footer>
    </div>
  )
}
