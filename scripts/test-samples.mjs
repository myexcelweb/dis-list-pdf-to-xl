// Parses every PDF in "sample of uplod files/", runs sanity checks on the
// extracted records and writes a combined Excel report.
//
//   npm run test:samples            (summary + checks)
//   npm run test:samples -- --dump  (also writes raw PDF text to test-output/)
import fs from 'fs'
import crypto from 'crypto'
import path from 'path'
import * as pdfjs from 'pdfjs-dist/legacy/build/pdf.mjs'
import {
  parseSideCsv, extractTextFromPdfDocument, parseDisposalRegister,
  ageInDays, ageInYears, ageCategory, AGE_CATEGORIES, buildWebReports, removeDuplicateEntries
} from '../src/parser.js'
import { buildWorkbook } from '../src/excelExport.js'

const root = path.resolve(path.dirname(new URL(import.meta.url).pathname.replace(/^\/(\w:)/, '$1')), '..')
const sampleDir = path.join(root, 'sample of uplod files')
const outDir = path.join(root, 'test-output')
const dump = process.argv.includes('--dump')
fs.mkdirSync(outDir, { recursive: true })

const sideMap = parseSideCsv(fs.readFileSync(path.join(root, 'src/side_classification_mappings.csv'), 'utf8'))
const files = fs.readdirSync(sampleDir).filter(f => /\.pdf$/i.test(f)).sort()

let all = []
let problems = 0
const warn = (file, msg) => { problems++; console.log(`  ! ${file}: ${msg}`) }

const seen = new Map()
const skippedFileRecords = []
const pad = n => String(n).padStart(2, '0')
const esc = s => s.replace(/[.*+?^${}()|[\]\\/]/g, '\\$&')

for (const file of files) {
  const bytes = fs.readFileSync(path.join(sampleDir, file))
  // Same rule as the web app: identical PDFs are only counted once
  const hash = crypto.createHash('sha256').update(bytes).digest('hex')
  const data = new Uint8Array(bytes)
  const pdf = await pdfjs.getDocument({ data, verbosity: 0 }).promise
  const text = await extractTextFromPdfDocument(pdf)
  if (dump) fs.writeFileSync(path.join(outDir, file.replace(/\.pdf$/i, '.txt')), text)
  const records = parseDisposalRegister(text, file, sideMap)
  if (seen.has(hash)) {
    console.log(`
${file}: skipped – identical to ${seen.get(hash)}`)
    skippedFileRecords.push(...records) // used below to test entry-level de-duplication
    continue
  }
  seen.set(hash, file)

  // Highest serial number printed in the PDF = expected record count
  const serials = [...text.matchAll(/^(\d{1,4})\s+\S/gm)].map(m => +m[1])
  const nos = records.map(r => r.s_no)
  const maxSerial = Math.max(0, ...nos)
  const missing = []
  for (let n = 1; n <= maxSerial; n++) if (!nos.includes(n)) missing.push(n)
  const dupes = nos.filter((n, i) => nos.indexOf(n) !== i)

  console.log(`\n${file}: ${pdf.numPages} pages, ${records.length} records (max S.No. ${maxSerial}, candidate lines ${serials.length})`)
  console.log(`  court: ${records[0]?.court_name}  | header judges: ${[...new Set(records.map(r => r.judge_in_header))].join(', ')}`)
  console.log(`  judges: ${JSON.stringify(Object.fromEntries([...new Set(records.map(r => r.judge_name))].map(j => [j, records.filter(r => r.judge_name === j).length])))}`)
  console.log(`  nature: ${JSON.stringify(Object.fromEntries([...new Set(records.map(r => r.disposal_nature))].map(j => [j, records.filter(r => r.disposal_nature === j).length])))}`)

  if (!records.length) warn(file, 'no records parsed')
  if (missing.length) warn(file, `missing S.No. ${missing.slice(0, 30).join(',')}${missing.length > 30 ? '…' : ''} (${missing.length})`)
  if (dupes.length) warn(file, `duplicate S.No. ${[...new Set(dupes)].slice(0, 20).join(',')}`)
  if (records[0]?.court_name === 'Unknown') warn(file, 'court name not found')
  const unmapped = [...new Set(records.filter(r => r.side === 'UNMAPPED').map(r => r.case_type))]
  if (unmapped.length) warn(file, `UNMAPPED case types: ${unmapped.join(' | ')}`)
  const badNature = records.filter(r => !/^(Contested|Uncontested)$/i.test(r.disposal_nature))
  if (badNature.length) warn(file, `${badNature.length} records with nature "${[...new Set(badNature.map(r => r.disposal_nature))].join('|')}"`)
  const negAge = records.filter(r => ageInDays(r.filing_date, r.disposal_date) === '' || ageInDays(r.filing_date, r.disposal_date) < 0)
  if (negAge.length) warn(file, `${negAge.length} records with invalid/negative age, e.g. S.No. ${negAge[0].s_no}`)
  const odd = records.filter(r => !r.judge_name || r.judge_name.length < 4 || /\d|DISPOSAL|COURT|JUDGE|NATURE/i.test(r.judge_name))
  if (odd.length) warn(file, `${odd.length} suspicious judge names: ${[...new Set(odd.map(r => r.judge_name))].slice(0, 5).join(' | ')}`)
  const oddType = records.filter(r => !r.disposal_type || /\d{2}-\d{2}-\d{4}|\bCOURT\b|JUDGE NAME/i.test(r.disposal_type))
  if (oddType.length) warn(file, `${oddType.length} suspicious disposal types: ${[...new Set(oddType.map(r => r.disposal_type))].slice(0, 5).join(' | ')}`)
  const oddCase = records.filter(r => /\d|\//.test(r.case_type) || r.case_type.length > 14)
  if (oddCase.length) warn(file, `${oddCase.length} suspicious case types: ${[...new Set(oddCase.map(r => r.case_type))].slice(0, 5).join(' | ')}`)

  // Independent cross-check: every record must appear in the raw text as
  // "S.No. CaseType… Reg.No. Filing Disposal DisposalType… YY MM DD", and the number of
  // "Reg.No. date date" rows in the text must equal the number of records.
  const flat = text.replace(/\s+/g, ' ')
  const rowsInText = (flat.match(/\d+\/\d{4} \d{2}-\d{2}-\d{4} \d{2}-\d{2}-\d{4}/g) || []).length
  if (rowsInText !== records.length) warn(file, `${rowsInText} rows in PDF text but ${records.length} parsed`)
  const notFound = records.filter(r => !new RegExp(
    `(?<![0-9])${r.s_no} ${esc(r.case_type.split(' ')[0])}[A-Z ]*? ${esc(r.reg_no)} ${r.filing_date} ${r.disposal_date} ` +
    `${esc(r.disposal_type.split(' ')[0])}.{0,60}?${pad(r.duration_y)} ${pad(r.duration_m)} ${pad(r.duration_d)}`
  ).test(flat))
  if (notFound.length) warn(file, `${notFound.length} records do not match the PDF text, e.g. S.No. ${notFound.slice(0, 5).map(r => r.s_no).join(',')}`)

  all = all.concat(records)
}

// Entry-level duplicates (same CASE NO + Judge Name), as in the web app
const dedup = removeDuplicateEntries(all)
if (dedup.duplicates.length) console.log(`
Skipped ${dedup.duplicates.length} duplicate entries (same CASE NO + Judge Name)`)
all = dedup.unique
// Records from a skipped identical PDF must all be caught by the entry-level rule too
if (skippedFileRecords.length) {
  const check = removeDuplicateEntries([...all, ...skippedFileRecords])
  if (check.duplicates.length !== skippedFileRecords.length || check.unique.length !== all.length) {
    warn('dedupe', `expected ${skippedFileRecords.length} duplicates, got ${check.duplicates.length}`)
  } else {
    console.log(`Duplicate-entry check: all ${skippedFileRecords.length} entries from the identical PDF were detected`)
  }
}

// Report totals must agree with the record count
const reports = buildWebReports(all)
const gt = rows => rows[rows.length - 1]
if (gt(reports.judgeSummary).total !== all.length) warn('reports', 'judge summary total mismatch')
if (gt(reports.ageWise).total !== all.length) warn('reports', 'age-wise total mismatch')
if (gt(reports.ageNature).counts.TOTAL.total !== all.length) warn('reports', 'age-nature total mismatch')
if (gt(reports.sourcePivot).total !== all.length) warn('reports', 'source pivot total mismatch')

const wb = await buildWorkbook(all, { ageInDays, ageInYears, ageCategory, AGE_CATEGORIES, buildWebReports })
const xlsx = path.join(outDir, 'Combined_Disposal_Report.xlsx')
await wb.xlsx.writeFile(xlsx)

console.log(`\nTOTAL: ${all.length} records from ${files.length} PDFs → ${path.relative(root, xlsx)}`)
console.log(problems ? `${problems} problem(s) found` : 'All checks passed')
process.exitCode = problems ? 1 : 0
