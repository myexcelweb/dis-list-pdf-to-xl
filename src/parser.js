// Pure parsing / reporting logic shared by the web app and scripts/test-samples.mjs.
// No browser or Vite-specific code here.

// ---------- SIDE classification (CASE TYPE → CIVIL / CRIMINAL) ----------
export function parseSideCsv(csvText) {
  return Object.fromEntries(
    csvText.trim().split(/\r?\n/).slice(1).filter(Boolean).map(line => {
      const [caseType, side] = line.split(',')
      return [caseType.trim().toUpperCase(), (side || '').trim().toUpperCase()]
    })
  )
}

export function getCaseSide(caseType, sideMap) {
  return sideMap[String(caseType || '').trim().toUpperCase()] || 'UNMAPPED'
}

function getCaseNumber(caseType, regNo) {
  return caseType && regNo ? `${caseType}/${regNo}` : ''
}

// ---------- PDF text extraction (takes an already-loaded pdf.js document) ----------
export async function extractTextFromPdfDocument(pdf) {
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

// "MS.S. .SYED" / "K. R. RATHOD" → "MS.S.SYED" / "K.R.RATHOD"
function normalizeName(name) {
  return String(name || '').replace(/\s*\.\s*/g, '.').replace(/\.{2,}/g, '.').replace(/\s+/g, ' ').trim()
}

// ---------- Parser (robust – correctly applies Contested / Uncontested) ----------
export function parseDisposalRegister(fullText, filename = 'uploaded.pdf', sideMap = {}) {
  const lines = fullText.split('\n').map(l => l.trim()).filter(Boolean)

  // ---- Header info ----
  // Typical header:
  //   CIVIL COURT, PORBANDAR
  //   IN THE COURT OFCHIEF JUDICIAL MAGISTRATE & ADDL. SR. CIVIL JUDGE   (designation may be blank)
  //   K.K.RATHOD                                                          (may be missing)
  let courtName = 'Unknown'
  let courtOf = ''
  let judgeInHeader = 'Unknown'
  for (let i = 0; i < Math.min(20, lines.length); i++) {
    const upper = lines[i].toUpperCase()
    if (upper.startsWith('IN THE COURT OF')) {
      if (!courtOf) courtOf = lines[i].slice('IN THE COURT OF'.length).trim()
      if (judgeInHeader === 'Unknown') {
        for (let j = i + 1; j < Math.min(i + 4, lines.length); j++) {
          if (/^[A-Z][A-Z.\s]+$/.test(lines[j]) && !lines[j].toUpperCase().includes('DISPOSAL')) {
            judgeInHeader = normalizeName(lines[j])
            break
          }
        }
      }
    } else if (courtName === 'Unknown' && upper.includes('COURT') && upper.includes(',')) {
      courtName = lines[i]
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

  // Record text already contains: filing date, disposal date, … , YY MM DD
  const DURATION_SEEN = /\d{2}-\d{2}-\d{4}\s+\d{2}-\d{2}-\d{4}\s+.*?\d{1,2}\s+\d{1,2}\s+\d{1,2}/

  // Known short case-type suffixes that appear on their own line
  const CASE_TYPE_SUFFIXES = new Set(['J', 'S', 'EN', 'SC', 'R'])

  // ---- Build record blocks, tracking Disposal Nature ----
  const recordBlocks = []   // { nature, text, judgeInHeader }
  let currentBlockLines = []
  let currentDisposalNature = 'Unknown'
  let currentHeaderJudge = judgeInHeader
  let justSawPageHeader = false   // skip pure-name lines that follow a page header

  const flushBlock = () => {
    if (currentBlockLines.length === 0) return
    const text = currentBlockLines.join(' ').replace(/\s+/g, ' ').trim()
    if (text) recordBlocks.push({ nature: currentDisposalNature, text, judgeInHeader: currentHeaderJudge })
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
      currentHeaderJudge = normalizeName(line)
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

    // Continuation of current record (wrapped case type, disposal type or judge name).
    // Until the duration "YY MM DD" has been seen, every line belongs to the record
    // (e.g. "EX-PARTE" / "JUDGEMENT", "QUASHING BY" / "HIGHCOURT"). After that, only
    // judge-name fragments such as ".SYED" are kept; pure uppercase names there are
    // header leftovers.
    if (currentBlockLines.length > 0) {
      const up = line.toUpperCase()
      const hasDuration = DURATION_SEEN.test(currentBlockLines.join(' '))
      const isPureName = /^[A-Z][A-Z.\s]{3,}$/.test(line) && !CASE_TYPE_SUFFIXES.has(up) && !CONTINUATION_WORDS.has(up)
      if (!isNoise(line) && !(hasDuration && isPureName)) {
        currentBlockLines.push(line)
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
    return normalizeName(j)
  }

  for (const { nature, text, judgeInHeader: recordHeaderJudge } of recordBlocks) {
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
      judgeName = recordHeaderJudge !== 'Unknown' ? recordHeaderJudge : judgeName
    }

    results.push({
      s_no: parseInt(sNo, 10),
      case_type: caseType,
      case_no: getCaseNumber(caseType, regNo),
      side: getCaseSide(caseType, sideMap),
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
      court_of: courtOf,
      judge_in_header: recordHeaderJudge,
      disposal_nature: nature
    })

    // A few source PDFs encode a single row's judge as two slash-joined names,
    // while that page's printed header identifies the actual sitting judge.
    const parsedRecord = results[results.length - 1]
    if (parsedRecord.judge_name.includes('/') && recordHeaderJudge !== 'Unknown') {
      parsedRecord.judge_name = recordHeaderJudge
    }
  }

  results.sort((a, b) => a.s_no - b.s_no)
  return results
}

// PDF dates are dd-mm-yyyy. Use UTC so the day count is not affected by
// daylight-saving or local timezone changes.
export function ageInDays(filingDate, disposalDate) {
  const parseDate = (value) => {
    const match = value?.match(/^(\d{2})-(\d{2})-(\d{4})$/)
    if (!match) return null
    const [, day, month, year] = match
    const timestamp = Date.UTC(Number(year), Number(month) - 1, Number(day))
    const date = new Date(timestamp)
    if (date.getUTCFullYear() !== Number(year) || date.getUTCMonth() !== Number(month) - 1 || date.getUTCDate() !== Number(day)) return null
    return timestamp
  }
  const filing = parseDate(filingDate)
  const disposal = parseDate(disposalDate)
  return filing === null || disposal === null ? '' : Math.floor((disposal - filing) / 86400000)
}

export function ageInYears(filingDate, disposalDate) {
  const days = ageInDays(filingDate, disposalDate)
  return days === '' ? '' : Math.round((days / 365.2425) * 100) / 100
}

export function ageCategory(filingDate, disposalDate) {
  const years = ageInYears(filingDate, disposalDate)
  if (years === '') return ''
  if (years < 5) return '0-5Y'
  if (years < 10) return '5-10Y'
  if (years < 20) return '10-20Y'
  if (years < 30) return '20-30Y'
  return 'MORE THAN 30Y'
}

export const AGE_CATEGORIES = ['0-5Y', '5-10Y', '10-20Y', '20-30Y', 'MORE THAN 30Y']

export function buildAgeNatureReport(records) {
  const countsByJudge = new Map()
  for (const r of records) {
    const category = ageCategory(r.filing_date, r.disposal_date)
    if (!category) continue
    const judge = r.judge_name || 'Unknown'
    if (!countsByJudge.has(judge)) {
      countsByJudge.set(judge, Object.fromEntries(
        [...AGE_CATEGORIES, 'TOTAL'].map(value => [value, { contested: 0, uncontested: 0, total: 0 }])
      ))
    }
    const counts = countsByJudge.get(judge)
    const nature = (r.disposal_nature || '').trim().toLowerCase()
    const natureKey = nature.includes('uncontested') ? 'uncontested' : nature.includes('contested') ? 'contested' : null
    counts[category].total += 1
    counts.TOTAL.total += 1
    if (natureKey) {
      counts[category][natureKey] += 1
      counts.TOTAL[natureKey] += 1
    }
  }

  const rows = [...countsByJudge.entries()]
    .sort((a, b) => b[1].TOTAL.total - a[1].TOTAL.total)
    .map(([judge, counts]) => ({ judge, counts }))
  const grandTotals = Object.fromEntries(
    [...AGE_CATEGORIES, 'TOTAL'].map(group => [group, {
      contested: rows.reduce((sum, row) => sum + row.counts[group].contested, 0),
      uncontested: rows.reduce((sum, row) => sum + row.counts[group].uncontested, 0),
      total: rows.reduce((sum, row) => sum + row.counts[group].total, 0)
    }])
  )
  rows.push({ judge: 'GRAND TOTAL', counts: grandTotals })
  return rows
}

export function buildWebReports(records) {
  const judgeCounts = new Map()
  const ageCounts = new Map()
  const judgeSideCounts = new Map()
  const sourceFiles = [...new Set(records.map(r => r.filename || 'Unknown'))].sort((a, b) => a.localeCompare(b))
  const sourceCounts = new Map()

  for (const r of records) {
    const judge = r.judge_name || 'Unknown'
    const category = ageCategory(r.filing_date, r.disposal_date)
    judgeCounts.set(judge, (judgeCounts.get(judge) || 0) + 1)

    if (!judgeSideCounts.has(judge)) judgeSideCounts.set(judge, { civil: 0, criminal: 0, total: 0 })
    const sideCounts = judgeSideCounts.get(judge)
    if (r.side === 'CIVIL') sideCounts.civil += 1
    if (r.side === 'CRIMINAL') sideCounts.criminal += 1
    sideCounts.total += 1

    if (category) {
      if (!ageCounts.has(judge)) ageCounts.set(judge, Object.fromEntries(AGE_CATEGORIES.map(value => [value, 0])))
      ageCounts.get(judge)[category] += 1
    }

    const sourceFile = r.filename || 'Unknown'
    if (!sourceCounts.has(judge)) sourceCounts.set(judge, new Map())
    const judgeSourceCounts = sourceCounts.get(judge)
    judgeSourceCounts.set(sourceFile, (judgeSourceCounts.get(sourceFile) || 0) + 1)
  }

  const judges = [...judgeCounts.keys()].sort((a, b) => a.localeCompare(b))
  const judgeSummary = judges.map(judge => ({ judge, ...judgeSideCounts.get(judge) }))
  judgeSummary.push({
    judge: 'GRAND TOTAL',
    civil: judges.reduce((sum, judge) => sum + judgeSideCounts.get(judge).civil, 0),
    criminal: judges.reduce((sum, judge) => sum + judgeSideCounts.get(judge).criminal, 0),
    total: records.length
  })

  const ageWise = judges.map(judge => {
    const counts = ageCounts.get(judge) || Object.fromEntries(AGE_CATEGORIES.map(value => [value, 0]))
    return { judge, ...counts, total: AGE_CATEGORIES.reduce((sum, category) => sum + counts[category], 0) }
  })
  ageWise.push({
    judge: 'GRAND TOTAL',
    ...Object.fromEntries(AGE_CATEGORIES.map(category => [
      category,
      judges.reduce((sum, judge) => sum + (ageCounts.get(judge)?.[category] || 0), 0)
    ])),
    total: AGE_CATEGORIES.reduce((sum, category) =>
      sum + judges.reduce((categorySum, judge) => categorySum + (ageCounts.get(judge)?.[category] || 0), 0), 0)
  })

  const ageNature = buildAgeNatureReport(records)
  const ageNatureCivil = buildAgeNatureReport(records.filter(r => r.side === 'CIVIL'))
  const ageNatureCriminal = buildAgeNatureReport(records.filter(r => r.side === 'CRIMINAL'))

  const sourcePivot = judges.map(judge => {
    const counts = sourceCounts.get(judge) || new Map()
    return {
      judge,
      ...Object.fromEntries(sourceFiles.map(sourceFile => [sourceFile, counts.get(sourceFile) || 0])),
      total: judgeCounts.get(judge)
    }
  })
  sourcePivot.push({
    judge: 'GRAND TOTAL',
    ...Object.fromEntries(sourceFiles.map(sourceFile => [
      sourceFile,
      records.filter(r => (r.filename || 'Unknown') === sourceFile).length
    ])),
    total: records.length
  })

  // Earliest / latest disposal date per judge per source PDF (no totals)
  const dateKey = d => { const m = String(d || '').match(/^(\d{2})-(\d{2})-(\d{4})$/); return m ? `${m[3]}${m[2]}${m[1]}` : '' }
  const ranges = new Map() // judge → Map(file → { from, to })
  for (const r of records) {
    if (!dateKey(r.disposal_date)) continue
    const judge = r.judge_name || 'Unknown'
    const file = r.filename || 'Unknown'
    if (!ranges.has(judge)) ranges.set(judge, new Map())
    const range = ranges.get(judge).get(file)
    if (!range) ranges.get(judge).set(file, { from: r.disposal_date, to: r.disposal_date })
    else {
      if (dateKey(r.disposal_date) < dateKey(range.from)) range.from = r.disposal_date
      if (dateKey(r.disposal_date) > dateKey(range.to)) range.to = r.disposal_date
    }
  }
  const sourceDateRange = judges.map(judge => ({
    judge,
    ranges: Object.fromEntries(sourceFiles.map(file => [file, ranges.get(judge)?.get(file) || null]))
  }))

  return { sourceFiles, judgeSummary, ageWise, ageNature, ageNatureCivil, ageNatureCriminal, sourcePivot, sourceDateRange }
}


// ---------- Duplicate entries ----------
// The same case disposed by the same judge can appear more than once (overlapping
// registers, the same PDF exported twice, …). Keep the first occurrence of each
// CASE NO + Judge Name pair; return the skipped ones so the UI can report them.
export function removeDuplicateEntries(records) {
  const seen = new Set()
  const unique = []
  const duplicates = []
  for (const r of records) {
    const key = `${String(r.case_no || '').trim().toUpperCase()}|${String(r.judge_name || '').trim().toUpperCase()}`
    if (r.case_no && seen.has(key)) {
      duplicates.push(r)
    } else {
      seen.add(key)
      unique.push(r)
    }
  }
  return { unique, duplicates }
}
