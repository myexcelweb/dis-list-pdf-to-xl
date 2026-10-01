# PDF TO EXCEL PRO – Web (Client-side)

Upload a Court **Disposal Register PDF** → get a properly formatted **Excel report**.

- **100% client-side** (no backend, no server, no data leaves the browser)
- React + Vite + pdf.js + ExcelJS
- Ready for **Netlify** deployment via Git push

## Features

- Drag & drop or click to upload PDF
- Separate **Upload Files** and **Reports Dashboard** pages; successful uploads open the dashboard automatically
- Clear button resets the upload and report data
- Report views are selected from a dashboard sidebar
- Extracts text with pdf.js
- Parses records in the same style as the original Python STEP1 tool
- Skips duplicate uploads (identical PDF content under any file name) and shows a warning
- Skips duplicate entries with the same **CASE NO** and **Judge Name** (first entry is kept) and lists them in the warning
- Adds a **SIDE** column by matching **CASE TYPE** against `src/side_classification_mappings.csv` (`UNMAPPED` when no mapping exists)
- Adds **CASE NO** by combining CASE TYPE and Reg. No./Year (for example, `CC/3122/2025`)
- Generates Excel with eight sheets:
  1. **Disposal Register** – full list of cases
  2. **Judge Summary** – Civil, Criminal, and total disposals per judge
  3. **Judge Age Wise** – judge counts by age category
  4. **Age Category Nature** – contested and uncontested counts by judge and age category
  5. **AGE CATEGORY NATURE(CIVIL)** – Civil side counts by judge, age, and disposal nature
  6. **AGE CATEGORY NATURE(CRIMINAL)** – Criminal side counts by judge, age, and disposal nature
  7. **Judge Source Pivot** – disposal counts by judge and source PDF
  8. **Judge Source Date Range** – earliest (From) and latest (To) disposal date per judge per source PDF
- Shows all report views in the web page and the full disposal register with search and filters
- Preview table before download

## Local Development

```bash
npm install
npm run dev
```

Open http://localhost:5173

## Test with the sample PDFs

```bash
npm run test:samples
```

Parses every PDF in `sample of uplod files/`, cross-checks each record against the raw PDF
text (serial numbers, dates, durations, row counts), checks report totals, and writes
`test-output/Combined_Disposal_Report.xlsx`. Add `-- --dump` to also save the extracted
text of each PDF. Identical PDFs (same bytes, any name) are counted only once, as in the app.

## Build

```bash
npm run build
```

Output goes to `dist/`.

## Deploy to Netlify via Git

1. Create a new empty repository (GitHub / GitLab / Bitbucket).
2. Push this folder:

```bash
git init
git add .
git commit -m "PDF TO EXCEL PRO – static React converter"
git branch -M main
git remote add origin <YOUR_REPO_URL>
git push -u origin main
```

3. In Netlify → **Add new site** → **Import an existing project**
4. Connect the repository.
5. Build settings are already in `netlify.toml`:
   - Build command: `npm run build`
   - Publish directory: `dist`
6. Deploy. Every future `git push` auto-deploys.

## Notes

- Designed for the standard Taluka Court / Disposal Register PDF layout (the format used by your existing Python scripts).
- For the absolute highest accuracy on complex multi-line fields, the original desktop Python pipeline (STEP1–STEP4) is still recommended.
- This web version is convenient, private, and requires zero infrastructure.
