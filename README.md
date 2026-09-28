# PDF TO EXCEL PRO – Web (Client-side)

Upload a Court **Disposal Register PDF** → get a properly formatted **Excel report**.

- **100% client-side** (no backend, no server, no data leaves the browser)
- React + Vite + pdf.js + SheetJS
- Ready for **Netlify** deployment via Git push

## Features

- Drag & drop or click to upload PDF
- Extracts text with pdf.js
- Parses records in the same style as the original Python STEP1 tool
- Generates Excel with two sheets:
  1. **Disposal Register** – full list of cases
  2. **Judge Summary** – count per judge
- Preview table before download

## Local Development

```bash
npm install
npm run dev
```

Open http://localhost:5173

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
