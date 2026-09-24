# SheetSense — Turn sheets into insights

![SheetSense Logo](./public/logo.jpg)

Modern, premium analytics for CSV, JSON, XLSX, XLS up to **20MB** — all local, no server upload.

**Live Demo:** Enable GitHub Pages from `main` branch → `/(root)` and your site is live.

### ✨ Features

- **Upload:** Drag-drop CSV, JSON, XLSX, XLS (20MB max, ~100k rows) — SheetJS 0.18.5 with fallback CDN, progress 15%→100% (fixed 0% freeze)
- **AI Format & Clean:** Header normalization (Title Case), trim whitespace, duplicate headers → Name 2, dates → YYYY-MM-DD ISO, $1,240.00 → 1240, $2400 → Amount + Currency split, blank removal, smart fill
- **Data:** Searchable table, type badges, export CSV, raw/formatted toggle, quality score
- **Analyze:** Stats, column cards, unique counts
- **Analytics (new):** KPIs (Rows/Cols/Quality/Missing/Dupes), auto-insights in natural language, distribution histograms, time analysis by month, pivot builder (Group by + Sum/Avg/Count), correlation, quality deep dive, outliers
- **Visuals (Pro):** 9 chart types (Bar, Line, Area, Pie, Donut, Scatter, Heatmap, Funnel, Treemap), X/Y/Y2 + Aggregation + Time bucket (Day/Week/Month/Quarter/Year), color schemes (Premium/Ocean/Sunset/Mono), filters (category, date, value), presets (Revenue by Category, Monthly Trend...), export PNG, compare mode, thumbnail gallery
- **Currency:** Auto-detect Amount/Currency, target conversion EUR/USD/INR/GBP demo rates
- **Auth:** Detailed premium split modal — Email + Password only (no Google), Name, Company, Role, password strength meter (weak/medium/strong), Terms checkbox, Remember me, Guest mode
- **Themes:** Light Premium (#FBFBF8 + #6D28D9 + #FF8A4C + Gold), Dark Midnight (#0A0A0F + #14141A + #8B5CF6), Executive Blue (#0F172A + #06B6D4) — toggle in header + Settings > Appearance with preview cards, localStorage persisted
- **Settings:** Profile, Appearance (themes), Preferences (currency, date format, number format, max rows), Data & Privacy (local-only, clear, export JSON), Danger Zone, About

### 🚀 Deploy to GitHub Pages (30 sec)

1. Create new repo on GitHub: `sheetsense`
2. Upload all files from this folder (or `git push`)
3. Go to repo **Settings → Pages**
4. Source: **Deploy from branch**, Branch: **main**, Folder: **/(root)**
5. Save — your site is live at `https://<username>.github.io/sheetsense/`

No build step needed — `index.html` is fully self-contained (logo embedded as base64, SheetJS via CDN with fallback).

### 💻 Local Dev

Just open `index.html` in browser. Everything runs locally, no server.

For React source dev:
```bash
npm create vite@latest sheetsense -- --template react-ts
cd sheetsense
npm install
# copy src/App.tsx from this repo
npm run dev
```

### 📁 Structure

```
.
├── index.html          # Production build (310KB, self-contained, ready for Pages)
├── public/logo.jpg     # Original logo
├── src/App.tsx         # Source (React + Tailwind)
├── README.md
└── .gitignore
```

### 🔒 Privacy

100% local-first. Files never leave your browser. XLSX parsed via SheetJS in memory. No tracking, no backend.

### 📄 License

MIT — free for personal & commercial.

Built with: React, Tailwind, SheetJS 0.18.5, Canvas API, localStorage.

— Logo: Document Analytics (bar chart inside doc) provided by you.
