# Sales & Inventory Unified Dashboard

Single-page dashboard to replace separate Excel files for sales and inventory tracking.

## Features
- Monthly & Weekly sales charts
- Product-wise sales
- Current stock position with search & low-stock filter
- Low-stock alerts
- Top-selling products
- Customer-wise sales analysis
- Profit / Margin analysis
- Monthly comparison (this vs last vs last year same month)
- CSV upload for your own Excel data (save Excel as CSV)

## Tech Stack
- React + Vite + TypeScript
- Tailwind CSS
- Recharts
- Lucide Icons

## Run Locally

```bash
npm install
npm run dev
```

## Build

```bash
npm run build
```

Output is in `dist/` folder.

## How to push to GitHub

1. Create a new empty repo on github.com (e.g., `sales-inventory-dashboard`)
2. In this project folder, run:

```bash
git init
git add .
git commit -m "Initial commit: sales inventory dashboard"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/sales-inventory-dashboard.git
git push -u origin main
```

Replace YOUR_USERNAME.

## Deploy for Free

**Option A - GitHub Pages:**
```bash
npm install gh-pages --save-dev
```
Add to `package.json` scripts:
```json
"homepage": "https://YOUR_USERNAME.github.io/sales-inventory-dashboard",
"predeploy": "npm run build",
"deploy": "gh-pages -d dist"
```
Then:
```bash
npm run deploy
```

Set base in vite.config.ts to `/sales-inventory-dashboard/` for GitHub Pages.

**Option B - Vercel (Easiest):**
1. Go to vercel.com -> Add New Project -> Import your GitHub repo
2. Click Deploy - done.

## Excel Format Expected

**Sales CSV:** `date,product,qty,customer,selling_price,cost_price`
**Inventory CSV:** `product,sku,current_stock,low_stock_threshold,cost_price`

All calculations are client-side. No data leaves the browser.