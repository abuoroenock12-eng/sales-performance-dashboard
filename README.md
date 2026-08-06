# Sales Performance Dashboard

A modern, responsive sales BI dashboard built with React, Tailwind CSS, and Recharts.
Includes KPI cards, six interactive charts, a filterable/searchable/sortable data table
with CSV and Excel export, and light/dark mode — running on 5,000 generated sample orders.

## Run it locally

```bash
npm install
npm run dev
```

Then open the local URL Vite prints (usually http://localhost:5173).

## Build for production

```bash
npm run build
```

This outputs a static site to the `dist/` folder — plain HTML/CSS/JS, ready to host anywhere.

## Deploy

### Option A — Netlify Drop (fastest, no account needed)
1. Run `npm install && npm run build`
2. Go to https://app.netlify.com/drop
3. Drag the `dist/` folder onto the page — you'll get a live URL immediately

### Option B — Vercel
1. Push this folder to a GitHub repo
2. Go to https://vercel.com/new and import the repo
3. Vercel auto-detects Vite, builds, and deploys — you get a live URL and
   automatic redeploys on every push

### Option C — Netlify (connected to GitHub)
1. Push this folder to a GitHub repo
2. Go to https://app.netlify.com, "Add new site" → "Import an existing project"
3. Build command: `npm run build`, publish directory: `dist`

## Project structure

```
sales-performance-dashboard/
├── index.html            # Vite entry HTML
├── package.json          # Dependencies and scripts
├── tailwind.config.js    # Tailwind setup
├── postcss.config.js     # PostCSS/Tailwind pipeline
├── vite.config.js        # Vite + React plugin config
└── src/
    ├── main.jsx           # React root / mounts the app
    ├── index.css          # Tailwind directives
    └── SalesDashboard.jsx # The dashboard itself (all components + sample data)
```

## Notes

- Sample data (5,000 orders) is generated in-memory with a seeded random
  function in `SalesDashboard.jsx`, so it's consistent across reloads.
  Swap `generateSalesData()` for a real API call when you're ready to
  connect live data.
- Dark mode is toggled via React state (top-right icon in the header),
  not persisted between sessions — wire it to `localStorage` or a user
  preference if you want it to stick.
- CSV export is built with the browser's Blob API; Excel export uses
  the `xlsx` (SheetJS) package.
