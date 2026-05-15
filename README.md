# TAT & MTD Dashboard

Streamlit dashboard for **Dispatch / Delivery Turnaround Time** + **Month-to-Date stock reports**, deployable to **Vercel** as a fully static site.

The app runs **entirely in your browser** via [stlite](https://github.com/whitphx/stlite) (Streamlit compiled to WebAssembly with Pyodide). No backend, no server, no file uploads to a remote system — your data never leaves the browser.

## Features

### 5 Parent Toggles

1. **Dispatch TAT** — 4 sub-views (Invoice Date / MIS Item / Sales Channel / Warehouse) with `0-2 / 3-5 / >5 Days` buckets
2. **Delivery TAT** — same 4 sub-views with `0-9 / 10-12 / >12 Days / Delivery Pending` buckets
3. **MTD Dispatch** — Sum of Stock Qty in Kg grouped by Warehouse
4. **MTD Sales Channel** — Sales Channel × Warehouse cross-tab (sum of stock qty)
5. **MTD Sale Invoice** — Invoice Date × Warehouse cross-tab (sum of stock qty)

Dispatch & Delivery TAT tabs include a **Customer drill-down** under each Sales Channel.

### Filters

- Date range with presets (All Time / This Month MTD / Last 7-15-30 Days / Custom)
- From Warehouse, MIS Item Group, Sales Channel, Customer — all with searchable multi-select, Select All / None, all-selected by default
- Cascading Customer filter from selected Sales Channels

### Output

- Color-coded percentage tables (Dispatch / Delivery TAT) — green ≤30%, amber 30-50%, red ≥50%
- Indian-style number formatting (1,87,567) on MTD tables
- Bold dark-navy headers on every table for high visibility
- One-click Excel download with all pivots from all toggles in separate sheets

## Deploy to Vercel

1. Go to https://vercel.com and sign in with GitHub
2. **Add New** → **Project** → import `Rameshwarnaik013/tat-mtd-dashboard`
3. **Configure Project**: leave everything as default (Framework: Other, no build command)
4. Click **Deploy**

App goes live at `https://<project-name>.vercel.app` in under a minute.

## Local development

This repo is browser-only. To run locally as a regular Streamlit app:

```bash
pip install streamlit pandas numpy plotly openpyxl
streamlit run streamlit_app.py
```

## Required columns in upload

| Column                | Used by                          |
| --------------------- | -------------------------------- |
| Invoice Date          | All views, date filter           |
| Sales_Channel         | All views, filter                |
| From Warehouse        | All views, filter                |
| New MIS Item Group    | TAT views, filter                |
| Customer              | TAT customer drill-down, filter  |
| Dis Days / Dis TAT    | Dispatch TAT bucketing           |
| Delivery Days         | Delivery TAT bucketing           |
| Stock Qty In Kg       | MTD views                        |
| Sale Invoice          | KPI count (optional)             |

## File map

- `index.html` — stlite browser bootstrap, mounts `streamlit_app.py`
- `streamlit_app.py` — the full Streamlit app
- `vercel.json` — explicit static-only deployment
- `.vercelignore` — keep README out of deployment
