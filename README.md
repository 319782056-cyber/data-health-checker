# Data Workbench

A browser-first, local data preparation and analysis workspace for Excel and CSV files.

## V1.0 features

- Import multiple XLSX, XLS and CSV files, including multiple worksheets
- Search across all columns, apply multiple text filters, and sort columns
- Batch cleaning: trim whitespace, remove fully duplicated rows, remove fully blank rows, fill a selected column's blank cells, uppercase/lowercase text
- Build basic pivot summaries with row grouping, optional column grouping, sum, count and average
- Export the current worksheet as XLSX or CSV; export pivot output as XLSX
- Undo recent data transformations within the current browser session

## Privacy

Files are read with the browser File API and processed in browser memory using SheetJS. The app does not upload spreadsheet contents to an application server or use cloud storage. GitHub Pages only serves the static application files. As with any browser application, browser extensions, malware, and future code changes are outside this guarantee.

## Important limitations in V1.0

- This is a data preparation workbench, not a full Excel replacement. Formulas, macros, charts, formatting, merged cells, and workbook features are not preserved when exporting the transformed table.
- Pivot tables are basic summaries; they do not yet support calculated fields, subtotals, slicers, or full Excel pivot-table metadata.
- Filters are text-contains filters; numeric/date operators and nested AND/OR conditions are planned.
- Sorting, filtering, and cleaning operate on the complete in-memory dataset, while the table preview is capped at 500 visible rows.
- Large files are constrained by browser memory and device performance.
- Undo history is session-only and is cleared when files are re-imported or the page is refreshed.

## Development

```bash
npm install
npm run dev
npm run build
```

## Stack

- React + Vite + TypeScript
- SheetJS (xlsx)
- Lucide React
