# Data Health Checker

A modern, browser-first Excel and CSV data quality checker.

## Product goals

- Import XLSX and CSV files locally in the browser
- Preview worksheets and tabular data
- Profile missing values, duplicate rows, and column types
- Provide a safe foundation for data cleaning and export

## Privacy

The first version reads and processes selected XLSX/CSV files in the user's browser using the File API and SheetJS. The application code does not upload file contents to an application server. No account or cloud storage is used.

The site no longer loads web fonts from a third-party font CDN. As with any website, the browser still requests the site's static application assets from GitHub Pages. Browser extensions, malware, and future code changes are outside this local-processing guarantee.

## Development

```bash
npm install
npm run dev
```

## Stack

- React + Vite + TypeScript
- SheetJS (xlsx) for spreadsheet parsing
- Lucide React icons

## Current stage

Phase 1: project foundation, upload, worksheet selection, preview, and starter profiling indicators.
