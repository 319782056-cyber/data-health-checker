# Data Health Checker

A modern, browser-first Excel and CSV data quality checker.

## Product goals

- Import XLSX and CSV files locally in the browser
- Preview worksheets and tabular data
- Profile missing values, duplicate rows, and column types
- Provide a safe foundation for data cleaning and export

## Privacy

The first version processes uploaded files locally in your browser. Files are not sent to an application server.

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
