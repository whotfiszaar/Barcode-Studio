# Uddesh's Barcode Studio

A modern, fast, client-side PWA utility to generate 1D and 2D barcode PDF label sheets from any Excel or CSV spreadsheet.

## Key Features

- **Spreadsheet Input**: Upload `.xlsx`, `.xls`, or `.csv` files with single or multiple worksheets.
- **Smart Field Mapping**: Place any column at the top header, middle barcode, or bottom label.
- **Symbology Support**:
  - 1D Barcodes: Code 128, Code 39, EAN-13, UPC-A, ITF-14
  - 2D Barcodes: QR Code, Data Matrix, PDF417
- **Grid Layout Presets**:
  - Warehouse Standard (2x5 Grid, 10 per A4 page)
  - Compact Shelf Tags (3x7 Grid, 21 per A4 page)
  - Pallet Cards (2x3 Grid, 6 per A4 page)
  - Avery 5160 & 5163 standard sheets
  - Thermal Roll 4x6 inch and 3x2 inch single stickers
  - Fully customizable rows, columns, margins, and gaps
- **Live WYSIWYG Preview**: Instant DOM rendering of printable sheets with pagination and zoom controls.
- **Fast PDF Export**: Vector-crisp multi-page PDF generation directly inside the browser.
- **100% Offline & PWA Installable**: Can be installed to desktop or mobile and run with zero external server dependencies.

## Project Structure

```text
├── index.html            # Main web application interface
├── style.css             # Light theme responsive design system
├── app.js                # Core state, Excel parser, barcode engine, PDF export
├── sw.js                 # Service worker for offline PWA caching
├── manifest.webmanifest  # PWA installation manifest
├── icons/                # App launcher icons
└── vendor/               # Bundled client-side engines (SheetJS, jsPDF, bwip-js)
```

## Quick Start

1. Double-click `index.html` to open directly in any web browser.
2. Click **Sample Data** to preview with sample warehouse locators, or drop your own `.xlsx` file.
3. Select your columns, adjust layout preset, and click **Export Single PDF**.
