# Excel Reader

A lightweight Electron desktop app for viewing Excel files. No editing, no bloat — just fast, accurate display of your spreadsheets.

## Features

- **View Excel files** — Open `.xlsx`, `.xls`, `.xlsm`, and `.csv` files
- **Formula results** — Displays cached formula results, with automatic column-sum computation for uncached totals
- **Style preservation** — Bold, italic, underline, font color, cell background color, and text alignment are all rendered
- **Custom color palettes** — Correctly reads workbook-specific indexed color palettes
- **Column widths** — Preserves original column widths from the spreadsheet
- **Clickable hyperlinks** — Displays hyperlink text with clickable links that open in your default browser (supports both `HYPERLINK()` formulas and native Excel hyperlinks)
- **Multiple sheets** — Switch between sheets using tabs
- **Multiple files** — Each file opens in its own window
- **Finder integration (macOS)** — Right-click any Excel file in Finder and open it with Excel Reader
- **Number formatting** — Respects Excel number formats (commas, decimals, percentages, accounting dashes for zeros, etc.)

## Prerequisites

- [Node.js](https://nodejs.org/) (v18 or later)
- npm (included with Node.js)

## Setup

```bash
git clone <repository-url>
cd excel-reader
npm install
```

## Development

Run the app in development mode:

```bash
npm start
```

## Building & Installation

### macOS

Build a universal (Intel + Apple Silicon) app:

```bash
npm run build
```

Install to Applications:

```bash
cp -r "dist/Excel Reader-darwin-universal/Excel Reader.app" /Applications/
```

To register Excel file associations with Finder immediately:

```bash
/System/Library/Frameworks/CoreServices.framework/Frameworks/LaunchServices.framework/Support/lsregister -f "/Applications/Excel Reader.app"
```

After this, you can right-click any `.xlsx` file in Finder, select **Open With**, and choose **Excel Reader**.

### Windows

Add a Windows build script to `package.json`:

```json
{
  "scripts": {
    "build:win": "electron-packager . \"Excel Reader\" --platform=win32 --arch=x64 --out=dist --overwrite --no-asar --ignore=\"^/dist$\" --ignore=\"^/screenshots$\""
  }
}
```

Then build:

```bash
npm run build:win
```

The packaged app will be at `dist/Excel Reader-win32-x64/Excel Reader.exe`. You can move this folder anywhere and run the `.exe` directly, or create a shortcut to it.

To associate `.xlsx` files with the app on Windows:

1. Right-click any `.xlsx` file in Explorer
2. Select **Open with** > **Choose another app**
3. Click **More apps** > **Look for another app on this PC**
4. Navigate to `Excel Reader.exe` and select it
5. Check **Always use this app** if desired

## Project Structure

```
excel-reader/
  main.js          — Electron main process (file parsing, window management, IPC)
  preload.js       — Secure bridge between main and renderer processes
  renderer.js      — UI rendering (tables, tabs, hyperlinks)
  index.html       — App shell
  styles.css       — macOS-style UI styling
  extend-info.plist — macOS file type associations
  package.json     — Dependencies and build scripts
```

## Dependencies

- **[ExcelJS](https://github.com/exceljs/exceljs)** — Reads Excel files with full style support (fonts, fills, alignment, hyperlinks)
- **[SSF](https://github.com/SheetJS/ssf)** — Excel number format parser (same engine used by SheetJS)
- **[JSZip](https://stuk.github.io/jszip/)** — Reads raw xlsx zip contents for custom color palettes and hyperlink formula extraction (included as a dependency of ExcelJS)
- **[Electron](https://www.electronjs.org/)** — Cross-platform desktop app framework

## License

MIT
