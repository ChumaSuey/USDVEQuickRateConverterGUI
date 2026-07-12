# USDVEQuickRateConverterGUI

**BCV Dollar Tracker - Contabilidad**

Track the official BCV (Banco Central de Venezuela) USD and EUR exchange rates. View historical rates, export to Excel/CSV, and get a complete calendar with Smart Fill for weekends and holidays.

---

## Features

- **USD and EUR support** — toggle between currencies
- **Smart Fill calendar** — weekends and holidays get the last known rate, so every day of the month has a value (perfect for accounting)
- **Monthly and custom views** — filter by month/year or see the last 30 days
- **Metrics panel** — min, max, and percentage variation for the selected period
- **Excel export** — formatted `.xlsx` with headers, borders, and alternating fills for filled days
- **CSV export** — plain `.csv` for sheets or external tools
- **Experimental CLI** — `python mainCLI.py -c eur -d 7 --csv tasas.csv`
- **Keyboard shortcuts** — `Ctrl+R` refresh, `Ctrl+E` export, `Ctrl+C` copy rate

---

## Quick Start

### GUI

```bash
pip install -r requirements.txt   # if you have the file, or install requests + openpyxl
python mainGUI.py
```

### CLI (prototype)

```bash
python mainCLI.py -c usd -d 10
python mainCLI.py -c eur -r --csv julio.csv
```

### Standalone executable

A pre-built `.exe` is available in `dist/BCV_Dollar_Tracker.exe` (Windows, built with PyInstaller).

---

## Data Source

Exchange rates are fetched from [DolarApi](https://ve.dolarapi.com) — the official BCV rates with historical data.

---

## Project Structure

```
mainGUI.py       GUI application (Tkinter)
mainCLI.py       CLI prototype (argparse)
backend.py       Data layer: API calls, Smart Fill logic, currency formatting
requirements.txt Python dependencies
dist/            Standalone Windows executable
```

---

> This tool was originally built for my mom, who handles USD and EUR calculations in Excel for the family's bookkeeping. The goal was to speed up her workflow — one click to get a complete monthly calendar with official rates, ready to export.

---
