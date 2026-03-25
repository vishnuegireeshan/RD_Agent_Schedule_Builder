# Post Office RD Schedule Automation System

A desktop application for India Post Recurring Deposit (RD) agents to manage account collections, generate optimized deposit lists, and export professional reports.

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Tech Stack](#tech-stack)
- [Installation](#installation)
- [Usage](#usage)
- [Application Windows](#application-windows)
- [Data Import Formats](#data-import-formats)
- [Export & Output](#export--output)
- [Architecture](#architecture)
- [Diagnostic Tools](#diagnostic-tools)
- [Configuration](#configuration)

---

## Overview

RD agents at post offices manage hundreds of recurring deposit accounts. Each month they must:

1. Collect installments from depositors
2. Group collections into deposit batches (max Rs 20,000 each)
3. Submit deposit lists to the post office

This application automates the entire workflow — from loading account data (PDF/Excel), tracking collections, generating optimized deposit batches, to exporting formatted Excel reports ready for submission.

---

## Features

### Data Import
- **Excel** (`.xlsx`, `.xls`) and **PDF** table extraction
- Flexible column name matching (case-insensitive, multiple aliases)
- Handles merged PDF rows at page boundaries via regex fallback
- Denomination parsing: `2,000.00 Cr.`, `Rs 1,500`, `1000`
- Multiple date formats: `DD-Mon-YYYY`, `YYYY-MM-DD`, `DD/MM/YYYY`
- Duplicate detection and data validation with warnings

### Account Management
- Auto-classification: **ACTIVE** / **MATURED** (60 months) / **DEFAULTED**
- Priority: **HIGH** (due date day <= 15) / **NORMAL**
- Months missed calculation for overdue accounts
- Individual and bulk payment marking

### Collection Tracking
- Mark/unmark accounts as collected (money with agent)
- Checkbox-based multi-select for batch operations
- Pay times (1–60) for multiple installment collection per account

### Deposit List Generation
- Greedy knapsack algorithm: packs accounts into batches up to configurable max (default Rs 20,000)
- HIGH priority accounts are placed first
- Sorted by denomination (descending) within priority groups
- Handles 1000+ accounts efficiently

### Search, Filter & Sort
- Real-time search by name or account number
- Status filter: All / Defaulted Only / Non-defaulted
- Due date filter: All / Due <= 15 This Month / Hide Due <= 15
- **Click any column header** to sort ascending (▲) / descending (▼)
- Smart sort: numeric for amounts/months, chronological for dates, alphabetical for text

### Reports & Export
- Multi-sheet Excel workbook with professional formatting
- Text file with comma-separated account numbers per list
- Timestamped filenames

---

## Tech Stack

| Component       | Technology                           |
|-----------------|--------------------------------------|
| Language        | Python 3.8+                          |
| GUI             | Tkinter / ttk                        |
| PDF Extraction  | pdfplumber                           |
| Data Processing | pandas                               |
| Excel Export    | openpyxl                             |
| Date Parsing    | python-dateutil                      |
| Packaging       | PyInstaller (standalone `.exe`)      |

---

## Installation

### From Source

```bash
cd D:\Vishnnu\RD_System\files

# Create and activate virtual environment
python -m venv .venv
.venv\Scripts\activate

# Install dependencies
pip install pdfplumber pandas openpyxl python-dateutil

# Launch GUI
python rd_schedule_system.py

# Or launch CLI mode
python rd_schedule_system.py --cli
```

### Build Standalone Executable

```bash
pip install pyinstaller
pyinstaller rd_schedule_system.spec
# Output: dist/rd_schedule_system.exe
```

---

## Usage

### Typical Workflow

```
1. Load Data      →  Click "Load Excel/PDF" → select account file
2. Review         →  Dashboard shows totals, status breakdown
3. Collect        →  Open "Collected Window" → check accounts → Mark Collected
4. Set Pay Times  →  Open "View Collected Accounts" → set repeat count (1–60)
5. Generate Lists →  Click "Generate" → optimized batches appear in preview
6. Export         →  Click "Export Excel" → save .xlsx report + .txt listing
```

### CLI Mode

```
Main Menu:
  1. Dashboard (Overview)
  2. View Payment Status
  3. Mark Payments (Paid/Unpaid)
  4. Generate Deposit List
  5. Export to Excel
  6. Search Accounts
  7. Exit
```

---

## Application Windows

### Main Window (1200 x 760)

| Section              | Description                                                  |
|----------------------|--------------------------------------------------------------|
| Toolbar              | Load file, open collected window, export buttons             |
| Search Bar           | Real-time filter by name/account number                      |
| Statistics Dashboard | Total, Active, Matured, Collected, Unpaid, Defaulters, etc.  |
| Accounts Table       | Color-coded rows — green (paid), red (defaulted), purple (matured) |
| Deposit List Preview | Grouped lists with subtotals and grand total                 |

### Collected by Agent Window (900 x 560)

Manage accounts where money has been collected but not yet deposited.

| Column        | Sortable | Description                          |
|---------------|----------|--------------------------------------|
| Select        | No       | Checkbox `[x]` / `[ ]`              |
| Account No    | Yes      | Unique account identifier            |
| Name          | Yes      | Depositor name                       |
| Amount        | Yes      | Monthly denomination (numeric sort)  |
| Months Missed | Yes      | Overdue installment count            |
| Due Date      | Yes      | Next due date (chronological sort)   |
| Status        | Yes      | ACTIVE / DEFAULTED                   |
| Collected?    | Yes      | YES / NO                             |

**Filters:** Status (All / Defaulted / Non-defaulted) · Due Date (All / Due <= 15 / Hide Due <= 15)

### Collected Accounts Window (920 x 560)

View collected accounts and configure pay-times multiplier.

| Column      | Sortable | Description                  |
|-------------|----------|------------------------------|
| Account No  | Yes      | Account identifier           |
| Name        | Yes      | Depositor name               |
| Amount      | Yes      | Denomination (numeric sort)  |
| Months Paid | Yes      | Installments completed       |
| Priority    | Yes      | HIGH / NORMAL                |
| Due Date    | Yes      | Next due (chronological)     |
| Pay Times   | Yes      | Repeat count 1–60            |

---

## Data Import Formats

### Supported Column Names

The system uses flexible matching. Any of these aliases work:

| Field        | Accepted Column Names                                                     |
|--------------|---------------------------------------------------------------------------|
| Account No   | `account no`, `account_no`, `acc no`, `account number`, `acct no`         |
| Name         | `account name`, `customer name`, `name`, `depositor`, `holder name`       |
| Denomination | `denomination`, `amount`, `monthly amount`, `deposit amount`, `instalment`|
| Months Paid  | `month paid upto`, `months paid`, `installments`, `paid months`           |
| Due Date     | `next rd installment due date`, `due date`, `next due date`               |

### PDF Handling

- Extracts tables from multi-page PDFs using `pdfplumber`
- Detects and skips repeated header rows across pages
- **Merged row recovery:** When rows at page boundaries collapse into a single string, a regex parser splits them back into proper columns by anchoring on the denomination pattern (`X,XXX.XX Cr.`) and date pattern (`DD-Mon-YYYY`)

---

## Export & Output

### Excel Report (`RD_Deposit_Report_YYYYMMDD.xlsx`)

| Sheet            | Contents                                                        |
|------------------|-----------------------------------------------------------------|
| Deposit Lists    | Grouped batches with subtotals, priority highlighting, grand total |
| Payment Summary  | Statistics table — totals, active, matured, defaulters, amount  |
| All Accounts     | Complete reference listing with status color coding             |

**Formatting:** Dark blue headers, currency formatting (Rs #,##0), color-coded status rows, thin borders, auto-width columns.

### Account Number Listing (`..._listwise_account_numbers.txt`)

```
List 1: 1234,1235,1236
List 2: 12347,12348
```

---

## Architecture

```
┌──────────────────┐
│  Excel / PDF     │   Input
└────────┬─────────┘
         ▼
┌──────────────────┐
│   DataLoader     │   Parse, validate, create RDAccount objects
└────────┬─────────┘
         ▼
┌──────────────────┐
│ AccountManager   │   Store accounts, track payments, compute stats
└────────┬─────────┘
         ▼
┌──────────────────┐
│ DepositList      │   Greedy knapsack: batch accounts ≤ max amount
│ Generator        │
└────────┬─────────┘
         ▼
┌──────────────────┐
│ ExcelExporter    │   Multi-sheet workbook + text listing
└──────────────────┘
```

### Key Classes

| Class                  | Responsibility                                         |
|------------------------|--------------------------------------------------------|
| `RDAccount`            | Domain entity — account data, status, priority         |
| `DataLoader`           | Parse Excel/PDF into `RDAccount` list                  |
| `AccountManager`       | Business logic — filtering, payments, statistics       |
| `DepositList`          | Single batch container with capacity tracking          |
| `DepositListGenerator` | Optimization algorithm for batch creation              |
| `ExcelExporter`        | Formatted `.xlsx` report generation                    |
| `GUIInterface`         | Tkinter desktop application                            |
| `CLIInterface`         | Command-line interface                                 |

---

## Diagnostic Tools

### PDF Debug Script

```bash
python pdf_debug.py "path/to/accounts.pdf"
```

**Output:**
1. Raw table extraction per page (marks merged rows with `** MERGED **`)
2. Parsed accounts via `load_from_pdf` with all fields
3. Warnings for any parsing issues

Useful for troubleshooting when accounts are missing or columns misaligned.

---

## Configuration

| Setting                     | Default    | Notes                              |
|-----------------------------|------------|------------------------------------|
| Max amount per deposit list | Rs 20,000  | Configurable in Generate dialog    |
| RD maturity period          | 60 months  | Standard India Post RD term        |
| High priority threshold     | Day <= 15  | Based on due date day of month     |
| Pay times range             | 1–60       | Multiple installments per cycle    |
| Window size (main)          | 1200 x 760 | Minimum 980 x 620                 |

---

## File Structure

```
rd_schedule_system.py      Main application (single-file, all classes)
rd_schedule_system.spec    PyInstaller build configuration
pdf_debug.py               PDF diagnostic tool
.venv/                     Python virtual environment
dist/                      Compiled executable (after build)
```

---

*Built for India Post RD agents. Offline-only, no network or database required.*
