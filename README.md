# Expense Extract

**Pulls a Profit & Loss report straight out of QuickBooks Desktop and turns it into a clean, formula-driven Excel expense report in one click.**

## Background

Built in 2020 to automate bookkeeping at a multi-location metals recycling company. Each month someone had to run a Profit & Loss report in every QuickBooks company file, export it, strip out sales, cost of goods sold and other income, and re-format what was left for review. Expense Extract does that end to end: pick a date range and a company file, and it writes a ready-to-review workbook.

It was one piece of a broader automation suite Weston built for the company's back office, which together saved roughly $100K–$120K a year in labor. Its output feeds the companion tool [Expense-RG](https://github.com/Weston-Matthews/Expense-RG), which combines the per-location reports into a year-over-year comparison.

## Features

- Talks to QuickBooks Desktop directly through Intuit's **QBFC13 SDK** (COM), so there's no manual export or copy-paste
- Requests a **Profit & Loss Standard** report for any date range, summarized **by month**
- Keeps only the **Expense** section (from "Expense" through "Total Expense") and drops income and COGS
- Rebuilds QuickBooks' nested account hierarchy in Excel with matching indentation
- Writes **live Excel formulas** for subtotals and totals (`=ROUND(SUM(...),5)` / `=SUM(...)`) instead of pasted numbers, so the workbook stays correct when someone edits a value
- QuickBooks-style formatting: Arial 8 pt, accounting number format, bold headers, borders on subtotal rows, and a page header with run time, date and report basis (Cash/Accrual)
- Simple GUI prompts (date range, company file picker) with input validation and readable error messages
- Ships as a stand-alone Windows `.exe` via PyInstaller

## How it works

1. **Input.** An `easygui` dialog asks for a start and end date (`MM/DD/YY` or `MM/DD/YYYY`), validated by `check_date()`, then a file picker selects the QuickBooks company file (`.qbw`).
2. **Query QuickBooks.** The script opens a `QBFC13.QBSessionManager` COM session under its own AppID, sends a `GeneralSummaryReportQuery` for `ProfitAndLossStandard` with `SummarizeColumnsBy = Month`, then closes the session.
3. **Flatten the response.** Each report row is classified as a header (`h`), data row (`r`), subtotal (`s`) or total (`t`) and stored with its column values.
4. **Keep only expenses.** `extract_expense()` slices the rows between the `Expense` and `Total Expense` markers.
5. **Lay out the sheet.** `calc_ws_col()` walks the row types to work out how deep the account hierarchy goes, which decides how many indent columns the sheet needs. Rows are then written with the right indentation, values go into every other column, and subtotal and total rows get SUM formulas that point at the cells above them.
6. **Save.** The workbook is written to the working directory as `<company file> Exp Report - <start> to <end>.xlsx`.

## Architecture

A single script, `Expense_Extract.py`:

| Function | Role |
|---|---|
| `check_date(date)` | Parses and validates the user's date text into a `datetime` |
| `calc_ws_col(rowTypes)` | Computes the indentation range of the account hierarchy from the row-type sequence |
| `extract_expense(rowData)` | Returns only the Expense section of the P&L rows, plus its start and end indexes |
| `main()` | GUI prompts → QBFC session and report query → row flattening → openpyxl layout, formulas and styling → save |

## Tech stack

Python 3.8 · pywin32 (COM) · Intuit QuickBooks SDK / QBFC13 · openpyxl · easygui · Poetry · PyInstaller

## Requirements

- Windows (the QuickBooks SDK is COM-based)
- QuickBooks Desktop (US edition) with access to the company file
- Intuit QuickBooks SDK 13 (QBFC13) installed
- Python 3.8
- Poetry

## Install

```bat
git clone https://github.com/Weston-Matthews/Expense-Extract.git
cd Expense-Extract
poetry install
```

Without Poetry: `pip install -r requirements.txt`

## Run

```bat
poetry run python Expense_Extract.py
```

The first time it connects to a company file, QuickBooks asks you to authorize the application ("Expense Report Extractor v1.0"). Approve it in QuickBooks. First-time authorization needs QuickBooks open and signed in as an admin.

## Build (stand-alone .exe)

```bat
poetry run pyinstaller -y Expense_Extract.spec
```

The executable is written to `dist\Expense_Extract\Expense_Extract.exe`. The spec file uses only paths relative to the repository, so it builds from any checkout location.

## Configuration

There's no configuration file. Everything is chosen at run time: the date range and the company file. The QuickBooks AppID and AppName are constants at the top of `Expense_Extract.py`.

## Sample output

For a company file named `Sample Co.qbw` and the range 01/01/2020–09/30/2020, the tool writes `Sample Co.qbw Exp Report - 1-1-2020 to 9-30-2020.xlsx` with one sheet, **Report**:

```
                                     Jan 20     Feb 20     Mar 20   ...    Sep 20      TOTAL
Expense
  Payroll Expenses
    Wages                          12,450.00  11,980.00  12,730.00  ...  12,105.00  110,882.00
    Payroll Taxes                   1,120.50   1,078.20   1,145.70  ...   1,089.45    9,979.38
  Total Payroll Expenses           13,570.50  13,058.20  13,875.70  ...  13,194.45  120,861.38   <- =ROUND(SUM(...),5)
  Rent Expense                      4,000.00   4,000.00   4,000.00  ...   4,000.00   36,000.00
  Utilities                           812.34     790.10     655.42  ...     702.88    6,701.25
Total Expense                      18,382.84  17,848.30  18,531.12  ...  17,897.33  163,562.63   <- =SUM(...)
```

(Fictional numbers, for illustration.) The page header shows the run time, the date and the report basis, with "Expense Report" centered.

## Limitations

- Windows only. It needs QuickBooks Desktop and the QBFC13 SDK; QuickBooks Online isn't supported.
- Handles one company file per run.
- Uses the US edition and qbXML 7.0 message set (`CreateMsgSetRequest("US", 7, 0)`).
- Warning and Error responses from QuickBooks are printed to the console but not shown in a dialog.
- The output path is built Windows-style and saved to the current working directory.
- Written in 2020 for Python 3.8 and PyInstaller 3.6. It hasn't been updated for newer versions.

## Author

**Weston Matthews** · [github.com/Weston-Matthews](https://github.com/Weston-Matthews)

## License

TBD
