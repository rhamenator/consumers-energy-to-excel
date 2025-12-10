# Consumers Energy Itemized Statement Processor

This program processes Consumers Energy Itemized Statement bills and converts them into a clean, tabular Excel file suitable for analysis.

It supports both:

- Original PDF statements downloaded from Consumers Energy.
- Existing cleaned Excel (`.xlsx`) files.

The script extracts billing information (dates, meter readings, charges, balances, etc.) and writes the normalized results to both `output.xlsx` (styled) and `output.csv` (raw data).

## What the Program Does

- Opens a file picker dialog so you can choose either a PDF or an Excel file.
- Reads the file into a `pandas` `DataFrame`:
  - PDF files are parsed using Python PDF utilities (via `pdfplumber`) to reconstruct the table layout.
  - Excel files are loaded directly with `pandas.read_excel`.
- Normalizes column names such as:
  - `Current Bill Month-Year`, `Due Date`, `Meter Number`
  - `Begin Date`, `End Date`, `Days Billed`, `Begin Read`, `End Read`, `Read Type`
  - `Energy Used`, `Use Per Day`
  - `Monthly Charge`, `Tax`, `Total Bill`
  - `Payment Received`, `Payment Date`, `Late Charge`, `Adjustment Date`
  - `Last Month Balance`, `Balance Before Current Charges`
- Converts dollars/negatives into proper accounting/number formats and parses date-like strings into real dates.
- Groups each billing segment into a single monthly summary row.
- Outputs a styled Excel workbook with alternating greenbar stripes, frozen header row, auto-fit columns, and a companion `output.csv` for downstream imports.

Sample files included in this folder (`Consumers Energy Itemized Statement 2024-06-14.pdf` and `.xlsx`) demonstrate the one-to-one correspondence between the PDF and Excel formats the script expects.

## Inputs and Outputs

### Inputs

- Consumers Energy Itemized Statement PDF (`*.pdf`).
- Consumers Energy Itemized Statement Excel (`*.xlsx`).

### Output

- `output.xlsx` — normalized, formatted Excel file (greenbar styling, frozen header, number/date formats) with one row per billing month.
- `output.csv` — the same normalized data with no styling for easy importing elsewhere.

## Dependencies

Key actively maintained libraries used by this script:

- `pandas` — data wrangling and Excel I/O.
- `openpyxl` — Excel engine used under the hood by `pandas`.
- `pdfplumber` — PDF parsing used to extract tables from the Consumers Energy statements.
- `tkinter` — built-in GUI toolkit for Windows that provides the file picker dialog.

If any PDF-specific dependency were to become unmaintained in the future, it can be isolated and replaced without altering the downstream data-mapping logic.

## Setup and Usage

### 1. Optional: create and activate a virtual environment

```pwsh
python -m venv .venv
.\.venv\Scripts\Activate.ps1
```

### 2. Install dependencies

```pwsh
pip install -r requirements.txt
```

### 3. Run the script

```pwsh
python process_consumers_itemized_statement.py
```

A file dialog appears. Choose either a PDF or an Excel statement. Once processing finishes, `output.xlsx` and `output.csv` will be created in the same directory.

## Notes and Future Extensions

- The PDF parsing logic is tuned to Consumers Energy Itemized Statements. Major layout changes in future statements may require adjustments.
- Keeping the Excel import path allows quick comparisons or fallbacks if a particular PDF variation does not parse cleanly.
- If more complex processing or services are required later, the PDF-handling portion can be moved into its own module or service while preserving the existing mapping code.
