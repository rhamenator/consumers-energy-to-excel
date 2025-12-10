import sys
from pathlib import Path

import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox

try:
    import pdfplumber
except ImportError:  # pragma: no cover - optional dependency warning shown later
    pdfplumber = None


COLUMN_NAMES = [
    "Current Bill Month-Year",
    "Due Date",
    "Meter Number",
    "Begin Date",
    "End Date",
    "Days Billed",
    "Begin Read",
    "End Read",
    "Read Type",
    "Energy Used",
    "Use Per Day",
    "Power Factor",
    "Billed KW",
    "Max KW",
    "Balance After Current Charges",
    "Winter Protection Payment",
    "Monthly Charge",
    "Tax",
    "Total Bill",
    "Payment Received",
    "Payment Date",
    "Late Charge",
    "Adjustment Date",
    "Last Month Balance",
    "Balance Before Current Charges",
]


def main() -> None:
    root = tk.Tk()
    root.withdraw()

    file_path = filedialog.askopenfilename(
        title="Select Itemized Statement (PDF or XLSX)",
        filetypes=(
            ("PDF and XLSX files", "*.pdf *.xlsx"),
            ("PDF files", "*.pdf"),
            ("XLSX files", "*.xlsx"),
            ("All files", "*.*"),
        ),
    )

    if not file_path:
        return

    try:
        itemized_statement_df = load_itemized_statement(Path(file_path))
    except Exception as exc:  # pragma: no cover - surfaced to the user
        messagebox.showerror("Error", f"Unable to load statement: {exc}")
        return

    mapped_df = map_monthly_data(itemized_statement_df)
    mapped_df.to_excel("output.xlsx", index=False)
    messagebox.showinfo("Success", "Mapped data saved to output.xlsx")


def load_itemized_statement(file_path: Path) -> pd.DataFrame:
    extension = file_path.suffix.lower()
    if extension == ".pdf":
        if pdfplumber is None:
            raise ImportError("pdfplumber is required to process PDF statements.")
        raw_df = read_pdf_tables(file_path)
    elif extension in {".xlsx", ".xls"}:
        raw_df = pd.read_excel(file_path, header=None)
    else:
        raise ValueError("Please select a PDF or XLSX file.")

    cleaned_df = clean_itemized_statement(raw_df)
    if "Description" not in cleaned_df.columns:
        raise ValueError("Unable to find a Description column in the statement.")
    cleaned_df["Description"] = (
        cleaned_df["Description"].astype(str).str.replace("\n", " ")
    )
    if "Use Per Day" not in cleaned_df.columns and "Use Per" in cleaned_df.columns:
        cleaned_df = cleaned_df.rename(columns={"Use Per": "Use Per Day"})
    elif "Use Per Day" not in cleaned_df.columns:
        cleaned_df["Use Per Day"] = None

    return cleaned_df


def read_pdf_tables(pdf_path: Path) -> pd.DataFrame:
    frames = []
    with pdfplumber.open(pdf_path) as pdf:
        for page in pdf.pages:
            tables = page.extract_tables() or []
            for table in tables:
                df = pd.DataFrame(table)
                df = _standardize_cells(df)
                if not _table_contains_key(df, "Description"):
                    continue
                frames.append(df)

    if not frames:
        raise ValueError("No recognizable tables were found in the PDF.")

    combined = pd.concat(frames, ignore_index=True)
    combined = combined.replace({"" : pd.NA})
    combined = combined.dropna(how="all")
    return combined


def _standardize_cells(df: pd.DataFrame) -> pd.DataFrame:
    df = df.replace({"\u00a0": " "}, regex=True)
    df = df.applymap(lambda value: value.strip() if isinstance(value, str) else value)
    return df


def _table_contains_key(df: pd.DataFrame, keyword: str) -> bool:
    keyword_lower = keyword.lower()
    for _, row in df.iterrows():
        for value in row:
            if isinstance(value, str) and keyword_lower in value.lower():
                return True
    return False


def clean_itemized_statement(raw_df: pd.DataFrame) -> pd.DataFrame:
    sanitized = raw_df.copy()
    sanitized = sanitized.replace({"\u00a0": " "})
    sanitized = sanitized.dropna(how="all")
    header_row_idx = find_header_row_index(sanitized)
    header_row = (
        sanitized.iloc[header_row_idx]
        .astype(str)
        .str.replace("\n", " ")
        .str.strip()
        .tolist()
    )
    data = sanitized.iloc[header_row_idx + 1 :].copy()
    data.columns = header_row
    data = data.dropna(how="all")
    data = data.reset_index(drop=True)
    return data


def find_header_row_index(df: pd.DataFrame) -> int:
    for idx, row in df.iterrows():
        values = [str(value).strip().lower() for value in row.tolist() if value is not None]
        if "description" in values and "monthly charge" in values:
            return idx
    return 0


def map_monthly_data(itemized_statement_df: pd.DataFrame) -> pd.DataFrame:
    mapped_df = pd.DataFrame(columns=COLUMN_NAMES)
    for i in range(0, len(itemized_statement_df), 12):
        month_data = itemized_statement_df.iloc[i : i + 12]
        if month_data.empty:
            continue
        mapped_row = map_single_month(month_data)
        if mapped_row:
            mapped_df = pd.concat([mapped_df, pd.DataFrame([mapped_row])], ignore_index=True)
    return mapped_df


def map_single_month(month_data: pd.DataFrame) -> dict | None:
    mapped_data: dict[str, object] = {}
    try:
        begin_date_row = month_data[month_data["Description"].str.contains("Gas", na=False)].iloc[0]
        mapped_data["Current Bill Month-Year"] = pd.to_datetime(
            begin_date_row["Begin Date"]
        ).strftime("%Y-%m-%d")
        mapped_data["Due Date"] = pd.to_datetime(
            month_data[month_data["Description"].str.contains("due on or before")]
            .iloc[0]["Description"].split()[-1],
            format="%m/%d/%Y",
        ).strftime("%Y-%m-%d")
        mapped_data["Meter Number"] = begin_date_row.get("Meter Number")
        mapped_data["Begin Date"] = begin_date_row.get("Begin Date")
        mapped_data["End Date"] = begin_date_row.get("End Date")
        mapped_data["Days Billed"] = begin_date_row.get("Days Billed")
        mapped_data["Begin Read"] = begin_date_row.get("Begin Read")
        mapped_data["End Read"] = begin_date_row.get("End Read")
        mapped_data["Read Type"] = begin_date_row.get("Read Type")
        mapped_data["Energy Used"] = begin_date_row.get("Energy Used")
        mapped_data["Use Per Day"] = begin_date_row.get("Use Per Day")
        mapped_data["Power Factor"] = None
        mapped_data["Billed KW"] = None
        mapped_data["Max KW"] = None
        mapped_data["Balance After Current Charges"] = _first_value(
            month_data,
            "BALANCE AFTER CURRENT CHARGES",
            "Monthly Charge",
        )
        mapped_data["Winter Protection Payment"] = _first_value(
            month_data,
            "WINTER PROTECTION PLAN PAYMENT",
            "Monthly Charge",
        )
        mapped_data["Monthly Charge"] = begin_date_row.get("Monthly Charge")
        mapped_data["Tax"] = _exact_value(month_data, "Tax", "Monthly Charge")
        mapped_data["Total Bill"] = _first_value(
            month_data,
            "due on or before",
            "Transaction Amt",
        )
        mapped_data["Payment Received"] = _first_value(
            month_data,
            "Payment Received",
            "Transaction Amt",
        )
        mapped_data["Payment Date"] = _first_value(
            month_data,
            "Payment Received",
            "Date",
        )
        mapped_data["Late Charge"] = _first_value(
            month_data,
            "LATE PAYMENT CHARGE",
            "Transaction Amt",
        )
        canceled_row = _first_row(month_data, "Canceled Late Payment Charge")
        if canceled_row is not None:
            mapped_data["Late Charge"] = canceled_row.get("Transaction Amt")
            mapped_data["Adjustment Date"] = canceled_row.get("Date")
        mapped_data["Last Month Balance"] = _first_value(
            month_data,
            "Last Month's Consumers Energy Account Balance",
            "Account Balance",
        )
        mapped_data["Balance Before Current Charges"] = _first_value(
            month_data,
            "Account Balance Before Current Charges",
            "Account Balance",
        )
        return mapped_data
    except Exception as exc:  # pragma: no cover - log and skip malformed month
        print(f"Skipping a month due to error: {exc}", file=sys.stderr)
        return None


def _first_row(df: pd.DataFrame, pattern: str):
    matching = df[df["Description"].str.contains(pattern, na=False)]
    return matching.iloc[0] if not matching.empty else None


def _first_value(df: pd.DataFrame, pattern: str, column: str):
    row = _first_row(df, pattern)
    return row.get(column) if row is not None and column in row else None


def _exact_value(df: pd.DataFrame, exact_description: str, column: str):
    matching = df[df["Description"] == exact_description]
    if matching.empty:
        return None
    row = matching.iloc[0]
    return row.get(column)


if __name__ == "__main__":
    main()