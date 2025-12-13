from __future__ import annotations

from datetime import date
from pathlib import Path
from unittest.mock import MagicMock

import pandas as pd

import process_consumers_itemized_statement as pcs


def test_get_input_path_from_argv_empty_returns_none():
    assert pcs.get_input_path_from_argv([]) is None


def test_get_input_path_from_argv_returns_existing_path(tmp_path: Path):
    statement = tmp_path / "statement.pdf"
    statement.write_bytes(b"%PDF-1.4\n%fake\n")
    assert pcs.get_input_path_from_argv([str(statement)]) == statement


def test_get_input_path_from_argv_missing_path_raises(tmp_path: Path):
    missing = tmp_path / "missing.pdf"
    try:
        pcs.get_input_path_from_argv([str(missing)])
        assert False, "Expected FileNotFoundError"
    except FileNotFoundError:
        assert True


def test_normalize_cell_trailing_minus_accounting():
    value, fmt = pcs.normalize_cell("$205.24-")
    assert value == -205.24
    assert fmt is not None
    assert fmt["type"] == "accounting"
    assert fmt["decimals"] == 2


def test_normalize_cell_date_parsing():
    value, fmt = pcs.normalize_cell("01/12/2023")
    assert value == date(2023, 1, 12)
    assert fmt is not None
    assert fmt["type"] == "date"


def test_split_into_months_uses_total_amount_due_boundary():
    df = pd.DataFrame(
        {
            "Description": [
                "Gas",
                "TotalAmountDue",
                "Gas",
                "Total Amount Due",
            ],
            "Meter Number": ["123", None, "456", None],
        }
    )
    months = pcs.split_into_months(df)
    assert len(months) == 2
    assert months[0].iloc[-1]["Description"].lower().replace(" ", "") == "totalamountdue"
    assert months[1].iloc[-1]["Description"].lower().replace(" ", "") == "totalamountdue"


def test_map_single_month_minimal_pdf_like():
    month_df = pd.DataFrame(
        {
            "Description": [
                "Gas",
                "Tax",
                "TotalCurrentBilldueonorbefore 02/07/2023",
                "TotalAmountDue",
            ],
            "Meter Number": ["95033411", None, None, None],
            "Begin Date": ["12/10/2022", None, None, None],
            "End Date": ["01/10/2023", None, None, None],
            "Days Billed": ["32", None, None, None],
            "Begin Read": ["2293", None, None, None],
            "End Read": ["2348", None, None, None],
            "Read Type": ["ACTUAL", None, None, None],
            "Energy Used": ["5.5", None, None, None],
            "Use Per Day": ["0.172", None, None, None],
            "Monthly Charge": ["$76.74", "$3.07", "$79.81", None],
            "Transaction Amt": [None, None, None, None],
            "Date": [None, None, None, None],
            "Account Balance": [None, None, None, "$82.80"],
        }
    )

    mapped = pcs.map_single_month(month_df)
    assert mapped is not None
    assert mapped["Meter Number"] == "95033411"
    assert mapped["Due Date"] == "2023-02-07"
    assert mapped["Total Bill"] == "$82.80"


def test_read_pdf_tables_reuses_header_across_tables(tmp_path: Path, monkeypatch):
    # Simulate a PDF with one header-only table followed by a data table.
    header = [
        "Description",
        "Meter\nNumber",
        "Begin\nDate",
        "End\nDate",
        "Days\nBilled",
        "Begin\nRead",
        "End\nRead",
        "Read\nType",
        "Energy\nUsed",
        "Use Per\nDay",
        "Power\nFactor",
        "Billed\nKW",
        "Max\nKW",
        "Monthly\nCharge",
        "Transaction\nAmt",
        "Date",
        "Account\nBalance",
    ]
    data_row = [
        "Gas",
        "95033411",
        "12/10/2022",
        "01/10/2023",
        "32",
        "2293",
        "2348",
        "ACTUAL",
        "5.5",
        "0.172",
        "",
        "",
        "",
        "$76.74",
        "",
        "",
        "",
    ]

    fake_page = MagicMock()
    fake_page.extract_tables.return_value = [
        [header],
        [data_row],
    ]
    fake_pdf = MagicMock()
    fake_pdf.pages = [fake_page]

    class FakeContext:
        def __enter__(self):
            return fake_pdf

        def __exit__(self, exc_type, exc, tb):
            return False

    monkeypatch.setattr(pcs, "pdfplumber", MagicMock(open=lambda _: FakeContext()))

    df = pcs.read_pdf_tables(tmp_path / "fake.pdf")
    assert len(df) == 1
    assert "Description" in df.columns
    assert df.iloc[0]["Description"] == "Gas"
    assert df.iloc[0]["Meter Number"] == "95033411"
