import pandas as pd
import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# Define the column names for the output DataFrame
column_names = [
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

# Create the main window
root = tk.Tk()
root.withdraw()  # Hide the main window

# Open the file picker dialog
file_path = filedialog.askopenfilename(
    title="Select Itemized Statement XLSX File",
    filetypes=(("XLSX files", "*.xlsx"), ("all files", "*.*")),
)

# Check if the user canceled the file picker dialog
if not file_path:
    exit()

try:
    # Load the itemized statement data directly from the XLSX file
    itemized_statement_df = pd.read_excel(file_path)
except FileNotFoundError:
    messagebox.showerror("Error", "File not found.")
    exit()
except Exception as e:
    messagebox.showerror("Error", f"An error occurred: {e}")
    exit()

# Assign the values in the first row as the header
itemized_statement_df.columns = itemized_statement_df.iloc[0].values

# Drop the first row
itemized_statement_df = itemized_statement_df.iloc[1:].copy()

# Rename the column
itemized_statement_df = itemized_statement_df.rename(
    columns={"Use Per": "Use Per Day"}
)

# Remove the newline character from the `Description` column
itemized_statement_df["Description"] = (
    itemized_statement_df["Description"].astype(str).str.replace("\n", " ")
)

# Create a new DataFrame for mapped data with the defined column names
mapped_df = pd.DataFrame(columns=column_names)

# Extract and map the data
for i in range(0, len(itemized_statement_df), 12):
    month_data = itemized_statement_df.iloc[i : i + 12]
    mapped_data = {}
    try:
        # Find the row containing the begin date based on the description containing 'Gas'
        begin_date_row = month_data[
            month_data["Description"].str.contains("Gas", na=False)
        ].iloc[0]
        mapped_data["Current Bill Month-Year"] = pd.to_datetime(
            begin_date_row["Begin Date"]
        ).strftime("%Y-%m-%d")
        mapped_data["Due Date"] = pd.to_datetime(
            month_data[month_data["Description"].str.contains("due on or before")].iloc[
                0
            ]["Description"].split()[-1],
            format="%m/%d/%Y",
        ).strftime("%Y-%m-%d")
        mapped_data["Meter Number"] = begin_date_row["Meter Number"]
        mapped_data["Begin Date"] = begin_date_row["Begin Date"]
        mapped_data["End Date"] = begin_date_row["End Date"]
        mapped_data["Days Billed"] = begin_date_row["Days Billed"]
        mapped_data["Begin Read"] = begin_date_row["Begin Read"]
        mapped_data["End Read"] = begin_date_row["End Read"]
        mapped_data["Read Type"] = begin_date_row["Read Type"]
        mapped_data["Energy Used"] = begin_date_row["Energy Used"]
        mapped_data["Use Per Day"] = begin_date_row["Use Per Day"]
        mapped_data["Power Factor"] = None
        mapped_data["Billed KW"] = None
        mapped_data["Max KW"] = None
        mapped_data["Balance After Current Charges"] = month_data[
            month_data["Description"].str.contains("BALANCE AFTER CURRENT CHARGES")
        ].iloc[0]["Monthly Charge"] if len(
            month_data[
                month_data["Description"].str.contains("BALANCE AFTER CURRENT CHARGES")
            ]
        ) > 0 else None
        mapped_data["Winter Protection Payment"] = month_data[
            month_data["Description"].str.contains("WINTER PROTECTION PLAN PAYMENT")
        ].iloc[0]["Monthly Charge"] if len(
            month_data[
                month_data["Description"].str.contains(
                    "WINTER PROTECTION PLAN PAYMENT"
                )
            ]
        ) > 0 else None
        mapped_data["Monthly Charge"] = begin_date_row["Monthly Charge"]
        mapped_data["Tax"] = month_data[month_data["Description"] == "Tax"].iloc[0][
            "Monthly Charge"
        ]
        mapped_data["Total Bill"] = month_data[
            month_data["Description"].str.contains("due on or before")
        ].iloc[0]["Transaction Amt"]
        mapped_data["Payment Received"] = month_data[
            month_data["Description"].str.contains("Payment Received")
        ].iloc[0]["Transaction Amt"] if len(
            month_data[month_data["Description"].str.contains("Payment Received")]
        ) > 0 else None
        mapped_data["Payment Date"] = month_data[
            month_data["Description"].str.contains("Payment Received")
        ].iloc[0]["Date"] if len(
            month_data[month_data["Description"].str.contains("Payment Received")]
        ) > 0 else None

        # Extract Late Charge and Adjustment Date
        late_charge_row = month_data[
            month_data["Description"].str.contains("LATE PAYMENT CHARGE")
        ]
        if not late_charge_row.empty:
            mapped_data["Late Charge"] = late_charge_row.iloc[0]["Transaction Amt"]

        canceled_late_charge_row = month_data[
            month_data["Description"].str.contains("Canceled Late Payment Charge")
        ]
        if not canceled_late_charge_row.empty:
            mapped_data["Late Charge"] = canceled_late_charge_row.iloc[0][
                "Transaction Amt"
            ]
            mapped_data["Adjustment Date"] = canceled_late_charge_row.iloc[0]["Date"]

        mapped_data["Last Month Balance"] = month_data[
            month_data["Description"].str.contains(
                "Last Month's Consumers Energy Account Balance"
            )
        ].iloc[0]["Account Balance"] if len(
            month_data[
                month_data["Description"].str.contains(
                    "Last Month's Consumers Energy Account Balance"
                )
            ]
        ) > 0 else None
        mapped_data["Balance Before Current Charges"] = month_data[
            month_data["Description"].str.contains(
                "Account Balance Before Current Charges"
            )
        ].iloc[0]["Account Balance"] if len(
            month_data[
                month_data["Description"].str.contains(
                    "Account Balance Before Current Charges"
                )
            ]
        ) > 0 else None

        # Filter out empty or all-NA columns before concatenation
        mapped_data_df = pd.DataFrame([mapped_data])
        mapped_data_df = mapped_data_df.dropna(axis=1, how="all")
        mapped_df = pd.concat([mapped_df, mapped_data_df], ignore_index=True)

    except Exception as e:
        print(f"An error occurred while processing month data: {e}")

# Save the mapped data to an Excel file
mapped_df.to_excel('output.xlsx', index=False)