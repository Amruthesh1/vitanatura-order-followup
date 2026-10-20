"""
Vitanatura Assessment — Step 3: Order Dataset Validation
Run:
    python validate_orders.py vitanatura_sample_orders.xlsx
    python validate_orders.py vitanatura_sample_orders.csv

Optional fixed evaluation timestamp for reproducible testing:
    python validate_orders.py vitanatura_sample_orders.xlsx --as-of "2026-10-19 12:00:00"

Outputs:
    validation_summary.csv
    valid_orders.csv
    records_needing_correction.csv
    duplicate_order_ids.csv
"""

from pathlib import Path
import argparse
import pandas as pd

REQUIRED_COLUMNS = [
    "Order ID",
    "Order Date",
    "Customer Name",
    "Phone Number",
    "Order Value (INR)",
    "Payment Type",
    "Order Status",
    "Confirmation Timestamp",
    "Dispatch Date",
    "Expected Delivery Date",
    "Delivered At",
    "Last Delivery Attempt Status",
    "Last Delivery Attempt At",
    "Assigned Employee",
]

# These fields must be present on every order row.
ROW_REQUIRED_FIELDS = [
    "Order ID",
    "Order Date",
    "Customer Name",
    "Phone Number",
    "Order Value (INR)",
    "Payment Type",
    "Order Status",
    "Assigned Employee",
]

DATE_COLUMNS = [
    "Order Date",
    "Confirmation Timestamp",
    "Dispatch Date",
    "Expected Delivery Date",
    "Delivered At",
    "Last Delivery Attempt At",
]

ALLOWED_PAYMENT_TYPES = {"prepaid", "cod"}
ALLOWED_STATUSES = {
    "unconfirmed",
    "confirmed",
    "dispatched",
    "out for delivery",
    "delivered",
    "delivery failed",
    "cancelled",
}


def load_file(file_path: str) -> pd.DataFrame:
    """Load an Excel or CSV file."""
    path = Path(file_path)
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path.resolve()}")

    suffix = path.suffix.lower()
    if suffix == ".csv":
        return pd.read_csv(path)
    if suffix in {".xlsx", ".xls"}:
        return pd.read_excel(path, sheet_name="Orders" if suffix == ".xlsx" else 0)
    raise ValueError("Unsupported file type. Please use .csv, .xlsx, or .xls.")


def validate_orders(raw_df: pd.DataFrame):
    """
    Return:
      clean_df: input rows with helper validation columns
      summary: summary metrics
      missing_columns: required columns absent from the uploaded file
    """
    df = raw_df.copy()
    df.columns = [str(c).strip() for c in df.columns]

    missing_columns = [c for c in REQUIRED_COLUMNS if c not in df.columns]
    if missing_columns:
        # Add absent columns as empty so the validation report can explain the issue.
        for column in missing_columns:
            df[column] = pd.NA

    # Normalize text for comparisons, but retain original displayed values.
    for column in ["Order ID", "Customer Name", "Phone Number", "Payment Type",
                   "Order Status", "Last Delivery Attempt Status", "Assigned Employee"]:
        df[column] = df[column].astype("string").str.strip()
        df[column] = df[column].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "<NA>": pd.NA})

    # Parse dates; errors='coerce' turns invalid non-empty date strings into NaT.
    invalid_date_columns = []
    for column in DATE_COLUMNS:
        original = df[column].copy()
        parsed = pd.to_datetime(original, errors="coerce")
        invalid_mask = original.notna() & original.astype("string").str.strip().ne("") & parsed.isna()
        df[column] = parsed
        flag_column = f"Invalid {column}"
        df[flag_column] = invalid_mask
        invalid_date_columns.append(flag_column)

    # Numeric validation for order value.
    original_value = df["Order Value (INR)"].copy()
    df["Order Value (INR)"] = pd.to_numeric(df["Order Value (INR)"], errors="coerce")
    df["Invalid Order Value"] = (
        original_value.notna()
        & original_value.astype("string").str.strip().ne("")
        & df["Order Value (INR)"].isna()
    )
    df["Negative Order Value"] = df["Order Value (INR)"].lt(0).fillna(False)

    # Row-level missing required fields.
    missing_field_names = []
    for _, row in df.iterrows():
        absent = [field for field in ROW_REQUIRED_FIELDS if pd.isna(row[field])]
        missing_field_names.append(", ".join(absent))
    df["Missing Mandatory Fields"] = missing_field_names
    df["Has Missing Mandatory Fields"] = df["Missing Mandatory Fields"].ne("")

    # Conditional requirements: required only when the status makes the field applicable.
    conditional_issues = []
    for _, row in df.iterrows():
        issues = []
        status = str(row["Order Status"]).strip().lower() if pd.notna(row["Order Status"]) else ""
        if status in {"confirmed", "dispatched", "out for delivery", "delivered", "delivery failed"}:
            if pd.isna(row["Confirmation Timestamp"]):
                issues.append("Missing Confirmation Timestamp")
        if status in {"dispatched", "out for delivery", "delivered", "delivery failed"}:
            if pd.isna(row["Dispatch Date"]):
                issues.append("Missing Dispatch Date")
            if pd.isna(row["Expected Delivery Date"]):
                issues.append("Missing Expected Delivery Date")
        if status == "delivered" and pd.isna(row["Delivered At"]):
            issues.append("Missing Delivered At")
        if status == "delivery failed":
            attempt = str(row["Last Delivery Attempt Status"]).strip().lower()
            if attempt != "failed":
                issues.append("Delivery Failed status but attempt status is not Failed")
            if pd.isna(row["Last Delivery Attempt At"]):
                issues.append("Missing Last Delivery Attempt At")
        conditional_issues.append(", ".join(issues))
    df["Conditional Data Issues"] = conditional_issues

    df["Has Invalid Date"] = df[invalid_date_columns].any(axis=1)
    df["Has Invalid Value"] = df["Invalid Order Value"] | df["Negative Order Value"]
    df["Duplicate Order ID"] = df["Order ID"].notna() & df["Order ID"].duplicated(keep=False)

    # Any row-level quality issue means we should not run exception rules on that row.
    df["Needs Data Correction"] = (
        df["Has Missing Mandatory Fields"]
        | df["Has Invalid Date"]
        | df["Has Invalid Value"]
        | df["Duplicate Order ID"]
        | df["Conditional Data Issues"].ne("")
        | df["Order Status"].str.lower().fillna("").isin(ALLOWED_STATUSES) .eq(False)
        | df["Payment Type"].str.lower().fillna("").isin(ALLOWED_PAYMENT_TYPES).eq(False)
    )

    def correction_reason(row):
        reasons = []
        if row["Has Missing Mandatory Fields"]:
            reasons.append("Missing mandatory fields: " + row["Missing Mandatory Fields"])
        invalid_dates = [col.replace("Invalid ", "") for col in invalid_date_columns if row[col]]
        if invalid_dates:
            reasons.append("Invalid date(s): " + ", ".join(invalid_dates))
        if row["Has Invalid Value"]:
            reasons.append("Invalid or negative order value")
        if row["Duplicate Order ID"]:
            reasons.append("Duplicate Order ID")
        if row["Conditional Data Issues"]:
            reasons.append(row["Conditional Data Issues"])
        if str(row["Order Status"]).lower() not in ALLOWED_STATUSES:
            reasons.append("Unrecognized Order Status")
        if str(row["Payment Type"]).lower() not in ALLOWED_PAYMENT_TYPES:
            reasons.append("Unrecognized Payment Type")
        return "; ".join(reasons)

    df["Correction Reason"] = df.apply(correction_reason, axis=1)
    df["Validation Status"] = df["Needs Data Correction"].map(
        {True: "Needs Data Correction", False: "Valid"}
    )

    summary = {
        "Total rows uploaded": int(len(df)),
        "Required columns missing from file": int(len(missing_columns)),
        "Rows with missing mandatory fields": int(df["Has Missing Mandatory Fields"].sum()),
        "Rows with invalid dates": int(df["Has Invalid Date"].sum()),
        "Rows with invalid/negative order values": int(df["Has Invalid Value"].sum()),
        "Rows involved in duplicate Order IDs": int(df["Duplicate Order ID"].sum()),
        "Rows with conditional data issues": int(df["Conditional Data Issues"].ne("").sum()),
        "Rows needing data correction (unique rows)": int(df["Needs Data Correction"].sum()),
        "Rows eligible for rule evaluation": int((~df["Needs Data Correction"]).sum()),
    }
    return df, summary, missing_columns


def main():
    parser = argparse.ArgumentParser(description="Validate Vitanatura order CSV/Excel data.")
    parser.add_argument("file", help="Path to .csv, .xlsx, or .xls file")
    parser.add_argument(
        "--as-of",
        default=None,
        help='Optional reference timestamp, e.g. "2026-10-19 12:00:00" (reserved for next step)',
    )
    args = parser.parse_args()

    raw_df = load_file(args.file)
    checked_df, summary, missing_columns = validate_orders(raw_df)

    out_dir = Path.cwd()
    summary_df = pd.DataFrame(
        [{"Check": key, "Result": value} for key, value in summary.items()]
    )
    summary_df.to_csv(out_dir / "validation_summary.csv", index=False)

    checked_df.loc[~checked_df["Needs Data Correction"]].to_csv(
        out_dir / "valid_orders.csv", index=False
    )
    checked_df.loc[checked_df["Needs Data Correction"]].to_csv(
        out_dir / "records_needing_correction.csv", index=False
    )
    checked_df.loc[checked_df["Duplicate Order ID"]].to_csv(
        out_dir / "duplicate_order_ids.csv", index=False
    )

    print("\n=== VITANATURA DATA VALIDATION SUMMARY ===")
    for key, value in summary.items():
        print(f"{key}: {value}")
    if missing_columns:
        print("\nMissing columns in uploaded file:")
        for column in missing_columns:
            print(f" - {column}")
    print("\nCreated files in:", out_dir.resolve())
    print(" - validation_summary.csv")
    print(" - valid_orders.csv")
    print(" - records_needing_correction.csv")
    print(" - duplicate_order_ids.csv")


if __name__ == "__main__":
    main()
