"""
Vitanatura Assessment — Automatic Exception Engine

Run:
    python exception_engine.py valid_orders.csv
    python exception_engine.py vitanatura_sample_orders.xlsx

Optional fixed evaluation time:
    python exception_engine.py valid_orders.csv --as-of "2026-10-19 12:00:00"

Outputs:
    generated_exceptions.csv
"""

from pathlib import Path
import argparse
import pandas as pd


def load_orders(file_path: str) -> pd.DataFrame:
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError(
            f"File not found: {path.resolve()}"
        )

    if path.suffix.lower() == ".csv":
        return pd.read_csv(path)

    if path.suffix.lower() in {".xlsx", ".xls"}:
        sheet = "Orders" if path.suffix.lower() == ".xlsx" else 0
        return pd.read_excel(path, sheet_name=sheet)

    raise ValueError("Use a CSV or Excel file.")


def parse_dates(df: pd.DataFrame) -> pd.DataFrame:
    date_columns = [
        "Order Date",
        "Confirmation Timestamp",
        "Dispatch Date",
        "Expected Delivery Date",
        "Delivered At",
        "Last Delivery Attempt At",
    ]

    for col in date_columns:
        if col in df.columns:
            df[col] = pd.to_datetime(
                df[col], errors="coerce"
            )

    return df


def evaluate_exceptions(
    orders: pd.DataFrame,
    as_of=None,
    dispatch_deadline_hours: int = 24,
) -> pd.DataFrame:
    """
    Evaluate order data against five business rules.

    dispatch_deadline_hours:
        Number of hours allowed between confirmation
        and dispatch. Defaults to 24.
    """

    if (
        not isinstance(dispatch_deadline_hours, int)
        or isinstance(dispatch_deadline_hours, bool)
        or dispatch_deadline_hours < 1
    ):
        raise ValueError(
            "dispatch_deadline_hours must be a positive integer."
        )

    df = orders.copy()
    df.columns = [str(c).strip() for c in df.columns]
    df = parse_dates(df)

    evaluation_time = (
        pd.Timestamp.now()
        if as_of is None
        else pd.Timestamp(as_of)
    )

    required = [
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

    missing = [
        col for col in required
        if col not in df.columns
    ]

    if missing:
        raise ValueError(
            "Missing required columns: "
            + ", ".join(missing)
        )

    # Exclude records already identified as needing correction.
    if "Needs Data Correction" in df.columns:
        needs_correction = (
            df["Needs Data Correction"]
            .fillna(True)
            .astype(bool)
        )
        df = df.loc[~needs_correction].copy()

    # Duplicate Order IDs must not generate tasks.
    df["Order ID"] = (
        df["Order ID"]
        .astype("string")
        .str.strip()
    )

    df = df.loc[
        ~df["Order ID"].duplicated(keep=False)
    ].copy()

    df["Order Value (INR)"] = pd.to_numeric(
        df["Order Value (INR)"],
        errors="coerce",
    )

    df["Order Status Normalized"] = (
        df["Order Status"]
        .astype("string")
        .str.strip()
        .str.lower()
    )

    df["Attempt Status Normalized"] = (
        df["Last Delivery Attempt Status"]
        .astype("string")
        .str.strip()
        .str.lower()
    )

    exceptions = []

    def add_exception(
        row,
        exception_type,
        reason,
        priority,
        action,
        deadline,
    ):
        exceptions.append({
            "Order ID": row["Order ID"],
            "Customer Name": row["Customer Name"],
            "Phone Number": row["Phone Number"],
            "Order Value (INR)": row["Order Value (INR)"],
            "Payment Type": row["Payment Type"],
            "Order Status": row["Order Status"],
            "Exception Type": exception_type,
            "Reason": reason,
            "Priority": priority,
            "Assigned Employee": row["Assigned Employee"],
            "Recommended Action": action,
            "Deadline": deadline,
            "Task Status": "Open",
            "Detected At": evaluation_time,
        })

    for _, row in df.iterrows():

        status = row["Order Status Normalized"]
        order_date = row["Order Date"]
        confirmed_at = row["Confirmation Timestamp"]
        dispatch_date = row["Dispatch Date"]
        expected_date = row["Expected Delivery Date"]
        delivered_at = row["Delivered At"]
        attempt_status = row["Attempt Status Normalized"]

        # RULE 1:
        # Unconfirmed orders older than 2 hours.
        # Order Date is used as a proxy because the sample
        # data does not contain an unconfirmed-since timestamp.
        if (
            status == "unconfirmed"
            and pd.notna(order_date)
        ):
            age_hours = (
                evaluation_time - order_date
            ).total_seconds() / 3600

            if age_hours > 2:
                add_exception(
                    row,
                    "Unconfirmed Order",
                    (
                        f"Order has remained unconfirmed for "
                        f"{age_hours:.1f} hours "
                        "(threshold: >2 hours)."
                    ),
                    "Medium",
                    "Contact the customer and confirm or cancel the order.",
                    evaluation_time + pd.Timedelta(hours=2),
                )

        # RULE 2:
        # Confirmed but not dispatched after the configured
        # dispatch deadline.
        if (
            status == "confirmed"
            and pd.isna(dispatch_date)
            and pd.notna(confirmed_at)
        ):
            age_hours = (
                evaluation_time - confirmed_at
            ).total_seconds() / 3600

            if age_hours > dispatch_deadline_hours:
                add_exception(
                    row,
                    "Dispatch Delay",
                    (
                        f"Order has been confirmed for "
                        f"{age_hours:.1f} hours but is not "
                        f"dispatched (threshold: "
                        f">{dispatch_deadline_hours} hours)."
                    ),
                    "High",
                    "Contact the warehouse and arrange dispatch.",
                    evaluation_time + pd.Timedelta(hours=4),
                )

        # RULE 3:
        # Dispatched, not delivered, and expected delivery
        # date has passed.
        if (
            status in {
                "dispatched",
                "out for delivery",
                "delivery failed",
            }
            and pd.notna(dispatch_date)
            and pd.notna(expected_date)
            and pd.isna(delivered_at)
            and expected_date < evaluation_time
        ):
            add_exception(
                row,
                "Shipment Delay",
                (
                    "Expected delivery date has passed "
                    "and the order is not marked delivered."
                ),
                "High",
                "Check courier status and arrange a delivery follow-up.",
                evaluation_time + pd.Timedelta(hours=4),
            )

        # RULE 4:
        # Failed delivery requires immediate follow-up.
        if (
            status == "delivery failed"
            or attempt_status == "failed"
        ):
            add_exception(
                row,
                "Failed Delivery",
                (
                    "The latest delivery attempt failed "
                    "and requires immediate follow-up."
                ),
                "High",
                (
                    "Contact the customer, verify delivery "
                    "details, and arrange another attempt."
                ),
                evaluation_time + pd.Timedelta(hours=1),
            )

        # RULE 5:
        # Order value > ₹5,000, overdue, and not delivered.
        if (
            pd.notna(row["Order Value (INR)"])
            and row["Order Value (INR)"] > 5000
            and status in {
                "dispatched",
                "out for delivery",
                "delivery failed",
            }
            and pd.notna(dispatch_date)
            and pd.notna(expected_date)
            and pd.isna(delivered_at)
            and expected_date < evaluation_time
        ):
            add_exception(
                row,
                "High-Value Overdue Order",
                (
                    f"Order value is "
                    f"₹{row['Order Value (INR)']:.2f} "
                    "and expected delivery date has passed."
                ),
                "High",
                (
                    "Escalate to the logistics manager and "
                    "contact the courier/customer."
                ),
                evaluation_time + pd.Timedelta(hours=2),
            )

    columns = [
        "Order ID",
        "Customer Name",
        "Phone Number",
        "Order Value (INR)",
        "Payment Type",
        "Order Status",
        "Exception Type",
        "Reason",
        "Priority",
        "Assigned Employee",
        "Recommended Action",
        "Deadline",
        "Task Status",
        "Detected At",
    ]

    return pd.DataFrame(
        exceptions,
        columns=columns,
    )


def main():
    parser = argparse.ArgumentParser(
        description="Generate order exceptions from validated order data."
    )

    parser.add_argument(
        "file",
        help="Input validated CSV or original Excel/CSV",
    )

    parser.add_argument(
        "--as-of",
        default=None,
        help='Fixed time, e.g. "2026-10-19 12:00:00"',
    )

    parser.add_argument(
        "--dispatch-deadline-hours",
        type=int,
        default=24,
        help="Hours allowed between confirmation and dispatch (default: 24)",
    )

    args = parser.parse_args()

    orders = load_orders(args.file)

    exceptions = evaluate_exceptions(
        orders,
        as_of=args.as_of,
        dispatch_deadline_hours=args.dispatch_deadline_hours,
    )

    output = Path.cwd() / "generated_exceptions.csv"

    exceptions.to_csv(
        output,
        index=False,
        date_format="%Y-%m-%d %H:%M:%S",
    )

    evaluation_time = (
        pd.Timestamp(args.as_of)
        if args.as_of
        else pd.Timestamp.now()
    )

    print("\n=== EXCEPTION ENGINE SUMMARY ===")
    print(f"Evaluation time: {evaluation_time}")
    print(
        "Dispatch deadline: "
        f"{args.dispatch_deadline_hours} hours"
    )
    print(f"Exceptions generated: {len(exceptions)}")

    if not exceptions.empty:
        print("\nExceptions by category:")
        print(
            exceptions["Exception Type"]
            .value_counts()
            .to_string()
        )

        print("\nExceptions by priority:")
        print(
            exceptions["Priority"]
            .value_counts()
            .to_string()
        )

    print(f"\nSaved: {output.resolve()}")


if __name__ == "__main__":
    main()