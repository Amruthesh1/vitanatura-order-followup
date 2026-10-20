"""
Vitanatura Assessment — Step 5: SQLite Task Management

Creates/uses vitanatura_tasks.db in the current working directory.

Run:
    python task_manager.py init
    python task_manager.py import-exceptions generated_exceptions.csv
    python task_manager.py list
    python task_manager.py update VN0004 "Failed Delivery" --status "In Progress"
    python task_manager.py reassign VN0004 "Failed Delivery" "Meera S" --by "Manager"

This is a starter database module. The Streamlit interface will call the same
functions in this file in a later step.
"""

import argparse
import sqlite3
from pathlib import Path
import pandas as pd

DB_PATH = Path.cwd() / "vitanatura_tasks.db"


def connect():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA foreign_keys = ON")
    return conn


def init_db():
    with connect() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS tasks (
                task_id INTEGER PRIMARY KEY AUTOINCREMENT,
                order_id TEXT NOT NULL,
                exception_type TEXT NOT NULL,
                customer_name TEXT,
                phone_number TEXT,
                order_value REAL,
                payment_type TEXT,
                order_status TEXT,
                reason TEXT NOT NULL,
                priority TEXT NOT NULL CHECK(priority IN ('High', 'Medium', 'Low')),
                assigned_employee TEXT,
                recommended_action TEXT NOT NULL,
                deadline TEXT,
                task_status TEXT NOT NULL DEFAULT 'Open'
                    CHECK(task_status IN ('Open', 'In Progress', 'Resolved')),
                detected_at TEXT,
                created_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                updated_at TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(order_id, exception_type)
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS audit_log (
                audit_id INTEGER PRIMARY KEY AUTOINCREMENT,
                task_id INTEGER NOT NULL,
                order_id TEXT NOT NULL,
                action TEXT NOT NULL,
                old_value TEXT,
                new_value TEXT,
                performed_by TEXT NOT NULL,
                action_timestamp TEXT NOT NULL DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY(task_id) REFERENCES tasks(task_id)
            )
        """)
        conn.commit()


def import_exceptions(csv_path: str, performed_by: str = "System") -> dict:
    init_db()
    df = pd.read_csv(csv_path)
    required = [
        "Order ID", "Exception Type", "Reason", "Priority",
        "Recommended Action", "Task Status"
    ]
    missing = [c for c in required if c not in df.columns]
    if missing:
        raise ValueError("Exception CSV missing columns: " + ", ".join(missing))

    inserted = 0
    updated = 0
    skipped = 0
    with connect() as conn:
        for _, row in df.iterrows():
            order_id = str(row["Order ID"]).strip()
            exception_type = str(row["Exception Type"]).strip()
            if not order_id or order_id.lower() == "nan":
                skipped += 1
                continue

            values = (
                order_id,
                exception_type,
                None if pd.isna(row.get("Customer Name")) else str(row.get("Customer Name")),
                None if pd.isna(row.get("Phone Number")) else str(row.get("Phone Number")),
                None if pd.isna(row.get("Order Value (INR)")) else float(row.get("Order Value (INR)")),
                None if pd.isna(row.get("Payment Type")) else str(row.get("Payment Type")),
                None if pd.isna(row.get("Order Status")) else str(row.get("Order Status")),
                str(row["Reason"]),
                str(row["Priority"]),
                None if pd.isna(row.get("Assigned Employee")) else str(row.get("Assigned Employee")),
                str(row["Recommended Action"]),
                None if pd.isna(row.get("Deadline")) else str(row.get("Deadline")),
                str(row.get("Task Status", "Open")),
                None if pd.isna(row.get("Detected At")) else str(row.get("Detected At")),
            )
            existing = conn.execute(
                "SELECT task_id, task_status, assigned_employee FROM tasks WHERE order_id=? AND exception_type=?",
                (order_id, exception_type)
            ).fetchone()

            if existing:
                # Do not overwrite task status or assignment on repeated uploads.
                # Refresh the descriptive fields while preserving task workflow.
                conn.execute("""
                    UPDATE tasks
                    SET customer_name=?, phone_number=?, order_value=?, payment_type=?,
                        order_status=?, reason=?, priority=?, recommended_action=?,
                        deadline=?, detected_at=?, updated_at=CURRENT_TIMESTAMP
                    WHERE task_id=?
                """, (
                    values[2], values[3], values[4], values[5], values[6],
                    values[7], values[8], values[10], values[11], values[13],
                    existing["task_id"]
                ))
                updated += 1
            else:
                cur = conn.execute("""
                    INSERT INTO tasks (
                        order_id, exception_type, customer_name, phone_number,
                        order_value, payment_type, order_status, reason, priority,
                        assigned_employee, recommended_action, deadline, task_status, detected_at
                    ) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
                """, values)
                task_id = cur.lastrowid
                conn.execute("""
                    INSERT INTO audit_log (task_id, order_id, action, old_value, new_value, performed_by)
                    VALUES (?, ?, ?, ?, ?, ?)
                """, (task_id, order_id, "Task Created", None, exception_type, performed_by))
                inserted += 1

        conn.commit()

    return {"inserted": inserted, "updated_existing": updated, "skipped": skipped}


def list_tasks():
    init_db()
    with connect() as conn:
        rows = conn.execute("""
            SELECT task_id, order_id, exception_type, customer_name, priority,
                   assigned_employee, task_status, deadline, reason
            FROM tasks
            ORDER BY CASE priority WHEN 'High' THEN 1 WHEN 'Medium' THEN 2 ELSE 3 END,
                     deadline
        """).fetchall()
    return pd.DataFrame([dict(r) for r in rows])


def update_status(order_id: str, exception_type: str, new_status: str, performed_by: str = "Manager"):
    allowed = {"Open", "In Progress", "Resolved"}
    if new_status not in allowed:
        raise ValueError(f"Status must be one of: {', '.join(sorted(allowed))}")
    init_db()
    with connect() as conn:
        row = conn.execute(
            "SELECT task_id, task_status FROM tasks WHERE order_id=? AND exception_type=?",
            (order_id, exception_type)
        ).fetchone()
        if not row:
            raise ValueError("Task not found for that Order ID and Exception Type.")
        old_status = row["task_status"]
        conn.execute("""
            UPDATE tasks SET task_status=?, updated_at=CURRENT_TIMESTAMP WHERE task_id=?
        """, (new_status, row["task_id"]))
        conn.execute("""
            INSERT INTO audit_log (task_id, order_id, action, old_value, new_value, performed_by)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (row["task_id"], order_id, "Status Changed", old_status, new_status, performed_by))
        conn.commit()


def reassign_task(order_id: str, exception_type: str, new_employee: str, performed_by: str = "Manager"):
    if not new_employee.strip():
        raise ValueError("Employee name cannot be empty.")
    init_db()
    with connect() as conn:
        row = conn.execute(
            "SELECT task_id, assigned_employee FROM tasks WHERE order_id=? AND exception_type=?",
            (order_id, exception_type)
        ).fetchone()
        if not row:
            raise ValueError("Task not found for that Order ID and Exception Type.")
        old_employee = row["assigned_employee"]
        conn.execute("""
            UPDATE tasks SET assigned_employee=?, updated_at=CURRENT_TIMESTAMP WHERE task_id=?
        """, (new_employee.strip(), row["task_id"]))
        conn.execute("""
            INSERT INTO audit_log (task_id, order_id, action, old_value, new_value, performed_by)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (row["task_id"], order_id, "Task Reassigned", old_employee, new_employee.strip(), performed_by))
        conn.commit()


def show_audit():
    init_db()
    with connect() as conn:
        rows = conn.execute("""
            SELECT audit_id, order_id, action, old_value, new_value,
                   performed_by, action_timestamp
            FROM audit_log ORDER BY audit_id DESC
        """).fetchall()
    return pd.DataFrame([dict(r) for r in rows])


def main():
    parser = argparse.ArgumentParser(description="Vitanatura SQLite task manager")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("init", help="Create the SQLite database and tables")

    imp = sub.add_parser("import-exceptions", help="Import generated exceptions from CSV")
    imp.add_argument("csv_file")
    imp.add_argument("--by", default="System")

    sub.add_parser("list", help="List stored tasks")
    sub.add_parser("audit", help="Show audit history")

    upd = sub.add_parser("update", help="Update a task status")
    upd.add_argument("order_id")
    upd.add_argument("exception_type")
    upd.add_argument("--status", required=True)
    upd.add_argument("--by", default="Manager")

    reas = sub.add_parser("reassign", help="Reassign a task")
    reas.add_argument("order_id")
    reas.add_argument("exception_type")
    reas.add_argument("employee")
    reas.add_argument("--by", default="Manager")

    args = parser.parse_args()

    if args.command == "init":
        init_db()
        print(f"Database ready: {DB_PATH.resolve()}")
    elif args.command == "import-exceptions":
        result = import_exceptions(args.csv_file, args.by)
        print("Import complete:", result)
        print(f"Database: {DB_PATH.resolve()}")
    elif args.command == "list":
        df = list_tasks()
        print(df.to_string(index=False) if not df.empty else "No tasks found.")
    elif args.command == "audit":
        df = show_audit()
        print(df.to_string(index=False) if not df.empty else "No audit entries found.")
    elif args.command == "update":
        update_status(args.order_id, args.exception_type, args.status, args.by)
        print("Task status updated.")
    elif args.command == "reassign":
        reassign_task(args.order_id, args.exception_type, args.employee, args.by)
        print("Task reassigned.")


if __name__ == "__main__":
    main()
