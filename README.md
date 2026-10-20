# Vitanatura Order Follow-up System

## Overview
A local prototype for identifying order exceptions, assigning follow-up tasks, tracking task status, and exporting reports.

## Technology
- Python
- Pandas
- Streamlit
- SQLite (`sqlite3`, included with Python)
- Excel/CSV input and CSV reports

## Project files
- `app.py` — Streamlit user interface
- `validate_orders.py` — required-column, missing-value, duplicate, date, numeric and conditional validation
- `exception_engine.py` — five rule-based exception checks
- `task_manager.py` — SQLite task storage, status changes, reassignment and audit log
- `vitanatura_sample_orders.xlsx` / `.csv` — synthetic sample data
- `requirements.txt` — Python package dependencies

## Setup
Use Python 3.10 or later if available.

Open a terminal in this folder and run:

```bash
pip install -r requirements.txt
```

## Run the application
```bash
streamlit run app.py
```

Open the local address shown in the terminal, usually `http://localhost:8501`.

## Optional command-line workflow
Create the database:
```bash
python task_manager.py init
```

Validate the sample workbook:
```bash
python validate_orders.py vitanatura_sample_orders.xlsx
```

Generate exceptions using a reproducible evaluation time:
```bash
python exception_engine.py valid_orders.csv --as-of "2026-10-19 12:00:00"
```

Import exception records into SQLite:
```bash
python task_manager.py import-exceptions generated_exceptions.csv
```

View tasks and audit history:
```bash
python task_manager.py list
python task_manager.py audit
```

## Business rules
1. Unconfirmed orders older than 2 hours.
2. Confirmed orders not dispatched more than 24 hours after confirmation.
3. Dispatched, undelivered orders whose expected delivery date has passed.
4. Failed delivery attempts requiring follow-up.
5. Undelivered orders worth more than ₹5,000 whose expected delivery date has passed.

The application uses a rule-based engine. Machine learning is not required for these deterministic conditions.

## Data quality behavior
Rows with missing mandatory values, invalid dates/values, duplicate order IDs, or conditional data issues are routed to data correction rather than rule evaluation. Review duplicate rows before deciding which record is authoritative.

## Persistence
The SQLite database file `vitanatura_tasks.db` is created in the current working directory. Keep this file to preserve task and audit data.

## Known prototype limitations
- The current app is a local prototype, not a deployed multi-user system.
- The responsible-user field is an audit label, not authenticated login.
- Rule thresholds and priority mapping are hard-coded in the first version.
- The sample dataset uses synthetic customer information and fixed example timestamps.
- Power BI and machine-learning extensions are not part of the critical working workflow.

## Five-minute demo outline
1. **0:00–0:40 — Purpose:** Explain that the system reduces manual order follow-up.
2. **0:40–1:20 — Upload and validation:** Upload the sample workbook and show eligible records and data-quality issues.
3. **1:20–2:20 — Exception engine:** Run the five rules and show reasons, priorities, actions and deadlines.
4. **2:20–3:20 — Task management:** Save exceptions, update a task status, reassign a task, and show the audit log.
5. **3:20–4:10 — Reporting:** Download the CSV report and describe dashboard metrics.
6. **4:10–5:00 — Business impact and limitations:** Explain time saved estimation, duplicate prevention approach, and remaining limitations.

## Submission checklist
- [ ] Application starts with `streamlit run app.py`
- [ ] Excel upload works
- [ ] CSV upload works
- [ ] Missing fields, duplicate IDs and invalid dates are identified
- [ ] All five rules have been tested
- [ ] Tasks persist after app refresh
- [ ] Status change appears in audit history
- [ ] Reassignment appears in audit history
- [ ] Re-uploading the same exceptions does not create duplicate task rows
- [ ] CSV report downloads successfully
- [ ] Source code, sample data and this README are included
