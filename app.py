"""
VITANATURA ELEMENTS
D2C ORDER FOLLOW-UP AUTOMATION SYSTEM

Run:
    streamlit run app.py

Required modules:
    validate_orders.py
    exception_engine.py
    task_manager.py

Features:
    - CSV and Excel upload
    - Data validation
    - Five-rule exception engine
    - SQLite task persistence
    - Duplicate-safe exception synchronization
    - Dark operations dashboard
    - Four uniform charts
    - Deadline-aware overdue calculations
    - Task assignment and status updates
    - Daily report CSV
    - Audit trail
"""

from pathlib import Path
import tempfile

import pandas as pd
import streamlit as st
import altair as alt

from validate_orders import load_file, validate_orders
from exception_engine import evaluate_exceptions

from task_manager import (
    init_db,
    import_exceptions,
    list_tasks,
    update_status,
    reassign_task,
    show_audit,
)


# ============================================================
# 1. PAGE CONFIGURATION
# ============================================================

st.set_page_config(
    page_title="Vitanatura | Order Operations",
    page_icon="📦",
    layout="wide",
    initial_sidebar_state="expanded",
)


# ============================================================
# 2. GLOBAL DARK THEME
# ============================================================

st.markdown(
    """
    <style>
    :root {
        color-scheme: dark;
    }

    .stApp,
    [data-testid="stAppViewContainer"],
    [data-testid="stMain"] {
        background: #0B1220 !important;
        color: #E8EEF7 !important;
    }

    [data-testid="stHeader"] {
        background: #0B1220 !important;
    }

    [data-testid="stSidebar"] {
        background: #101A2B !important;
        border-right: 1px solid #26344B;
    }

    [data-testid="stSidebar"] * {
        color: #E8EEF7 !important;
    }

    .block-container {
        max-width: 1550px;
        padding-top: 1.5rem;
        padding-bottom: 3rem;
    }

    h1, h2, h3, h4, p, label {
        color: #E8EEF7;
    }

    /* Header banner */

    .hero {
        background: linear-gradient(
            115deg,
            #10283B 0%,
            #124A49 55%,
            #116E64 100%
        );
        border: 1px solid #286D69;
        border-radius: 17px;
        padding: 25px 30px;
        margin-bottom: 22px;
        box-shadow: 0 10px 30px rgba(0,0,0,.18);
    }

    .hero-kicker {
        color: #71E3CA;
        font-size: 11px;
        font-weight: 800;
        letter-spacing: 2px;
        margin-bottom: 10px;
    }

    .hero-title {
        color: #FFFFFF;
        font-size: 30px;
        font-weight: 800;
        line-height: 1.3;
    }

    .hero-subtitle {
        color: #C4E4E1;
        font-size: 14px;
        line-height: 1.6;
        margin-top: 10px;
    }

    /* Section headings */

    .section-title {
        color: #F1F5F9;
        font-size: 22px;
        font-weight: 750;
        margin-top: 20px;
        margin-bottom: 4px;
    }

    .section-subtitle {
        color: #91A4BC;
        font-size: 13px;
        margin-bottom: 16px;
    }

    /* KPI cards */

    .kpi-card {
        background: linear-gradient(
            145deg,
            #151F31,
            #111A29
        );
        border: 1px solid #2A3A52;
        border-radius: 15px;
        padding: 17px;
        min-height: 125px;
        box-shadow: 0 5px 16px rgba(0,0,0,.13);
    }

    .kpi-label {
        color: #A2B3C9;
        font-size: 11px;
        font-weight: 700;
        letter-spacing: .3px;
    }

    .kpi-value {
        color: #F8FAFC;
        font-size: 30px;
        font-weight: 800;
        line-height: 1.3;
        margin-top: 10px;
    }

    .kpi-note {
        color: #8397B0;
        font-size: 11px;
        margin-top: 5px;
    }

    /* Chart panels */

    .chart-heading {
        color: #F1F5F9;
        font-size: 16px;
        font-weight: 750;
        margin-bottom: 8px;
    }

    /* Streamlit components */

    div[data-testid="stMetric"] {
        background: #111B2B;
        border: 1px solid #29384F;
        border-radius: 12px;
        padding: 12px 15px;
    }

    div[data-testid="stMetricLabel"] {
        color: #A2B3C9 !important;
    }

    div[data-testid="stMetricValue"] {
        color: #F8FAFC !important;
    }

    div[data-testid="stDataFrame"] {
        border: 1px solid #29384F;
        border-radius: 10px;
    }

    div[data-testid="stExpander"] {
        border: 1px solid #29384F;
        border-radius: 12px;
    }

    .stButton button[kind="primary"] {
        background: #0F8B78;
        border-color: #0F8B78;
        color: white;
    }

    .stButton button[kind="primary"]:hover {
        background: #0B6D60;
        border-color: #0B6D60;
    }

    hr {
        border-color: #29384F !important;
    }

    .footer {
        color: #72849B;
        text-align: center;
        font-size: 11px;
        padding-top: 25px;
    }
    </style>
    """,
    unsafe_allow_html=True,
)


# ============================================================
# 3. INITIALIZE DATABASE
# ============================================================

init_db()


# ============================================================
# 4. HELPER FUNCTIONS
# ============================================================

def section(title, subtitle=None):
    st.markdown(
        f'<div class="section-title">{title}</div>',
        unsafe_allow_html=True,
    )

    if subtitle:
        st.markdown(
            f'<div class="section-subtitle">{subtitle}</div>',
            unsafe_allow_html=True,
        )


def render_kpi(label, value, note, icon):
    st.markdown(
        f"""
        <div class="kpi-card">
            <div class="kpi-label">{icon} &nbsp; {label}</div>
            <div class="kpi-value">{value}</div>
            <div class="kpi-note">{note}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )


def parse_dates(series):
    return pd.to_datetime(
        series,
        errors="coerce",
    )


def style_chart(chart):
    return (
        chart.configure_view(
            strokeOpacity=0,
        )
        .configure_axis(
            labelColor="#B8C7DA",
            titleColor="#B8C7DA",
            gridColor="#27364A",
            domainColor="#42516A",
            tickColor="#42516A",
            labelFontSize=11,
            titleFontSize=11,
        )
        .configure_legend(
            labelColor="#D8E2EF",
            titleColor="#D8E2EF",
            labelFontSize=11,
        )
    )


def get_task_data():
    data = list_tasks()

    if data is None:
        return pd.DataFrame()

    if not isinstance(data, pd.DataFrame):
        return pd.DataFrame(data)

    return data.copy()


def get_audit_data():
    data = show_audit()

    if data is None:
        return pd.DataFrame()

    if not isinstance(data, pd.DataFrame):
        return pd.DataFrame(data)

    return data.copy()


def calculate_overdue(tasks_df, evaluation_timestamp):
    """
    A task is overdue when:
    1. Its deadline is valid.
    2. Its deadline is earlier than the evaluation timestamp.
    3. Its status is Open or In Progress.

    The supplied evaluation timestamp is used consistently.
    """

    if tasks_df.empty:
        return pd.Series(
            False,
            index=tasks_df.index,
            dtype=bool,
        )

    if "deadline" not in tasks_df.columns:
        return pd.Series(
            False,
            index=tasks_df.index,
            dtype=bool,
        )

    deadlines = pd.to_datetime(
        tasks_df["deadline"],
        errors="coerce",
    )

    unresolved = tasks_df["task_status"].isin(
        ["Open", "In Progress"]
    )

    # Remove timezone information when comparing timestamps,
    # so a timezone-aware deadline doesn't cause a comparison error.
    if getattr(deadlines.dt, "tz", None) is not None:
        deadlines = deadlines.dt.tz_convert(None)

    if evaluation_timestamp.tzinfo is not None:
        evaluation_timestamp = (
            evaluation_timestamp
            .tz_convert(None)
            if hasattr(evaluation_timestamp, "tz_convert")
            else evaluation_timestamp.replace(tzinfo=None)
        )

    return (
        deadlines.notna()
        & (deadlines < evaluation_timestamp)
        & unresolved
    )


# ============================================================
# 5. SIDEBAR
# ============================================================

with st.sidebar:

    st.markdown("## 📦 VITANATURA")
    st.caption("ORDER OPERATIONS CENTER")

    st.divider()

    st.markdown("### Evaluation settings")

    evaluation_date = st.date_input(
        "Evaluation date",
        value=pd.Timestamp.now().date(),
    )

    evaluation_time = st.time_input(
        "Evaluation time",
        value=pd.Timestamp.now().time().replace(
            second=0,
            microsecond=0,
        ),
    )

    dispatch_deadline_hours = st.selectbox(
        "Dispatch deadline",
        options=[24, 12],
        index=0,
        format_func=lambda value: f"{value} hours",
    )

    current_user = st.text_input(
        "Responsible user",
        value="Manager",
    )

    st.divider()

    st.markdown("### 📁 Import orders")

    uploaded_file = st.file_uploader(
        "Upload Excel or CSV",
        type=["csv", "xlsx", "xls"],
    )

    st.caption(
        "Task records and audit history are stored in SQLite."
    )


# ============================================================
# 6. HEADER
# ============================================================

# Native Streamlit heading prevents HTML from appearing as code.

st.markdown(
    """
    <div style="
        background: linear-gradient(115deg, #10283B, #116E64);
        border: 1px solid #286D69;
        border-radius: 17px;
        padding: 24px 30px;
        margin-bottom: 22px;
    ">
        <div style="
            color: #71E3CA;
            font-size: 11px;
            font-weight: 800;
            letter-spacing: 2px;
        ">
            D2C · FULFILMENT · CUSTOMER EXPERIENCE
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    """
    # 📦 Order Operations Dashboard

    Monitor order exceptions, prioritize customer follow-ups,
    balance team workload, and keep orders moving.
    """
)


# ============================================================
# 7. UPLOAD AND VALIDATE ORDERS
# ============================================================

if uploaded_file is not None:

    try:
        suffix = Path(uploaded_file.name).suffix.lower()

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=suffix,
        ) as temp_file:

            temp_file.write(
                uploaded_file.getvalue()
            )

            temp_path = temp_file.name

        raw_orders = load_file(temp_path)

        (
            checked_orders,
            validation_summary,
            missing_columns,
        ) = validate_orders(raw_orders)

        valid_orders = checked_orders.loc[
            ~checked_orders["Needs Data Correction"]
        ].copy()

        st.session_state["valid_orders"] = valid_orders
        st.session_state["uploaded_order_count"] = len(
            raw_orders
        )
        st.session_state["uploaded_file_name"] = (
            uploaded_file.name
        )

        with st.expander(
            "📋 Data validation results",
            expanded=False,
        ):

            a, b, c, d = st.columns(4)

            a.metric(
                "Rows uploaded",
                validation_summary["Total rows uploaded"],
            )

            b.metric(
                "Eligible rows",
                validation_summary[
                    "Rows eligible for rule evaluation"
                ],
            )

            c.metric(
                "Needs correction",
                validation_summary[
                    "Rows needing data correction (unique rows)"
                ],
            )

            d.metric(
                "Duplicate ID rows",
                validation_summary[
                    "Rows involved in duplicate Order IDs"
                ],
            )

            if missing_columns:
                st.error(
                    "Required columns missing: "
                    + ", ".join(missing_columns)
                )

            correction_rows = checked_orders.loc[
                checked_orders["Needs Data Correction"]
            ].copy()

            if not correction_rows.empty:

                display_columns = [
                    column
                    for column in [
                        "Order ID",
                        "Customer Name",
                        "Validation Status",
                        "Correction Reason",
                    ]
                    if column in correction_rows.columns
                ]

                st.dataframe(
                    correction_rows[display_columns],
                    use_container_width=True,
                    hide_index=True,
                )

        st.success(
            f"Loaded {len(raw_orders)} orders. "
            f"{len(valid_orders)} eligible for evaluation."
        )

    except Exception as error:
        st.error(
            f"File processing failed: {error}"
        )


# ============================================================
# 8. EXCEPTION ENGINE
# ============================================================

evaluation_timestamp = pd.Timestamp.combine(
    evaluation_date,
    evaluation_time,
)

with st.expander(
    "⚙️ Exception engine",
    expanded=False,
):

    section(
        "Generate exceptions",
        "Run the five business rules against valid orders.",
    )

    st.caption(
        f"Evaluation timestamp: {evaluation_timestamp} · "
        f"Dispatch deadline: {dispatch_deadline_hours} hours"
    )

    if st.button(
        "Run exception rules",
        type="primary",
    ):

        orders_to_evaluate = st.session_state.get(
            "valid_orders"
        )

        if orders_to_evaluate is None:

            st.warning(
                "Upload an order file first."
            )

        else:

            try:
                generated = evaluate_exceptions(
                    orders_to_evaluate,
                    as_of=evaluation_timestamp,
                    dispatch_deadline_hours=int(
                        dispatch_deadline_hours
                    ),
                )

                st.session_state["latest_exceptions"] = (
                    generated
                )

                st.success(
                    f"Generated {len(generated)} exception records."
                )

            except Exception as error:
                st.error(
                    f"Exception generation failed: {error}"
                )

    if "latest_exceptions" in st.session_state:

        generated = st.session_state[
            "latest_exceptions"
        ]

        st.dataframe(
            generated,
            use_container_width=True,
            hide_index=True,
        )

        if not generated.empty:

            generated_csv = generated.to_csv(
                index=False
            ).encode("utf-8")

            left, right = st.columns(2)

            left.download_button(
                "⬇ Download exceptions CSV",
                generated_csv,
                "generated_exceptions.csv",
                "text/csv",
                use_container_width=True,
            )

            if right.button(
                "💾 Save exceptions to database",
                type="primary",
                use_container_width=True,
            ):

                with tempfile.NamedTemporaryFile(
                    mode="wb",
                    delete=False,
                    suffix=".csv",
                ) as temp_csv:

                    temp_csv.write(generated_csv)
                    exception_path = temp_csv.name

                try:
                    result = import_exceptions(
                        exception_path,
                        performed_by=current_user or "Manager",
                    )

                    st.success(
                        f"Database synchronization complete: {result}"
                    )

                    st.rerun()

                except Exception as error:
                    st.error(
                        f"Database save failed: {error}"
                    )


# ============================================================
# 9. LOAD TASKS
# ============================================================

tasks = get_task_data()

if not tasks.empty and "deadline" in tasks.columns:
    tasks["deadline"] = parse_dates(
        tasks["deadline"]
    )


# ============================================================
# 10. MAIN DASHBOARD
# ============================================================

if not tasks.empty:

    required_columns = [
        "exception_type",
        "assigned_employee",
        "priority",
        "task_status",
    ]

    missing_columns = [
        column
        for column in required_columns
        if column not in tasks.columns
    ]

    if missing_columns:

        st.error(
            "The task database is missing required columns: "
            + ", ".join(missing_columns)
        )

    else:

        section(
            "Operations overview",
            "Filter the task list to explore exceptions, priorities, and workload.",
        )

        # ----------------------------------------------------
        # FILTERS
        # ----------------------------------------------------

        with st.container(border=True):

            f1, f2, f3 = st.columns(
                [1.3, 1.2, 0.8]
            )

            exception_options = sorted(
                tasks["exception_type"]
                .dropna()
                .astype(str)
                .unique()
            )

            employee_options = sorted(
                tasks["assigned_employee"]
                .dropna()
                .astype(str)
                .unique()
            )

            priority_options = sorted(
                tasks["priority"]
                .dropna()
                .astype(str)
                .unique()
            )

            with f1:
                selected_exceptions = st.multiselect(
                    "Exception category",
                    exception_options,
                    default=exception_options,
                    key="dashboard_exception_filter",
                )

            with f2:
                selected_employees = st.multiselect(
                    "Assigned employee",
                    employee_options,
                    default=employee_options,
                    key="dashboard_employee_filter",
                )

            with f3:
                selected_priorities = st.multiselect(
                    "Priority",
                    priority_options,
                    default=priority_options,
                    key="dashboard_priority_filter",
                )

            use_deadline_filter = st.checkbox(
                "Filter by deadline date range",
                value=False,
            )

            start_date = None
            end_date = None

            if use_deadline_filter:

                d1, d2 = st.columns(2)

                with d1:
                    start_date = st.date_input(
                        "Deadline from",
                        value=(
                            evaluation_timestamp.date()
                            - pd.Timedelta(days=30)
                        ),
                        key="deadline_start",
                    )

                with d2:
                    end_date = st.date_input(
                        "Deadline to",
                        value=evaluation_timestamp.date(),
                        key="deadline_end",
                    )

        # ----------------------------------------------------
        # APPLY FILTERS
        # ----------------------------------------------------

        filtered = tasks.copy()

        filtered = filtered[
            filtered["exception_type"].isin(
                selected_exceptions
            )
        ]

        filtered = filtered[
            filtered["assigned_employee"].isin(
                selected_employees
            )
        ]

        filtered = filtered[
            filtered["priority"].isin(
                selected_priorities
            )
        ]

        if (
            use_deadline_filter
            and start_date is not None
            and end_date is not None
            and "deadline" in filtered.columns
        ):

            if start_date > end_date:

                st.warning(
                    "The start date must not be after the end date."
                )

            else:

                filtered = filtered[
                    filtered["deadline"].dt.date.between(
                        start_date,
                        end_date,
                    )
                ]

        # ----------------------------------------------------
        # KPI CALCULATIONS
        # ----------------------------------------------------

        active_mask = filtered["task_status"].isin(
            ["Open", "In Progress"]
        )

        resolved_mask = filtered["task_status"].isin(
            ["Resolved", "Completed", "Closed"]
        )

        overdue_mask = calculate_overdue(
            filtered,
            evaluation_timestamp,
        )

        task_count = len(filtered)
        active_count = int(active_mask.sum())
        high_count = int(
            filtered["priority"].eq("High").sum()
        )
        overdue_count = int(overdue_mask.sum())
        resolved_count = int(resolved_mask.sum())

        # ----------------------------------------------------
        # KPI CARDS
        # ----------------------------------------------------

        k1, k2, k3, k4, k5 = st.columns(5)

        with k1:
            render_kpi(
                "TASKS IN VIEW",
                f"{task_count:,}",
                "Matching current filters",
                "▦",
            )

        with k2:
            render_kpi(
                "ACTIVE EXCEPTIONS",
                f"{active_count:,}",
                "Open or in progress",
                "↻",
            )

        with k3:
            render_kpi(
                "HIGH PRIORITY",
                f"{high_count:,}",
                "Requires close attention",
                "◆",
            )

        with k4:
            render_kpi(
                "OVERDUE TASKS",
                f"{overdue_count:,}",
                "Past deadline and unresolved",
                "◷",
            )

        with k5:
            render_kpi(
                "RESOLVED",
                f"{resolved_count:,}",
                "Completed follow-ups",
                "✓",
            )

        # ----------------------------------------------------
        # RESOLUTION PROGRESS
        # ----------------------------------------------------

        section(
            "Operational performance",
            "Monitor task resolution across the current selection.",
        )

        progress_left, progress_right = st.columns(
            [2, 1]
        )

        resolution_rate = (
            resolved_count / task_count
            if task_count > 0
            else 0
        )

        with progress_left:

            st.markdown("**Resolution progress**")

            st.progress(resolution_rate)

            st.caption(
                f"{resolved_count} of {task_count} tasks resolved."
            )

        with progress_right:

            st.metric(
                "Resolution rate",
                f"{resolution_rate:.1%}",
            )

        # ====================================================
        # 11. FOUR EQUAL-SIZED CHARTS
        # ====================================================

        section(
            "Exception analytics",
            "Compare exception categories, priorities, task status, and employee workload.",
        )

        CHART_HEIGHT = 250

        if not filtered.empty:

            # Chart data

            exception_data = (
                filtered.groupby(
                    "exception_type",
                    as_index=False,
                )
                .size()
                .rename(columns={"size": "Tasks"})
                .sort_values("Tasks", ascending=True)
            )

            priority_data = (
                filtered.groupby(
                    "priority",
                    as_index=False,
                )
                .size()
                .rename(columns={"size": "Tasks"})
            )

            status_data = (
                filtered.groupby(
                    "task_status",
                    as_index=False,
                )
                .size()
                .rename(columns={"size": "Tasks"})
            )

            employee_data = (
                filtered.groupby(
                    "assigned_employee",
                    as_index=False,
                )
                .size()
                .rename(columns={"size": "Tasks"})
                .sort_values("Tasks", ascending=True)
            )

            # CHART 1: EXCEPTION CATEGORY

            exception_chart = (
                alt.Chart(exception_data)
                .mark_bar(
                    cornerRadiusEnd=5,
                    color="#36C9A5",
                )
                .encode(
                    x=alt.X(
                        "Tasks:Q",
                        title="Number of tasks",
                        scale=alt.Scale(
                            domainMin=0,
                            nice=True,
                        ),
                    ),
                    y=alt.Y(
                        "exception_type:N",
                        sort="-x",
                        title=None,
                        axis=alt.Axis(
                            labelLimit=150,
                            labelFontSize=10,
                        ),
                    ),
                    tooltip=[
                        alt.Tooltip(
                            "exception_type:N",
                            title="Exception",
                        ),
                        alt.Tooltip(
                            "Tasks:Q",
                            title="Tasks",
                        ),
                    ],
                )
                .properties(
                    height=CHART_HEIGHT,
                )
            )

            # CHART 2: PRIORITY DONUT

            priority_colors = alt.Scale(
                domain=[
                    "High",
                    "Medium",
                    "Low",
                ],
                range=[
                    "#FF647C",
                    "#F7B955",
                    "#36C9A5",
                ],
            )

            priority_donut = (
                alt.Chart(priority_data)
                .mark_arc(
                    innerRadius=60,
                    outerRadius=90,
                    stroke="#111B2B",
                    strokeWidth=3,
                )
                .encode(
                    theta=alt.Theta(
                        "Tasks:Q",
                        stack=True,
                    ),
                    color=alt.Color(
                        "priority:N",
                        scale=priority_colors,
                        legend=alt.Legend(
                            title=None,
                            orient="bottom",
                            direction="horizontal",
                        ),
                    ),
                    tooltip=[
                        alt.Tooltip(
                            "priority:N",
                            title="Priority",
                        ),
                        alt.Tooltip(
                            "Tasks:Q",
                            title="Tasks",
                        ),
                    ],
                )
                .properties(
                    height=CHART_HEIGHT,
                )
            )

            donut_center = (
                alt.Chart(
                    pd.DataFrame({
                        "label": [str(task_count)]
                    })
                )
                .mark_text(
                    fontSize=25,
                    fontWeight="bold",
                    color="#F8FAFC",
                )
                .encode(
                    text="label:N",
                )
            )

            priority_chart = (
                priority_donut + donut_center
            ).properties(
                height=CHART_HEIGHT,
            )

            # CHART 3: TASK STATUS

            status_colors = alt.Scale(
                domain=[
                    "Open",
                    "In Progress",
                    "Resolved",
                    "Completed",
                    "Closed",
                ],
                range=[
                    "#F7B955",
                    "#61A8FF",
                    "#36C9A5",
                    "#36C9A5",
                    "#36C9A5",
                ],
            )

            status_chart = (
                alt.Chart(status_data)
                .mark_bar(
                    cornerRadiusEnd=5,
                )
                .encode(
                    y=alt.Y(
                        "task_status:N",
                        sort="-x",
                        title=None,
                        axis=alt.Axis(
                            labelLimit=130,
                            labelFontSize=10,
                        ),
                    ),
                    x=alt.X(
                        "Tasks:Q",
                        title="Number of tasks",
                        scale=alt.Scale(
                            domainMin=0,
                            nice=True,
                        ),
                    ),
                    color=alt.Color(
                        "task_status:N",
                        scale=status_colors,
                        legend=None,
                    ),
                    tooltip=[
                        alt.Tooltip(
                            "task_status:N",
                            title="Status",
                        ),
                        alt.Tooltip(
                            "Tasks:Q",
                            title="Tasks",
                        ),
                    ],
                )
                .properties(
                    height=CHART_HEIGHT,
                )
            )

            # CHART 4: EMPLOYEE WORKLOAD

            employee_chart = (
                alt.Chart(employee_data)
                .mark_bar(
                    cornerRadiusEnd=5,
                    color="#8290FF",
                )
                .encode(
                    y=alt.Y(
                        "assigned_employee:N",
                        sort="-x",
                        title=None,
                        axis=alt.Axis(
                            labelLimit=130,
                            labelFontSize=10,
                        ),
                    ),
                    x=alt.X(
                        "Tasks:Q",
                        title="Assigned tasks",
                        scale=alt.Scale(
                            domainMin=0,
                            nice=True,
                        ),
                    ),
                    tooltip=[
                        alt.Tooltip(
                            "assigned_employee:N",
                            title="Employee",
                        ),
                        alt.Tooltip(
                            "Tasks:Q",
                            title="Assigned tasks",
                        ),
                    ],
                )
                .properties(
                    height=CHART_HEIGHT,
                )
            )

            # DISPLAY THE CHART GRID

            left_column, right_column = st.columns(2)

            with left_column:

                with st.container(border=True):

                    st.markdown(
                        "#### Exception category"
                    )

                    st.altair_chart(
                        style_chart(exception_chart),
                        use_container_width=True,
                    )

            with right_column:

                with st.container(border=True):

                    st.markdown(
                        "#### Priority distribution"
                    )

                    st.altair_chart(
                        style_chart(priority_chart),
                        use_container_width=True,
                    )

            left_column, right_column = st.columns(2)

            with left_column:

                with st.container(border=True):

                    st.markdown(
                        "#### Task status"
                    )

                    st.altair_chart(
                        style_chart(status_chart),
                        use_container_width=True,
                    )

            with right_column:

                with st.container(border=True):

                    st.markdown(
                        "#### Employee workload"
                    )

                    st.altair_chart(
                        style_chart(employee_chart),
                        use_container_width=True,
                    )

        else:

            st.info(
                "No matching tasks are available for the selected filters."
            )

        # ====================================================
        # 12. PRIORITY QUEUE
        # ====================================================

        section(
            "Priority queue",
            "Review unresolved tasks and identify overdue work.",
        )

        priority_queue = filtered.loc[
            active_mask
        ].copy()

        if not priority_queue.empty:

            priority_queue["_overdue"] = calculate_overdue(
                priority_queue,
                evaluation_timestamp,
            )

            priority_queue["_priority_order"] = (
                priority_queue["priority"]
                .map(
                    {
                        "High": 0,
                        "Medium": 1,
                        "Low": 2,
                    }
                )
                .fillna(3)
            )

            sort_columns = [
                "_overdue",
                "_priority_order",
            ]

            sort_ascending = [
                False,
                True,
            ]

            if "deadline" in priority_queue.columns:
                sort_columns.append("deadline")
                sort_ascending.append(True)

            priority_queue = priority_queue.sort_values(
                by=sort_columns,
                ascending=sort_ascending,
                na_position="last",
            )

            priority_queue = priority_queue.drop(
                columns=[
                    "_overdue",
                    "_priority_order",
                ]
            )

            queue_columns = [
                column
                for column in [
                    "order_id",
                    "exception_type",
                    "priority",
                    "assigned_employee",
                    "task_status",
                    "deadline",
                    "reason",
                ]
                if column in priority_queue.columns
            ]

            st.dataframe(
                priority_queue[queue_columns],
                use_container_width=True,
                hide_index=True,
            )

        else:

            st.success(
                "No unresolved tasks in the current view."
            )

        # ====================================================
        # 13. TASK SEARCH
        # ====================================================

        with st.expander(
            "🔎 Search and inspect tasks",
            expanded=False,
        ):

            search_text = st.text_input(
                "Search order ID, customer, exception, or employee",
                key="task_search",
            )

            visible_tasks = filtered.copy()

            if search_text.strip():

                searchable_columns = [
                    column
                    for column in [
                        "order_id",
                        "customer_name",
                        "exception_type",
                        "assigned_employee",
                        "task_status",
                    ]
                    if column in visible_tasks.columns
                ]

                search_mask = pd.Series(
                    False,
                    index=visible_tasks.index,
                )

                for column in searchable_columns:

                    search_mask |= (
                        visible_tasks[column]
                        .fillna("")
                        .astype(str)
                        .str.contains(
                            search_text.strip(),
                            case=False,
                            regex=False,
                        )
                    )

                visible_tasks = visible_tasks[
                    search_mask
                ]

            st.caption(
                f"Showing {len(visible_tasks)} of {len(filtered)} tasks."
            )

            st.dataframe(
                visible_tasks,
                use_container_width=True,
                hide_index=True,
            )

else:

    st.markdown(
        """
        <div style="
            background:#111B2B;
            border:1px solid #29384F;
            border-radius:14px;
            padding:24px;
        ">
            <h3 style="color:#F1F5F9;">
                No task data available
            </h3>
            <p style="color:#91A4BC;">
                Upload your order file, run the exception rules,
                and save the generated exceptions to populate the dashboard.
            </p>
        </div>
        """,
        unsafe_allow_html=True,
    )


# ============================================================
# 14. TASK MANAGEMENT
# ============================================================

st.divider()

section(
    "Task management",
    "Update task progress and employee assignments.",
)

management_tasks = get_task_data()

if management_tasks.empty:

    st.info(
        "Save generated exceptions before managing tasks."
    )

else:

    with st.expander(
        "View and update tasks",
        expanded=False,
    ):

        st.dataframe(
            management_tasks,
            use_container_width=True,
            hide_index=True,
        )

        task_labels = [
            f'{row["order_id"]} | {row["exception_type"]}'
            for _, row in management_tasks.iterrows()
        ]

        selected_label = st.selectbox(
            "Select task",
            task_labels,
            key="management_task",
        )

        selected_task = management_tasks.iloc[
            task_labels.index(selected_label)
        ]

        selected_order_id = str(
            selected_task["order_id"]
        )

        selected_exception_type = str(
            selected_task["exception_type"]
        )

        left, right = st.columns(2)

        # STATUS UPDATE

        with left:

            valid_statuses = [
                "Open",
                "In Progress",
                "Resolved",
            ]

            current_status = selected_task["task_status"]

            status_index = (
                valid_statuses.index(current_status)
                if current_status in valid_statuses
                else 0
            )

            new_status = st.selectbox(
                "New status",
                valid_statuses,
                index=status_index,
                key="management_status",
            )

            if st.button(
                "Update task status",
                type="primary",
            ):

                try:

                    update_status(
                        selected_order_id,
                        selected_exception_type,
                        new_status,
                        performed_by=current_user or "Manager",
                    )

                    st.success(
                        f"Task {selected_order_id} updated to {new_status}."
                    )

                    st.rerun()

                except Exception as error:

                    st.error(
                        f"Status update failed: {error}"
                    )

        # EMPLOYEE REASSIGNMENT

        with right:

            employees = [
                "Anjali Nair",
                "Rahul Menon",
                "Fathima K",
                "Arjun Das",
                "Meera S",
            ]

            current_employee = selected_task.get(
                "assigned_employee"
            )

            employee_index = (
                employees.index(current_employee)
                if current_employee in employees
                else 0
            )

            new_employee = st.selectbox(
                "Assign employee",
                employees,
                index=employee_index,
                key="management_employee",
            )

            if st.button("Reassign task"):

                try:

                    reassign_task(
                        selected_order_id,
                        selected_exception_type,
                        new_employee,
                        performed_by=current_user or "Manager",
                    )

                    st.success(
                        f"Task {selected_order_id} reassigned to {new_employee}."
                    )

                    st.rerun()

                except Exception as error:

                    st.error(
                        f"Reassignment failed: {error}"
                    )


# ============================================================
# 15. DAILY EXCEPTION REPORT
# ============================================================

st.divider()

section(
    "Daily exception report",
    "Export task records and operational indicators to CSV.",
)

report_tasks = get_task_data()

if not report_tasks.empty:

    report = report_tasks.copy()

    if "deadline" in report.columns:

        report["deadline"] = parse_dates(
            report["deadline"]
        )

        report["is_overdue"] = calculate_overdue(
            report,
            evaluation_timestamp,
        )

    else:

        report["is_overdue"] = False

    report["is_unresolved"] = report[
        "task_status"
    ].isin(
        ["Open", "In Progress"]
    )

    report["is_high_priority"] = report[
        "priority"
    ].eq("High")

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "Unresolved tasks",
        int(report["is_unresolved"].sum()),
    )

    c2.metric(
        "Overdue unresolved",
        int(report["is_overdue"].sum()),
    )

    c3.metric(
        "High-priority tasks",
        int(report["is_high_priority"].sum()),
    )

    st.download_button(
        "⬇ Download daily exception report",
        report.to_csv(index=False).encode("utf-8"),
        "daily_exception_report.csv",
        "text/csv",
        type="primary",
    )

else:

    st.info(
        "No task records available for the report."
    )


# ============================================================
# 16. AUDIT TRAIL
# ============================================================

st.divider()

section(
    "Audit trail",
    "Review task creation, status changes, and reassignments.",
)

audit_data = get_audit_data()

if audit_data.empty:

    st.info(
        "No audit events recorded yet."
    )

else:

    with st.expander(
        "View audit history",
        expanded=False,
    ):

        st.dataframe(
            audit_data,
            use_container_width=True,
            hide_index=True,
        )

        st.download_button(
            "⬇ Download audit trail",
            audit_data.to_csv(index=False).encode("utf-8"),
            "vitanatura_audit_trail.csv",
            "text/csv",
        )


# ============================================================
# 17. FOOTER
# ============================================================

st.markdown(
    """
    <div class="footer">
        Vitanatura Elements · D2C Order Follow-up Automation System
        <br>
        Python · Pandas · Streamlit · Altair · SQLite
    </div>
    """,
    unsafe_allow_html=True,
)