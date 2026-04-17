
import os
import pandas as pd
import openpyxl 
import streamlit as st  
from datetime import datetime, time, timedelta
from streamlit_calendar import calendar
import requests
from io import BytesIO
import sqlite3
import io


# Create database connection

def init_db():
    conn = sqlite3.connect("notes.db")
    c = conn.cursor()

    c.execute("""
        CREATE TABLE IF NOT EXISTS notes (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            name TEXT,
            date TEXT,
            type TEXT,
            note TEXT
        )
    """)

    conn.commit()
    conn.close()

init_db()

# Functions to load and save notes
def get_note(name, date, event_type):
    conn = sqlite3.connect("notes.db")
    c = conn.cursor()

    c.execute("""
        SELECT note FROM notes
        WHERE name=? AND date=? AND type=?
    """, (name, date, event_type))

    row = c.fetchone()
    conn.close()

    return row[0] if row else ""


def save_note(name, date, event_type, note):
    conn = sqlite3.connect("notes.db")
    c = conn.cursor()

    # delete old note if exists
    c.execute("""
        DELETE FROM notes
        WHERE name=? AND date=? AND type=?
    """, (name, date, event_type))

    # insert new note
    c.execute("""
        INSERT INTO notes (name, date, type, note)
        VALUES (?, ?, ?, ?)
    """, (name, date, event_type, note))

    conn.commit()
    conn.close()


    # *****Build report function that combines attendance data with notes*****
def build_report(filtered_df, first_punch_df, cutoff_time, start_date, end_date, conn):

    # Load notes from SQLite
    notes_rows = conn.execute(
        "SELECT name, date, type, note FROM notes"
    ).fetchall()

    notes_data = [
        {"Name": n[0], "Date": n[1], "Type": n[2], "Note": n[3]}
        for n in notes_rows
    ]

    report_rows = []

    all_days = pd.date_range(start=start_date, end=end_date).date
    weekdays = [d for d in all_days if d.weekday() < 5]

    for employee in filtered_df["Name"].unique():

        emp_first = first_punch_df[first_punch_df["Name"] == employee]
        emp_data = filtered_df[filtered_df["Name"] == employee]

        punched_days = emp_data["Date"].unique()

        for day in weekdays:

            status = "Present"
            note_text = ""

            # Check Late
            late_row = emp_first[
                (emp_first["Date"] == day) &
                (emp_first["Time"] > cutoff_time)
            ]

            if not late_row.empty:
                status = "Late"

            # Check Absent
            elif day not in punched_days:
                status = "Absent"

            # Get matching note
            note_match = next(
                (n["Note"] for n in notes_data
                 if n["Name"] == employee
                 and n["Date"] == str(day)
                 and n["Type"] == status),
                ""
            )

            report_rows.append({
                "Date": str(day),
                "Employee": employee,
                "Status": status,
                "Note": note_match
            })

    report_df = pd.DataFrame(report_rows)

    # Sort nicely
    report_df = report_df.sort_values(by=["Date", "Employee"])

    return report_df

 

# EXCEL FILE URL (DIRECT DOWNLOAD LINK)

file_path = "https://weldexperts-my.sharepoint.com/:x:/g/personal/reception_weldexperts_ca/IQDIifpuW4GRSapra5DshmwBAdIVkzSGdMBvmBkLwfAdpAA?download=1"


st.set_page_config(layout="wide")


@st.cache_data(ttl=300)
def load_data():
    try:
        response = requests.get(file_path)
        response.raise_for_status()
        return pd.read_excel(BytesIO(response.content), engine="openpyxl")
    except Exception as e:
        st.error(f"Error loading Excel file: {e}")
        return pd.DataFrame()

df = load_data()


if df.empty:
    st.stop()



st.title("Attendance Analysis")

# --- Cutoff time ---
cutoff_time = time(7, 31, 0)

# --- Format Date and Time ---
df["Date"] = pd.to_datetime(df["Date"]).dt.date
df["Time"] = pd.to_datetime(df["Time"], format="%H:%M:%S", errors="coerce").dt.time

# --- Sidebar Filters ---

today = datetime.today().date()
# Find last Monday
last_monday = today - timedelta(days=today.weekday() + 7)
# Last Friday
last_friday = last_monday + timedelta(days=4)

st.sidebar.header("Filters")
start_date = st.sidebar.date_input("Start date", last_monday)
end_date = st.sidebar.date_input("End date", last_friday)

employees = df["Name"].unique()
selected_employees = st.sidebar.multiselect(
    "Select Employee(s)",
    options=employees,
    default=list(employees)
)

# --- Apply Filters ---
filtered_df = df[
    (df["Date"] >= start_date) &
    (df["Date"] <= end_date) &
    (df["Name"].isin(selected_employees))
]

# --- First punch logic ---
first_punch_df = (
    filtered_df.sort_values(by=["Date", "Time"])
    .groupby(["Name", "Date"], as_index=False)
    .first()
)

first_punch_idx = (
    filtered_df.sort_values(by=["Name", "Date", "Time"])
    .groupby(["Name", "Date"])["Time"]
    .idxmin()
)

# --- Highlight late punches ---
def highlight_time(row):
    if row["Name"] == "Shawn":
        return ['' for _ in row.index]
    if row.name in first_punch_idx.values and row["Time"] > cutoff_time:
        return ['background-color: orange' if col == "Time" else '' for col in row.index]
    else:
        return ['' for _ in row.index]

# =========================
# ✅ SIDE BY SIDE LAYOUT
# =========================
col1, col2 = st.columns([3, 1], gap="large")

# --- Left: Attendance Table ---
with col1:
    st.subheader(f"{start_date} -> {end_date}")
    st.dataframe(filtered_df.style.apply(highlight_time, axis=1), use_container_width=True)
    

# =====================
# 📅 CALENDAR
# =====================
with col1:

    calendar_events = []

    # Late events
    for employee in selected_employees:
        if employee == "Shawn":
            continue
        emp_first = first_punch_df[first_punch_df["Name"] == employee]
        late_days = emp_first[emp_first["Time"] > cutoff_time]["Date"]

        for day in late_days:
            calendar_events.append({
                "start": str(day),
                "end": str(day),
                "title": f"{employee}: Late",
                "color": "orange"
            })

    # Absent events
    for employee in selected_employees:
        emp_data = filtered_df[filtered_df["Name"] == employee]
        all_days = pd.date_range(start=start_date, end=end_date).date
        weekdays = [d for d in all_days if d.weekday() < 5]

        punched_days = emp_data["Date"].unique()
        absent_days = [d for d in weekdays if d not in punched_days]

        for day in absent_days:
            calendar_events.append({
                "start": str(day),
                "end": str(day),
                "title": f"{employee}: Absent",
                "color": "red"
            })

    st.subheader("Late & Absent Calendar")
    calendar_response = calendar(events=calendar_events)

    # =========================
    # 🧠 STORE CLICKED EVENT (FIX)
    # =========================
    if calendar_response and calendar_response.get("eventClick"):

        event = calendar_response["eventClick"]["event"]

        st.session_state["selected_name"] = event["title"].split(": ")[0]
        st.session_state["selected_type"] = event["title"].split(": ")[1]
        st.session_state["selected_date"] = event["start"][:10]


# =========================
# 📝 NOTES SYSTEM (FIXED)
# =========================
with col1:

    st.subheader("Notes")

    if "selected_name" in st.session_state:

        name = st.session_state["selected_name"]
        event_type = st.session_state["selected_type"]
        date = st.session_state["selected_date"]

        st.write(f"Editing: **{name}** ({event_type}) on {date}")

        default_text = get_note(name, date, event_type)

        note = st.text_input(
            "Enter note",
            value=default_text,
            key="note_input"
        )

        if st.button("Save Note"):

            save_note(name, date, event_type, note)

            st.success("Note saved!")

# --- Right: Summary ---
with col2:
    st.subheader("Summary")

    for employee in sorted(selected_employees):
    
        emp_first = first_punch_df[first_punch_df["Name"] == employee]
        late_count = 0 if employee == "Shawn" else sum(emp_first["Time"] > cutoff_time)

        emp_data = filtered_df[filtered_df["Name"] == employee]
        all_days = pd.date_range(start=start_date, end=end_date).date
        weekdays = [d for d in all_days if d.weekday() < 5]

        punched_days = emp_data["Date"].unique()
        absent_days = [d for d in weekdays if d not in punched_days]
        absent_count = len(absent_days)

        st.markdown(f"""
        <div style="
            padding:10px;
            margin-bottom:10px;
            border-radius:10px;
            background-color:#f5f5f5;
        ">
            <b>{employee}</b><br>
            🟠 Late: {late_count}<br>
            🔴 Absent: {absent_count}
        </div>
        """, unsafe_allow_html=True)

        

if st.button("Generate Excel Report", key="report_btn"):

    conn = sqlite3.connect("notes.db")

    report_df = build_report(
        filtered_df,
        first_punch_df,
        cutoff_time,
        start_date,
        end_date,
        conn
    )

    conn.close()

    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        report_df.to_excel(writer, index=False, sheet_name="Attendance Report")

    st.download_button(
        label="⬇ Download Excel File",
        data=buffer.getvalue(),
        file_name="attendance_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )