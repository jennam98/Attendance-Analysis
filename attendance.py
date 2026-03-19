import pandas as pd
import openpyxl
import streamlit as st  
from datetime import datetime, time, timedelta
from streamlit_calendar import calendar

st.title("Attendance Analysis")

# --- Load Excel Data ---
df = pd.read_excel(r"S:\Reception\Attendance\Punch-Clock-Attendance.xlsm", engine="openpyxl")

# --- Cutoff time for late punches ---
cutoff_time = time(7, 31, 0)

# --- Format Date and Time columns ---
df["Date"] = pd.to_datetime(df["Date"]).dt.date
df["Time"] = pd.to_datetime(df["Time"], format="%H:%M:%S", errors="coerce").dt.time

# --- Sidebar Filters ---
st.sidebar.header("Filters")
start_date = st.sidebar.date_input("Start date", datetime.today().date())
end_date = st.sidebar.date_input("End date", datetime.today().date())
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

# --- First punch logic for late ---
first_punch_df = filtered_df.sort_values(by=["Date", "Time"]).groupby(["Name", "Date"], as_index=False).first()
first_punch_idx = filtered_df.sort_values(by=["Name", "Date", "Time"]).groupby(["Name", "Date"])["Time"].idxmin()

# --- Function to highlight late first punches only ---
def highlight_time(row):
    if row.name in first_punch_idx.values and row["Time"] > cutoff_time:
        return ['background-color: orange' if col == "Time" else '' for col in row.index]
    else:
        return ['' for _ in row.index]

# --- Display Punch Log Table with late highlights ---
st.subheader(f"Showing attendance from {start_date} to {end_date}")
st.dataframe(filtered_df.style.apply(highlight_time, axis=1))

# --- Build Calendar Events ---
calendar_events = []

# Add late punches
for employee in selected_employees:
    emp_first = first_punch_df[first_punch_df["Name"] == employee]
    late_days = emp_first[emp_first["Time"] > cutoff_time]["Date"]
    for day in late_days:
        calendar_events.append({
            "start": str(day),
            "end": str(day),
            "title": f"{employee}: Late",
            "color" : "orange"
        })

# Add absent days based on weekday with no punches
for employee in selected_employees:
    emp_data = filtered_df[filtered_df["Name"] == employee]
    all_days = pd.date_range(start=start_date, end=end_date).date
    weekdays = [d for d in all_days if d.weekday() < 5]  # 0=Mon, ..., 4=Fri
    
    # Find days with no punches
    punched_days = emp_data["Date"].unique()
    absent_days = [d for d in weekdays if d not in punched_days]
    
    for day in absent_days:
        calendar_events.append({
            "start": str(day),
            "end": str(day),
            "title": f"{employee}: Absent",
            "color": "red"
        })

# --- Display Calendar ---
st.subheader("Late & Absent Calendar")
calendar(events=calendar_events)

