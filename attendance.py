import pandas as pd
import openpyxl
import streamlit as st  
from datetime import datetime, time, timedelta
from streamlit_calendar import calendar
import os

file_path = r"S:\Reception\Attendance\Punch-Clock-Attendance.xlsm"

# --- Load Excel Data ---
df = pd.read_excel(file_path, engine="openpyxl")

st.set_page_config(layout="wide")

# --- Last refreshed timestamp ---
last_modified = os.path.getmtime(file_path)
last_refreshed = datetime.fromtimestamp(last_modified)
st.caption(f"Punch Log Data Last Updated: {last_refreshed.strftime('%Y-%m-%d %I:%M:%S %p')}")

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


    st.dataframe(filtered_df.style.apply(highlight_time, axis=1))
    
    
# =========================
# 📅 CALENDAR
# =========================

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
# 📝 NOTES SYSTEM
# =========================
with col1:
    if calendar_response and calendar_response.get("eventClick"):
        event = calendar_response["eventClick"]["event"]

        title = event["title"]
        date = event["start"][:10]

        name, event_type = title.split(": ")

        st.subheader(f"Add/View Note for {name} ({event_type}) on {date}")

        # Load notes safely
        if not os.path.exists("notes.csv") or os.stat("notes.csv").st_size == 0:
            notes_df = pd.DataFrame(columns=["Name", "Date", "Type", "Note"])
        else:
            notes_df = pd.read_csv("notes.csv")

        # Existing note
        existing_note = notes_df[
            (notes_df["Name"] == name) &
            (notes_df["Date"] == date) &
            (notes_df["Type"] == event_type)
        ]

        default_text = existing_note.iloc[0]["Note"] if not existing_note.empty else ""

        note = st.text_input("Enter note", value=default_text)

        if st.button("Save Note"):
            notes_df = notes_df[
                ~(
                    (notes_df["Name"] == name) &
                    (notes_df["Date"] == date) &
                    (notes_df["Type"] == event_type)
                )
            ]

            new_row = pd.DataFrame([{
                "Name": name,
                "Date": date,
                "Type": event_type,
                "Note": note
            }])

            notes_df = pd.concat([notes_df, new_row], ignore_index=True)
            notes_df.to_csv("notes.csv", index=False)

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