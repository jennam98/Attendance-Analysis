import pandas as pd
import openpyxl
import streamlit as st  
from datetime import datetime, timedelta, time

st.title("Attendance Analysis")

def load_data():
    return pd.read_excel(r"S:\Reception\Attendance\Punch-Clock-Attendance.xlsm", engine="openpyxl")

df = load_data()

# streamlit run attendance.py

#setting cut off punch in time
cutoff_time = time(7,31,0)


df["Date"] = pd.to_datetime(df["Date"]).dt.date
df["Time"] = pd.to_datetime(df["Time"], format="%H:%M:%S").dt.time

# --- User selects date range ---
st.sidebar.header("Filters")
start_date = st.sidebar.date_input("Start date", datetime.today().date())
end_date = st.sidebar.date_input("End date", datetime.today().date())
employees = df["Name"].unique()
selected_employees = st.sidebar.multiselect(
    "Select Employee(s)",
    options=employees,
    default=list(employees)  
)

# Filter by date range and employee
filtered_df = df[
    (df["Date"] >= start_date) &
    (df["Date"] <= end_date) &
    (df["Name"].isin(selected_employees))
]



sorted_df = filtered_df.sort_values(by=["Name", "Date", "Time"])

# Group by employee + date, get only first punch row indices
first_punch_idx = sorted_df.groupby(["Name", "Date"])["Time"].idxmin()

def highlight_first_late_consistent(row):
    # Only highlight if this row is the first punch of the day AND it's late
    if row.name in first_punch_idx.values and row["Time"] > cutoff_time:
        return ['background-color: red' if col == "Time" else '' for col in row.index]
    else:
        return ['' for _ in row.index]


# Display filtered data
st.subheader(f"Showing attendance from {start_date} to {end_date}")
st.dataframe(filtered_df.style.apply(highlight_first_late_consistent, axis=1))

first_punch_df = (
    filtered_df.sort_values(by=["Date", "Time"])  # make sure times are ascending
    .groupby(["Name", "Date"], as_index=False)   # group by employee + date
    .first()                                     # take first row in each group
)

st.subheader("Late Punch Summary")
for employee in selected_employees:
    emp_data = first_punch_df[first_punch_df["Name"] == employee]
    num_late = sum(emp_data["Time"] > cutoff_time)
    st.write(f"**{employee}:** {num_late} late punch{'s' if num_late != 1 else ''}")

# Optional: Export filtered data to Excel
if st.button("Export to Excel"):
    filtered_df.to_excel("filtered_attendance.xlsx", index=False)
    st.success("Filtered data exported to filtered_attendance.xlsx")