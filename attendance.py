import pandas as pd
import openpyxl
import streamlit as st  
from datetime import datetime, timedelta, time

st.title("Attendance Analysis")

df = pd.read_excel(r"S:\Reception\Attendance\Punch-Clock-Attendance.xlsm", engine="openpyxl")

# print(df)

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


#Highlight Late Days Orange-----------------------------------------------
sorted_df = filtered_df.sort_values(by=["Name", "Date", "Time"])

# Group by employee + date, get only first punch row indices
first_punch_idx = sorted_df.groupby(["Name", "Date"])["Time"].idxmin()

def highlight_first_late_consistent(row):
    # Only highlight if this row is the first punch of the day AND it's late
    if row.name in first_punch_idx.values and row["Time"] > cutoff_time:
        return ['background-color: orange' if col == "Time" else '' for col in row.index]
    else:
        return ['' for _ in row.index]

first_punch_df = (
    filtered_df.sort_values(by=["Date", "Time"])  # make sure times are ascending
    .groupby(["Name", "Date"], as_index=False)   # group by employee + date
    .first()                                     # take first row in each group
)


# Display filtered data table-------------------------------------------------
st.subheader(f"Showing attendance from {start_date} to {end_date}")
st.dataframe(filtered_df.style.apply(highlight_first_late_consistent, axis=1))


st.subheader("Late Punch Summary")
for employee in selected_employees:
    emp_data = first_punch_df[first_punch_df["Name"] == employee]
    num_late = sum(emp_data["Time"] > cutoff_time)
    st.write(f"**{employee}:** {num_late} late punch{'s' if num_late != 1 else ''}")

st.subheader("Punch Error Summary")

for employee in selected_employees:
    emp_data = filtered_df[filtered_df["Name"] == employee]
    
    # Group by Date and count IN and OUT
    daily_counts = emp_data.groupby("Date")["Action"].value_counts().unstack(fill_value=0)
    
    # Ensure both 'IN' and 'OUT' columns exist
    if "IN" not in daily_counts.columns:
        daily_counts["IN"] = 0
    if "OUT" not in daily_counts.columns:
        daily_counts["OUT"] = 0
    
    # Find dates where counts are not exactly 2
    mispunched_dates = daily_counts[
        (daily_counts["IN"] != 2) | (daily_counts["OUT"] != 2)
    ].index.tolist()
    
    if mispunched_dates:
        mispunched_str = ", ".join([str(d) for d in mispunched_dates])
        st.write(f"**{employee}:** Punch Error on {mispunched_str}")
    else:
        st.write(f"**{employee}:** No punch errors")



# Optional: Export filtered data to Excel
if st.button("Export to Excel"):
    filtered_df.to_excel("filtered_attendance.xlsx", index=False)
    st.success("Filtered data exported to filtered_attendance.xlsx")