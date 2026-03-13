import pandas as pd
import openpyxl
import streamlit as st  
from datetime import datetime, timedelta, time

st.title("Attendance Analysis")

df = pd.read_excel(r"S:\Reception\Attendance\Punch-Clock-Attendance.xlsm", engine="openpyxl")
# print(df)

# streamlit

df["Date"] = pd.to_datetime(df["Date"]).dt.date
df["Time"] = pd.to_datetime(df["Time"], format="%H:%M:%S").dt.time

# --- User selects date range ---
st.sidebar.header("Filters")
start_date = st.sidebar.date_input("Start date", df["Date"].min())
end_date = st.sidebar.date_input("End date", df["Date"].max())
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

# Display filtered data
st.subheader(f"Showing attendance from {start_date} to {end_date}")
st.dataframe(filtered_df)

# Optional: Export filtered data to Excel
if st.button("Export to Excel"):
    filtered_df.to_excel("filtered_attendance.xlsx", index=False)
    st.success("Filtered data exported to filtered_attendance.xlsx")