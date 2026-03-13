import pandas as pd
import openpyxl

df = pd.read_excel(r"S:\Reception\Attendance\Punch-Clock-Attendance.xlsm", engine="openpyxl")
print(df)