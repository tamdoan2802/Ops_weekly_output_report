import pandas as pd
import sys
import os

sys.stdout.reconfigure(encoding='utf-8')
base_dir = r"g:\My Drive\Dữ liệu nhân sự"
path = os.path.join(base_dir, "Data", "Attendance", "HR_Fact_Attendance.xlsx")
df = pd.read_excel(path, sheet_name='FACT_Attendance_Daily')
print("Columns:")
print(df.columns.tolist())

print("\nUnique values in 'Type of Date':")
if 'Type of Date' in df.columns:
    print(df['Type of Date'].unique())
else:
    print("Column not found")
