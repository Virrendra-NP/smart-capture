import pandas as pd
import json

path = r'D:\JKD_Folder\JKD-PROJECT SITE\DPR\Mar_DPR\Weekly_Dashboard_JKD - DPR - 10.03.2026.xlsx'
xl = pd.ExcelFile(path)
print(f"Sheets: {xl.sheet_names}")

if 'Sheet1' in xl.sheet_names:
    df = pd.read_excel(path, sheet_name='Sheet1', header=None, nrows=20)
    print("First 20 rows of Sheet1:")
    print(df.to_string())
else:
    print("Sheet1 not found!")
