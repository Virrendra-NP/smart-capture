import pandas as pd
import sys

path = r'D:\JKD_Folder\JKD-PROJECT SITE\DPR\Mar_DPR\Weekly_Dashboard_JKD - DPR - 10.03.2026.xlsx'
try:
    xl = pd.ExcelFile(path)
    with open('inspect_results.txt', 'w', encoding='utf-8') as f:
        f.write(f"Sheets: {xl.sheet_names}\n\n")
        if 'Sheet1' in xl.sheet_names:
            df = pd.read_excel(path, sheet_name='Sheet1', header=None, nrows=50)
            f.write("First 50 rows of Sheet1:\n")
            f.write(df.to_string())
        else:
            f.write("Sheet1 not found!")
except Exception as e:
    with open('inspect_results.txt', 'w', encoding='utf-8') as f:
        f.write(f"Error: {str(e)}")
