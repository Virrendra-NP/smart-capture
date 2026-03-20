import sys
import pandas as pd
import json

file_path = r'D:\JKD_Folder\JKD-PROJECT SITE\DPR\Excel_DPR\JKD - DPR - 28.02.2026 (1) (1) (1).xlsx'
try:
    df = pd.read_excel(file_path, sheet_name="Feb'26", header=None)
    data = {}
    for r in range(6):
        data[f"Row_{r}"] = [str(df.iloc[r, c]) for c in range(20)]
    
    with open('dpr_headers.json', 'w') as f:
        json.dump(data, f, indent=2)
except Exception as e:
    print(e)
