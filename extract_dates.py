import sys
import pandas as pd
from datetime import datetime

file_path = r'D:\JKD_Folder\JKD-PROJECT SITE\DPR\Excel_DPR\JKD - DPR - 28.02.2026 (1) (1) (1).xlsx'
try:
    df = pd.read_excel(file_path, sheet_name="Feb'26", header=None)
    date_cols = {}
    for c in range(3, df.shape[1]):
        val = df.iloc[4, c]
        if pd.notna(val):
            # Try to parse it as date or check if it's already datetime
            if isinstance(val, datetime):
                date_cols[c] = val.strftime('%d-%b-%Y')
            elif isinstance(val, str):
                try:
                    # try to parse str
                    parsed = pd.to_datetime(val)
                    date_cols[c] = parsed.strftime('%d-%b-%Y')
                except:
                    pass
    print("Found dates:")
    print(date_cols)
except Exception as e:
    print(e)
