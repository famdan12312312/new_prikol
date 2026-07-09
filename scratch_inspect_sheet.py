import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

f_name = 'individ_plan_Архипов.xlsx'
df = pd.read_excel(f_name, sheet_name='1.1. ', header=None)

with open("scratch_sheet_1.1_all.txt", "w", encoding="utf-8") as f_out:
    f_out.write(f"Shape: {df.shape}\n")
    for r in range(df.shape[0]):
        row_vals = [f"Col {c}: {str(df.iloc[r, c]).strip()}" for c in range(df.shape[1]) if pd.notna(df.iloc[r, c])]
        f_out.write(f"Row {r}:\n  " + "\n  ".join(row_vals) + "\n")
print("Done writing all sheet details")
