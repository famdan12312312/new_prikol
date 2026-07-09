import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

f_name = 'individ_plan_Архипов.xlsx'
df = pd.read_excel(f_name, sheet_name='Сводные данные', header=None)

for r in range(df.shape[0]):
    row_vals = [f"Col {c}: {str(df.iloc[r, c]).strip()}" for c in range(df.shape[1]) if pd.notna(df.iloc[r, c])]
    print(f"Row {r}:", ", ".join(row_vals))
