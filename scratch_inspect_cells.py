import pandas as pd
import sys

sys.stdout.reconfigure(encoding='utf-8')

f_name = 'individ_plan_Архипов.xlsx'
df = pd.read_excel(f_name, sheet_name='1.1. ', header=None)

with open("scratch_cells_debug.txt", "w", encoding="utf-8") as f_out:
    # Print columns header first
    f_out.write("Col indices: " + "\t".join(f"Col{i}" for i in range(df.shape[1])) + "\n")
    for r in range(4, 21):
        vals = []
        for c in range(df.shape[1]):
            v = df.iloc[r, c]
            vals.append(str(v) if pd.notna(v) else "")
        f_out.write(f"Row {r:02d}: " + "\t".join(vals) + "\n")
print("Done")
