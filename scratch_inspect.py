import pandas as pd
import os
import sys

# Ensure stdout uses utf-8 just in case
sys.stdout.reconfigure(encoding='utf-8')

files = [f for f in os.listdir('.') if f.endswith('.xlsx')]

with open("scratch_output.txt", "w", encoding="utf-8") as f_out:
    f_out.write(f"Found Excel files: {files}\n")
    for f_name in files:
        f_out.write(f"\n==================== FILE: {f_name} ====================\n")
        try:
            xls = pd.ExcelFile(f_name)
            f_out.write(f"Sheets: {xls.sheet_names}\n")
            for sheet in xls.sheet_names:
                df = pd.read_excel(f_name, sheet_name=sheet, header=None)
                f_out.write(f"  Sheet: {sheet}, Shape: {df.shape}\n")
                f_out.write("  Preview (first 10 rows, first 5 cols):\n")
                for r in range(min(10, df.shape[0])):
                    row_vals = [str(df.iloc[r, c]).strip() for c in range(min(5, df.shape[1])) if pd.notna(df.iloc[r, c])]
                    if any(row_vals):
                        f_out.write(f"    Row {r}: {row_vals}\n")
        except Exception as e:
            f_out.write(f"  Error reading {f_name}: {e}\n")
print("Done writing to scratch_output.txt")
