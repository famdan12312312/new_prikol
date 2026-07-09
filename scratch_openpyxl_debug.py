import openpyxl
import sys

sys.stdout.reconfigure(encoding='utf-8')

wb = openpyxl.load_workbook('individ_plan_Архипов.xlsx', data_only=True)
sheet = wb['1.1. ']

print("Merged ranges:")
for r in sheet.merged_cells.ranges:
    print(r)

print("\nAll cells with values:")
for row in sheet.iter_rows(min_row=1, max_row=35, min_col=1, max_col=17):
    row_vals = []
    for cell in row:
        if cell.value is not None:
            row_vals.append(f"{cell.coordinate}: {cell.value}")
    if row_vals:
        print(", ".join(row_vals))
