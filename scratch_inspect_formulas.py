import openpyxl

wb = openpyxl.load_workbook('individ_plan_Архипов.xlsx', data_only=False)
sheet = wb['1.1. ']

for r in [19, 32, 33, 34, 35]:
    row_vals = []
    for c in range(1, 18):
        cell = sheet.cell(row=r, column=c)
        if cell.value is not None:
            row_vals.append(f"{cell.coordinate}: {cell.value}")
    print(f"Row {r}:", ", ".join(row_vals))
