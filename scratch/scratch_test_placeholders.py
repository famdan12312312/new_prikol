import pandas as pd
import openpyxl
import os
import sys
from parser_engine import get_cell_val_by_coord, STRICT_PLAN_COORDINATES

sys.stdout.reconfigure(encoding='utf-8')

template_path = "individ_plan_Архипов.xlsx"

print("Reading placeholders from:", template_path)
for key, coord_info in STRICT_PLAN_COORDINATES.items():
    sheet_name = coord_info["sheet"]
    coord = coord_info["coordinate"]
    
    df_resolve = pd.read_excel(template_path, sheet_name=sheet_name, header=None)
    if "общие" in sheet_name.lower():
        df_resolve = df_resolve.T
        
    val = get_cell_val_by_coord(df_resolve, coord)
    print(f"{key} ({sheet_name} {coord}): '{val}'")
