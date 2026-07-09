import parser_engine
from local_db import MongoClient
import os
import sys
import json

sys.stdout.reconfigure(encoding='utf-8')

f_name = 'individ_plan_Архипов.xlsx'
db = MongoClient()

print("Testing is_individual_plan_file:")
try:
    with open(f_name, "rb") as f:
        is_plan = parser_engine.is_individual_plan_file(f)
        print("is_individual_plan_file:", is_plan)
except Exception as e:
    print("Error in is_individual_plan_file:", e)

print("\nTesting parse_individual_plan_file:")
try:
    with open(f_name, "rb") as f:
        res = parser_engine.parse_individual_plan_file(f)
        # Convert objects to serializable form for printing
        print("Success! Parsed structure:")
        print("FIO:", res.get("fio"))
        print("Position:", res.get("position"))
        print("Rate / Conditions:", res.get("employment_conditions"))
        print("Degree:", res.get("degree"))
        print("Title:", res.get("title"))
        print("Contract:", res.get("contract"))
        print("Education:", res.get("education"))
        print("Number of loads:", len(res.get("loads", [])))
        if res.get("loads"):
            print("First 3 loads:")
            for l in res["loads"][:3]:
                print("  -", l)
except Exception as e:
    print("Error in parse_individual_plan_file:", e)
    import traceback
    traceback.print_exc()
