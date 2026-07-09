filepath = "parser_engine.py"
with open(filepath, "r", encoding="utf-8") as f:
    content = f.read()

start_marker = "def parse_individual_plan_dynamic("
end_marker = "def run_profile_parsing_pipeline("

start_idx = content.find(start_marker)
if start_idx == -1:
    print("Error: start marker not found")
    exit(1)

end_idx = content.find(end_marker)
if end_idx == -1:
    print("Error: end marker not found")
    exit(1)

new_function_code = """def parse_individual_plan_dynamic(file_stream, mapping):
    \"\"\"
    Парсит XLSX файл индивидуального плана с использованием динамического маппинга от ИИ.
    \"\"\"
    try:
        if hasattr(file_stream, "seek"):
            file_stream.seek(0)
        xls = pd.ExcelFile(file_stream)
        sheet_names = xls.sheet_names
        xls.close()
    except Exception as e:
        raise ValueError(f"Не удалось прочитать Excel-файл: {e}")

    tp = mapping.get("teacher_profile", {}) or {}
    metadata = mapping.get("metadata", {}) or {}
    
    # Разворачиваем словари метаданных и профиля в строки
    def get_str_val(src, key, default=""):
        val = src.get(key)
        if isinstance(val, dict):
            return val.get("value") or default
        return str(val) if val is not None else default
        
    profile_data = {
        "fio": get_str_val(tp, "fio") or get_str_val(metadata, "employee_fio") or "Не указан",
        "position": get_str_val(tp, "position") or get_str_val(metadata, "employee_position") or "преподаватель",
        "employment_conditions": get_str_val(tp, "employment_conditions") or get_str_val(metadata, "employee_conditions") or get_str_val(metadata, "employee_rate") or "штатный",
        "degree": get_str_val(tp, "degree") or get_str_val(metadata, "employee_degree") or "",
        "title": get_str_val(tp, "title") or get_str_val(metadata, "employee_title") or "",
        "contract": {
            "number": get_str_val(metadata, "employee_contract_num"),
            "date": get_str_val(metadata, "employee_contract_date"),
            "duration": get_str_val(metadata, "employee_contract_duration")
        },
        "education": {
            "institution": get_str_val(metadata, "employee_edu_institution"),
            "year": get_str_val(metadata, "employee_edu_year"),
            "specialty": get_str_val(metadata, "employee_edu_specialty"),
            "qualification": get_str_val(metadata, "employee_edu_qualification")
        }
    }
    
    col_map = mapping.get("column_mapping", {})
    header_row = mapping.get("header_row_index", 0)
    
    loads = []
    
    active_cols = [c_val for c_val in col_map.values() if c_val is not None]
    max_col_idx = max(active_cols) if active_cols else 0

    def parse_hours_safe(val):
        if pd.isna(val):
            return 0.0
        try:
            return float(str(val).replace(',', '.').strip())
        except:
            return 0.0

    # Проходим по всем листам книги Excel индивидуального плана
    for s_name in sheet_names:
        try:
            if hasattr(file_stream, "seek"):
                file_stream.seek(0)
            df = pd.read_excel(file_stream, sheet_name=s_name, header=None)
        except Exception:
            continue
            
        if df.shape[0] <= header_row + 1:
            continue
            
        data_start = header_row + 1
        
        # Пропускаем строку с числовыми индексами колонок
        for r in range(header_row + 1, min(header_row + 5, df.shape[0])):
            row_vals = [str(df.iloc[r, c]).strip() for c in range(min(5, df.shape[1])) if pd.notna(df.iloc[r, c])]
            if all(v.isdigit() for v in row_vals if v):
                data_start = r + 1
                break
                
        current_semester = "1"
        s_lower = s_name.lower()
        if "весен" in s_lower or "2 семестр" in s_lower or "ii" in s_lower or "2сем" in s_lower:
            current_semester = "2"
        elif "осен" in s_lower or "1 семестр" in s_lower or "i" in s_lower or "1сем" in s_lower:
            current_semester = "1"
            
        sub_col = col_map.get("subject_name") or col_map.get("subject")
        if sub_col is not None and sub_col < df.shape[1]:
            current_subject = ""
            current_group = ""
            
            for r in range(data_start, df.shape[0]):
                first_cell = df.iloc[r, sub_col]
                is_empty_subject = pd.isna(first_cell) or not str(first_cell).strip()
                
                if not is_empty_subject:
                    sub_name = str(first_cell).strip()
                    sub_lower = sub_name.lower()
                    
                    if "осенний" in sub_lower:
                        current_semester = "1"
                        continue
                    if "весенний" in sub_lower:
                        current_semester = "2"
                        continue
                    if sub_lower.startswith("а)") or sub_lower.startswith("б)"):
                        if "осенн" in sub_lower:
                            current_semester = "1"
                        elif "весен" in sub_lower:
                            current_semester = "2"
                        continue
                    if sub_name.startswith("II.") or "научно-" in sub_lower or "организационно-" in sub_lower:
                        break
                    if any(kw in sub_lower for kw in ["итого", "всего", "фактически выполнено"]):
                        continue
                    if sub_name.isdigit():
                        continue
                        
                    current_subject = sub_name
                    
                    def get_col_val(key):
                        c_idx = col_map.get(key)
                        if c_idx is not None and c_idx < df.shape[1]:
                            return df.iloc[r, c_idx]
                        return None
                        
                    group_val = get_col_val("group_name") or get_col_val("group")
                    if pd.notna(group_val) and str(group_val).strip():
                        current_group = str(group_val).strip()
                    else:
                        current_group = ""
                        
                    sem_val = get_col_val("semester_number") or get_col_val("semester")
                    if pd.notna(sem_val) and str(sem_val).strip():
                        current_semester = str(sem_val).strip()
                
                if not current_subject:
                    continue
                    
                def get_row_col_val(key):
                    c_idx = col_map.get(key)
                    if c_idx is not None and c_idx < df.shape[1]:
                        return df.iloc[r, c_idx]
                    return None
                    
                hours_dict = {
                    "lectures": parse_hours_safe(get_row_col_val("hours_lectures") or get_row_col_val("lectures")),
                    "practicals": parse_hours_safe(get_row_col_val("hours_practicals") or get_row_col_val("practicals")),
                    "laboratories": parse_hours_safe(get_row_col_val("hours_laboratories") or get_row_col_val("laboratories")),
                    "consultations": parse_hours_safe(get_row_col_val("hours_consultations") or get_row_col_val("consultations")),
                    "exams": parse_hours_safe(get_row_col_val("hours_exams") or get_row_col_val("exams")),
                    "zachets": parse_hours_safe(get_row_col_val("hours_zachets") or get_row_col_val("zachets")),
                    "coursework": parse_hours_safe(get_row_col_val("hours_coursework") or get_row_col_val("coursework")),
                    "practice": parse_hours_safe(get_row_col_val("hours_practice") or get_row_col_val("practice")),
                    "vkr": parse_hours_safe(get_row_col_val("hours_vkr") or get_row_col_val("vkr")),
                    "gek": parse_hours_safe(get_row_col_val("hours_gek") or get_row_col_val("gek")),
                    "additional": parse_hours_safe(get_row_col_val("hours_additional") or get_row_col_val("additional")),
                }
                
                total_hours = parse_hours_safe(get_row_col_val("hours_total") or get_row_col_val("total"))
                if total_hours == 0.0:
                    total_hours = sum(hours_dict.values())
                    
                if total_hours == 0.0:
                    continue
                    
                row_group_val = get_row_col_val("group_name") or get_row_col_val("group")
                row_group = str(row_group_val).strip() if pd.notna(row_group_val) and str(row_group_val).strip() else current_group
                
                row_sem_val = get_row_col_val("semester_number") or get_row_col_val("semester")
                row_semester = str(row_sem_val).strip() if pd.notna(row_sem_val) and str(row_sem_val).strip() else current_semester
                
                merged = False
                for existing in loads:
                    if (existing["subject"].lower().strip() == current_subject.lower().strip() and
                        existing["group"].lower().strip() == row_group.lower().strip() and
                        existing["semester"].lower().strip() == row_semester.lower().strip()):
                        
                        for h_key in hours_dict:
                            existing["hours"][h_key] += hours_dict[h_key]
                        existing["total"] += total_hours
                        merged = True
                        break
                        
                if not merged:
                    loads.append({
                        "subject": current_subject,
                        "group": row_group,
                        "direction": metadata.get("department_direction", "09.03.01"),
                        "semester": row_semester,
                        "hours": hours_dict,
                        "total": total_hours
                    })
                    
    return {
        "fio": profile_data["fio"],
        "employment_conditions": profile_data["employment_conditions"],
        "position": profile_data["position"],
        "degree": profile_data["degree"],
        "title": profile_data["title"],
        "contract": profile_data["contract"],
        "education": profile_data["education"],
        "loads": loads
    }

"""

new_content = content[:start_idx] + new_function_code + "\n" + content[end_idx:]

with open(filepath, "w", encoding="utf-8") as f:
    f.write(new_content)

print("Dynamic replacement successful")
