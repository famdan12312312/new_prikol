import re

filepath = "parser_engine.py"
with open(filepath, "r", encoding="utf-8") as f:
    content = f.read()

start_marker = "def parse_individual_plan_file("
end_marker = "def run_parsing_pipeline_dynamic("

start_idx = content.find(start_marker)
if start_idx == -1:
    print("Error: start marker not found")
    exit(1)

end_idx = content.find(end_marker)
if end_idx == -1:
    print("Error: end marker not found")
    exit(1)

# Let's extract the correct new code
new_function_code = """def parse_individual_plan_file(file_stream, target_fio=None):
    \"\"\"
    Парсит XLSX файл индивидуального плана.
    Динамически ищет:
      - Лист с профилем преподавателя (Ф.И.О., Должность, Ставка, Степень, Звание)
      - Лист с учебной нагрузкой (таблица с дисциплинами и часами)
    \"\"\"
    try:
        if hasattr(file_stream, "seek"):
            file_stream.seek(0)
        xls = pd.ExcelFile(file_stream)
        sheet_names = xls.sheet_names
        xls.close()
    except Exception as e:
        raise ValueError(f"Не удалось прочитать Excel-файл: {e}")

    # ====== 1. ПОИСК ЛИСТА С ПРОФИЛЕМ ПРЕПОДАВАТЕЛЯ ======
    profile_data = {
        "fio": "Не указан",
        "position": "преподаватель",
        "employment_conditions": "штатный",
        "degree": "",
        "title": "",
        "rate": "",
        "contract_number": "",
        "contract_date": "",
        "contract_duration": "",
        "edu_institution": "",
        "edu_year": "",
        "edu_specialty": "",
        "edu_qualification": ""
    }
    
    profile_found = False
    for meta_key, coord_info in STRICT_PLAN_COORDINATES.items():
        sheet_name = coord_info["sheet"]
        coord = coord_info["coordinate"]
        if sheet_name in sheet_names:
            try:
                if hasattr(file_stream, "seek"):
                    file_stream.seek(0)
                df_resolve = pd.read_excel(file_stream, sheet_name=sheet_name, header=None)
                if "общие" in sheet_name.lower():
                    df_resolve = df_resolve.T
                
                cell_val = get_cell_val_by_coord(df_resolve, coord)
                if cell_val:
                    val = str(cell_val).strip()
                    if meta_key == "employee_fio":
                        profile_data["fio"] = val
                        profile_found = True
                    elif meta_key == "employee_position":
                        profile_data["position"] = val
                    elif meta_key == "employee_rate":
                        profile_data["rate"] = val
                    elif meta_key == "employee_conditions":
                        profile_data["employment_conditions"] = val
                    elif meta_key == "employee_degree":
                        profile_data["degree"] = val
                    elif meta_key == "employee_title":
                        profile_data["title"] = val
                    elif meta_key == "employee_contract_num":
                        profile_data["contract_number"] = val
                    elif meta_key == "employee_contract_date":
                        profile_data["contract_date"] = val
                    elif meta_key == "employee_contract_duration":
                        profile_data["contract_duration"] = val
                    elif meta_key == "employee_edu_institution":
                        profile_data["edu_institution"] = val
                    elif meta_key == "employee_edu_year":
                        profile_data["edu_year"] = val
                    elif meta_key == "employee_edu_specialty":
                        profile_data["edu_specialty"] = val
                    elif meta_key == "employee_edu_qualification":
                        profile_data["edu_qualification"] = val
            except Exception:
                pass

    if not profile_found:
        for sheet in sheet_names:
            try:
                if hasattr(file_stream, "seek"):
                    file_stream.seek(0)
                df = pd.read_excel(file_stream, sheet_name=sheet, header=None)
                
                if df.shape[0] < 2 or df.shape[1] < 2:
                    continue
                
                # Ищем ключевые слова в любых позициях первых 20 строк
                for r in range(min(20, df.shape[0])):
                    for c in range(min(10, df.shape[1])):
                        cell_val = df.iloc[r, c]
                        if pd.isna(cell_val):
                            continue
                        cell_str = str(cell_val).strip().lower()
                        
                        # Ищем значение справа от ключевого слова (в колонках c+1..c+5)
                        def get_value_right(row, col_start):
                            for cc in range(col_start + 1, min(col_start + 6, df.shape[1])):
                                v = df.iloc[row, cc]
                                if pd.notna(v) and str(v).strip():
                                    return str(v).strip()
                            return ""
                        
                        if "ф.и.о" in cell_str or "фамилия" in cell_str:
                            val = get_value_right(r, c)
                            if val and len(val) > 2:
                                profile_data["fio"] = val
                                profile_found = True
                        elif "должность" in cell_str and "наименование" not in cell_str:
                            val = get_value_right(r, c)
                            if val:
                                profile_data["position"] = val
                        elif "размер ставки" in cell_str or "ставк" in cell_str:
                            val = get_value_right(r, c)
                            if val:
                                profile_data["rate"] = val
                        elif "условия привлечения" in cell_str or "условия" in cell_str:
                            val = get_value_right(r, c)
                            if val and "привлечения" not in val.lower():
                                profile_data["employment_conditions"] = val
                        elif "ученая степень" in cell_str:
                            val = get_value_right(r, c)
                            if val:
                                profile_data["degree"] = val
                        elif "ученое звание" in cell_str:
                            val = get_value_right(r, c)
                            if val:
                                profile_data["title"] = val
                        elif "договор" in cell_str or "контракт" in cell_str:
                            val = get_value_right(r, c)
                            if "номер" in cell_str or "№" in cell_str:
                                profile_data["contract_number"] = val
                            elif "дата" in cell_str or "от" in cell_str:
                                profile_data["contract_date"] = val
                            elif "срок" in cell_str:
                                profile_data["contract_duration"] = val
                        elif "окончил" in cell_str or "образовательн" in cell_str:
                            profile_data["edu_institution"] = get_value_right(r, c)
                        elif "год окончания" in cell_str or "год оконч" in cell_str:
                            profile_data["edu_year"] = get_value_right(r, c)
                        elif "специальност" in cell_str:
                            profile_data["edu_specialty"] = get_value_right(r, c)
                        elif "квалификаци" in cell_str:
                            profile_data["edu_qualification"] = get_value_right(r, c)
                
                if profile_found:
                    break
            except Exception:
                continue
    
    # Если нашли ставку, но не условия — используем ставку
    if profile_data["rate"] and profile_data["employment_conditions"] == "штатный":
        profile_data["employment_conditions"] = profile_data["rate"]
    # Если нашли оба — объединяем
    elif profile_data["rate"] and profile_data["employment_conditions"] != "штатный":
        profile_data["employment_conditions"] = f"{profile_data['rate']} ставки, {profile_data['employment_conditions']}"

    # Проверяем target_fio на соответствие
    if target_fio and profile_found:
        cleaned_target = clean_fio(target_fio).lower()
        cleaned_fio_val = clean_fio(profile_data["fio"]).lower()
    
    # ====== 2. ПОИСК ЛИСТА С НАГРУЗКОЙ ======
    loads = []
    
    for sheet in sheet_names:
        try:
            if hasattr(file_stream, "seek"):
                file_stream.seek(0)
            df = pd.read_excel(file_stream, sheet_name=sheet, header=None)
            
            if df.shape[0] < 2 or df.shape[1] < 2:
                continue
            
            # Ищем строку заголовка с ключевыми словами (лекции, группа, дисциплина)
            header_row = None
            for r in range(min(15, df.shape[0])):
                row_text = " ".join([str(df.iloc[r, c]).strip().lower() for c in range(df.shape[1]) if pd.notna(df.iloc[r, c])])
                has_disc = any(kw in row_text for kw in ["дисциплин", "наимен"])
                has_group = "групп" in row_text
                has_hours = any(kw in row_text for kw in ["лекц", "практич", "лабор", "всего"])
                if has_disc and (has_group or has_hours):
                    header_row = r
                    break
            
            # Если строгий заголовок не найден, ищем просто дисциплину/наименование
            if header_row is None:
                for r in range(min(15, df.shape[0])):
                    row_text = " ".join([str(df.iloc[r, c]).strip().lower() for c in range(df.shape[1]) if pd.notna(df.iloc[r, c])])
                    if any(kw in row_text for kw in ["дисциплин", "наимен"]):
                        header_row = r
                        break
            
            # Если все равно не нашли, не пропускаем лист! Ставим header_row = 0
            if header_row is None:
                header_row = 0
            
            # Маппинг колонок по ключевым словам (проверяем первые 15 строк)
            col_map = {}
            for c in range(df.shape[1]):
                parts = []
                for scan_r in range(min(15, df.shape[0])):
                    v = df.iloc[scan_r, c]
                    if pd.notna(v):
                        parts.append(str(v).strip().lower())
                val_str = " ".join(parts)
                
                if "дисциплин" in val_str or "наимен" in val_str:
                    col_map["subject"] = c
                elif "групп" in val_str:
                    col_map["group"] = c
                elif "семестр" in val_str:
                    col_map["semester"] = c
                elif "лекц" in val_str:
                    col_map["lectures"] = c
                elif any(kw in val_str for kw in ["лаб", "лабораторн"]):
                    col_map["laboratories"] = c
                elif "консульт" in val_str:
                    col_map["consultations"] = c
                elif "экзам" in val_str:
                    col_map["exams"] = c
                elif "зач" in val_str:
                    col_map["zachets"] = c
                elif any(kw in val_str for kw in ["курсов", "кп", "кр"]):
                    col_map["coursework"] = c
                elif "вкр" in val_str or "выпускн" in val_str:
                    col_map["vkr"] = c
                elif "гэк" in val_str or "гос" in val_str:
                    col_map["gek"] = c
                elif "дополн" in val_str or "доп" in val_str:
                    col_map["additional"] = c
                elif "всего" in val_str and "часов" in val_str:
                    col_map["total"] = c
                elif "всего" in val_str or "итого" in val_str:
                    if "total" not in col_map:
                        col_map["total"] = c
                elif any(kw in val_str for kw in ["руковод", "производств", "преддиплом"]):
                    col_map["practice"] = c
                elif val_str.strip() == "практика":
                    col_map["practice"] = c
                elif any(kw in val_str for kw in ["практич", "семин"]):
                    col_map["practicals"] = c
            
            if "subject" not in col_map:
                if df.shape[1] > 1:
                    col_map["subject"] = 1
                else:
                    col_map["subject"] = 0
                    
            data_start = header_row + 1
            for r in range(header_row + 1, min(header_row + 5, df.shape[0])):
                row_vals = [str(df.iloc[r, c]).strip() for c in range(min(5, df.shape[1])) if pd.notna(df.iloc[r, c])]
                row_str = " ".join(row_vals).lower()
                if all(v.isdigit() for v in row_vals if v):
                    data_start = r + 1
                    continue
                if "всего по видам" in row_str:
                    data_start = r + 1
                    continue
                if "осенний" in row_str or "весенний" in row_str:
                    data_start = r + 1
                    continue
                break
                
            def parse_hours_safe(val):
                if pd.isna(val):
                    return 0.0
                try:
                    return float(str(val).replace(',', '.').strip())
                except:
                    return 0.0
            
            def get_col_val(row_idx, key):
                if key in col_map and col_map[key] < df.shape[1]:
                    return df.iloc[row_idx, col_map[key]]
                return None
                
            # Парсим строки данных
            current_semester = "1"
            current_subject = ""
            current_group = ""
            
            for r in range(data_start, df.shape[0]):
                first_cell = df.iloc[r, col_map["subject"]] if col_map["subject"] < df.shape[1] else None
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
                    
                    group_val = get_col_val(r, "group")
                    if pd.notna(group_val) and str(group_val).strip():
                        current_group = str(group_val).strip()
                    else:
                        current_group = ""
                        
                    sem_val = get_col_val(r, "semester")
                    if pd.notna(sem_val) and str(sem_val).strip():
                        current_semester = str(sem_val).strip()
                
                if not current_subject:
                    continue
                    
                hours_dict = {
                    "lectures": parse_hours_safe(get_col_val(r, "lectures")),
                    "practicals": parse_hours_safe(get_col_val(r, "practicals")),
                    "laboratories": parse_hours_safe(get_col_val(r, "laboratories")),
                    "consultations": parse_hours_safe(get_col_val(r, "consultations")),
                    "exams": parse_hours_safe(get_col_val(r, "exams")),
                    "zachets": parse_hours_safe(get_col_val(r, "zachets")),
                    "coursework": parse_hours_safe(get_col_val(r, "coursework")),
                    "practice": parse_hours_safe(get_col_val(r, "practice")),
                    "vkr": parse_hours_safe(get_col_val(r, "vkr")),
                    "gek": parse_hours_safe(get_col_val(r, "gek")),
                    "additional": parse_hours_safe(get_col_val(r, "additional")),
                }
                
                total_hours = parse_hours_safe(get_col_val(r, "total"))
                if total_hours == 0.0:
                    total_hours = sum(hours_dict.values())
                    
                if total_hours == 0.0:
                    continue
                    
                row_group_val = get_col_val(r, "group")
                row_group = str(row_group_val).strip() if pd.notna(row_group_val) and str(row_group_val).strip() else current_group
                
                row_sem_val = get_col_val(r, "semester")
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
                        "direction": "09.03.01",
                        "semester": row_semester,
                        "hours": hours_dict,
                        "total": total_hours
                    })
        except Exception:
            continue

    return {
        "fio": profile_data["fio"],
        "employment_conditions": profile_data["employment_conditions"],
        "position": profile_data["position"],
        "degree": profile_data["degree"],
        "title": profile_data["title"],
        "contract": {
            "number": profile_data.get("contract_number", ""),
            "date": profile_data.get("contract_date", ""),
            "duration": profile_data.get("contract_duration", "")
        },
        "education": {
            "institution": profile_data.get("edu_institution", ""),
            "year": profile_data.get("edu_year", ""),
            "specialty": profile_data.get("edu_specialty", ""),
            "qualification": profile_data.get("edu_qualification", "")
        },
        "loads": loads
    }

"""

new_content = content[:start_idx] + new_function_code + "\n" + content[end_idx:]

with open(filepath, "w", encoding="utf-8") as f:
    f.write(new_content)

print("Replacement successful")
