import os
import re
import datetime
import math
import pandas as pd

STRICT_PLAN_COORDINATES = {
    "employee_fio": {"coordinate": "D3", "sheet": "Общие сведения"},
    "employee_fio_title": {"coordinate": "A31", "sheet": "Тит. лист"},
    "employee_position": {"coordinate": "D4", "sheet": "Общие сведения"},
    "employee_position_title": {"coordinate": "A34", "sheet": "Тит. лист"},
    "employee_rate": {"coordinate": "D5", "sheet": "Общие сведения"},
    "employee_conditions": {"coordinate": "D6", "sheet": "Общие сведения"},
    "employee_degree": {"coordinate": "D7", "sheet": "Общие сведения"},
    "employee_title": {"coordinate": "D8", "sheet": "Общие сведения"},
    "employee_contract": {"coordinate": "D10", "sheet": "Общие сведения"},
    "employee_contract_duration": {"coordinate": "D11", "sheet": "Общие сведения"},
    "employee_edu_inst_year": {"coordinate": "D12", "sheet": "Общие сведения"},
    "employee_edu_specialty_1": {"coordinate": "D14", "sheet": "Общие сведения"},
    "employee_edu_specialty_2": {"coordinate": "D15", "sheet": "Общие сведения"},
    "employee_edu_qualification": {"coordinate": "D18", "sheet": "Общие сведения"},
    "department_name": {"coordinate": "A25", "sheet": "Тит. лист"},
    "department_head": {"coordinate": "F16", "sheet": "Тит. лист"},
    "institute_name": {"coordinate": "A28", "sheet": "Тит. лист"},
    "study_year": {"coordinate": "A37", "sheet": "Тит. лист"}
}

def clean_fio(name):
    """
    Очистка и форматирование ФИО.
    """
    if not name or not isinstance(name, str):
        return ""
    name = re.sub(r'\s+', ' ', name).strip()
    return name

def find_similar_teacher(parsed_fio, db):
    """
    Ищет существующего преподавателя в БД по сходству ФИО.
    Например, сопоставляет "Обухов" или "Обухов А.Д." с "Обухов Андрей Дмитриевич".
    """
    if not parsed_fio:
        return None
        
    cleaned_parsed = clean_fio(parsed_fio)
    if not cleaned_parsed:
        return None
        
    # Сначала проверяем точное совпадение
    if hasattr(db, "employees") and hasattr(db.employees, "find_one") and callable(db.employees.find_one):
        try:
            exact = db.employees.find_one({"fio": cleaned_parsed})
            if exact:
                return exact
        except:
            pass
            
    # Загружаем всех преподавателей из БД для поиска
    all_teachers = []
    if hasattr(db, "employees") and hasattr(db.employees, "find") and callable(db.employees.find):
        try:
            all_teachers = list(db.employees.find())
        except:
            pass
            
    if not all_teachers:
        return None
        
    # Сначала проверяем совпадение по сохраненным псевдонимам (aliases)
    for teacher in all_teachers:
        aliases = teacher.get("aliases", [])
        cleaned_aliases = [clean_fio(a).lower() for a in aliases if a]
        if cleaned_parsed.lower() in cleaned_aliases:
            return teacher
            
    cleaned_parsed_norm = cleaned_parsed.replace(".", " ")
    parsed_words = [w.strip().lower() for w in cleaned_parsed_norm.split() if w.strip()]
    if not parsed_words:
        return None
        
    parsed_lastname = parsed_words[0]
    
    matches = []
    for teacher in all_teachers:
        teacher_fio = teacher.get("fio", "")
        teacher_fio_norm = teacher_fio.replace(".", " ")
        teacher_words = [w.strip().lower() for w in teacher_fio_norm.split() if w.strip()]
        if not teacher_words:
            continue
        teacher_lastname = teacher_words[0]
        
        # Если совпадает фамилия
        if parsed_lastname == teacher_lastname:
            if len(parsed_words) == 1:
                # Если в пришедшем файле только фамилия (например, "Обухов")
                matches.append(teacher)
            else:
                # Проверяем совместимость по остальным словам (именам/инициалам)
                compatible = True
                for i in range(1, min(len(parsed_words), len(teacher_words))):
                    p_w = parsed_words[i]
                    t_w = teacher_words[i]
                    if len(p_w) == 1 or len(t_w) == 1:
                        if p_w[0] != t_w[0]:
                            compatible = False
                            break
                    else:
                        if p_w != t_w:
                            compatible = False
                            break
                if compatible:
                    matches.append(teacher)
                    
    # Если нашлось ровно одно совпадение, возвращаем его
    if len(matches) == 1:
        return matches[0]
        
    return None

def is_individual_plan_file(file_stream):
    """
    Автоматически определяет, является ли файл индивидуальным планом преподавателя
    или же файлом общей кафедральной нагрузки (САПР).
    """
    try:
        if hasattr(file_stream, "seek"):
            file_stream.seek(0)
        xls = pd.ExcelFile(file_stream)
        sheet_names = xls.sheet_names
        xls.close()
    except Exception:
        return False

    # 1. Если имя файла содержит "сапр" (sapr), это точно кафедральный отчет САПР
    filename = getattr(file_stream, "name", "")
    if filename:
        filename_lower = filename.lower()
        if any(kw in filename_lower for kw in ["сапр", "sapr"]):
            return False

    # 2. Если есть специфические листы, характерные для индивидуальных планов
    has_profile_sheets = any(any(kw in s.lower() for kw in ["общие сведения", "тит. лист", "титульн"]) for s in sheet_names)
    if has_profile_sheets:
        return True

    # 3. Сканируем листы на наличие профильных ключевых слов
    profile_keywords_count = 0
    has_fio = False

    for s in sheet_names:
        try:
            if hasattr(file_stream, "seek"):
                file_stream.seek(0)
            df_temp = pd.read_excel(file_stream, sheet_name=s, nrows=15, header=None)
            if df_temp.shape[0] >= 2 and df_temp.shape[1] >= 2:
                for r in range(min(15, df_temp.shape[0])):
                    for c in range(min(10, df_temp.shape[1])):
                        val = str(df_temp.iloc[r, c]).strip().lower() if pd.notna(df_temp.iloc[r, c]) else ""
                        if "ф.и.о" in val or ("фамилия" in val and "имя" in val):
                            has_fio = True
                        if any(kw in val for kw in ["должность", "условия привлечения", "ученая степень", "ученое звание", "договор", "контракт"]):
                            profile_keywords_count += 1
        except Exception:
            continue

    return has_fio and profile_keywords_count >= 2


def run_parsing_pipeline(file_stream, db):
    """
    Основной конвейер парсинга загруженного Excel-файла распределения нагрузки кафедры.
    Записывает данные в эмулированную базу данных SQLite.
    Колонки жестко зафиксированы по спецификации пользователя:
    A (0) - Номер
    B (1) - Наименование дисциплины / вида работы
    C (2) - Группа
    D (3) - Специальность/Направление
    E (4) - Специализация/Программа/Профиль
    F (5) - Контингент
    G (6) - Семестр
    H (7) - Поток
    I (8) - Лекции
    J (9) - Практические занятия
    K (10) - Лабораторные работы
    L (11) - Консультации
    M (12) - Экзамен
    N (13) - Зачет
    O (14) - КП, КР
    P (15) - Практика
    Q (16) - ВКР
    R (17) - ГЭК
    S (18) - Дополнительная работа
    T (19) - Итого:
    U (20) - Преподаватель (ФИО)
    """
    try:
        if hasattr(file_stream, "seek"):
            file_stream.seek(0)
        xls = pd.ExcelFile(file_stream)
        sheet_names = xls.sheet_names
        xls.close()
    except Exception as e:
        raise ValueError(f"Не удалось прочитать Excel-файл: {e}")
        
    # Проверяем, является ли загруженный файл индивидуальным планом преподавателя
    is_individual_plan = is_individual_plan_file(file_stream)


    if is_individual_plan:
        parsed = parse_individual_plan_file(file_stream)
        parsed_fio = parsed["fio"]
        parsed_conditions = parsed["employment_conditions"]
        parsed_position = parsed["position"]
        parsed_degree = parsed["degree"]
        parsed_title = parsed["title"]
        parsed_loads = parsed["loads"]
        
        # Автозамена фамилии на ФИО с помощью find_similar_teacher
        similar_teacher = find_similar_teacher(parsed_fio, db)
        if similar_teacher:
            parsed_fio = similar_teacher["fio"]
            emp_id = similar_teacher["_id"]
            db_loads = similar_teacher.get("loads", [])
        else:
            exact = None
            if hasattr(db, "employees") and hasattr(db.employees, "find_one") and callable(db.employees.find_one):
                try:
                    exact = db.employees.find_one({"fio": parsed_fio})
                except:
                    pass
            if exact:
                emp_id = exact["_id"]
                db_loads = exact.get("loads", [])
            else:
                emp_doc = {"fio": parsed_fio, "loads": []}
                if hasattr(db, "employees") and hasattr(db.employees, "insert_one") and callable(db.employees.insert_one):
                    emp_id = db.employees.insert_one(emp_doc).inserted_id
                else:
                    from local_db import ObjectId
                    emp_id = ObjectId()
                db_loads = []
                
        # Слияние нагрузок
        active_loads = list(db_loads)
        for nl in parsed_loads:
            found_idx = -1
            for idx, al in enumerate(active_loads):
                if (str(al.get("subject", "")).strip().lower() == str(nl["subject"]).strip().lower() and 
                    str(al.get("group", "")).strip().lower() == str(nl["group"]).strip().lower() and 
                    str(al.get("semester", "")).strip().lower() == str(nl["semester"]).strip().lower()):
                    found_idx = idx
                    break
            if found_idx != -1:
                active_loads[found_idx] = nl
            else:
                active_loads.append(nl)
                
        # Обновляем профиль и сохраняем историю
        if hasattr(db, "employees") and hasattr(db.employees, "update_one") and callable(db.employees.update_one):
            db.employees.update_one(
                {"_id": emp_id},
                {"$set": {
                    "loads": active_loads,
                    "position": parsed_position,
                    "employment_conditions": parsed_conditions,
                    "degree": parsed_degree,
                    "title": parsed_title
                }}
            )
        return {parsed_fio: active_loads}

    employees_cache = {}  # FIO -> ObjectId
    employees_loads = {}  # ObjectId -> list of loads
    parsed_count = 0
    
    def parse_hours(val):
        if pd.isna(val):
            return 0.0
        try:
            val_str = str(val).replace(',', '.').strip()
            return float(val_str)
        except:
            return 0.0

    # Проходим по всем листам книги Excel без исключения
    for target_sheet in sheet_names:
        if hasattr(file_stream, "seek"):
            file_stream.seek(0)
        try:
            df = pd.read_excel(file_stream, sheet_name=target_sheet, header=None)
        except Exception:
            continue
            
        # Поиск строки с заголовками, после которой начинаются данные
        header_idx = None
        for idx, row in df.iterrows():
            row_vals = [str(v).lower() for v in row.values if pd.notna(v)]
            row_str = " ".join(row_vals)
            if any(kw in row_str for kw in ["наименование дисциплины", "лекции", "лабораторные работы"]):
                header_idx = idx
                break
                
        if header_idx is None:
            # Если не нашли, ищем первую строку с заполненными ячейками
            for idx, row in df.iterrows():
                filled_cells = [v for v in row.values if pd.notna(v) and len(str(v).strip()) > 1]
                if len(filled_cells) >= 3:
                    header_idx = idx
                    break
                    
        if header_idx is None:
            header_idx = 0
            
        # Строки с данными находятся после строки заголовков
        data_df = df.iloc[header_idx + 1:]
        
        for _, row in data_df.iterrows():
            vals = row.values
            
            # Проверяем, что строка достаточно длинная
            if len(vals) < 20:
                continue
                
            # Дисциплина находится в столбце B (индекс 1)
            sub_name_raw = vals[1]
            if pd.isna(sub_name_raw) or not str(sub_name_raw).strip():
                continue
            sub_name = str(sub_name_raw).strip()
            if any(kw in sub_name.lower() for kw in ["итого", "всего", "все учебные работы"]):
                continue
            # Пропускаем строки-разделители (только подчёркивания, тире, точки)
            if all(c in "_-=. " for c in sub_name):
                continue
                
            # Группа находится в столбце C (индекс 2)
            grp_name = str(vals[2]).strip() if pd.notna(vals[2]) else "Неизвестная"
            
            # Специальность/Направление находится в столбце D (индекс 3)
            dir_code = str(vals[3]).strip() if pd.notna(vals[3]) else "09.03.01"
            
            # Семестр находится в столбце G (индекс 6)
            sem_val = str(vals[6]).strip() if pd.notna(vals[6]) else "1"
            
            # Преподаватель
            teacher_name = ""
            if len(vals) >= 21 and pd.notna(vals[20]):
                teacher_name = str(vals[20]).strip()
                
            teacher_name = clean_fio(teacher_name)
            
            # Фильтрация мусорных "имён" преподавателей (итоговые строки, разделители и т.п.)
            def is_garbage_teacher(name):
                if not name or name == "Не указан" or name.isdigit() or len(name) < 2:
                    return True
                name_lower = name.lower().strip()
                # Итоговые строки
                if any(kw in name_lower for kw in ["итого", "всего", "___", "---"]):
                    return True
                # Строки из одних символов-разделителей
                if all(c in "_-=. " for c in name):
                    return True
                # Слишком короткое ФИО (менее 3 символов без пробелов)
                if len(name.replace(" ", "")) < 3:
                    return True
                # Категории САПР и агрегаты
                garbage_keywords = [
                    "нагрузка", "семестр", "бакалавриат", "магистратура", 
                    "аспирантура", "специалитет", "заведующий", "начальник",
                    "декан", "ректор", "проректор", "заместитель"
                ]
                if any(kw in name_lower for kw in garbage_keywords):
                    return True
                # Оканчивается на двоеточие — это заголовок раздела
                if name.endswith(":"):
                    return True
                # Чисто аббревиатура группы (3-4 заглавных буквы, без строчных)
                cleaned = name.replace(" ", "")
                if len(cleaned) <= 5 and cleaned.isalpha() and cleaned.isupper():
                    return True
                # Содержит скобки — это скорее всего название дисциплины
                if "(" in name and ")" in name:
                    return True
                # Название дисциплины, а не ФИО
                discipline_keywords = [
                    "практика", "подготовка", "нормоконтроль", "выпускн",
                    "квалификацион", "курсов", "дисциплин"
                ]
                if any(kw in name_lower for kw in discipline_keywords):
                    return True
                # Слишком длинное "имя" — ФИО обычно не длиннее 50 символов
                if len(name) > 50:
                    return True
                # ФИО должно начинаться с заглавной кириллической буквы
                if name and name[0].isalpha():
                    first_word = name.split()[0]
                    if first_word and not first_word[0].isupper():
                        return True
                return False
            
            # Доп. проверка: если "имя преподавателя" совпадает с названием предмета — это не ФИО
            if teacher_name and teacher_name.lower() == sub_name.lower():
                teacher_name = ""
            
            if is_garbage_teacher(teacher_name):
                if len(vals) >= 21 and (not teacher_name or teacher_name == "Не указан" or teacher_name.isdigit() or len(teacher_name) < 2):
                    # Попробуем сделать поиск с конца в качестве фоллбека
                    fallback_name = "Не указан"
                    for v in reversed(vals):
                        if pd.notna(v):
                            v_str = str(v).strip()
                            if v_str and len(v_str) >= 2:
                                try:
                                    float(v_str.replace(',', '.'))
                                    continue
                                except ValueError:
                                    if any(char.isdigit() for char in v_str):
                                        continue
                                    if not is_garbage_teacher(v_str):
                                        fallback_name = v_str
                                        break
                    if fallback_name != "Не указан" and not is_garbage_teacher(fallback_name):
                        teacher_name = clean_fio(fallback_name)
                    else:
                        teacher_name = ""
                else:
                    teacher_name = ""
                
            hours_dict = {
                "lectures": parse_hours(vals[8]),
                "practicals": parse_hours(vals[9]),
                "laboratories": parse_hours(vals[10]),
                "consultations": parse_hours(vals[11]),
                "exams": parse_hours(vals[12]),
                "zachets": parse_hours(vals[13]),
                "coursework": parse_hours(vals[14]),
                "practice": parse_hours(vals[15]),
                "vkr": parse_hours(vals[16]),
                "gek": parse_hours(vals[17]),
                "additional": parse_hours(vals[18])
            }
            
            total_hours = parse_hours(vals[19])
            if total_hours == 0.0:
                total_hours = sum(hours_dict.values())
        
            if teacher_name:
                similar_teacher = find_similar_teacher(teacher_name, db)
                if similar_teacher:
                    teacher_name = similar_teacher["fio"]
                if teacher_name not in employees_cache:
                    existing_emp = None
                    if hasattr(db.employees, "find_one") and callable(db.employees.find_one):
                        try:
                            existing_emp = db.employees.find_one({"fio": teacher_name})
                        except Exception:
                            pass
                    
                    if existing_emp:
                        emp_id = existing_emp["_id"]
                        employees_cache[teacher_name] = emp_id
                        employees_loads[emp_id] = existing_emp.get("loads", [])
                    else:
                        emp_doc = {
                            "fio": teacher_name,
                            "loads": []
                        }
                        emp_id = db.employees.insert_one(emp_doc).inserted_id
                        employees_cache[teacher_name] = emp_id
                        employees_loads[emp_id] = []
                else:
                    emp_id = employees_cache[teacher_name]
                    
                load_entry = {
                    "subject": sub_name,
                    "group": grp_name,
                    "direction": dir_code,
                    "semester": sem_val,
                    "hours": hours_dict,
                    "total": total_hours
                }
                
                found_idx = -1
                for idx, existing_load in enumerate(employees_loads[emp_id]):
                    match_subj = str(existing_load.get("subject", "")).strip().lower() == str(load_entry["subject"]).strip().lower()
                    match_group = str(existing_load.get("group", "")).strip().lower() == str(load_entry["group"]).strip().lower()
                    match_sem = str(existing_load.get("semester", "")).strip().lower() == str(load_entry["semester"]).strip().lower()
                    if match_subj and match_group and match_sem:
                        found_idx = idx
                        break
                        
                if found_idx != -1:
                    employees_loads[emp_id][found_idx] = load_entry
                else:
                    employees_loads[emp_id].append(load_entry)
            else:
                unassigned_entry = {
                    "subject": sub_name,
                    "group": grp_name,
                    "direction": dir_code,
                    "semester": sem_val,
                    "hours": hours_dict,
                    "total": total_hours
                }
                
                if hasattr(db, "unassigned_loads"):
                    existing_unassigned = None
                    if hasattr(db.unassigned_loads, "find_one") and callable(db.unassigned_loads.find_one):
                        try:
                            existing_unassigned = db.unassigned_loads.find_one({
                                "subject": sub_name,
                                "group": grp_name,
                                "semester": sem_val
                            })
                        except Exception:
                            pass
                    
                    if existing_unassigned:
                        db.unassigned_loads.update_one(
                            {"_id": existing_unassigned["_id"]},
                            {"$set": unassigned_entry}
                        )
                    else:
                        db.unassigned_loads.insert_one(unassigned_entry)
                
        parsed_count += 1

    # Сохраняем собранную нагрузку преподавателям в БД
    result_loads = {}
    for fio, emp_id in employees_cache.items():
        loads = employees_loads.get(emp_id, [])
        db.employees.update_one({"_id": emp_id}, {"$set": {"loads": loads}})
        result_loads[fio] = loads
        
    print(f"Парсинг успешно завершен. Обработано строк: {parsed_count}")
    return result_loads


def parse_individual_plan_file(file_stream, target_fio=None):
    """
    Парсит XLSX файл индивидуального плана.
    Динамически ищет:
      - Лист с профилем преподавателя (Ф.И.О., Должность, Ставка, Степень, Звание)
      - Лист с учебной нагрузкой (таблица с дисциплинами и часами)
    """
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
        "fio_title": "",
        "position": "преподаватель",
        "position_title": "",
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
        "edu_qualification": "",
        "department_name": "",
        "institute_name": "",
        "year": "",
        "specialty_1": "",
        "specialty_2": ""
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
                    elif meta_key == "employee_fio_title":
                        profile_data["fio_title"] = val
                    elif meta_key == "employee_position":
                        profile_data["position"] = val
                    elif meta_key == "employee_position_title":
                        profile_data["position_title"] = val
                    elif meta_key == "employee_rate":
                        profile_data["rate"] = val
                    elif meta_key == "employee_conditions":
                        profile_data["employment_conditions"] = val
                    elif meta_key == "employee_degree":
                        profile_data["degree"] = val
                    elif meta_key == "employee_title":
                        profile_data["title"] = val
                    elif meta_key == "employee_contract":
                        val_str = str(val).strip()
                        if "от" in val_str:
                            parts = val_str.split("от")
                            profile_data["contract_number"] = parts[0].replace("№", "").replace("номер", "").strip()
                            profile_data["contract_date"] = parts[1].strip()
                        else:
                            profile_data["contract_number"] = val_str
                    elif meta_key == "employee_contract_duration":
                        profile_data["contract_duration"] = val
                    elif meta_key == "employee_edu_inst_year":
                        val_str = str(val).strip()
                        year_match = re.search(r'\b(\d{4})\b', val_str)
                        if year_match:
                            profile_data["edu_year"] = year_match.group(1)
                            inst_part = val_str.replace(year_match.group(1), "").strip()
                            inst_part = re.sub(r'[\s,г\.уодед]+$', '', inst_part).strip()
                            profile_data["edu_institution"] = inst_part
                        else:
                            profile_data["edu_institution"] = val_str
                    elif meta_key == "employee_edu_specialty_1":
                        profile_data["specialty_1"] = val
                    elif meta_key == "employee_edu_specialty_2":
                        profile_data["specialty_2"] = val
                    elif meta_key == "employee_edu_qualification":
                        profile_data["edu_qualification"] = val
                    elif meta_key == "department_name":
                        profile_data["department_name"] = val
                    elif meta_key == "institute_name":
                        profile_data["institute_name"] = val
                    elif meta_key == "study_year":
                        year_str = str(val).strip()
                        years = re.findall(r'\b\d{4}\b', year_str)
                        if len(years) >= 2:
                            profile_data["year"] = f"{years[0]}/{years[1]}"
                        elif len(years) == 1:
                            profile_data["year"] = years[0]
                        else:
                            profile_data["year"] = year_str
            except Exception:
                pass

    # Объединение специальности
    spec1 = profile_data.get("specialty_1") or ""
    spec2 = profile_data.get("specialty_2") or ""
    if spec1 and spec2:
        profile_data["edu_specialty"] = f"{spec1} {spec2}".strip()
    elif spec1:
        profile_data["edu_specialty"] = spec1
    elif spec2:
        profile_data["edu_specialty"] = spec2

    # Объединение и дедупликация ФИО
    fio_gen = profile_data.get("fio") or ""
    fio_title = profile_data.get("fio_title") or ""
    fio_gen_clean = clean_fio(fio_gen)
    fio_title_clean = clean_fio(fio_title)
    
    if fio_gen_clean and fio_title_clean:
        if fio_gen_clean.lower() == fio_title_clean.lower():
            profile_data["fio"] = fio_gen_clean
        else:
            profile_data["fio"] = fio_gen_clean if len(fio_gen_clean) >= len(fio_title_clean) else fio_title_clean
        profile_found = True
    elif fio_title_clean:
        profile_data["fio"] = fio_title_clean
        profile_found = True
    elif fio_gen_clean:
        profile_data["fio"] = fio_gen_clean
        profile_found = True

    # Объединение и дедупликация Должности
    pos_gen = profile_data.get("position") or ""
    pos_title = profile_data.get("position_title") or ""
    pos_gen_clean = pos_gen.strip()
    pos_title_clean = pos_title.strip()
    
    if pos_gen_clean and pos_title_clean:
        if pos_gen_clean.lower() == pos_title_clean.lower():
            profile_data["position"] = pos_gen_clean
        else:
            profile_data["position"] = pos_gen_clean if len(pos_gen_clean) >= len(pos_title_clean) else pos_title_clean
    elif pos_title_clean:
        profile_data["position"] = pos_title_clean
    elif pos_gen_clean:
        profile_data["position"] = pos_gen_clean

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
        "department_name": profile_data.get("department_name", ""),
        "institute_name": profile_data.get("institute_name", ""),
        "year": profile_data.get("year", ""),
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


def run_parsing_pipeline_dynamic(file_stream, mapping, db, fio_resolutions=None):
    """
    Динамический парсинг кафедрального отчета на основе ИИ-маппинга.
    Поддерживает полную структуру сущностей БД (кафедры, группы, студенты, практики, предметы).
    """
    if fio_resolutions is None:
        fio_resolutions = {}

    try:
        if hasattr(file_stream, "seek"):
            file_stream.seek(0)
        xls = pd.ExcelFile(file_stream)
        sheet_names = xls.sheet_names
        xls.close()
    except Exception as e:
        raise ValueError(f"Не удалось прочитать Excel-файл: {e}")

    col_map = mapping.get("column_mapping", {})
    header_idx = mapping.get("header_row_index", 0)
    metadata = mapping.get("metadata", {}) or {}
    
    # Разворачиваем метаданные (если они представлены в виде словарей с координатами) в простые строки
    flat_metadata = {}
    for k, v in metadata.items():
        if isinstance(v, dict):
            flat_metadata[k] = v.get("value") or ""
        else:
            flat_metadata[k] = str(v) if v is not None else ""
    metadata = flat_metadata
    
    # Применяем резолюции ФИО к метаданным
    meta_fio = metadata.get("employee_fio") or ""
    if meta_fio in fio_resolutions:
        metadata["employee_fio"] = fio_resolutions[meta_fio]
    
    # Сохраняем информацию о кафедре
    dep_name = metadata.get("department_name")
    if dep_name:
        dep_doc = {
            "name": dep_name,
            "head": metadata.get("department_head", ""),
            "direction": metadata.get("department_direction", "")
        }
        if hasattr(db, "departments"):
            try:
                existing_dep = db.departments.find_one({"name": dep_name})
                if existing_dep:
                    db.departments.update_one({"_id": existing_dep["_id"]}, {"$set": dep_doc})
                else:
                    db.departments.insert_one(dep_doc)
            except Exception:
                pass

    # Сохраняем информацию об институте
    inst_name = metadata.get("institute_name")
    if inst_name:
        inst_doc = {
            "name": inst_name,
            "director": metadata.get("institute_director", "")
        }
        if hasattr(db, "institutes"):
            try:
                existing_inst = db.institutes.find_one({"name": inst_name})
                if existing_inst:
                    db.institutes.update_one({"_id": existing_inst["_id"]}, {"$set": inst_doc})
                else:
                    db.institutes.insert_one(inst_doc)
            except Exception:
                pass
    
    employees_cache = {}  # FIO -> ObjectId
    employees_loads = {}  # ObjectId -> list of loads
    parsed_count = 0

    active_cols = [c_val for c_val in col_map.values() if c_val is not None]
    max_col_idx = max(active_cols) if active_cols else 0

    def parse_hours_safe(row_vals, col_key):
        c_idx = col_map.get(col_key)
        if c_idx is not None and c_idx < len(row_vals):
            val = row_vals[c_idx]
            if pd.notna(val):
                try:
                    return float(str(val).replace(',', '.').strip())
                except:
                    pass
        return 0.0

    def get_col_val(row_vals, col_key, default=""):
        c_idx = col_map.get(col_key)
        if c_idx is not None and c_idx < len(row_vals):
            val = row_vals[c_idx]
            if pd.notna(val):
                return str(val).strip()
        return default

    def get_col_val_both(row_vals, new_key, old_key, default=""):
        val = get_col_val(row_vals, new_key)
        if not val:
            val = get_col_val(row_vals, old_key)
        return val or default

    # Вспомогательная функция фильтрации мусора
    def is_garbage_teacher(name):
        if not name or name == "Не указан" or name.isdigit() or len(name) < 2:
            return True
        name_lower = name.lower().strip()
        if any(kw in name_lower for kw in ["итого", "всего", "___", "---"]):
            return True
        if all(c in "_-=. " for c in name):
            return True
        if len(name.replace(" ", "")) < 3:
            return True
        garbage_keywords = [
            "нагрузка", "семестр", "бакалавриат", "магистратура", 
            "аспирантура", "специалитет", "заведующий", "начальник",
            "декан", "ректор", "проректор", "заместитель"
        ]
        if any(kw in name_lower for kw in garbage_keywords):
            return True
        if name.endswith(":"):
            return True
        cleaned = name.replace(" ", "")
        if len(cleaned) <= 5 and cleaned.isalpha() and cleaned.isupper():
            return True
        if "(" in name and ")" in name:
            return True
        discipline_keywords = [
            "практика", "подготовка", "нормоконтроль", "выпускн",
            "квалификацион", "курсов", "дисциплин"
        ]
        if any(kw in name_lower for kw in discipline_keywords):
            return True
        if len(name) > 50:
            return True
        if name and name[0].isalpha():
            first_word = name.split()[0]
            if first_word and not first_word[0].isupper():
                return True
        return False

    # Перебираем ВСЕ листы файла нагрузки
    for s_name in sheet_names:
        try:
            if hasattr(file_stream, "seek"):
                file_stream.seek(0)
            df = pd.read_excel(file_stream, sheet_name=s_name, header=None)
        except Exception:
            continue
            
        if df.shape[0] <= header_idx + 1:
            continue
            
        data_df = df.iloc[header_idx + 1:]

        for _, row in data_df.iterrows():
            vals = row.values
                
            sub_name = get_col_val_both(vals, "subject_name", "subject")
            if not sub_name:
                continue
            if any(kw in sub_name.lower() for kw in ["итого", "всего", "все учебные работы"]):
                continue
            if all(c in "_-=. " for c in sub_name):
                continue
                
            grp_name = get_col_val_both(vals, "group_name", "group", "Неизвестная")
            sem_val = get_col_val_both(vals, "semester_number", "semester", "1")
            
            teacher_name = get_col_val_both(vals, "teacher_fio", "teacher", "")
            teacher_name = clean_fio(teacher_name)
            
            if teacher_name and teacher_name.lower() == sub_name.lower():
                teacher_name = ""
                
            if is_garbage_teacher(teacher_name):
                # Попробуем сделать поиск с конца в качестве фоллбека
                fallback_name = "Не указан"
                for v in reversed(vals):
                    if pd.notna(v):
                        v_str = str(v).strip()
                        if v_str and len(v_str) >= 2:
                            try:
                                float(v_str.replace(',', '.'))
                                continue
                            except ValueError:
                                if any(char.isdigit() for char in v_str):
                                    continue
                                if not is_garbage_teacher(v_str):
                                    fallback_name = v_str
                                    break
                if fallback_name != "Не указан" and not is_garbage_teacher(fallback_name):
                    teacher_name = clean_fio(fallback_name)
                else:
                    teacher_name = ""
                    
            # Применяем резолюцию ФИО
            if teacher_name in fio_resolutions:
                teacher_name = fio_resolutions[teacher_name]
                
            # Если в строке нет ФИО, но документ является планом конкретного преподавателя
            if not teacher_name and metadata.get("employee_fio"):
                teacher_name = metadata.get("employee_fio")
                
            hours_dict = {
                "lectures": parse_hours_safe(vals, "hours_lectures") or parse_hours_safe(vals, "lectures"),
                "practicals": parse_hours_safe(vals, "hours_practicals") or parse_hours_safe(vals, "practicals"),
                "laboratories": parse_hours_safe(vals, "hours_laboratories") or parse_hours_safe(vals, "laboratories"),
                "consultations": parse_hours_safe(vals, "hours_consultations") or parse_hours_safe(vals, "consultations"),
                "exams": parse_hours_safe(vals, "hours_exams") or parse_hours_safe(vals, "exams"),
                "zachets": parse_hours_safe(vals, "hours_zachets") or parse_hours_safe(vals, "zachets"),
                "coursework": parse_hours_safe(vals, "hours_coursework") or parse_hours_safe(vals, "coursework"),
                "practice": parse_hours_safe(vals, "hours_practice") or parse_hours_safe(vals, "practice"),
                "vkr": parse_hours_safe(vals, "hours_vkr") or parse_hours_safe(vals, "vkr"),
                "gek": parse_hours_safe(vals, "hours_gek") or parse_hours_safe(vals, "gek"),
                "additional": parse_hours_safe(vals, "hours_additional") or parse_hours_safe(vals, "additional")
            }
            
            total_hours = parse_hours_safe(vals, "hours_total") or parse_hours_safe(vals, "total")
            if total_hours == 0.0:
                total_hours = sum(hours_dict.values())
                
            # Сохранение группы
            if grp_name and grp_name != "Неизвестная":
                course_match = re.search(r'-(\d)', grp_name)
                course_val = int(course_match.group(1)) if course_match else 1
                grp_doc = {
                    "name": grp_name,
                    "course": course_val,
                    "direction": metadata.get("department_direction", "09.03.01")
                }
                if hasattr(db, "groups"):
                    try:
                        existing_grp = db.groups.find_one({"name": grp_name})
                        if existing_grp:
                            db.groups.update_one({"_id": existing_grp["_id"]}, {"$set": grp_doc})
                        else:
                            db.groups.insert_one(grp_doc)
                    except Exception:
                        pass
                        
            # Сохранение студентов
            stud_fio = get_col_val(vals, "student_fio", "")
            stud_profile = get_col_val(vals, "student_profile", "")
            if stud_fio:
                stud_doc = {
                    "fio": stud_fio,
                    "group": grp_name,
                    "profile": stud_profile
                }
                if hasattr(db, "students"):
                    try:
                        existing_stud = db.students.find_one({"fio": stud_fio, "group": grp_name})
                        if not existing_stud:
                            db.students.insert_one(stud_doc)
                    except Exception:
                        pass
                        
            # Сохранение практики
            pract_type = get_col_val(vals, "practice_type", "")
            pract_org = get_col_val(vals, "practice_org", "")
            pract_kind = get_col_val(vals, "practice_kind", "")
            if pract_type or pract_org or pract_kind or metadata.get("practice_order_date") or mapping.get("document_type") == "practice_order":
                p_doc = {
                    "employee_fio": teacher_name or metadata.get("employee_fio", ""),
                    "group": grp_name,
                    "order_date": metadata.get("practice_order_date", ""),
                    "order_signer": metadata.get("practice_order_signer", ""),
                    "type": pract_type or metadata.get("practice_type", ""),
                    "org": pract_org or metadata.get("practice_org", ""),
                    "kind": pract_kind or metadata.get("practice_kind", "")
                }
                if hasattr(db, "practices"):
                    try:
                        existing_prac = db.practices.find_one({
                            "employee_fio": p_doc["employee_fio"],
                            "group": p_doc["group"],
                            "type": p_doc["type"]
                        })
                        if not existing_prac:
                            db.practices.insert_one(p_doc)
                    except Exception:
                        pass
                        
            # Сохранение предмета
            if sub_name:
                sub_doc = {
                    "name": sub_name,
                    "rpd": "",
                    "guidelines": ""
                }
                if hasattr(db, "subjects"):
                    try:
                        existing_sub = db.subjects.find_one({"name": sub_name})
                        if not existing_sub:
                            db.subjects.insert_one(sub_doc)
                    except Exception:
                        pass
                        
            if teacher_name:
                similar_teacher = find_similar_teacher(teacher_name, db)
                if similar_teacher:
                    teacher_name = similar_teacher["fio"]
                    
                if teacher_name not in employees_cache:
                    existing_emp = None
                    if hasattr(db.employees, "find_one") and callable(db.employees.find_one):
                        try:
                            existing_emp = db.employees.find_one({"fio": teacher_name})
                        except Exception:
                            pass
                            
                    if existing_emp:
                        emp_id = existing_emp["_id"]
                        employees_cache[teacher_name] = emp_id
                        employees_loads[emp_id] = existing_emp.get("loads", [])
                    else:
                        emp_doc = {
                            "fio": teacher_name,
                            "position": metadata.get("employee_position", "") or "преподаватель",
                            "employment_conditions": metadata.get("employee_conditions") or metadata.get("employee_rate", "") or "штатный",
                            "degree": metadata.get("employee_degree", ""),
                            "title": metadata.get("employee_title", ""),
                            "contract": {
                                "number": metadata.get("employee_contract_num", ""),
                                "date": metadata.get("employee_contract_date", ""),
                                "duration": metadata.get("employee_contract_duration", "")
                            },
                            "education": {
                                "institution": metadata.get("employee_edu_institution", ""),
                                "year": metadata.get("employee_edu_year", ""),
                                "specialty": metadata.get("employee_edu_specialty", ""),
                                "qualification": metadata.get("employee_edu_qualification", "")
                            },
                            "loads": []
                        }
                        emp_id = db.employees.insert_one(emp_doc).inserted_id
                        employees_cache[teacher_name] = emp_id
                        employees_loads[emp_id] = []
                else:
                    emp_id = employees_cache[teacher_name]
                    
                load_entry = {
                    "subject": sub_name,
                    "group": grp_name,
                    "direction": metadata.get("department_direction", "09.03.01"),
                    "semester": sem_val,
                    "hours": hours_dict,
                    "total": total_hours
                }
                
                # Слияние нагрузок
                found_idx = -1
                for idx, existing_load in enumerate(employees_loads[emp_id]):
                    match_subj = str(existing_load.get("subject", "")).strip().lower() == str(load_entry["subject"]).strip().lower()
                    match_group = str(existing_load.get("group", "")).strip().lower() == str(load_entry["group"]).strip().lower()
                    match_sem = str(existing_load.get("semester", "")).strip().lower() == str(load_entry["semester"]).strip().lower()
                    if match_subj and match_group and match_sem:
                        found_idx = idx
                        break
                        
                if found_idx != -1:
                    employees_loads[emp_id][found_idx] = load_entry
                else:
                    employees_loads[emp_id].append(load_entry)
            else:
                unassigned_entry = {
                    "subject": sub_name,
                    "group": grp_name,
                    "direction": metadata.get("department_direction", "09.03.01"),
                    "semester": sem_val,
                    "hours": hours_dict,
                    "total": total_hours
                }
                
                if hasattr(db, "unassigned_loads"):
                    existing_unassigned = None
                    if hasattr(db.unassigned_loads, "find_one") and callable(db.unassigned_loads.find_one):
                        try:
                            existing_unassigned = db.unassigned_loads.find_one({
                                "subject": sub_name,
                                "group": grp_name,
                                "semester": sem_val
                            })
                        except Exception:
                            pass
                            
                    if existing_unassigned:
                        db.unassigned_loads.update_one(
                            {"_id": existing_unassigned["_id"]},
                            {"$set": unassigned_entry}
                        )
                    else:
                        db.unassigned_loads.insert_one(unassigned_entry)
                        
            parsed_count += 1

    # Сохраняем собранную нагрузку и метаданные преподавателям в БД
    result_loads = {}
    for fio, emp_id in employees_cache.items():
        loads = employees_loads.get(emp_id, [])
        update_doc = {
            "loads": loads
        }
        
        # Обновляем профиль при наличии новых данных в метаданных
        for f, meta_k in [("position", "employee_position"), 
                          ("employment_conditions", "employee_conditions"), 
                          ("degree", "employee_degree"), 
                          ("title", "employee_title")]:
            if metadata.get(meta_k):
                update_doc[f] = metadata.get(meta_k)
                
        contract_doc = {}
        for f, meta_k in [("number", "employee_contract_num"), 
                          ("date", "employee_contract_date"), 
                          ("duration", "employee_contract_duration")]:
            if metadata.get(meta_k):
                contract_doc[f] = metadata.get(meta_k)
        if contract_doc:
            update_doc["contract"] = contract_doc

        edu_doc = {}
        for f, meta_k in [("institution", "employee_edu_institution"), 
                          ("year", "employee_edu_year"), 
                          ("specialty", "employee_edu_specialty"), 
                          ("qualification", "employee_edu_qualification")]:
            if metadata.get(meta_k):
                edu_doc[f] = metadata.get(meta_k)
        if edu_doc:
            update_doc["education"] = edu_doc

        db.employees.update_one({"_id": emp_id}, {"$set": update_doc})
        result_loads[fio] = loads
        
    print(f"Динамический парсинг успешно завершен. Обработано строк: {parsed_count}")
    return result_loads


def get_cell_val_by_coord(df, coord_str):
    if not coord_str or not isinstance(coord_str, str):
        return None
        
    coord_str = coord_str.strip().upper()
    # Если датафрейм транспонирован (строк мало), переводим старые координаты в реальные transposed координаты!
    if df is not None and df.shape[0] < 15:
        coord_map = {
            "D3": "C4",      # employee_fio (col D, row 3 -> row 4, col C)
            "D4": "D4",      # employee_position
            "D5": "E4",      # employee_rate
            "D6": "F4",      # employee_conditions
            "D7": "G4",      # employee_degree
            "D8": "H4",      # employee_title
            "D10": "J4",     # employee_contract
            "D11": "K4",     # employee_contract_duration
            "D12": "L4",     # employee_edu_inst_year
            "D14": "N4",     # employee_edu_specialty_1
            "D15": "O4",     # employee_edu_specialty_2
            "D18": "R4"      # employee_edu_qualification
        }
        coord_str = coord_map.get(coord_str, coord_str)
        
    import re
    match = re.match(r'^([A-Z]+)(\d+)$', coord_str)
    if not match:
        return None
    col_str, row_str = match.groups()
    col_idx = 0
    for char in col_str:
        col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
    col_idx -= 1
    row_idx = int(row_str) - 1
    if row_idx < df.shape[0] and col_idx < df.shape[1]:
        val = df.iloc[row_idx, col_idx]
        if pd.notna(val):
            return str(val).strip()
    return None

def resolve_meta_value(df, meta_item, ignore_fallback=True):
    """
    Разрешает значение метаданных: если есть координата, считывает её из df.
    """
    if not meta_item:
        return ""
        
    coordinate = None
    fallback_val = ""
    
    # 1. Если это словарь (например, из сохраненного обучения)
    if isinstance(meta_item, dict):
        coordinate = meta_item.get("coordinate")
        if not ignore_fallback:
            fallback_val = meta_item.get("value") or ""
    # 2. Если это строка
    elif isinstance(meta_item, str):
        meta_item = meta_item.strip()
        import re
        match = re.search(r'\((?:found in cell|cell)?\s*([A-Za-z]+\d+)\)', meta_val if 'meta_val' in locals() else meta_item)
        if match:
            coordinate = match.group(1)
            fallback_val = meta_item[:match.start()].strip()
        else:
            if re.match(r'^[A-Za-z]+\d+$', meta_item):
                coordinate = meta_item
            else:
                fallback_val = meta_item
                
    if coordinate and df is not None:
        val = get_cell_val_by_coord(df, coordinate)
        if val is not None:
            return val
            
    return fallback_val if not ignore_fallback else ""


def parse_individual_plan_dynamic(file_stream, mapping):
    """
    Парсит XLSX файл индивидуального плана с использованием динамического маппинга от ИИ.
    """
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
        
    fio_gen = get_str_val(tp, "fio") or get_str_val(metadata, "employee_fio") or ""
    fio_title = get_str_val(metadata, "employee_fio_title") or ""
    fio_gen_clean = clean_fio(fio_gen)
    fio_title_clean = clean_fio(fio_title)
    
    final_fio = "Не указан"
    if fio_gen_clean and fio_title_clean:
        if fio_gen_clean.lower() == fio_title_clean.lower():
            final_fio = fio_gen_clean
        else:
            final_fio = fio_gen_clean if len(fio_gen_clean) >= len(fio_title_clean) else fio_title_clean
    elif fio_title_clean:
        final_fio = fio_title_clean
    elif fio_gen_clean:
        final_fio = fio_gen_clean

    pos_gen = get_str_val(tp, "position") or get_str_val(metadata, "employee_position") or ""
    pos_title = get_str_val(metadata, "employee_position_title") or ""
    pos_gen_clean = pos_gen.strip()
    pos_title_clean = pos_title.strip()
    
    final_position = "преподаватель"
    if pos_gen_clean and pos_title_clean:
        if pos_gen_clean.lower() == pos_title_clean.lower():
            final_position = pos_gen_clean
        else:
            final_position = pos_gen_clean if len(pos_gen_clean) >= len(pos_title_clean) else pos_title_clean
    elif pos_title_clean:
        final_position = pos_title_clean
    elif pos_gen_clean:
        final_position = pos_gen_clean

    # Извлечение и чистка года
    raw_year = get_str_val(metadata, "study_year")
    cleaned_yr = ""
    if raw_year:
        years = re.findall(r'\b\d{4}\b', raw_year)
        if len(years) >= 2:
            cleaned_yr = f"{years[0]}/{years[1]}"
        elif len(years) == 1:
            cleaned_yr = years[0]
        else:
            cleaned_yr = raw_year

    profile_data = {
        "fio": final_fio,
        "position": final_position,
        "employment_conditions": get_str_val(tp, "employment_conditions") or get_str_val(metadata, "employee_conditions") or get_str_val(metadata, "employee_rate") or "штатный",
        "degree": get_str_val(tp, "degree") or get_str_val(metadata, "employee_degree") or "",
        "title": get_str_val(tp, "title") or get_str_val(metadata, "employee_title") or "",
        "department_name": get_str_val(metadata, "department_name"),
        "institute_name": get_str_val(metadata, "institute_name"),
        "year": cleaned_yr,
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
        "department_name": profile_data.get("department_name", ""),
        "institute_name": profile_data.get("institute_name", ""),
        "year": profile_data.get("year", ""),
        "contract": profile_data["contract"],
        "education": profile_data["education"],
        "loads": loads
    }


def process_resolved_metadata(meta_dict):
    """
    Постобработка словаря извлеченных метаданных:
    1. Слияние и дедупликация ФИО (employee_fio и employee_fio_title)
    2. Слияние и дедупликация Должности (employee_position и employee_position_title)
    3. Разделение контракта (employee_contract -> contract_num и contract_date)
    4. Разделение образования (employee_edu_inst_year -> edu_institution и edu_year)
    5. Объединение специальности (employee_edu_specialty_1 и specialty_2 -> edu_specialty)
    6. Стандартизация года (study_year)
    """
    # Вспомогательные функции получения значений
    def get_val(key):
        item = meta_dict.get(key)
        if isinstance(item, dict):
            return item.get("value") or ""
        return str(item) if item is not None else ""

    def set_val(key, val, base_key_for_coord=None):
        if base_key_for_coord and base_key_for_coord in meta_dict and isinstance(meta_dict[base_key_for_coord], dict):
            meta_dict[key] = {
                "value": val,
                "coordinate": meta_dict[base_key_for_coord].get("coordinate") or "",
                "sheet": meta_dict[base_key_for_coord].get("sheet") or ""
            }
            meta_dict[key] = {"value": val, "coordinate": "", "sheet": ""}

    # 1. ФИО
    fio_gen = get_val("employee_fio")
    fio_title = get_val("employee_fio_title")
    fio_gen_clean = clean_fio(fio_gen)
    fio_title_clean = clean_fio(fio_title)
    final_fio = ""
    if fio_gen_clean and fio_title_clean:
        if fio_gen_clean.lower() == fio_title_clean.lower():
            final_fio = fio_gen_clean
        else:
            final_fio = fio_gen_clean if len(fio_gen_clean) >= len(fio_title_clean) else fio_title_clean
    elif fio_title_clean:
        final_fio = fio_title_clean
    elif fio_gen_clean:
        final_fio = fio_gen_clean
    
    if final_fio:
        set_val("employee_fio", final_fio, "employee_fio")
    if "employee_fio_title" in meta_dict:
        del meta_dict["employee_fio_title"]

    # 2. Должность
    pos_gen = get_val("employee_position")
    pos_title = get_val("employee_position_title")
    pos_gen_clean = pos_gen.strip()
    pos_title_clean = pos_title.strip()
    final_pos = ""
    if pos_gen_clean and pos_title_clean:
        if pos_gen_clean.lower() == pos_title_clean.lower():
            final_pos = pos_gen_clean
        else:
            final_pos = pos_gen_clean if len(pos_gen_clean) >= len(pos_title_clean) else pos_title_clean
    elif pos_title_clean:
        final_pos = pos_title_clean
    elif pos_gen_clean:
        final_pos = pos_gen_clean
    
    if final_pos:
        set_val("employee_position", final_pos, "employee_position")
    if "employee_position_title" in meta_dict:
        del meta_dict["employee_position_title"]

    # 3. Контракт (D10 -> num & date)
    if "employee_contract" in meta_dict:
        contract_raw = get_val("employee_contract")
        num_val = ""
        date_val = ""
        if contract_raw:
            if "от" in contract_raw:
                parts = contract_raw.split("от")
                num_val = parts[0].replace("№", "").replace("номер", "").strip()
                date_val = parts[1].strip()
            else:
                num_val = contract_raw
        
        set_val("employee_contract_num", num_val, "employee_contract")
        set_val("employee_contract_date", date_val, "employee_contract")
        del meta_dict["employee_contract"]

    # 4. Образование (D12 -> inst & year)
    if "employee_edu_inst_year" in meta_dict:
        edu_raw = get_val("employee_edu_inst_year")
        inst_val = edu_raw
        year_val = ""
        if edu_raw:
            year_match = re.search(r'\b(\d{4})\b', edu_raw)
            if year_match:
                year_val = year_match.group(1)
                inst_part = edu_raw.replace(year_match.group(1), "").strip()
                inst_part = re.sub(r'[\s,г\.уодед]+$', '', inst_part).strip()
                inst_val = inst_part
        
        set_val("employee_edu_institution", inst_val, "employee_edu_inst_year")
        set_val("employee_edu_year", year_val, "employee_edu_inst_year")
        del meta_dict["employee_edu_inst_year"]

    # 5. Специальность (D14 + D15)
    if "employee_edu_specialty_1" in meta_dict or "employee_edu_specialty_2" in meta_dict:
        spec1 = get_val("employee_edu_specialty_1")
        spec2 = get_val("employee_edu_specialty_2")
        final_spec = f"{spec1} {spec2}".strip()
        
        set_val("employee_edu_specialty", final_spec, "employee_edu_specialty_1")
        
        if "employee_edu_specialty_1" in meta_dict:
            del meta_dict["employee_edu_specialty_1"]
        if "employee_edu_specialty_2" in meta_dict:
            del meta_dict["employee_edu_specialty_2"]

    # 6. Учебный год (A37)
    if "study_year" in meta_dict:
        raw_yr = get_val("study_year")
        cleaned_yr = ""
        if raw_yr:
            years = re.findall(r'\b\d{4}\b', str(raw_yr))
            if len(years) >= 2:
                cleaned_yr = f"{years[0]}/{years[1]}"
            elif len(years) == 1:
                cleaned_yr = years[0]
            else:
                cleaned_yr = str(raw_yr)
        
        set_val("study_year", cleaned_yr, "study_year")

    return meta_dict


def run_profile_parsing_pipeline(file_stream, db):
    """
    Упрощенный парсинг профиля преподавателя из файла индивидуального плана,
    игнорируя подробные часы/нагрузку (loads).
    """
    parsed = parse_individual_plan_file(file_stream)
    
    profile_doc = {
        "fio": parsed["fio"],
        "position": parsed["position"],
        "employment_conditions": parsed["employment_conditions"],
        "degree": parsed["degree"],
        "title": parsed["title"],
        "contract": parsed["contract"],
        "education": parsed["education"],
        "subjects": [l["subject"] for l in parsed["loads"] if l.get("subject")],
        "department_name": parsed.get("department_name", ""),
        "institute_name": parsed.get("institute_name", ""),
        "year": parsed.get("year", "")
    }
    
    # Удаляем дубликаты дисциплин
    profile_doc["subjects"] = list(set(profile_doc["subjects"]))
    
    similar_teacher = find_similar_teacher(parsed["fio"], db)
    if similar_teacher:
        db.employees.update_one({"_id": similar_teacher["_id"]}, {"$set": profile_doc})
        emp_id = similar_teacher["_id"]
    else:
        existing = db.employees.find_one({"fio": parsed["fio"]})
        if existing:
            db.employees.update_one({"_id": existing["_id"]}, {"$set": profile_doc})
            emp_id = existing["_id"]
        else:
            emp_id = db.employees.insert_one(profile_doc).inserted_id
            
    return {parsed["fio"]: profile_doc}


def run_profile_parsing_pipeline_dynamic(file_stream, mapping, db):
    """
    Динамический парсинг профиля преподавателя с помощью ИИ-маппинга для сохранения в workload_db.employees.
    """
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
    
    # Запускаем постобработку для разделения/слияния полей перед сохранением в БД
    metadata = process_resolved_metadata(metadata)
    
    # Пытаемся получить лист метаданных/профиля по умолчанию
    meta_sheet = mapping.get("sheet_name") or (sheet_names[0] if sheet_names else None)
    
    # Кэш датафреймов по листам
    df_cache = {}
    def get_df_for_sheet(s_name):
        if not s_name:
            return None
        if s_name not in df_cache:
            try:
                if hasattr(file_stream, "seek"):
                    file_stream.seek(0)
                df = pd.read_excel(file_stream, sheet_name=s_name, header=None)
                if "общие" in s_name.lower():
                    df = df.T
                df_cache[s_name] = df
            except:
                df_cache[s_name] = None
        return df_cache[s_name]

    def get_metadata_val(key, default=""):
        item = metadata.get(key)
        if isinstance(item, dict):
            val = item.get("value")
            if val is not None and str(val).strip():
                return str(val).strip()
            # fallback to resolving coordinate if value is empty
            coord = item.get("coordinate")
            s_name = item.get("sheet") or meta_sheet
            if coord:
                df_target = get_df_for_sheet(s_name)
                res_val = resolve_meta_value(df_target, item)
                if res_val is not None:
                    return str(res_val).strip()
            return ""
        return str(item).strip() if item is not None else default

    # Извлекаем все поля, приоритет отдается значению из metadata (включая ручные правки)
    fio = get_metadata_val("employee_fio") or tp.get("fio") or "Не указан"
    position = get_metadata_val("employee_position") or tp.get("position") or "преподаватель"
    employment_conditions = get_metadata_val("employee_rate") or get_metadata_val("employee_conditions") or tp.get("employment_conditions") or "штатный"
    degree = get_metadata_val("employee_degree") or tp.get("degree") or ""
    title = get_metadata_val("employee_title") or tp.get("title") or ""
    
    contract_number = get_metadata_val("employee_contract_num")
    contract_date = get_metadata_val("employee_contract_date")
    contract_duration = get_metadata_val("employee_contract_duration")
    
    edu_institution = get_metadata_val("employee_edu_institution")
    edu_year = get_metadata_val("employee_edu_year")
    edu_specialty = get_metadata_val("employee_edu_specialty")
    edu_qualification = get_metadata_val("employee_edu_qualification")
    
    department_name = get_metadata_val("department_name")
    department_head = get_metadata_val("department_head")
    department_direction = get_metadata_val("department_direction")
    institute_name = get_metadata_val("institute_name")
    institute_director = get_metadata_val("institute_director")
    
    raw_year = get_metadata_val("study_year")
    cleaned_yr = ""
    if raw_year:
        years = re.findall(r'\b\d{4}\b', str(raw_year))
        if len(years) >= 2:
            cleaned_yr = f"{years[0]}/{years[1]}"
        elif len(years) == 1:
            cleaned_yr = years[0]
        else:
            cleaned_yr = str(raw_year)

    profile_doc = {
        "fio": fio,
        "position": position,
        "employment_conditions": employment_conditions,
        "degree": degree,
        "title": title,
        "contract": {
            "number": contract_number,
            "date": contract_date,
            "duration": contract_duration
        },
        "education": {
            "institution": edu_institution,
            "year": edu_year,
            "specialty": edu_specialty,
            "qualification": edu_qualification
        },
        "subjects": mapping.get("metadata", {}).get("subjects", []) if isinstance(mapping.get("metadata", {}).get("subjects"), list) else [],
        "department_name": department_name,
        "department_head": department_head,
        "department_direction": department_direction,
        "institute_name": institute_name,
        "institute_director": institute_director,
        "study_year": cleaned_yr,
        "year": cleaned_yr
    }
    
    # Если subjects пустой, попробуем динамически спарсить предметы по колонкам
    if not profile_doc["subjects"]:
        try:
            parsed_plan = parse_individual_plan_dynamic(file_stream, mapping)
            profile_doc["subjects"] = list(set([l["subject"] for l in parsed_plan.get("loads", []) if l.get("subject")]))
        except:
            pass
            
    similar_teacher = find_similar_teacher(fio, db)
    if similar_teacher:
        db.employees.update_one({"_id": similar_teacher["_id"]}, {"$set": profile_doc})
    else:
        existing = db.employees.find_one({"fio": fio})
        if existing:
            db.employees.update_one({"_id": existing["_id"]}, {"$set": profile_doc})
        else:
            db.employees.insert_one(profile_doc)
            
    return {fio: profile_doc}
