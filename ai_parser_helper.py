import urllib.request
import urllib.error
import json
import re
import pandas as pd
import os

def load_training_examples():
    """
    Загружает сохраненные примеры обучения ИИ из файла ai_training_examples.json.
    """
    file_path = "ai_training_examples.json"
    if os.path.exists(file_path):
        try:
            with open(file_path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            return []
    return []

def format_training_examples_for_prompt(doc_type=None):
    """
    Форматирует примеры для вставки в системный промпт ИИ.
    """
    examples = load_training_examples()
    if not examples:
        return ""
    
    formatted_list = []
    for idx, ex in enumerate(examples):
        # Если задан конкретный тип документа, фильтруем примеры
        if doc_type and ex.get("document_type") != doc_type:
            continue
            
        ex_str = f"Example {idx + 1}:\n"
        ex_str += f"- Document Type: {ex.get('document_type', 'unknown')}\n"
        ex_str += f"- Sheet Name: {ex.get('sheet_name', '')}\n"
        if ex.get("columns_present"):
            ex_str += "- Excel Columns:\n"
            for col in ex["columns_present"][:15]:  # Ограничиваем, чтобы не раздувать промпт
                ex_str += f"  * {col}\n"
        
        meta_clean = {}
        for k, v in ex.get("correct_metadata", {}).items():
            if isinstance(v, dict):
                val = v.get("value")
                coord = v.get("coordinate")
                if coord:
                    meta_clean[k] = f"{val} (found in cell {coord})"
                else:
                    meta_clean[k] = val
            else:
                meta_clean[k] = v
                
        target_json = {
            "sheet_name": ex.get("sheet_name", ""),
            "document_type": ex.get("document_type", ""),
            "header_row_index": ex.get("header_row_index", 0),
            "transpose": ex.get("transpose", False),
            "column_mapping": ex.get("correct_mapping", {}),
            "metadata": meta_clean
        }
        ex_str += f"- Expected JSON Response:\n{json.dumps(target_json, ensure_ascii=False, indent=2)}\n"
        formatted_list.append(ex_str)
        
    if not formatted_list:
        return ""
        
    return "\n### FEW-SHOT EXAMPLES (USE THEM TO GUIDE YOUR MAPPING):\n" + "\n---\n".join(formatted_list) + "\n---\n"

def get_descriptive_col_names(df, header_row_idx=None):
    col_names = []
    # Ищем строку, содержащую ключевые слова, если header_row_idx не передан
    if header_row_idx is None:
        header_row_idx = 0
        for r in range(min(15, df.shape[0])):
            row_vals = [str(df.iloc[r, c]).strip().lower() for c in range(df.shape[1]) if pd.notna(df.iloc[r, c])]
            row_str = " ".join(row_vals)
            if any(kw in row_str for kw in ["дисциплин", "наимен", "лекц", "групп", "часов"]):
                header_row_idx = r
                break
            
    for c in range(df.shape[1]):
        letter = chr(65 + c) if c < 26 else str(c)
        # Ищем непустой текст в колонке c начиная со строки заголовка
        header_text = ""
        for scan_r in range(header_row_idx, min(header_row_idx + 4, df.shape[0])):
            val = df.iloc[scan_r, c]
            if pd.notna(val) and str(val).strip():
                header_text = str(val).strip().replace("\n", " ").replace("\r", " ")
                break
        if not header_text:
            # Попробуем просканировать первые 10 строк
            for scan_r in range(min(10, df.shape[0])):
                val = df.iloc[scan_r, c]
                if pd.notna(val) and str(val).strip():
                    header_text = str(val).strip().replace("\n", " ").replace("\r", " ")
                    break
                    
        if header_text:
            if len(header_text) > 40:
                header_text = header_text[:37] + "..."
            col_names.append(f"Столбец {c} (Буква {letter}) - \"{header_text}\"")
        else:
            col_names.append(f"Столбец {c} (Буква {letter})")
    return col_names

def heuristic_column_mapping(col_names, ai_mapping=None):
    if ai_mapping is None:
        ai_mapping = {}
        
    schema_fields = {
        "subject_name": ["дисциплина", "наименование", "предмет", "вид работы"],
        "group_name": ["группа", "групп", "гр."],
        "semester_number": ["семестр", "сем."],
        "teacher_fio": ["преподаватель", "фио", "ф.и.о", "руководитель"],
        "hours_lectures": ["лекции", "лекц"],
        "hours_practicals": ["практические", "семинары", "практ", "семин"],
        "hours_laboratories": ["лабораторные", "лаб"],
        "hours_consultations": ["консультации", "конс"],
        "hours_exams": ["экзамен", "экз"],
        "hours_zachets": ["зачет", "зач"],
        "hours_coursework": ["курсов", "кп", "кр"],
        "hours_practice": ["практика", "производств", "руковод"],
        "hours_vkr": ["вкр", "выпускн", "диплом"],
        "hours_gek": ["гэк", "гос"],
        "hours_additional": ["дополнительн", "доп"],
        "hours_total": ["итого", "всего"]
    }
    
    mapping = {}
    # Сначала берем то, что определил ИИ
    if isinstance(ai_mapping, dict):
        for k, v in ai_mapping.items():
            if v is not None:
                try:
                    mapping[k] = int(v)
                except:
                    pass
                    
    # Для пропущенных полей запускаем эвристику
    for field, keywords in schema_fields.items():
        if mapping.get(field) is None:
            # Ищем совпадение в col_names
            for idx, col_name in enumerate(col_names):
                col_lower = col_name.lower()
                # Ищем точное вхождение любого ключевого слова
                if any(kw in col_lower for kw in keywords):
                    # Дополнительная проверка для практики vs практических
                    if field == "hours_practice" and "практическ" in col_lower:
                        continue
                    if field == "hours_practicals" and "производств" in col_lower:
                        continue
                    mapping[field] = idx
                    break
                    
    return mapping


def check_ollama_status(api_url):
    """
    Проверяет, запущен ли Ollama по указанному URL-адресу.
    """
    url = f"{api_url.rstrip('/')}/api/tags"
    try:
        req = urllib.request.Request(url, method="GET")
        with urllib.request.urlopen(req, timeout=3) as response:
            if response.status == 200:
                data = json.loads(response.read().decode('utf-8'))
                models = [model['name'] for model in data.get("models", [])]
                return True, models
            return False, []
    except Exception:
        return False, []

def pull_model_stream(api_url, model_name):
    """
    Потоково скачивает модель с Ollama API и возвращает генератор с ходом загрузки.
    """
    url = f"{api_url.rstrip('/')}/api/pull"
    payload = {"name": model_name, "stream": True}
    data = json.dumps(payload).encode('utf-8')
    req = urllib.request.Request(url, data=data, headers={"Content-Type": "application/json"})
    try:
        # Увеличиваем таймаут для загрузки модели
        response = urllib.request.urlopen(req, timeout=3600)
        buffer = b""
        while True:
            chunk = response.read(1024)
            if not chunk:
                break
            buffer += chunk
            while b"\n" in buffer:
                line, buffer = buffer.split(b"\n", 1)
                line_str = line.decode('utf-8').strip()
                if line_str:
                    try:
                        yield json.loads(line_str)
                    except Exception:
                        pass
    except Exception as e:
        yield {"error": str(e)}

def prepare_excel_preview(file_stream, max_rows=15, target_sheet=None, header_row_idx=None):
    """
    Читает начало Excel-файла и форматирует его для промпта ИИ.
    Если target_sheet задан, возвращает превью этого листа.
    Если target_sheet не задан (None), возвращает консолидированное превью всех листов.
    """
    try:
        if hasattr(file_stream, "seek"):
            file_stream.seek(0)
        xls = pd.ExcelFile(file_stream)
        sheet_names = xls.sheet_names
        xls.close()
        
        # Находим регистронезависимое совпадение для target_sheet
        matched_sheet = None
        if target_sheet is not None:
            target_clean = str(target_sheet).strip().lower().replace("\"", "").replace("'", "")
            for s in sheet_names:
                s_clean = s.strip().lower()
                if s_clean == target_clean:
                    matched_sheet = s
                    break
                    
        # Если конкретный лист задан и совпал
        if matched_sheet is not None:
            if hasattr(file_stream, "seek"):
                file_stream.seek(0)
            df = pd.read_excel(file_stream, sheet_name=matched_sheet, header=None)
            
            preview_lines = []
            for idx, row in df.iloc[:max_rows].iterrows():
                row_vals = []
                for val in row.values:
                    if pd.isna(val):
                        row_vals.append("")
                    else:
                        val_str = str(val).replace("\n", " ").replace("\r", " ").strip()
                        if len(val_str) > 30:
                            val_str = val_str[:27] + "..."
                        row_vals.append(val_str)
                if any(v for v in row_vals):
                    preview_lines.append(f"Row {idx}: " + " | ".join(row_vals))
                    
            preview_text = f"Sheet Name: \"{matched_sheet}\"\n" + "\n".join(preview_lines)
            col_names = get_descriptive_col_names(df, header_row_idx=header_row_idx)
            return preview_text, col_names, matched_sheet
            
        # Если лист не задан, формируем сводный предпросмотр по всем листам
        all_previews = []
        default_sheet = sheet_names[0]
        
        # Пытаемся найти наиболее вероятный лист по ключевым словам
        for s in sheet_names:
            if hasattr(file_stream, "seek"):
                file_stream.seek(0)
            try:
                df_test = pd.read_excel(file_stream, sheet_name=s, nrows=6, header=None)
                row_str = " ".join([str(v).lower() for v in df_test.values.flatten() if pd.notna(v)])
                if any(kw in row_str for kw in ["наименование", "дисциплина", "лекци", "препод", "групп", "кафедр"]):
                    default_sheet = s
                    break
            except Exception:
                pass
                
        # Если совпадений по словам нет, выберем лист с максимальным числом колонок (это основная таблица)
        if default_sheet == sheet_names[0]:
            max_cols = 0
            best_sheet = sheet_names[0]
            for s in sheet_names:
                if hasattr(file_stream, "seek"):
                    file_stream.seek(0)
                try:
                    df_temp = pd.read_excel(file_stream, sheet_name=s, nrows=2, header=None)
                    if df_temp.shape[1] > max_cols:
                        max_cols = df_temp.shape[1]
                        best_sheet = s
                except:
                    pass
            default_sheet = best_sheet
                
        # Собираем строки из каждого листа для ИИ
        for s in sheet_names:
            if hasattr(file_stream, "seek"):
                file_stream.seek(0)
            try:
                s_low = s.lower()
                is_profile_sheet = any(kw in s_low for kw in ["титульн", "сведения", "анкета", "общие", "profile"])
                rows_to_read = 35 if is_profile_sheet else 12
                df = pd.read_excel(file_stream, sheet_name=s, nrows=rows_to_read, header=None)
                preview_lines = []
                for idx, row in df.iterrows():
                    row_vals = []
                    for val in row.values:
                        if pd.isna(val):
                            row_vals.append("")
                        else:
                            val_str = str(val).replace("\n", " ").replace("\r", " ").strip()
                            if len(val_str) > 25:
                                val_str = val_str[:22] + "..."
                            row_vals.append(val_str)
                    if any(v for v in row_vals):
                        preview_lines.append(f"  Row {idx}: " + " | ".join(row_vals))
                if preview_lines:
                    all_previews.append(f"Sheet Name: \"{s}\"\n" + "\n".join(preview_lines))
            except Exception:
                pass
                
        consolidated_preview = "\n\n".join(all_previews)
        
        if hasattr(file_stream, "seek"):
            file_stream.seek(0)
        df_def = pd.read_excel(file_stream, sheet_name=default_sheet, header=None)
        col_names = get_descriptive_col_names(df_def)
        
        return consolidated_preview, col_names, default_sheet
        
    except Exception as e:
        return f"Ошибка подготовки превью: {e}", [], None

def analyze_excel_structure(file_stream, model_name, api_url, target_sheet=None):
    """
    Отправляет первые строки листов ИИ-модели для автоматического определения листа и его структуры.
    """
    preview_text, col_names, default_sheet = prepare_excel_preview(file_stream, target_sheet=target_sheet)
    if not col_names:
        raise ValueError("Не удалось прочитать структуру колонок файла.")
        
    if hasattr(file_stream, "seek"):
        file_stream.seek(0)
    xls = pd.ExcelFile(file_stream)
    sheet_names = xls.sheet_names
    xls.close()
    
    if target_sheet and target_sheet in sheet_names:
        sheet_instruction = f'You MUST analyze the sheet named "{target_sheet}" and set "sheet_name" to "{target_sheet}".'
        sheets_list_str = f'"{target_sheet}"'
    else:
        sheet_instruction = 'Identify which worksheet from the workbook represents the main workload/plan data to parse.'
        sheets_list_str = ", ".join([f'"{name}"' for name in sheet_names])
        
    prompt = f"""You are a specialized parser assistant for university department workload, practice sheets and individual plans.
Your task is to:
1. {sheet_instruction}
2. Analyze the structure of that selected worksheet and map its columns and cell metadata to our complete database schema.

Available sheets list: [{sheets_list_str}]

Excel Spreadsheet Preview (previews of sheets, where each line starts with 'Row index: column0 | column1 | ...'):
---
{preview_text}
---

Our target schema fields (match them with 0-based column indices of Excel in the chosen sheet, or return null if not present as a column):
- `subject_name`: Name of the discipline/subject (e.g. "Математика", "Физика", "Информатика", "Программирование").
- `group_name`: Code/name of student group(s) (e.g., "ИБ-11", "ПР-41", "АСУб-21").
- `semester_number`: Semester number (typically "1", "2" or "I", "II" or "осенний", "весенний").
- `teacher_fio`: Teacher FIO / name (typically contains FIO like "Обухов А.Д.", "Архипов" or similar).
- `hours_lectures`: Lecture hours (numeric).
- `hours_practicals`: Practical/seminar hours (numeric).
- `hours_laboratories`: Lab hours (numeric).
- `hours_consultations`: Consultation hours (numeric).
- `hours_exams`: Exam hours (numeric).
- `hours_zachets`: Zachet/credit hours (numeric).
- `hours_coursework`: Coursework / projects / КП / КР (numeric).
- `hours_practice`: Practice hours (numeric).
- `hours_vkr`: VKR / выпускная квалификационная работа (numeric).
- `hours_gek`: GEK / ГЭК / государственная комиссия (numeric).
- `hours_additional`: Additional workload hours (numeric).
- `hours_total`: Total hours column.
- `student_fio`: Student's FIO / name (if present in columns).
- `student_profile`: Student's study profile/direction (if present in columns).
- `practice_type`: Type of practice (e.g. "ознакомительная", "технологическая", "проектно-технологическая").
- `practice_org`: Name of profiling organization/company where practice is held (e.g. "ТГТУ", "Компания").
- `practice_kind`: Kind of practice (e.g. "учебная", "производственная").

Target schema metadata (if found in the selected sheet cells, extract their values; if not, return null):
- `employee_fio`: Teacher's FIO (often found in sheet cells near "Ф.И.О" or "ФИО" or "Преподаватель").
- `employee_position`: Academic position (e.g. "доцент", "профессор", "ассистент", "зав. кафедрой").
- `employee_rate`: Employment rate / conditions (e.g. "1.0 ставки", "0.5 ставки").
- `employee_degree`: Academic degree (e.g. "к.т.н.", "д.т.н.").
- `employee_title`: Academic title (e.g. "доцент", "профессор").
- `employee_contract_num`: Contract number if present in cells.
- `employee_contract_date`: Contract date if present in cells.
- `employee_contract_duration`: Contract validity / duration if present in cells.
- `employee_edu_institution`: Education institution name if present.
- `employee_edu_year`: Education graduation year if present.
- `employee_edu_specialty`: Education specialty/direction if present.
- `employee_edu_qualification`: Education qualification if present.
- `employee_conditions`: Specific conditions of recruitment (e.g. "штатный работник", "внутренний совместитель", "внешний совместитель", "почасовая оплата").
- `department_name`: Department name if present (e.g., "САПР").
- `department_head`: Head of department FIO if present.
- `department_direction`: Specialty direction of department if present.
- `institute_name`: Institute name if present (e.g., "ИИТ").
- `institute_director`: Director of institute FIO if present.
- `practice_order_date`: Date of practice order if present in cells.
- `practice_order_signer`: FIO of signer of practice order if present.

Determine:
1. `sheet_name`: The EXACT name of the sheet you selected from the available sheets list. It must be one of the sheets listed in the available sheets list: [{sheets_list_str}].
2. `document_type`: "department_load" (general department workload list), "individual_plan" (single teacher's plan/profile), "practice_order" (order / list of practices), "student_list" (group/student list) or "other".
3. `header_row_index`: 0-based index of the row where column headers are located.
4. `transpose`: boolean (true if variables are in rows and data columns are read horizontally; false if standard vertical columns).
5. `column_mapping`: 0-based column indices (numbers) for each schema field in the selected sheet. If a field is not present in that sheet, return null.
6. `metadata`: Extracted metadata details from the selected sheet cells (or null if not found).

You MUST respond with a valid JSON object matching the following structure:
{{
  "sheet_name": "string",
  "document_type": "department_load" | "individual_plan" | "practice_order" | "student_list" | "other",
  "header_row_index": number,
  "transpose": boolean,
  "column_mapping": {{
    "subject_name": number | null,
    "group_name": number | null,
    "semester_number": number | null,
    "teacher_fio": number | null,
    "hours_lectures": number | null,
    "hours_practicals": number | null,
    "hours_laboratories": number | null,
    "hours_consultations": number | null,
    "hours_exams": number | null,
    "hours_zachets": number | null,
    "hours_coursework": number | null,
    "hours_practice": number | null,
    "hours_vkr": number | null,
    "hours_gek": number | null,
    "hours_additional": number | null,
    "hours_total": number | null,
    "student_fio": number | null,
    "student_profile": number | null,
    "practice_type": number | null,
    "practice_org": number | null,
    "practice_kind": number | null
  }},
  "metadata": {{
    "employee_fio": "string" | null,
    "employee_position": "string" | null,
    "employee_rate": "string" | null,
    "employee_degree": "string" | null,
    "employee_title": "string" | null,
    "employee_contract_num": "string" | null,
    "employee_contract_date": "string" | null,
    "employee_contract_duration": "string" | null,
    "employee_edu_institution": "string" | null,
    "employee_edu_year": "string" | null,
    "employee_edu_specialty": "string" | null,
    "employee_edu_qualification": "string" | null,
    "employee_conditions": "string" | null,
    "department_name": "string" | null,
    "department_head": "string" | null,
    "department_direction": "string" | null,
    "institute_name": "string" | null,
    "institute_director": "string" | null,
    "practice_order_date": "string" | null,
    "practice_order_signer": "string" | null
  }}
}}

{format_training_examples_for_prompt()}

Ensure that indices are numbers, not strings. Ensure that the response contains only the JSON, no explanations, no markdown blocks.
"""

    url = f"{api_url.rstrip('/')}/api/generate"
    payload = {
        "model": model_name,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.0
        }
    }
    
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode('utf-8'),
        headers={"Content-Type": "application/json"}
    )
    
    try:
        with urllib.request.urlopen(req, timeout=1000) as response:
            res_data = json.loads(response.read().decode('utf-8'))
            content = res_data.get("response", "").strip()
            
            # Извлекаем JSON из ответа
            start_idx = content.find('{')
            if start_idx != -1:
                try:
                    decoder = json.JSONDecoder()
                    parsed_mapping, _ = decoder.raw_decode(content[start_idx:])
                except Exception:
                    # Резервный вариант с regex
                    match = re.search(r'\{.*\}', content, re.DOTALL)
                    if match:
                        parsed_mapping = json.loads(match.group(0))
                    else:
                        raise ValueError(f"Ответ ИИ не содержит валидного JSON блока. Ответ: {content}")
                
                # Добавляем или перезаписываем имя листа, если ИИ вернул нечто невалидное
                if "sheet_name" not in parsed_mapping or parsed_mapping["sheet_name"] not in sheet_names:
                    parsed_mapping["sheet_name"] = target_sheet or default_sheet
                
                mapped_sheet = parsed_mapping["sheet_name"]
                
                # Дозаполняем пропущенные или нераспознанные колонки с помощью эвристического маппинга
                if hasattr(file_stream, "seek"):
                    file_stream.seek(0)
                try:
                    df_target = pd.read_excel(file_stream, sheet_name=mapped_sheet, header=None)
                    descriptive_cols = get_descriptive_col_names(df_target)
                except:
                    descriptive_cols = col_names
                    
                parsed_mapping["column_mapping"] = heuristic_column_mapping(descriptive_cols, parsed_mapping.get("column_mapping"))
                return parsed_mapping
            else:
                raise ValueError(f"Ответ ИИ не содержит валидного JSON блока. Ответ: {content}")
    except Exception as e:
        raise ValueError(f"Ошибка при работе с ИИ: {e}")


def analyze_teacher_profile_with_ai(file_stream, model_name, api_url, target_sheet=None):
    """
    Отправляет превью листов индивидуального плана ИИ-модели для извлечения анкетных данных преподавателя.
    """
    preview_text, col_names, default_sheet = prepare_excel_preview(file_stream, max_rows=40, target_sheet=target_sheet)
    
    if hasattr(file_stream, "seek"):
        file_stream.seek(0)
    xls = pd.ExcelFile(file_stream)
    sheet_names = xls.sheet_names
    xls.close()
    
    # Пытаемся найти лист 'Общие сведения' или аналогичный для анкеты
    profile_sheet = target_sheet
    if not profile_sheet:
        for s in sheet_names:
            if any(kw in s.lower() for kw in ["общие", "сведения", "титульн", "тит", "profile", "анкета"]):
                profile_sheet = s
                break
        if not profile_sheet:
            profile_sheet = default_sheet

    prompt = f"""You are a specialized parser assistant for university teacher individual plans.
Your task is to extract the teacher's profile/metadata (general info) from the worksheet representing the teacher profile.

We have detected that the sheet "{profile_sheet}" is most likely the profile sheet.
Here is the preview of the sheet:
---
{preview_text}
---

Please analyze the spreadsheet preview above and extract the following details. If a detail is not found in the sheet, return null for it.

Target schema fields to extract:
- `employee_fio`: Teacher's Full Name / FIO (e.g., "Обухов Артем Дмитриевич").
- `employee_position`: Academic position (e.g. "доцент", "профессор", "ассистент", "зав. кафедрой", "старший преподаватель").
- `employee_rate`: Employment rate / conditions (e.g. "1.0", "0.5", "1.5").
- `employee_degree`: Academic degree (e.g. "к.т.н.", "д.т.н.").
- `employee_title`: Academic title (e.g. "доцент", "профессор", "без ученого звания").
- `employee_contract_num`: Contract/Labor agreement number if present (e.g. "№ТД173/21" or "K-2026-99").
- `employee_contract_date`: Contract agreement date if present (e.g. "29.10.2021г." or "01.09.2025").
- `employee_contract_duration`: Contract validity / duration if present (e.g. "по 31.08.2028 года" or "5 лет").
- `employee_edu_institution`: Base education institution name if present.
- `employee_edu_year`: Education graduation year if present.
- `employee_edu_specialty`: Education specialty/direction if present.
- `employee_edu_qualification`: Education qualification if present.
- `employee_conditions`: Specific conditions of recruitment (e.g. "штатный работник", "внутренний совместитель", "внешний совместитель", "почасовая оплата").
- `department_name`: Department name if present (e.g., "САПР").
- `department_head`: Head of department FIO if present.
- `department_direction`: Specialty direction of department if present.
- `institute_name`: Institute name if present.
- `institute_director`: Director of institute FIO if present.
- `subjects`: A JSON array of strings containing unique subject names that this teacher teaches (listed in the document, e.g. ["Офисные технологии", "Интернет-технологии", "Программирование"]).

You MUST respond with a valid JSON object matching the following structure.
For each field, provide an object with "value" (the extracted text) and "coordinate" (the Excel cell address like "B5" where the value was found). If a cell coordinate cannot be determined, use an empty string for "coordinate".
{{
  "employee_fio": {{"value": "string" | null, "coordinate": "string"}},
  "employee_position": {{"value": "string" | null, "coordinate": "string"}},
  "employee_rate": {{"value": "string" | null, "coordinate": "string"}},
  "employee_degree": {{"value": "string" | null, "coordinate": "string"}},
  "employee_title": {{"value": "string" | null, "coordinate": "string"}},
  "employee_contract_num": {{"value": "string" | null, "coordinate": "string"}},
  "employee_contract_date": {{"value": "string" | null, "coordinate": "string"}},
  "employee_contract_duration": {{"value": "string" | null, "coordinate": "string"}},
  "employee_edu_institution": {{"value": "string" | null, "coordinate": "string"}},
  "employee_edu_year": {{"value": "string" | null, "coordinate": "string"}},
  "employee_edu_specialty": {{"value": "string" | null, "coordinate": "string"}},
  "employee_edu_qualification": {{"value": "string" | null, "coordinate": "string"}},
  "employee_conditions": {{"value": "string" | null, "coordinate": "string"}},
  "department_name": {{"value": "string" | null, "coordinate": "string"}},
  "department_head": {{"value": "string" | null, "coordinate": "string"}},
  "department_direction": {{"value": "string" | null, "coordinate": "string"}},
  "institute_name": {{"value": "string" | null, "coordinate": "string"}},
  "institute_director": {{"value": "string" | null, "coordinate": "string"}},
  "subjects": ["string"]
}}

The preview above shows columns labeled as "Столбец N (Буква X)" where X is the Excel column letter. Row numbers in the preview correspond to Excel row numbers (1-based). Combine the column letter and row number to form the cell coordinate (e.g., column "Буква B" at row 5 = "B5").

{format_training_examples_for_prompt(doc_type="individual_plan")}

Ensure that the response contains only the JSON, no explanations, no markdown blocks.
"""

    url = f"{api_url.rstrip('/')}/api/generate"
    payload = {
        "model": model_name,
        "prompt": prompt,
        "stream": False,
        "options": {
            "temperature": 0.0
        }
    }
    
    req = urllib.request.Request(
        url,
        data=json.dumps(payload).encode('utf-8'),
        headers={"Content-Type": "application/json"}
    )
    
    try:
        with urllib.request.urlopen(req, timeout=1000) as response:
            res_data = json.loads(response.read().decode('utf-8'))
            content = res_data.get("response", "").strip()
            
            # Извлекаем JSON из ответа
            start_idx = content.find('{')
            if start_idx != -1:
                try:
                    decoder = json.JSONDecoder()
                    parsed_profile, _ = decoder.raw_decode(content[start_idx:])
                    return parsed_profile
                except Exception:
                    # Резервный вариант с regex
                    match = re.search(r'\{.*\}', content, re.DOTALL)
                    if match:
                        return json.loads(match.group(0))
                    else:
                        raise ValueError(f"Ответ ИИ не содержит валидного JSON блока. Ответ: {content}")
            else:
                raise ValueError(f"Ответ ИИ не содержит валидного JSON блока. Ответ: {content}")
    except Exception as e:
        raise ValueError(f"Ошибка при работе с ИИ при извлечении профиля: {e}")


def find_best_training_example(filename, sheet_name=None):
    """
    Ищет наиболее подходящий обучающий пример по имени файла и (опционально) по имени листа.
    """
    examples = load_training_examples()
    if not examples:
        return None
    
    filename_clean = filename.lower().strip()
    sheet_clean = sheet_name.lower().strip() if sheet_name else None
    
    # 1. Точное совпадение
    for ex in examples:
        if ex.get("filename", "").lower().strip() == filename_clean:
            if not sheet_clean or ex.get("sheet_name", "").lower().strip() == sheet_clean:
                return ex
            
    # 2. Совпадение по общим префиксам/паттернам
    prefix_patterns = ["individ_plrpr2024", "individ_plan", "сапр"]
    for pattern in prefix_patterns:
        if pattern in filename_clean:
            for ex in examples:
                ex_name = ex.get("filename", "").lower()
                if pattern in ex_name:
                    if not sheet_clean or ex.get("sheet_name", "").lower().strip() == sheet_clean:
                        return ex
                    
    # 3. Совпадение по схожести имени
    from difflib import SequenceMatcher
    best_ex = None
    best_ratio = 0.0
    for ex in examples:
        ex_name = ex.get("filename", "").lower()
        if sheet_clean and ex.get("sheet_name", "").lower().strip() != sheet_clean:
            continue
        ratio = SequenceMatcher(None, filename_clean.split(".")[0], ex_name.split(".")[0]).ratio()
        if ratio > best_ratio:
            best_ratio = ratio
            best_ex = ex
            
    if best_ratio > 0.5:
        return best_ex
        
    return None


def get_merged_training_template(filename):
    """
    Ищет все обучающие примеры для данного файла/типа и объединяет их:
    - метаданные берутся с листа, где они заполнены (например, Тит. лист)
    - маппинг колонок берется с листа, где он заполнен (например, 1 семестр)
    Возвращает единый merged шаблон для авто-импорта.
    """
    examples = load_training_examples()
    if not examples:
        return None
        
    filename_clean = filename.lower().strip()
    matching_examples = []
    
    # 1. Поиск точных совпадений или по паттерну
    prefix_patterns = ["individ_plrpr2024", "individ_plan", "сапр"]
    matched_pattern = None
    for pattern in prefix_patterns:
        if pattern in filename_clean:
            matched_pattern = pattern
            break
            
    if matched_pattern:
        matching_examples = [ex for ex in examples if matched_pattern in ex.get("filename", "").lower()]
    else:
        # Проверяем точное совпадение
        exact_matches = [ex for ex in examples if ex.get("filename", "").lower().strip() == filename_clean]
        if exact_matches:
            matching_examples = exact_matches
        else:
            # Fuzzy поиск
            from difflib import SequenceMatcher
            best_ex_name = None
            best_ratio = 0.0
            for ex in examples:
                ex_name = ex.get("filename", "").lower()
                ratio = SequenceMatcher(None, filename_clean.split(".")[0], ex_name.split(".")[0]).ratio()
                if ratio > best_ratio:
                    best_ratio = ratio
                    best_ex_name = ex_name
            if best_ratio > 0.5 and best_ex_name:
                matching_examples = [ex for ex in examples if ex.get("filename", "").lower() == best_ex_name]
                
    if not matching_examples:
        return None
        
    # Сливаем найденные примеры
    merged_metadata = {}
    merged_mapping = {}
    meta_sheet = None
    header_row = 0
    transpose = False
    doc_type = "individual_plan"
    
    for ex in matching_examples:
        # Сливаем метаданные
        meta = ex.get("correct_metadata") or {}
        has_real_meta = False
        for k, v in meta.items():
            if isinstance(v, dict) and v.get("coordinate"):
                merged_metadata[k] = v
                has_real_meta = True
            elif isinstance(v, str) and v.strip():
                if k not in merged_metadata:
                    merged_metadata[k] = v
                    has_real_meta = True
                    
        if has_real_meta and not meta_sheet:
            meta_sheet = ex.get("sheet_name")
            
        # Сливаем маппинг колонок (если на этом листе он есть)
        col_map = ex.get("correct_mapping") or {}
        has_real_col_map = any(v is not None for v in col_map.values())
        if has_real_col_map and not merged_mapping:
            merged_mapping = col_map
            header_row = ex.get("header_row_index", 0)
            transpose = ex.get("transpose", False)
            doc_type = ex.get("document_type", "individual_plan")
            
    if not meta_sheet and matching_examples:
        meta_sheet = matching_examples[0].get("sheet_name")
        
    return {
        "filename": matching_examples[0].get("filename"),
        "sheet_name": meta_sheet,
        "document_type": doc_type,
        "header_row_index": header_row,
        "transpose": transpose,
        "correct_mapping": merged_mapping,
        "correct_metadata": merged_metadata
    }


def get_best_template_by_type(doc_type):
    examples = load_training_examples()
    if not examples:
        return None
        
    filtered = [ex for ex in examples if ex.get("document_type") == doc_type]
    if not filtered:
        return None
        
    import re
    def score_example(ex):
        valid_coords = 0
        meta = ex.get("correct_metadata") or {}
        for val in meta.values():
            coord = ""
            if isinstance(val, dict):
                coord = val.get("coordinate") or ""
            elif isinstance(val, str):
                coord = val
            if coord and re.match(r'^[A-Z]+\d+(?::[A-Z]+\d+|-[A-Z]+\d+)?$', str(coord).strip().upper()):
                valid_coords += 1
        return valid_coords

    filtered.sort(key=lambda x: (score_example(x), x.get("filename", "")), reverse=True)
    best_ex = filtered[0]
    return {
        "filename": best_ex.get("filename"),
        "sheet_name": best_ex.get("sheet_name"),
        "document_type": best_ex.get("document_type"),
        "header_row_index": best_ex.get("header_row_index", 0),
        "transpose": best_ex.get("transpose", False),
        "correct_mapping": best_ex.get("correct_mapping"),
        "correct_metadata": best_ex.get("correct_metadata")
    }


def get_best_template_for_sheet(doc_type, sheet_name):
    examples = load_training_examples()
    if not examples:
        return None
        
    sheet_clean = sheet_name.lower().strip()
    filtered = [ex for ex in examples if ex.get("document_type") == doc_type and ex.get("sheet_name", "").lower().strip() == sheet_clean]
    if not filtered:
        return None
        
    import re
    def score_example(ex):
        valid_coords = 0
        meta = ex.get("correct_metadata") or {}
        for val in meta.values():
            coord = ""
            if isinstance(val, dict):
                coord = val.get("coordinate") or ""
            elif isinstance(val, str):
                coord = val
            if coord and re.match(r'^[A-Z]+\d+(?::[A-Z]+\d+|-[A-Z]+\d+)?$', str(coord).strip().upper()):
                valid_coords += 1
        return valid_coords

    filtered.sort(key=lambda x: (score_example(x), x.get("filename", "")), reverse=True)
    best_ex = filtered[0]
    return {
        "filename": best_ex.get("filename"),
        "sheet_name": best_ex.get("sheet_name"),
        "document_type": best_ex.get("document_type"),
        "header_row_index": best_ex.get("header_row_index", 0),
        "transpose": best_ex.get("transpose", False),
        "correct_mapping": best_ex.get("correct_mapping"),
        "correct_metadata": best_ex.get("correct_metadata")
    }

