import os
import re
import pandas as pd
import streamlit as st
import datetime
import json
from local_db import MongoClient, ObjectId



# Импортируем наши вспомогательные модули
import importlib
import ai_parser_helper
importlib.reload(ai_parser_helper)

import parser_engine
importlib.reload(parser_engine)
from parser_engine import (
    run_parsing_pipeline, 
    parse_individual_plan_file, 
    run_parsing_pipeline_dynamic, 
    parse_individual_plan_dynamic,
    find_similar_teacher,
    run_profile_parsing_pipeline,
    run_profile_parsing_pipeline_dynamic,
    is_individual_plan_file
)
from template_helper import fill_template, create_default_template, read_docx

# Настройка страницы
st.set_page_config(
    page_title="Автоматизация кафедры: Парсер и Шаблоны",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Стилизация интерфейса (Premium CSS)
st.markdown("""
<style>
    .main {
        background-color: #f8f9fa;
    }
    .reportview-container {
        background: #f8f9fa;
    }
    h1, h2, h3 {
        color: #1e3d59;
        font-family: 'Inter', sans-serif;
    }
    .stButton>button {
        background-color: #17b978;
        color: white;
        border-radius: 8px;
        font-weight: bold;
        border: none;
        padding: 0.5rem 1rem;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #086972;
        color: white;
        box-shadow: 0px 4px 10px rgba(0,0,0,0.1);
    }
    .card {
        background: white;
        padding: 1.5rem;
        border-radius: 12px;
        box-shadow: 0 4px 6px rgba(0,0,0,0.05);
        margin-bottom: 1rem;
    }
</style>
""", unsafe_allow_html=True)

# Подключение к локальной базе данных SQLite (Основная БД — САПР)
@st.cache_resource
def get_db_client():
    try:
        client = MongoClient()
        # Проверяем подключение
        client.server_info()
        return client, True
    except Exception as e:
        return None, False

# Подключение ко Второй БД (Нагрузка — не-САПР файлы)
@st.cache_resource
def get_workload_db_client():
    try:
        client = MongoClient(db_file="workload_db.sqlite")
        client.server_info()
        return client, True
    except Exception as e:
        return None, False

client, db_connected = get_db_client()
workload_client, workload_db_connected = get_workload_db_client()

st.title("🎓 Система автоматизации документооборота кафедры")
st.markdown("Интеллектуальный разбор Excel-планов нагрузки, импорт в **локальную БД SQLite** и генерация документов по шаблонам.")

if not db_connected:
    st.error("❌ Не удалось инициализировать локальную базу данных SQLite (university_db).")
    st.stop()

if not workload_db_connected:
    st.warning("⚠️ Не удалось инициализировать Вторую базу данных (workload_db). Файлы без 'САПР' в имени не будут обрабатываться.")

db = client["university_db"]
workload_db = workload_client["workload_db"] if workload_db_connected else None

# Заголовок
st.header("Управление базой данных SQLite")

# Боковая панель
# Боковая панель
st.sidebar.title("🧭 Навигация")
page = st.sidebar.radio(
    "Выберите раздел:",
    ["📂 Просмотр и импорт", "🔗 Распределение нагрузки", "✏️ Редактирование нагрузки"]
)

# Раздел локального ИИ в боковой панели (скрыт визуально, переменные сохранены в фоновом режиме во избежание NameError)
if "ollama_url" not in st.session_state:
    st.session_state["ollama_url"] = "http://localhost:11434"

ollama_url = st.session_state["ollama_url"]

# Проверка статуса подключения к Ollama в фоновом режиме
connected, models = ai_parser_helper.check_ollama_status(ollama_url)
selected_model = "llama3"

if connected:
    if models:
        default_idx = 0
        if "llama3" in models:
            default_idx = models.index("llama3")
        elif "llama3:latest" in models:
            default_idx = models.index("llama3:latest")
        selected_model = models[default_idx]


# --- СТРАНИЦА 1: ПРОСМОТР И ИМПОРТ ---
if page == "📂 Просмотр и импорт":
    unassigned_count = 0
    if hasattr(db, "unassigned_loads"):
        try:
            unassigned_count = len(list(db.unassigned_loads.find()))
        except:
            pass
    unassigned_wl_count = 0
    if workload_db and hasattr(workload_db, "unassigned_loads"):
        try:
            unassigned_wl_count = len(list(workload_db.unassigned_loads.find()))
        except:
            pass
    if unassigned_count > 0:
        st.warning(f"⚠️ **Внимание:** В системе (БД САПР) обнаружено нераспределенное распределение (**{unassigned_count}** строк нагрузки). Перейдите в раздел **🔗 Распределение нагрузки** для привязки к преподавателям.")
    if unassigned_wl_count > 0:
        st.warning(f"⚠️ **Внимание:** Во Второй БД (Нагрузка) обнаружено **{unassigned_wl_count}** нераспределенных строк. Перейдите в **🔗 Распределение нагрузки**.")
        
    col1, col2 = st.columns([1, 3])
    
    with col1:
        st.subheader("Импорт из Excel")
        st.write("Загрузите Excel-файл распределения учебной нагрузки:")
        
        # Отображение отчета о последнем импорте
        if "last_imported_report" in st.session_state:
            report = st.session_state["last_imported_report"]
            st.markdown("### 📊 Отчет о последнем импорте")
            st.markdown(f"**Файл:** `{report['filename']}`")
            
            if report.get("type") == "sapr":
                st.success("🎉 Успешно импортирована учебная нагрузка!")
                
                preview_rows = []
                for teacher, loads in report["data"].items():
                    total_h = sum(l.get("total", 0.0) for l in loads)
                    preview_rows.append({
                        "Преподаватель": teacher if teacher else "Не указан",
                        "Дисциплин": len(loads),
                        "Всего часов": f"{total_h:.1f} ч."
                    })
                if preview_rows:
                    st.dataframe(pd.DataFrame(preview_rows), use_container_width=True, hide_index=True)
                else:
                    st.info("В файле не найдено записей нагрузки с часами > 0.")
            else:
                st.success("🎉 Успешно импортирована анкета преподавателя!")
                for teacher, profile in report["data"].items():
                    st.markdown(f"#### 👤 Преподаватель: **{teacher}**")
                    st.markdown(f"**💼 Должность:** {profile.get('position', '—')} ({profile.get('employment_conditions', '—')})")
                    st.markdown(f"**🎓 Степень / Звание:** {profile.get('degree', '—')} / {profile.get('title', '—')}")
                    if profile.get('contract'):
                        st.markdown(f"**📄 Контракт:** №{profile['contract'].get('number', '—')} от {profile['contract'].get('date', '—')} ({profile['contract'].get('duration', '—')})")
                    if profile.get('education'):
                        edu = profile['education']
                        st.markdown(f"**🏫 Образование:** {edu.get('institution', '—')} ({edu.get('year', '—')}), спец.: {edu.get('specialty', '—')}")
                    if profile.get('subjects'):
                        st.markdown(f"**📖 Дисциплины:** {', '.join(profile.get('subjects', []))}")
            
            if st.button("🗑️ Закрыть отчет"):
                del st.session_state["last_imported_report"]
                st.rerun()
            st.markdown("---")

        uploaded_file = st.file_uploader("Выберите файл (XLSX, XLS)", type=["xlsx", "xls"])
        
        if uploaded_file is not None:
            # === РОУТИНГ ПО ИМЕНИ ФАЙЛА ===
            is_sapr = "сапр" in uploaded_file.name.lower()
            if is_sapr:
                active_db = db
                db_label = "📗 БД САПР (university_db)"
            else:
                if workload_db is not None:
                    active_db = workload_db
                    db_label = "📘 БД Нагрузка (workload_db)"
                else:
                    active_db = db
                    db_label = "📗 БД САПР (university_db) — Вторая БД недоступна"
            
            st.info(f"🗄️ Целевая база данных: **{db_label}**")
            
            # Чтение имен листов
            if hasattr(uploaded_file, "seek"):
                uploaded_file.seek(0)
            xls_inspect = pd.ExcelFile(uploaded_file)
            sheet_names = xls_inspect.sheet_names
            xls_inspect.close()
            
            # Инспекция загруженного файла
            try:
                report_uploaded = []
                report_uploaded.append(f"=== Uploaded File: {uploaded_file.name} ===")
                report_uploaded.append(f"Sheets: {sheet_names}")
                for sname in sheet_names:
                    if hasattr(uploaded_file, "seek"):
                        uploaded_file.seek(0)
                    df = pd.read_excel(uploaded_file, sheet_name=sname, header=None)
                    report_uploaded.append(f"\nSheet: {sname} (Shape: {df.shape})")
                    import openpyxl
                    for r in range(min(50, df.shape[0])):
                        row_vals = []
                        for c in range(df.shape[1]):
                            v = df.iloc[r, c]
                            if pd.notna(v) and str(v).strip():
                                col_letter = openpyxl.utils.get_column_letter(c + 1)
                                row_vals.append(f"{col_letter}{r+1}: '{str(v).strip()}'")
                        if row_vals:
                            report_uploaded.append(f"Row {r+1}: " + " | ".join(row_vals))
                structure_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "obukhov_structure.txt")
                with open(structure_file_path, "w", encoding="utf-8") as f_out:
                    f_out.write("\n".join(report_uploaded))
            except Exception as e_upload_dbg:
                structure_file_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "obukhov_structure.txt")
                with open(structure_file_path, "w", encoding="utf-8") as f_out:
                    f_out.write(f"Upload inspect error: {e_upload_dbg}\n")

            st.write("📁 **Листы в файле:**", sheet_names)
            
            final_mapping = None
            used_ai = False
            best_example = None
            
            # Определяем тип документа для автоматического наложения шаблона
            doc_type = "individual_plan" if not is_sapr else "department_load"
            
            # 1. Ищем и объединяем сохраненные шаблоны для всех листов этого файла
            sheet_templates = {}
            merged_metadata = {}
            merged_column_mapping = {}
            best_filename = None
            primary_sheet = None
            primary_header_row = 0
            primary_transpose = False
            
            if not is_sapr and doc_type == "individual_plan":
                # Задаем строгие координаты разбора для анкеты (индивидуального плана)
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
                
                # Сначала определяем best_filename и прочие параметры по шаблону
                for sheet_name in sheet_names:
                    tpl = ai_parser_helper.get_best_template_for_sheet(doc_type, sheet_name)
                    if tpl:
                        sheet_templates[sheet_name] = tpl
                        if not best_filename:
                            best_filename = tpl.get("filename")
                            primary_sheet = tpl.get("sheet_name")
                            primary_header_row = tpl.get("header_row_index", 0)
                            primary_transpose = tpl.get("transpose", False)
                            
                # Читаем строго по координатам из листов
                for meta_key, coord_info in STRICT_PLAN_COORDINATES.items():
                    coord = coord_info["coordinate"]
                    sheet_name = coord_info["sheet"]
                    
                    if sheet_name in sheet_names:
                        df_resolve = None
                        try:
                            if hasattr(uploaded_file, "seek"):
                                uploaded_file.seek(0)
                            df_resolve = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
                            if "общие" in sheet_name.lower():
                                df_resolve = df_resolve.T
                        except Exception:
                            pass
                            
                        new_value = ""
                        if coord and df_resolve is not None:
                            from parser_engine import get_cell_val_by_coord
                            cell_val = get_cell_val_by_coord(df_resolve, coord)
                            if cell_val:
                                new_value = str(cell_val).strip()
                                
                        merged_metadata[meta_key] = {
                            "value": new_value,
                            "coordinate": coord,
                            "sheet": sheet_name
                        }
                
                # Не производим постобработку при выводе в UI, чтобы сохранить все 20 полей раздельными
                # merged_metadata = process_resolved_metadata(merged_metadata)
                        
                # Мержим маппинг колонок по шаблону
                for sheet_name in sheet_names:
                    tpl = sheet_templates.get(sheet_name)
                    if tpl:
                        col_map = tpl.get("correct_mapping") or {}
                        if any(v is not None for v in col_map.values()):
                            merged_column_mapping.update(col_map)
            elif not is_sapr:
                # SAPR и прочие типы документов используют стандартный луп по шаблонам
                for sheet_name in sheet_names:
                    tpl = ai_parser_helper.get_best_template_for_sheet(doc_type, sheet_name)
                    if tpl:
                        sheet_templates[sheet_name] = tpl
                        if not best_filename:
                            best_filename = tpl.get("filename")
                            primary_sheet = tpl.get("sheet_name")
                            primary_header_row = tpl.get("header_row_index", 0)
                            primary_transpose = tpl.get("transpose", False)
                        
                        df_resolve = None
                        try:
                            if hasattr(uploaded_file, "seek"):
                                uploaded_file.seek(0)
                            df_resolve = pd.read_excel(uploaded_file, sheet_name=sheet_name, header=None)
                        except Exception:
                            pass
                        
                        template_metadata = tpl.get("correct_metadata", {}) or {}
                        for meta_key, meta_val in template_metadata.items():
                            if isinstance(meta_val, dict):
                                coord = meta_val.get("coordinate", "")
                                new_value = ""
                                if coord and df_resolve is not None:
                                    from parser_engine import get_cell_val_by_coord
                                    cell_val = get_cell_val_by_coord(df_resolve, coord)
                                    if cell_val:
                                        new_value = cell_val
                                
                                if new_value or coord:
                                    merged_metadata[meta_key] = {
                                        "value": new_value,
                                        "coordinate": coord,
                                        "sheet": sheet_name
                                    }
                            else:
                                merged_metadata[meta_key] = meta_val
                                
                        col_map = tpl.get("correct_mapping") or {}
                        if any(v is not None for v in col_map.values()):
                            merged_column_mapping.update(col_map)
            
            if sheet_templates:
                best_example = {
                    "filename": best_filename,
                    "sheet_name": primary_sheet or sheet_names[0],
                    "document_type": doc_type,
                    "header_row_index": primary_header_row,
                    "transpose": primary_transpose,
                    "correct_mapping": merged_column_mapping,
                    "correct_metadata": merged_metadata
                }
                
                final_mapping = {
                    "sheet_name": primary_sheet or sheet_names[0],
                    "document_type": doc_type,
                    "header_row_index": primary_header_row,
                    "transpose": primary_transpose,
                    "column_mapping": merged_column_mapping,
                    "metadata": merged_metadata
                }
                st.success(f"🎯 Применены сохраненные шаблоны из файла `{best_filename}` для листов: {', '.join(sheet_templates.keys())}! Данные автоматически извлечены из соответствующих листов.")
            # 2. Если шаблонов нет, но подключен Ollama, пробуем автоматический ИИ-анализ (с обучением few-shot)
            if connected and not is_sapr:
                cache_key = f"ai_map_{uploaded_file.name}_auto"
                if cache_key not in st.session_state:
                    with st.spinner("🤖 ИИ автоматически анализирует структуру Excel-файла..."):
                        try:
                            if hasattr(uploaded_file, "seek"):
                                uploaded_file.seek(0)
                            if is_sapr:
                                mapping = ai_parser_helper.analyze_excel_structure(
                                    uploaded_file, selected_model, ollama_url, target_sheet=None
                                )
                            else:
                                mapping = ai_parser_helper.analyze_teacher_profile_with_ai(
                                    uploaded_file, selected_model, ollama_url, target_sheet=None
                                )
                            st.session_state[cache_key] = mapping
                            st.success("🤖 Структура успешно распознана ИИ на основе обучающих примеров!")
                        except Exception as e_ai:
                            st.warning(f"⚠️ Ошибка ИИ-анализа: {e_ai}. Будет использован стандартный парсер.")
                            st.session_state[cache_key] = None
                
                temp_mapping = st.session_state.get(cache_key)
                if temp_mapping:
                    final_mapping = temp_mapping
                    used_ai = True
            
            # 3. Если ИИ нет и шаблонов нет, используем пустой словарь (стандартный парсер)
            if not final_mapping and not is_sapr:
                if not connected:
                    st.warning("⚠️ Локальный ИИ (Ollama) недоступен и подходящих шаблонов не найдено. Будет применен стандартный эвристический парсер.")
                final_mapping = {}

            if final_mapping:
                # Нормализуем структуру: если метаданные лежат в корне, переносим их в ключ "metadata"
                if "metadata" not in final_mapping:
                    meta_keys = ["employee_fio", "employee_position", "employee_rate", "employee_degree", "employee_title", 
                                 "employee_contract_num", "employee_contract_date", "employee_contract_duration", 
                                 "employee_edu_institution", "employee_edu_year", "employee_edu_specialty", "employee_edu_qualification", 
                                 "department_name", "department_head", "department_direction", "institute_name", "institute_director"]
                    found_meta = {}
                    for k in list(final_mapping.keys()):
                        if k in meta_keys:
                            found_meta[k] = final_mapping.pop(k)
                    if found_meta:
                        final_mapping["metadata"] = found_meta
                
                # Разрешаем координаты метаданных (читаем реальные значения из нового файла по координатам)
                if "metadata" in final_mapping:
                    from parser_engine import resolve_meta_value
                    resolved_metadata = {}
                    for meta_key, meta_val in final_mapping["metadata"].items():
                        # Определяем лист для этого метаданного
                        meta_sheet = sheet_names[0]
                        if isinstance(meta_val, dict) and meta_val.get("sheet"):
                            meta_sheet = meta_val.get("sheet")
                        elif final_mapping.get("sheet_name"):
                            meta_sheet = final_mapping.get("sheet_name")
                        
                        # Читаем лист из файла
                        df_resolve = None
                        try:
                            if hasattr(uploaded_file, "seek"):
                                uploaded_file.seek(0)
                            df_resolve = pd.read_excel(uploaded_file, sheet_name=meta_sheet, header=None)
                            if "общие" in meta_sheet.lower():
                                df_resolve = df_resolve.T
                        except Exception:
                            pass
                        
                        resolved_val = resolve_meta_value(df_resolve, meta_val)
                        
                        # Вытаскиваем координату
                        coord = ""
                        if isinstance(meta_val, dict):
                            coord = meta_val.get("coordinate") or ""
                        elif isinstance(meta_val, str):
                            import re
                            match = re.search(r'\((?:found in cell|cell)?\s*([A-Za-z]+\d+)\)', meta_val)
                            if match:
                                coord = match.group(1)
                            elif re.match(r'^[A-Za-z]+\d+$', meta_val.strip()):
                                coord = meta_val.strip()
                        
                        resolved_metadata[meta_key] = {
                            "value": resolved_val,
                            "coordinate": coord,
                            "sheet": meta_sheet
                        }
                    final_mapping["metadata"] = resolved_metadata
            
            # Не производим постобработку при выводе в UI, чтобы сохранить все 20 полей раздельными
            # if final_mapping and "metadata" in final_mapping:
            #     from parser_engine import process_resolved_metadata
            #     final_mapping["metadata"] = process_resolved_metadata(final_mapping["metadata"])
            
            metadata_labels = {
                "employee_fio": "ФИО Преподавателя (Общие сведения)",
                "employee_fio_title": "ФИО Преподавателя (Титульный лист)",
                "employee_position": "Должность (Общие сведения)",
                "employee_position_title": "Должность (Титульный лист)",
                "employee_rate": "Размер ставки",
                "employee_conditions": "Условия привлечения",
                "employee_degree": "Ученая степень",
                "employee_title": "Ученое звание",
                "employee_contract": "Сведения о контракте (номер и дата)",
                "employee_contract_duration": "Срок действия договора",
                "employee_edu_inst_year": "Образование и год окончания",
                "employee_edu_specialty_1": "Направление подготовки (код)",
                "employee_edu_specialty_2": "Направление подготовки (наименование)",
                "employee_edu_qualification": "Квалификация",
                "department_name": "Кафедра",
                "department_head": "Заведующий кафедрой",
                "department_direction": "Направление кафедры",
                "institute_name": "Институт",
                "institute_director": "Директор института",
                "study_year": "Учебный год"
            }

            if final_mapping and "metadata" in final_mapping:
                import hashlib
                import json
                meta_hash = hashlib.md5(json.dumps(final_mapping.get("metadata", {}), default=str).encode('utf-8')).hexdigest() if final_mapping else ""
                cache_import_key = f"{uploaded_file.name}_{meta_hash}"
                
                if "import_metadata_edits" not in st.session_state or st.session_state.get("import_metadata_cache_key") != cache_import_key:
                    st.session_state["import_metadata_cache_key"] = cache_import_key
                    st.session_state["import_metadata_file"] = uploaded_file.name
                    st.session_state["import_metadata_edits"] = {}
                    for m_key in metadata_labels.keys():
                        m_val = final_mapping["metadata"].get(m_key)
                        if isinstance(m_val, dict):
                            st.session_state["import_metadata_edits"][m_key] = m_val.get("value") or ""
                        elif m_val is not None:
                            st.session_state["import_metadata_edits"][m_key] = str(m_val)
                        else:
                            st.session_state["import_metadata_edits"][m_key] = ""

            if is_sapr:
                st.info(f"📂 **Тип документа:** Нагрузка кафедры (САПР)\n\n🔍 **Найдено листов:** {len(sheet_names)}")
                st.warning("⚠️ Для файлов САПР используется стандартный табличный парсер (колонки A-U). Автоанализ структуры ИИ и редактирование анкет отключены.")
                fio_resolutions = {}
            else:
                # Отображаем карточку с тем, что понял ИИ
                st.markdown("### 🔮 Результат автоанализа ИИ")
                doc_type_label = "Анкета / Индивидуальный план преподавателя"
                    
                st.info(f"📂 **Тип документа:** {doc_type_label}\n\n🔍 **Найдено листов:** {len(sheet_names)}")

                # Получаем имена колонок для отображения маппинга
                active_sheet = final_mapping.get("sheet_name") or sheet_names[0]
                if hasattr(uploaded_file, "seek"):
                    uploaded_file.seek(0)
                try:
                    _preview_text, col_names, _active_sheet = ai_parser_helper.prepare_excel_preview(uploaded_file, target_sheet=active_sheet)
                except Exception:
                    col_names = []

                schema_fields_labels = {
                    "subject_name": "Дисциплина (subject_name)",
                    "group_name": "Группа (group_name)",
                    "semester_number": "Семестр (semester_number)",
                    "hours_lectures": "Часы лекций (hours_lectures)",
                    "hours_practicals": "Часы практик (hours_practicals)",
                    "hours_laboratories": "Часы лабораторий (hours_laboratories)",
                    "hours_consultations": "Часы консультаций (hours_consultations)",
                    "hours_exams": "Часы экзаменов (hours_exams)",
                    "hours_zachets": "Часы зачетов (hours_zachets)",
                    "hours_coursework": "Часы КП/КР (hours_coursework)",
                    "hours_practice": "Часы практики (hours_practice)",
                    "hours_vkr": "Часы ВКР (hours_vkr)",
                    "hours_gek": "Часы ГЭК (hours_gek)",
                    "hours_additional": "Дополнительно (hours_additional)",
                    "hours_total": "Итого учебных часов (hours_total)",
                    "teacher_fio": "ФИО преподавателя в таблице (teacher_fio)",
                    "student_fio": "ФИО Студента (student_fio)",
                    "student_profile": "Профиль студента (student_profile)",
                    "practice_type": "Тип практики (practice_type)",
                    "practice_org": "Организация практики (practice_org)",
                    "practice_kind": "Вид практики (practice_kind)"
                }
                
                mapping_rows = []
                col_map_data = final_mapping.get("column_mapping", {}) or {}
                for field_key, field_name in schema_fields_labels.items():
                    col_idx = col_map_data.get(field_key)
                    if col_idx is not None:
                        try:
                            col_idx = int(col_idx)
                            if col_idx < len(col_names):
                                mapping_rows.append(f"- **{field_name}:** `{col_names[col_idx]}`")
                            else:
                                mapping_rows.append(f"- **{field_name}:** Столбец {col_idx}")
                        except:
                            mapping_rows.append(f"- **{field_name}:** `{col_idx}`")
                
                col_source_label = "📍 Сопоставление колонок из обучающего шаблона:" if best_example else "📍 Автоматически сопоставленные колонки ИИ:"
                if mapping_rows:
                    with st.expander(col_source_label, expanded=True):
                        st.markdown("\n".join(mapping_rows))

                # Разделяем ключи метаданных на два списка в зависимости от листов
                title_keys = []
                general_keys = []
                
                meta_data = final_mapping.get("metadata", {}) or {}
                for m_key in metadata_labels.keys():
                    val = meta_data.get(m_key)
                    sheet = ""
                    if isinstance(val, dict):
                        sheet = val.get("sheet") or ""
                    
                    sheet_lower = sheet.lower()
                    if "титульн" in sheet_lower or "тит" in sheet_lower:
                        title_keys.append(m_key)
                    elif "общие" in sheet_lower or "сведения" in sheet_lower:
                        general_keys.append(m_key)
                    else:
                        if m_key in ["employee_fio_title", "employee_position_title", "department_name", "department_head", "institute_name", "institute_director", "study_year"]:
                            title_keys.append(m_key)
                        else:
                            general_keys.append(m_key)
                
                st.markdown("### 📝 Извлеченные анкетные данные преподавателя (доступно для редактирования)")
                c_meta1, c_meta2 = st.columns(2)
                with c_meta1:
                    st.markdown("#### 📄 Данные из Титульного листа")
                    if title_keys:
                        for m_key in title_keys:
                            m_name = metadata_labels.get(m_key, m_key)
                            m_val = meta_data.get(m_key) or {}
                            coord = m_val.get("coordinate") if isinstance(m_val, dict) else ""
                            sheet = m_val.get("sheet") if isinstance(m_val, dict) else ""
                            help_txt = f"Ячейка {coord} на листе '{sheet}'" if coord else "Значение по умолчанию"
                            
                            current_val = st.session_state["import_metadata_edits"].get(m_key, "")
                            new_val = st.text_input(
                                m_name,
                                value=current_val,
                                key=f"edit_input_{m_key}",
                                help=help_txt
                            )
                            st.session_state["import_metadata_edits"][m_key] = new_val
                    else:
                        st.caption("Данные не найдены на Титульном листе.")
                        
                with c_meta2:
                    st.markdown("#### 📋 Данные из Общих сведений")
                    if general_keys:
                        for m_key in general_keys:
                            m_name = metadata_labels.get(m_key, m_key)
                            m_val = meta_data.get(m_key) or {}
                            coord = m_val.get("coordinate") if isinstance(m_val, dict) else ""
                            sheet = m_val.get("sheet") if isinstance(m_val, dict) else ""
                            help_txt = f"Ячейка {coord} на листе '{sheet}'" if coord else "Значение по умолчанию"
                            
                            current_val = st.session_state["import_metadata_edits"].get(m_key, "")
                            new_val = st.text_input(
                                m_name,
                                value=current_val,
                                key=f"edit_input_{m_key}",
                                help=help_txt
                            )
                            st.session_state["import_metadata_edits"][m_key] = new_val
                    else:
                        st.caption("Данные не найдены на листе Общих сведений.")

                # Разрешение конфликтов ФИО (только для САПР файлов с ИИ)
                fio_resolutions = {}
                if is_sapr and connected and final_mapping:
                    try:
                        class MockDB:
                            def __getattr__(self, name): return self
                            def insert_one(self, doc):
                                class MockResult: inserted_id = ObjectId()
                                return MockResult()
                            def update_one(self, q, u): pass
                            def find_one(self, query=None): return None
                            def find(self, query=None, projection=None): return []
                        mock_db = MockDB()
                        
                        if hasattr(uploaded_file, "seek"):
                            uploaded_file.seek(0)
                        temp_data = run_parsing_pipeline_dynamic(uploaded_file, final_mapping, mock_db)
                        
                        if temp_data:
                            unique_teachers = [t for t in temp_data.keys() if t and t != "Не указан"]
                            conflicts_found = False
                            for teacher in unique_teachers:
                                similar = find_similar_teacher(teacher, active_db)
                                if similar and similar["fio"] != teacher:
                                    conflicts_found = True
                                    break
                                    
                            if conflicts_found:
                                st.markdown("#### ⚠️ Разрешение конфликтов ФИО преподавателей")
                                st.caption("Найдены совпадения в базе данных. Выберите, заменить ли сокращенные имена на полные ФИО:")
                                
                                for teacher in unique_teachers:
                                    similar = find_similar_teacher(teacher, active_db)
                                    if similar and similar["fio"] != teacher:
                                        res_opt = st.selectbox(
                                            f"ФИО в файле: '{teacher}'",
                                            options=[
                                                f"Заменить на '{similar['fio']}' (из БД)",
                                                f"Оставить '{teacher}' как есть",
                                                "Ввести полное ФИО вручную"
                                            ],
                                            key=f"fio_res_import_{teacher}"
                                        )
                                        if res_opt.startswith("Заменить на"):
                                            fio_resolutions[teacher] = similar["fio"]
                                        elif res_opt == "Ввести полное ФИО вручную":
                                            custom_val = st.text_input(f"Полное ФИО для '{teacher}':", value=teacher, key=f"fio_custom_import_{teacher}")
                                            fio_resolutions[teacher] = custom_val
                                        else:
                                            fio_resolutions[teacher] = teacher
                    except Exception:
                        pass

            # Кнопка запуска импорта в один клик
            if st.button("📥 Запустить автоматический импорт в БД", use_container_width=True):
                with st.spinner("Выполняется автоматический парсинг и импорт..."):
                    try:
                        if hasattr(uploaded_file, "seek"):
                            uploaded_file.seek(0)
                            
                        # Применяем резолюции ФИО к маппингу
                        if fio_resolutions and final_mapping:
                            for short_name, full_name in fio_resolutions.items():
                                if short_name != full_name:
                                    emp = active_db.employees.find_one({"fio": full_name})
                                    if emp:
                                        aliases = emp.get("aliases", [])
                                        if short_name not in aliases:
                                            aliases.append(short_name)
                                            active_db.employees.update_one({"_id": emp["_id"]}, {"$set": {"aliases": aliases}})

                        # Применяем отредактированные пользователем метаданные из UI (все 20 полей)
                        if final_mapping and "metadata" in final_mapping:
                            for m_key in metadata_labels.keys():
                                new_val = st.session_state.get("import_metadata_edits", {}).get(m_key, "")
                                if m_key not in final_mapping["metadata"]:
                                    final_mapping["metadata"][m_key] = {"value": new_val, "coordinate": "", "sheet": ""}
                                else:
                                    if isinstance(final_mapping["metadata"][m_key], dict):
                                        final_mapping["metadata"][m_key]["value"] = new_val
                                    else:
                                        final_mapping["metadata"][m_key] = new_val

                        # Импортируем данные
                        if is_sapr:
                            if connected and final_mapping:
                                import_data = run_parsing_pipeline_dynamic(uploaded_file, final_mapping, active_db, fio_resolutions=fio_resolutions)
                            else:
                                import_data = run_parsing_pipeline(uploaded_file, active_db)
                                
                            st.session_state["last_imported_report"] = {
                                "filename": uploaded_file.name,
                                "type": "sapr",
                                "data": import_data
                            }
                        else:
                            if connected and final_mapping:
                                import_data = run_profile_parsing_pipeline_dynamic(uploaded_file, final_mapping, active_db)
                            else:
                                import_data = run_profile_parsing_pipeline(uploaded_file, active_db)
                                
                            st.session_state["last_imported_report"] = {
                                "filename": uploaded_file.name,
                                "type": "profile",
                                "data": import_data
                            }
                            
                        st.success(f"🎉 Данные из `{uploaded_file.name}` успешно импортированы!")
                        st.rerun()
                    except Exception as e_import:
                        st.error(f"Ошибка при парсинге: {e_import}")
        else:
            st.info("Пожалуйста, выберите Excel-файл для импорта.")
            
    with col2:
        st.subheader("📋 Распределение нагрузки по преподавателям")
        
        # Переключатель БД для просмотра
        view_db_options = ["📗 БД САПР (university_db)"]
        if workload_db is not None:
            view_db_options.append("📘 БД Нагрузка (workload_db)")
        
        selected_view_db = st.radio("Просмотр базы данных:", view_db_options, horizontal=True, key="view_db_radio")
        
        if selected_view_db.startswith("📘") and workload_db is not None:
            view_db = workload_db
            view_db_name = "БД Нагрузка"
        else:
            view_db = db
            view_db_name = "БД САПР"
        
        teachers = list(view_db.employees.find())
        
        if teachers:
            # Вычисление метрик в зависимости от типа БД
            is_view_sapr = not selected_view_db.startswith("📘")
            
            if is_view_sapr:
                total_hours_sum = 0
                total_subjects = set()
                for t in teachers:
                    for load in t.get("loads", []):
                        total_hours_sum += load.get("total", 0.0)
                        total_subjects.add(load.get("subject", ""))
                
                m_col1, m_col2, m_col3 = st.columns(3)
                with m_col1:
                    st.metric("Преподавателей в системе", len(teachers))
                with m_col2:
                    st.metric("Всего учебных часов", f"{total_hours_sum:.1f} ч.")
                with m_col3:
                    st.metric("Уникальных дисциплин", len(total_subjects))
            else:
                # Вторая БД (профили/анкеты)
                total_subjects = set()
                for t in teachers:
                    for s in t.get("subjects", []):
                        total_subjects.add(s)
                
                m_col1, m_col2, m_col3 = st.columns(3)
                with m_col1:
                    st.metric("Анкет преподавателей", len(teachers))
                with m_col2:
                    st.metric("Уникальных предметов", len(total_subjects))
                with m_col3:
                    positions = [t.get("position", "") for t in teachers if t.get("position")]
                    unique_pos = len(set(positions))
                    st.metric("Различных должностей", unique_pos)
                
            st.write("---")
            st.markdown(f"### 📂 Личные папки преподавателей ({view_db_name})")
            
            for t in teachers:
                fio = t.get("fio", "Не указан")
                
                if not is_view_sapr:
                    # Вторая БД: отображаем анкетные данные преподавателя (с возможностью редактирования)
                    position = t.get("position", "Не указана")
                    degree = t.get("degree", "—")
                    title = t.get("title", "—")
                    with st.expander(f"👤 {fio} — {position} ({degree} / {title})"):
                        with st.form(key=f"edit_form_{t['_id']}"):
                            p_col1, p_col2 = st.columns(2)
                            with p_col1:
                                new_fio = st.text_input("ФИО Преподавателя", value=fio)
                                new_position = st.text_input("Должность", value=position)
                                new_conditions = st.text_input("Условия привлечения / Ставка", value=t.get('employment_conditions', ''))
                                new_degree = st.text_input("Ученая степень", value=degree)
                                new_title = st.text_input("Ученое звание", value=title)
                                
                                contract = t.get("contract", {}) or {}
                                new_contract_num = st.text_input("Номер контракта", value=contract.get('number', ''))
                                new_contract_date = st.text_input("Дата контракта", value=contract.get('date', ''))
                                new_contract_dur = st.text_input("Срок действия контракта", value=contract.get('duration', ''))
                            with p_col2:
                                edu = t.get("education", {}) or {}
                                new_edu_inst = st.text_input("Образование (учреждение)", value=edu.get('institution', ''))
                                new_edu_year = st.text_input("Год окончания", value=edu.get('year', ''))
                                new_edu_spec = st.text_input("Специальность", value=edu.get('specialty', ''))
                                new_edu_qual = st.text_input("Квалификация", value=edu.get('qualification', ''))
                                
                                subjects = t.get("subjects", [])
                                subjects_str = ", ".join(subjects) if isinstance(subjects, list) else str(subjects)
                                new_subjects_str = st.text_input("Преподаваемые дисциплины (через запятую)", value=subjects_str)

                            # Кнопка сохранения изменений
                            save_btn = st.form_submit_button("💾 Сохранить изменения")
                            if save_btn:
                                new_subjects = [s.strip() for s in new_subjects_str.split(",") if s.strip()]
                                updated_doc = {
                                    "fio": new_fio,
                                    "position": new_position,
                                    "employment_conditions": new_conditions,
                                    "degree": new_degree,
                                    "title": new_title,
                                    "contract": {
                                        "number": new_contract_num,
                                        "date": new_contract_date,
                                        "duration": new_contract_dur
                                    },
                                    "education": {
                                        "institution": new_edu_inst,
                                        "year": new_edu_year,
                                        "specialty": new_edu_spec,
                                        "qualification": new_edu_qual
                                    },
                                    "subjects": new_subjects
                                }
                                view_db.employees.update_one({"_id": t["_id"]}, {"$set": updated_doc})
                                st.success(f"🎉 Данные преподавателя {new_fio} успешно обновлены!")
                                st.rerun()
                else:
                    # Первая БД (САПР): отображаем учебную нагрузку
                    loads = t.get("loads", [])
                    teacher_hours = sum(load.get("total", 0.0) for load in loads)
                    
                    with st.expander(f"📁 {fio} — Общая нагрузка: {teacher_hours:.1f} ч."):
                        if loads:
                            rows = []
                            for load in loads:
                                row = {
                                    "Дисциплина": load.get("subject", ""),
                                    "Группа": load.get("group", ""),
                                    "Направление": load.get("direction", ""),
                                    "Семестр": load.get("semester", "1"),
                                    "Лекции": load.get("hours", {}).get("lectures", 0.0),
                                    "Практики": load.get("hours", {}).get("practicals", 0.0),
                                    "Лабораторные": load.get("hours", {}).get("laboratories", 0.0),
                                    "Консультации": load.get("hours", {}).get("consultations", 0.0),
                                    "Экзамен": load.get("hours", {}).get("exams", 0.0),
                                    "Зачет": load.get("hours", {}).get("zachets", 0.0),
                                    "КР/КП": load.get("hours", {}).get("coursework", 0.0),
                                    "Практика": load.get("hours", {}).get("practice", 0.0),
                                    "ВКР": load.get("hours", {}).get("vkr", 0.0),
                                    "ГЭК": load.get("hours", {}).get("gek", 0.0),
                                    "Доп. часы": load.get("hours", {}).get("additional", 0.0),
                                    "Итого": load.get("total", 0.0)
                                }
                                rows.append(row)
                            
                            df_loads = pd.DataFrame(rows)
                            st.dataframe(df_loads, use_container_width=True, hide_index=True)
                        else:
                            st.info("У данного преподавателя нет назначенной нагрузки.")
            
            st.write("---")
            with st.expander("🔍 Показать сырые таблицы базы данных для проверки"):
                # Переключатель БД для инспектора
                inspector_db_opts = ["📗 БД САПР"]
                if workload_db is not None:
                    inspector_db_opts.append("📘 БД Нагрузка")
                inspector_db_sel = st.radio("Инспектировать БД:", inspector_db_opts, horizontal=True, key="inspector_db_radio")
                
                if inspector_db_sel.startswith("📘") and workload_db is not None:
                    inspect_db = workload_db
                else:
                    inspect_db = db
                
                collections = {
                    "Преподаватели и нагрузка (employees)": "employees",
                    "Кафедры (departments)": "departments",
                    "Институты (institutes)": "institutes",
                    "Группы (groups)": "groups",
                    "Студенты (students)": "students",
                    "Практика (practices)": "practices",
                    "Предметы (subjects)": "subjects",
                    "Черновик нераспределенной нагрузки (unassigned_loads)": "unassigned_loads",
                    "История планов (individual_plans)": "individual_plans"
                }
                
                selected_col_name = st.selectbox("Выберите коллекцию для просмотра:", list(collections.keys()))
                col_id = collections[selected_col_name]
                
                raw_data = list(inspect_db[col_id].find())
                if raw_data:
                    flat_data = []
                    for doc in raw_data:
                        flat_doc = {}
                        for k, v in doc.items():
                            if isinstance(v, ObjectId):
                                flat_doc[k] = str(v)
                            elif isinstance(v, dict):
                                for sub_k, sub_v in v.items():
                                    flat_doc[f"{k}_{sub_k}"] = str(sub_v)
                            elif isinstance(v, list):
                                flat_doc[k] = ", ".join([str(item) for item in v])
                            else:
                                flat_doc[k] = v
                        flat_data.append(flat_doc)
                    st.dataframe(flat_data, use_container_width=True)
                else:
                    st.info("Выбранная коллекция пуста.")
        else:
            st.info(f"В {view_db_name} нет загруженных данных. Пожалуйста, импортируйте Excel-файл нагрузки в левой колонке.")

# --- СТРАНИЦА 2: РАСПРЕДЕЛЕНИЕ НАГРУЗКИ ---
elif page == "🔗 Распределение нагрузки":
    st.subheader("🔗 Распределение неназначенной нагрузки")
    st.markdown("В этом разделе показаны строки учебной нагрузки, у которых не был указан преподаватель. Выберите вкладку базы данных.")
    
    # Вкладки для двух БД
    dist_tab_names = ["📗 САПР-нагрузка (university_db)"]
    if workload_db is not None:
        dist_tab_names.append("📘 Общая нагрузка (workload_db)")
    
    dist_tabs = st.tabs(dist_tab_names)
    
    def render_distribution_tab(tab_db, tab_label, tab_key_prefix):
        """Отрисовка содержимого вкладки распределения нагрузки для заданной БД."""
        unassigned = list(tab_db.unassigned_loads.find())
        
        if not unassigned:
            st.success(f"✅ В {tab_label} все нагрузки успешно распределены по преподавателям!")
        else:
            rows = []
            for idx, item in enumerate(unassigned):
                rows.append({
                    "ID": str(item["_id"]),
                    "Дисциплина": item.get("subject", ""),
                    "Группа": item.get("group", ""),
                    "Направление": item.get("direction", ""),
                    "Семестр": item.get("semester", ""),
                    "Лекции": item.get("hours", {}).get("lectures", 0.0),
                    "Практики": item.get("hours", {}).get("practicals", 0.0),
                    "Лабораторные": item.get("hours", {}).get("laboratories", 0.0),
                    "Консультации": item.get("hours", {}).get("consultations", 0.0),
                    "Экзамен": item.get("hours", {}).get("exams", 0.0),
                    "Зачет": item.get("hours", {}).get("zachets", 0.0),
                    "КР/КП": item.get("hours", {}).get("coursework", 0.0),
                    "Итого": item.get("total", 0.0)
                })
            df_unassigned = pd.DataFrame(rows)
            
            st.write(f"Найдено нераспределенных строк нагрузки в **{tab_label}**: **{len(unassigned)}**")
            st.write("Выберите строки нагрузки для привязки к преподавателю:")
            
            df_unassigned.insert(0, "Выбрать", False)
            edited_df = st.data_editor(
                df_unassigned,
                column_config={"Выбрать": st.column_config.CheckboxColumn(required=True)},
                disabled=[c for c in df_unassigned.columns if c != "Выбрать"],
                use_container_width=True,
                hide_index=True,
                key=f"{tab_key_prefix}_data_editor"
            )
            
            chosen_rows = edited_df[edited_df["Выбрать"] == True]
            
            if not chosen_rows.empty:
                selected_ids = chosen_rows["ID"].tolist()
                st.write(f"Выбрано строк нагрузки: **{len(selected_ids)}**")
                
                teachers = list(tab_db.employees.find())
                teacher_names = [t["fio"] for t in teachers]
                teacher_names.insert(0, "-- Создать нового преподавателя --")
                
                dcol1, dcol2 = st.columns(2)
                with dcol1:
                    selected_teacher = st.selectbox("Привязать к преподавателю:", teacher_names, key=f"{tab_key_prefix}_teacher_sel")
                with dcol2:
                    new_teacher_name = st.text_input("Или введите ФИО нового преподавателя:", value="", key=f"{tab_key_prefix}_new_teacher")
                    
                if st.button("🚀 Выполнить привязку нагрузки", key=f"{tab_key_prefix}_assign_btn"):
                    target_fio = ""
                    if selected_teacher == "-- Создать нового преподавателя --":
                        target_fio = new_teacher_name.strip()
                        if not target_fio:
                            st.error("Пожалуйста, укажите ФИО нового преподавателя!")
                            st.stop()
                    else:
                        target_fio = selected_teacher
                        
                    existing_emp = tab_db.employees.find_one({"fio": target_fio})
                    if existing_emp:
                        emp_id = existing_emp["_id"]
                        emp_loads = existing_emp.get("loads", [])
                    else:
                        emp_doc = {
                            "fio": target_fio,
                            "loads": []
                        }
                        emp_id = tab_db.employees.insert_one(emp_doc).inserted_id
                        emp_loads = []
                        
                    count_transferred = 0
                    for u_row in unassigned:
                        if str(u_row["_id"]) in selected_ids:
                            load_entry = {
                                "subject": u_row.get("subject"),
                                "group": u_row.get("group"),
                                "direction": u_row.get("direction"),
                                "semester": u_row.get("semester"),
                                "hours": u_row.get("hours"),
                                "total": u_row.get("total")
                            }
                            
                            found_idx = -1
                            for idx, existing_load in enumerate(emp_loads):
                                match_subj = str(existing_load.get("subject", "")).strip().lower() == str(load_entry["subject"]).strip().lower()
                                match_group = str(existing_load.get("group", "")).strip().lower() == str(load_entry["group"]).strip().lower()
                                match_sem = str(existing_load.get("semester", "")).strip().lower() == str(load_entry["semester"]).strip().lower()
                                if match_subj and match_group and match_sem:
                                    found_idx = idx
                                    break
                                    
                            if found_idx != -1:
                                emp_loads[found_idx] = load_entry
                            else:
                                emp_loads.append(load_entry)
                                
                            tab_db.unassigned_loads.delete_many({"_id": u_row["_id"]})
                            count_transferred += 1
                            
                    tab_db.employees.update_one({"_id": emp_id}, {"$set": {"loads": emp_loads}})
                    st.success(f"🎉 Успешно перенесено {count_transferred} строк нагрузки преподавателю **{target_fio}** в {tab_label}!")
                    st.rerun()
            else:
                st.info("Пожалуйста, отметьте галочками строки нагрузки в таблице выше, чтобы привязать их к преподавателю.")
                
            st.write("---")
            if st.button("🗑️ Очистить все нераспределенные нагрузки (сбросить черновик)", key=f"{tab_key_prefix}_clear_btn"):
                tab_db.unassigned_loads.delete_many({})
                st.success(f"Черновик нераспределенной нагрузки в {tab_label} успешно очищен.")
                st.rerun()
    
    # Рендерим первую вкладку (САПР)
    with dist_tabs[0]:
        render_distribution_tab(db, "БД САПР", "sapr")
    
    # Рендерим вторую вкладку (Нагрузка), если доступна
    if workload_db is not None and len(dist_tabs) > 1:
        with dist_tabs[1]:
            render_distribution_tab(workload_db, "БД Нагрузка", "workload")

# --- СТРАНИЦА 3: РЕДАКТИРОВАНИЕ НАГРУЗКИ ---
elif page == "✏️ Редактирование нагрузки":
    st.subheader("✏️ Перенос нагрузки и генерация индивидуального плана")
    st.markdown("В этом разделе вы можете выбрать дополнительные дисциплины из БД САПР, привязать их к преподавателю из БД и сгенерировать готовый файл индивидуального плана.")
    
    # 1. Загружаем список преподавателей из обеих БД
    db_sapr_teachers = [t["fio"] for t in db.employees.find() if t.get("fio")]
    db_wl_teachers = []
    if workload_db is not None:
        db_wl_teachers = [t["fio"] for t in workload_db.employees.find() if t.get("fio")]
        
    unique_teachers = sorted(list(set(db_sapr_teachers + db_wl_teachers)))
    
    if not unique_teachers:
        st.info("В базах данных нет преподавателей. Сначала импортируйте файлы нагрузок на первой вкладке.")
    else:
        selected_fio = st.selectbox("Шаг 1: Выберите преподавателя для генерации индивидуального плана:", unique_teachers)
        
        # Загружаем анкетные данные и текущую нагрузку преподавателя из БД
        profile_doc = {}
        if workload_db is not None:
            matched_wl = find_similar_teacher(selected_fio, workload_db)
            if matched_wl:
                profile_doc = matched_wl
                
        matched_sapr = find_similar_teacher(selected_fio, db)
        
        fio = selected_fio
        position = profile_doc.get("position") or (matched_sapr.get("position") if matched_sapr else "") or "преподаватель"
        employment_conditions = profile_doc.get("employment_conditions") or (matched_sapr.get("employment_conditions") if matched_sapr else "") or "штатный"
        degree = profile_doc.get("degree") or (matched_sapr.get("degree") if matched_sapr else "") or ""
        title = profile_doc.get("title") or (matched_sapr.get("title") if matched_sapr else "") or ""
        
        contract_number = ""
        contract_date = ""
        contract_duration = ""
        if profile_doc.get("contract"):
            contract_number = profile_doc["contract"].get("number", "")
            contract_date = profile_doc["contract"].get("date", "")
            contract_duration = profile_doc["contract"].get("duration", "")
        elif matched_sapr and matched_sapr.get("contract"):
            contract_number = matched_sapr["contract"].get("number", "")
            contract_date = matched_sapr["contract"].get("date", "")
            contract_duration = matched_sapr["contract"].get("duration", "")
            
        edu_institution = ""
        edu_year = ""
        edu_specialty = ""
        edu_qualification = ""
        if profile_doc.get("education"):
            edu_institution = profile_doc["education"].get("institution", "")
            edu_year = profile_doc["education"].get("year", "")
            edu_specialty = profile_doc["education"].get("specialty", "")
            edu_qualification = profile_doc["education"].get("qualification", "")
        elif matched_sapr and matched_sapr.get("education"):
            edu_institution = matched_sapr["education"].get("institution", "")
            edu_year = matched_sapr["education"].get("year", "")
            edu_specialty = matched_sapr["education"].get("specialty", "")
            edu_qualification = matched_sapr["education"].get("qualification", "")
            
        current_loads = profile_doc.get("loads", []) or (matched_sapr.get("loads", []) if matched_sapr else [])
        
        st.markdown(f"**👤 Текущий профиль преподавателя:** должность: *{position}*, степень/звание: *{degree or 'нет'} / {title or 'нет'}*")
        st.write(f"Текущая нагрузка в БД: **{len(current_loads)}** строк.")
        
        st.write("---")
        st.markdown("### Шаг 2: Выберите дополнительные дисциплины из БД САПР")
        
        # Собираем все нагрузки из БД САПР
        sapr_loads = []
        for emp in db.employees.find():
            emp_fio = emp.get("fio", "Не указан")
            for load in emp.get("loads", []):
                sapr_loads.append({
                    "subject": load.get("subject", ""),
                    "group": load.get("group", ""),
                    "direction": load.get("direction", ""),
                    "semester": str(load.get("semester", "1")),
                    "hours": load.get("hours", {}),
                    "total": load.get("total", 0.0),
                    "teacher_fio": emp_fio
                })
        for load in db.unassigned_loads.find():
            sapr_loads.append({
                "subject": load.get("subject", ""),
                "group": load.get("group", ""),
                "direction": load.get("direction", ""),
                "semester": str(load.get("semester", "1")),
                "hours": load.get("hours", {}),
                "total": load.get("total", 0.0),
                "teacher_fio": "Не распределена"
            })
            
        if not sapr_loads:
            st.warning("⚠️ В БД САПР нет строк нагрузки.")
        else:
            # Вычисляем дефолтные чекбоксы для выбранного преподавателя
            from parser_engine import clean_fio
            default_checks = []
            cleaned_selected = clean_fio(selected_fio).lower()
            for item in sapr_loads:
                is_same = False
                if item.get("teacher_fio"):
                    is_same = clean_fio(item.get("teacher_fio")).lower() == cleaned_selected
                default_checks.append(is_same)
                
            rows = []
            for idx, item in enumerate(sapr_loads):
                rows.append({
                    "Index": idx,
                    "Дисциплина": item.get("subject", ""),
                    "Группа": item.get("group", ""),
                    "Направление": item.get("direction", ""),
                    "Семестр": item.get("semester", ""),
                    "Лекции": float(item.get("hours", {}).get("lectures", 0.0)),
                    "Практики": float(item.get("hours", {}).get("practicals", 0.0)),
                    "Лабораторные": float(item.get("hours", {}).get("laboratories", 0.0)),
                    "Консультации": float(item.get("hours", {}).get("consultations", 0.0)),
                    "Экзамен": float(item.get("hours", {}).get("exams", 0.0)),
                    "Зачет": float(item.get("hours", {}).get("zachets", 0.0)),
                    "КР/КП": float(item.get("hours", {}).get("coursework", 0.0)),
                    "Практика": float(item.get("hours", {}).get("practice", 0.0)),
                    "ВКР": float(item.get("hours", {}).get("vkr", 0.0)),
                    "ГЭК": float(item.get("hours", {}).get("gek", 0.0)),
                    "Доп. часы": float(item.get("hours", {}).get("additional", 0.0)),
                    "Итого": float(item.get("total", 0.0)),
                    "Преподаватель в САПР": item.get("teacher_fio", "")
                })
            df_sapr = pd.DataFrame(rows)
            df_sapr.insert(0, "Выбрать", default_checks)
            
            # Динамический ключ виджета привязан к ФИО преподавателя, чтобы сбрасывать и обновлять чекбоксы
            widget_key = f"p3_sapr_editor_{selected_fio.replace(' ', '_')}"
            
            edited_df = st.data_editor(
                df_sapr,
                column_config={
                    "Выбрать": st.column_config.CheckboxColumn(required=True),
                    "Index": None
                },
                disabled=[c for c in df_sapr.columns if c != "Выбрать"],
                use_container_width=True,
                hide_index=True,
                key=widget_key
            )
            
            chosen_rows = edited_df[edited_df["Выбрать"] == True]
            st.write(f"Выбрано дополнительных строк нагрузки из САПР: **{len(chosen_rows)}**")
            
            st.write("---")
            st.markdown("### Шаг 3: Шаблон индивидуального плана и сопоставление заглушек")
            
            uploaded_template = st.file_uploader(
                "📂 Загрузите файл-шаблон индивидуального плана (.xlsx):",
                type=["xlsx"],
                key="p3_template_uploader"
            )
            
            default_template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_uploaded_p3_Копия individ_plrprGOD Шаблон для заполнения с переменными.xlsx")
            if not os.path.exists(default_template_path):
                default_template_path = "C:/Users/vrlab/Desktop/Копия individ_plrprGOD Шаблон для заполнения с переменными.xlsx"
            if not os.path.exists(default_template_path) and os.path.exists("individ_plan_Архипов.xlsx"):
                default_template_path = "individ_plan_Архипов.xlsx"
                
            template_path = None
            if uploaded_template is not None:
                template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"_uploaded_p3_{uploaded_template.name}")
                with open(template_path, "wb") as _f:
                    _f.write(uploaded_template.getbuffer())
                st.success(f"✅ Готов шаблон: `{uploaded_template.name}`")
            elif os.path.exists(default_template_path):
                template_path = default_template_path
                st.info(f"Используется стандартный шаблон: `{os.path.basename(template_path)}`")
            else:
                st.warning("⚠️ Загрузите файл-шаблон (.xlsx) через поле выше.")
                
            # Инициализируем стандартными заглушками по умолчанию
            placeholders = {
                "employee_fio": "{{ employee_fio }}",
                "employee_fio_title": "{{ employee_fio_title }}",
                "employee_position": "{{ employee_position }}",
                "employee_position_title": "{{ employee_position_title }}",
                "employee_rate": "{{ employee_rate }}",
                "employee_conditions": "{{ employee_conditions }}",
                "employee_degree": "{{ employee_degree }}",
                "employee_title": "{{ employee_title }}",
                "employee_contract": "{{ employee_contract }}",
                "employee_contract_duration": "{{ employee_contract_duration }}",
                "employee_edu_inst_year": "{{ employee_edu_inst_year }}",
                "employee_edu_specialty_1": "{{ employee_edu_specialty_1 }}",
                "employee_edu_specialty_2": "{{ employee_edu_specialty_2 }}",
                "employee_edu_qualification": "{{ employee_edu_qualification }}",
                "department_name": "{{ department_name }}",
                "department_head": "{{ department_head }}",
                "institute_name": "{{ institute_name }}",
                "study_year": "{{ study_year }}"
            }
            if template_path and os.path.exists(template_path):
                try:
                    from parser_engine import get_cell_val_by_coord, STRICT_PLAN_COORDINATES
                    for key, coord_info in STRICT_PLAN_COORDINATES.items():
                        sheet_name = coord_info["sheet"]
                        coord = coord_info["coordinate"]
                        
                        df_resolve = pd.read_excel(template_path, sheet_name=sheet_name, header=None)
                        if "общие" in sheet_name.lower():
                            df_resolve = df_resolve.T
                            
                        val = get_cell_val_by_coord(df_resolve, coord)
                        if val is not None and str(val).strip():
                            placeholders[key] = str(val).strip()
                except Exception:
                    pass
            
            metadata_labels = {
                "employee_fio": "ФИО Преподавателя",
                "employee_fio_title": "ФИО Преподавателя (Титульный)",
                "employee_position": "Должность",
                "employee_position_title": "Должность (Титульный)",
                "employee_rate": "Размер ставки",
                "employee_conditions": "Условия привлечения",
                "employee_degree": "Ученая степень",
                "employee_title": "Ученое звание",
                "employee_contract": "Договор (дата, номер)",
                "employee_contract_duration": "Срок действия договора",
                "employee_edu_inst_year": "Образовательная организация и год окончания",
                "employee_edu_specialty_1": "Направление подготовки (код)",
                "employee_edu_specialty_2": "Направление подготовки (наименование)",
                "employee_edu_qualification": "Квалификация",
                "department_name": "Кафедра",
                "department_head": "Заведующий кафедрой",
                "institute_name": "Институт",
                "study_year": "Учебный год"
            }
            
            # Разделяем специальность на код и наименование
            spec_parts = edu_specialty.split(" ", 1) if edu_specialty else []
            spec1 = spec_parts[0] if len(spec_parts) > 0 else edu_specialty
            spec2 = spec_parts[1] if len(spec_parts) > 1 else ""
            
            db_values = {
                "employee_fio": selected_fio,
                "employee_fio_title": selected_fio,
                "employee_position": position,
                "employee_position_title": position,
                "employee_rate": employment_conditions,
                "employee_conditions": employment_conditions,
                "employee_degree": degree,
                "employee_title": title,
                "employee_contract": contract_number + (" от " + contract_date if contract_date else ""),
                "employee_contract_duration": contract_duration,
                "employee_edu_inst_year": edu_institution + (" " + edu_year if edu_year else ""),
                "employee_edu_specialty_1": spec1,
                "employee_edu_specialty_2": spec2,
                "employee_edu_qualification": edu_qualification,
                "department_name": matched_sapr.get("department_name", "") if matched_sapr else "",
                "department_head": matched_sapr.get("department_head", "") if matched_sapr else "",
                "institute_name": matched_sapr.get("institute_name", "") if matched_sapr else "",
                "study_year": datetime.datetime.now().strftime("%Y")
            }
            
            st.markdown("#### Параметры сопоставления заглушек:")
            st.caption("Левое поле — заглушка в шаблоне (авто-поиск и замена по файлу), правое поле — значение из БД для вставки.")
            
            placeholder_inputs = {}
            
            # Отображаем форму сопоставления полей
            col_hdr1, col_hdr2 = st.columns(2)
            with col_hdr1:
                st.caption("**Заглушка в шаблоне**")
            with col_hdr2:
                st.caption("**Значение для вставки**")
                
            for key, label in metadata_labels.items():
                st.markdown(f"*{label}*")
                col_l, col_r = st.columns(2)
                with col_l:
                    p_val = st.text_input(
                        f"Заглушка для {label}",
                        value=placeholders.get(key, ""),
                        key=f"tpl_p_{key}_{selected_fio}",
                        label_visibility="collapsed"
                    )
                with col_r:
                    n_val = st.text_input(
                        f"Значение для {label}",
                        value=db_values.get(key, ""),
                        key=f"tpl_n_{key}_{selected_fio}",
                        label_visibility="collapsed"
                    )
                placeholder_inputs[key] = {
                    "placeholder": p_val,
                    "value": n_val
                }
                
            if st.button("🚀 Перенести нагрузку и сгенерировать индивидуальный план", use_container_width=True, key="p3_generate_btn"):
                if not template_path or not os.path.exists(template_path):
                    st.error("Пожалуйста, загрузите или укажите шаблон индивидуального плана.")
                else:
                    with st.spinner("Генерация индивидуального плана..."):
                        try:
                            # Извлекаем выбранные нагрузки
                            chosen_indices = chosen_rows["Index"].tolist()
                            transferred_loads = []
                            for idx in chosen_indices:
                                orig_item = sapr_loads[idx]
                                transferred_loads.append({
                                    "subject": orig_item["subject"],
                                    "group": orig_item["group"],
                                    "direction": orig_item["direction"],
                                    "semester": orig_item["semester"],
                                    "hours": orig_item["hours"],
                                    "total": orig_item["total"]
                                })
                            
                            # Используем только отмеченные пользователем дисциплины
                            combined_loads = transferred_loads
                            
                            placeholder_map = {}
                            for key in metadata_labels.keys():
                                placeholder_map[key] = {
                                    "placeholder": placeholder_inputs[key]["placeholder"],
                                    "value": placeholder_inputs[key]["value"]
                                }
                            
                            import excel_filler
                            import importlib
                            importlib.reload(excel_filler)
                            from excel_filler import fill_teacher_plan
                            output_path = f"individ_plan_{selected_fio.replace(' ', '_')}.xlsx"
                            
                            fill_teacher_plan(
                                template_path=template_path,
                                output_path=output_path,
                                loads=combined_loads,
                                placeholder_map=placeholder_map
                            )
                            
                            st.success(f"🎉 Документ индивидуального плана успешно сгенерирован: `{output_path}`!")
                            
                            with open(output_path, "rb") as f:
                                file_bytes = f.read()
                                
                            st.download_button(
                                label="📥 Скачать индивидуальный план",
                                data=file_bytes,
                                file_name=output_path,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        except Exception as e_gen:
                            st.error("Ошибка при генерации документа:")
                            st.exception(e_gen)


elif page == "📄 Индивидуальные планы и История":
    st.subheader("📄 Индивидуальные планы преподавателей и История")
    st.write("В этом разделе вы можете управлять индивидуальными планами преподавателей: загружать новые планы, отслеживать историю изменений (изменение ставок, должностей) и генерировать заполненные файлы по шаблонам.")
    
    teachers = list(db.employees.find())
    
    if teachers:
        col3_1, col3_2 = st.columns([1, 1])
        
        teacher_names = [t["fio"] for t in teachers]
        selected_fio = st.sidebar.selectbox("Выберите преподавателя:", teacher_names, key="tab3_fio_sel")
        teacher_doc = next(t for t in teachers if t["fio"] == selected_fio)
        
        # Загружаем анкету из второй БД (workload_db)
        profile_doc = {}
        if workload_db is not None:
            matched_profile = find_similar_teacher(selected_fio, workload_db)
            if matched_profile:
                profile_doc = matched_profile
                
        # Объединяем данные
        position = profile_doc.get("position") or teacher_doc.get("position") or "Не указана"
        conditions = profile_doc.get("employment_conditions") or teacher_doc.get("employment_conditions") or "Не указаны"
        degree = profile_doc.get("degree") or teacher_doc.get("degree") or "—"
        title = profile_doc.get("title") or teacher_doc.get("title") or "—"
        
        st.markdown(f"### 👤 Текущий профиль: **{selected_fio}**")
        col_m1, col_m2, col_m3, col_m4 = st.columns(4)
        with col_m1:
            st.metric("Должность", position)
        with col_m2:
            st.metric("Условия / Ставка", conditions)
        with col_m3:
            st.metric("Степень / Звание", f"{degree} / {title}")
        with col_m4:
            st.metric("Активных предметов", len(teacher_doc.get("loads", [])))
            
        st.write("---")
        
        with col3_1:
            st.markdown("### 📥 Загрузка нового плана и История")
            
            history = list(db.individual_plans.find({"employee_id": str(teacher_doc["_id"])}))
            
            if history:
                st.markdown("**💾 Сохраненная история индивидуальных планов:**")
                history_rows = []
                for plan in history:
                    history_rows.append({
                        "Учебный год": plan.get("year", "Не указан"),
                        "Должность": plan.get("position", ""),
                        "Ставка": plan.get("employment_conditions", ""),
                        "Кол-во предметов": len(plan.get("loads", [])),
                        "Загружен": plan.get("timestamp", "")[:10]
                    })
                st.dataframe(pd.DataFrame(history_rows), use_container_width=True, hide_index=True)
            else:
                st.info("История планов для этого преподавателя пуста.")
                
            st.write("---")
            st.markdown("#### Загрузить индивидуальный план преподавателя (.xlsx)")
            
            year_input = st.text_input("Введите учебный год плана (например, 2026):", value="2026", key="tab3_year_input")
            uploaded_plan = st.file_uploader(
                "📂 Выберите файл индивидуального плана:",
                type=["xlsx"],
                key="tab3_plan_uploader"
            )
            
            # Чекбокс для ИИ в индивидуальном плане
            use_ai_plan = st.checkbox("🔮 Использовать ИИ для разбора структуры индивидуального плана", value=True, key="tab3_use_ai_plan")
            
            if uploaded_plan is not None:
                # Чтение имен листов в плане
                if hasattr(uploaded_plan, "seek"):
                    uploaded_plan.seek(0)
                xls_plan_inspect = pd.ExcelFile(uploaded_plan)
                plan_sheets = xls_plan_inspect.sheet_names
                xls_plan_inspect.close()
                
                final_plan_mapping = None
                
                if use_ai_plan:
                    if not connected:
                        st.warning("⚠️ Локальный ИИ (Ollama) недоступен. Будет применен стандартный парсер планов.")
                        use_ai_plan = False
                    else:
                        user_selected_plan_sheet = st.session_state.get("plan_sheet_sel")
                        if user_selected_plan_sheet and user_selected_plan_sheet not in plan_sheets:
                            user_selected_plan_sheet = None
                            if "plan_sheet_sel" in st.session_state:
                                del st.session_state["plan_sheet_sel"]
                                
                        if user_selected_plan_sheet:
                            cache_key_plan = f"ai_plan_map_{uploaded_plan.name}_{user_selected_plan_sheet}"
                            target_plan_sheet_to_analyze = user_selected_plan_sheet
                        else:
                            cache_key_plan = f"ai_plan_map_{uploaded_plan.name}_auto"
                            target_plan_sheet_to_analyze = None
                            
                        if cache_key_plan not in st.session_state:
                            with st.spinner("🤖 ИИ анализирует индивидуальный план..."):
                                try:
                                    if hasattr(uploaded_plan, "seek"):
                                        uploaded_plan.seek(0)
                                    mapping = ai_parser_helper.analyze_excel_structure(
                                        uploaded_plan, selected_model, ollama_url, target_sheet=target_plan_sheet_to_analyze
                                    )
                                    st.session_state[cache_key_plan] = mapping
                                    if mapping and mapping.get("sheet_name"):
                                        sheet_name = mapping["sheet_name"]
                                        st.session_state[f"ai_plan_map_{uploaded_plan.name}_{sheet_name}"] = mapping
                                    st.success("🤖 Структура успешно распознана ИИ!")
                                except Exception as e_ai:
                                    st.error(f"Ошибка ИИ-анализа плана: {e_ai}. Применен стандартный парсер.")
                                    st.session_state[cache_key_plan] = None
                                    
                        ai_plan_mapping = st.session_state.get(cache_key_plan)
                        if ai_plan_mapping:
                            active_sheet_plan = st.session_state.get("plan_sheet_sel") or ai_plan_mapping.get("sheet_name")
                            if hasattr(uploaded_plan, "seek"):
                                uploaded_plan.seek(0)
                            _, col_names, active_sheet = ai_parser_helper.prepare_excel_preview(uploaded_plan, target_sheet=active_sheet_plan)
                            
                            with st.expander("🛠️ ИИ-настройка разметки индивидуального плана (дополнительно)", expanded=False):
                                st.caption("Вы можете изменить ИИ-метаданные и маппинг колонок:")
                                
                                # Метаданные профиля и другие сущности
                                tp_info = ai_plan_mapping.get("teacher_profile", {}) or {}
                                meta_info = ai_plan_mapping.get("metadata", {}) or {}
                                
                                tab3_meta_col1, tab3_meta_col2 = st.columns(2)
                                with tab3_meta_col1:
                                    p_fio = st.text_input("ФИО Преподавателя:", value=tp_info.get("fio") or meta_info.get("employee_fio") or selected_fio, key="tab3_p_fio")
                                    p_pos = st.text_input("Должность:", value=tp_info.get("position") or meta_info.get("employee_position") or "доцент", key="tab3_p_pos")
                                    p_cond = st.text_input("Условия / Ставка:", value=tp_info.get("employment_conditions") or meta_info.get("employee_conditions") or meta_info.get("employee_rate") or "1.0 ставки, штатный", key="tab3_p_cond")
                                    p_deg = st.text_input("Степень:", value=tp_info.get("degree") or meta_info.get("employee_degree") or "", key="tab3_p_deg")
                                    p_tit = st.text_input("Звание:", value=tp_info.get("title") or meta_info.get("employee_title") or "", key="tab3_p_tit")
                                    p_c_num = st.text_input("Номер контракта:", value=meta_info.get("employee_contract_num") or "", key="tab3_p_c_num")
                                    p_c_date = st.text_input("Дата контракта:", value=meta_info.get("employee_contract_date") or "", key="tab3_p_c_date")
                                    p_c_dur = st.text_input("Срок действия контракта:", value=meta_info.get("employee_contract_duration") or "", key="tab3_p_c_dur")
                                with tab3_meta_col2:
                                    p_e_inst = st.text_input("Образовательная организация:", value=meta_info.get("employee_edu_institution") or "", key="tab3_p_e_inst")
                                    p_e_year = st.text_input("Год окончания образования:", value=meta_info.get("employee_edu_year") or "", key="tab3_p_e_year")
                                    p_e_spec = st.text_input("Специальность образования:", value=meta_info.get("employee_edu_specialty") or "", key="tab3_p_e_spec")
                                    p_e_qual = st.text_input("Квалификация образования:", value=meta_info.get("employee_edu_qualification") or "", key="tab3_p_e_qual")
                                    p_dep_name = st.text_input("Название кафедры:", value=meta_info.get("department_name") or "", key="tab3_p_dep_name")
                                    p_dep_head = st.text_input("Зав. кафедрой:", value=meta_info.get("department_head") or "", key="tab3_p_dep_head")
                                    p_dep_dir = st.text_input("Направление кафедры:", value=meta_info.get("department_direction") or "", key="tab3_p_dep_dir")
                                    p_inst_name = st.text_input("Название института:", value=meta_info.get("institute_name") or "", key="tab3_p_inst_name")
                                    p_inst_dir = st.text_input("Директор института:", value=meta_info.get("institute_director") or "", key="tab3_p_inst_dir")
                                
                                # Схема колонок
                                schema_fields = {
                                    "subject_name": "Дисциплина",
                                    "group_name": "Группа",
                                    "semester_number": "Семестр",
                                    "hours_lectures": "Лекции",
                                    "hours_practicals": "Практики",
                                    "hours_laboratories": "Лабораторные",
                                    "consultations": "Консультации",
                                    "exams": "Экзамен",
                                    "zachets": "Зачет",
                                    "coursework": "КР/КП",
                                    "practice": "Практика",
                                    "vkr": "ВКР",
                                    "gek": "ГЭК",
                                    "additional": "Дополнительно",
                                    "total": "Итого часов"
                                }
                                
                                dropdown_options = ["-- Отсутствует --"] + col_names
                                updated_column_mapping = {}
                                
                                map_col1, map_col2 = st.columns(2)
                                for idx, (field_key, field_name) in enumerate(schema_fields.items()):
                                    target_col = map_col1 if idx % 2 == 0 else map_col2
                                    with target_col:
                                        current_idx = ai_plan_mapping.get("column_mapping", {}).get(field_key)
                                        # Обратная совместимость со старыми ключами
                                        if current_idx is None:
                                            old_key_map = {
                                                "subject_name": "subject", 
                                                "group_name": "group", 
                                                "semester_number": "semester",
                                                "hours_lectures": "lectures",
                                                "hours_practicals": "practicals",
                                                "hours_laboratories": "laboratories",
                                                "hours_consultations": "consultations",
                                                "hours_exams": "exams",
                                                "hours_zachets": "zachets",
                                                "hours_coursework": "coursework",
                                                "hours_practice": "practice",
                                                "hours_vkr": "vkr",
                                                "hours_gek": "gek",
                                                "hours_additional": "additional",
                                                "hours_total": "total"
                                            }
                                            if field_key in old_key_map:
                                                current_idx = ai_plan_mapping.get("column_mapping", {}).get(old_key_map[field_key])
                                                
                                        # Приведение типов для строковых индексов (например, "1" -> 1)
                                        if current_idx is not None:
                                            try:
                                                current_idx = int(current_idx)
                                            except (ValueError, TypeError):
                                                if isinstance(current_idx, str):
                                                    import re
                                                    dig_match = re.search(r'\d+', current_idx)
                                                    if dig_match:
                                                        current_idx = int(dig_match.group(0))
                                                    else:
                                                        current_idx = None
                                                else:
                                                    current_idx = None
                                                    
                                        default_val_idx = 0
                                        if current_idx is not None and isinstance(current_idx, int) and current_idx < len(col_names):
                                            default_val_idx = current_idx + 1
                                            
                                        selected_opt = st.selectbox(
                                            f"📍 {field_name}:",
                                            options=dropdown_options,
                                            index=default_val_idx,
                                            key=f"sel_plan_{field_key}"
                                        )
                                        
                                        if selected_opt == "-- Отсутствует --":
                                            updated_column_mapping[field_key] = None
                                        else:
                                            updated_column_mapping[field_key] = col_names.index(selected_opt)
                                            
                                st.write("---")
                                header_row_input = st.number_input(
                                    "Строка заголовка (0-based индекс):", 
                                    min_value=0, 
                                    value=int(ai_plan_mapping.get("header_row_index", 0)),
                                    key="plan_header_row"
                                )
                                
                                try:
                                    sheet_index = plan_sheets.index(ai_plan_mapping.get("sheet_name", active_sheet))
                                except:
                                    sheet_index = 0
                                selected_sheet = st.selectbox("Лист с нагрузкой:", plan_sheets, index=sheet_index, key="plan_sheet_sel")
                                
                                final_plan_mapping = {
                                    "document_type": "individual_plan",
                                    "header_row_index": header_row_input,
                                    "sheet_name": selected_sheet,
                                    "column_mapping": updated_column_mapping,
                                    "teacher_profile": {
                                        "fio": p_fio,
                                        "position": p_pos,
                                        "employment_conditions": p_cond,
                                        "degree": p_deg,
                                        "title": p_tit
                                    },
                                    "metadata": {
                                        "employee_fio": p_fio,
                                        "employee_position": p_pos,
                                        "employee_conditions": p_cond,
                                        "employee_rate": p_cond,
                                        "employee_degree": p_deg,
                                        "employee_title": p_tit,
                                        "employee_contract_num": p_c_num,
                                        "employee_contract_date": p_c_date,
                                        "employee_contract_duration": p_c_dur,
                                        "employee_edu_institution": p_e_inst,
                                        "employee_edu_year": p_e_year,
                                        "employee_edu_specialty": p_e_spec,
                                        "employee_edu_qualification": p_e_qual,
                                        "department_name": p_dep_name,
                                        "department_head": p_dep_head,
                                        "department_direction": p_dep_dir,
                                        "institute_name": p_inst_name,
                                        "institute_director": p_inst_dir
                                    }
                                }
                import hashlib
                mapping_hash = hashlib.md5(json.dumps(final_plan_mapping, default=str).encode('utf-8')).hexdigest() if final_plan_mapping else ""
                cache_trigger_key = f"{uploaded_plan.name}_{mapping_hash}_{year_input}"
                
                if "parsed_individual_plan" not in st.session_state or st.session_state.get("uploaded_plan_cache_trigger") != cache_trigger_key:
                    try:
                        if hasattr(uploaded_plan, "seek"):
                            uploaded_plan.seek(0)
                        
                        if use_ai_plan and final_plan_mapping:
                            parsed_plan = parse_individual_plan_dynamic(uploaded_plan, final_plan_mapping)
                        else:
                            parsed_plan = parse_individual_plan_file(uploaded_plan, selected_fio)
                            
                        st.session_state["parsed_individual_plan"] = parsed_plan
                        st.session_state["parsed_individual_year"] = parsed_plan.get("year") or year_input
                        st.session_state["uploaded_plan_name"] = uploaded_plan.name
                        st.session_state["uploaded_plan_cache_trigger"] = cache_trigger_key
                        st.success("Файл успешно прочитан и распарсен. Ознакомьтесь со сравнением ниже!")
                    except Exception as e_parse:
                        st.error(f"Не удалось распарсить индивидуальный план: {e_parse}")
                        
            # Если есть спарсенный план в состоянии - показываем сравнение
            if "parsed_individual_plan" in st.session_state:
                parsed_plan = st.session_state["parsed_individual_plan"]
                parsed_year = st.session_state["parsed_individual_year"]
                
                # Проверяем конфликты ФИО для индивидуального плана
                plan_teacher_fio = parsed_plan["fio"]
                similar = find_similar_teacher(plan_teacher_fio, db)
                
                resolved_plan_fio = plan_teacher_fio
                if similar and similar["fio"] != plan_teacher_fio:
                    st.warning(f"⚠️ В плане указано ФИО: **{plan_teacher_fio}**. В базе данных найден похожего преподаватель: **{similar['fio']}**.")
                    res_opt = st.selectbox(
                        "Выберите действие для ФИО в индивидуальном плане:",
                        options=[
                            f"Заменить на '{similar['fio']}' (из БД)",
                            f"Оставить '{plan_teacher_fio}' как есть",
                            "Ввести вручную"
                        ],
                        key="fio_res_plan"
                    )
                    if res_opt.startswith("Заменить на"):
                        resolved_plan_fio = similar["fio"]
                    elif res_opt == "Ввести вручную":
                        resolved_plan_fio = st.text_input("Введите ФИО вручную:", value=plan_teacher_fio, key="fio_custom_plan")
                    else:
                        resolved_plan_fio = plan_teacher_fio
                
                parsed_plan["fio"] = resolved_plan_fio
                
                st.markdown("### 🔍 Сравнение изменений с БД")
                
                col_prop, col_db, col_new = st.columns(3)
                with col_prop:
                    st.markdown("**Характеристика**")
                    st.write("Ставка / Условия")
                    st.write("Должность")
                    st.write("Ученая степень")
                    st.write("Ученое звание")
                with col_db:
                    st.markdown("**В базе данных**")
                    st.write(conditions)
                    st.write(position)
                    st.write(degree)
                    st.write(title)
                with col_new:
                    st.markdown("**Из нового файла**")
                    cond_changed = conditions != parsed_plan["employment_conditions"]
                    st.markdown(f"<span style='color:{'green' if cond_changed else 'black'}'>{parsed_plan['employment_conditions']}</span>", unsafe_allow_html=True)
                    
                    pos_changed = position != parsed_plan["position"]
                    st.markdown(f"<span style='color:{'green' if pos_changed else 'black'}'>{parsed_plan['position']}</span>", unsafe_allow_html=True)
                    
                    deg_changed = degree != parsed_plan["degree"]
                    st.markdown(f"<span style='color:{'green' if deg_changed else 'black'}'>{parsed_plan['degree']}</span>", unsafe_allow_html=True)
                    
                    title_changed = title != parsed_plan["title"]
                    st.markdown(f"<span style='color:{'green' if title_changed else 'black'}'>{parsed_plan['title']}</span>", unsafe_allow_html=True)
                
                db_loads = teacher_doc.get("loads", [])
                new_loads = parsed_plan["loads"]
                
                added = []
                updated = []
                
                for nl in new_loads:
                    found = None
                    for dl in db_loads:
                        if (str(dl.get("subject", "")).strip().lower() == str(nl["subject"]).strip().lower() and 
                            str(dl.get("group", "")).strip().lower() == str(nl["group"]).strip().lower() and 
                            str(dl.get("semester", "")).strip().lower() == str(nl["semester"]).strip().lower()):
                            found = dl
                            break
                    if found:
                        if found.get("total") != nl["total"]:
                            updated.append((found, nl))
                    else:
                        added.append(nl)
                        
                st.markdown("#### Изменения в нагрузке:")
                if added:
                    st.markdown(f"➕ **Новые предметы ({len(added)}):**")
                    for a in added:
                        st.write(f"- {a['subject']} (Гр: {a['group']}, Итого: {a['total']} ч.)")
                if updated:
                    st.markdown(f"🔄 **Изменились часы ({len(updated)}):**")
                    for old_l, new_l in updated:
                        st.write(f"- {new_l['subject']} (Гр: {new_l['group']}): было {old_l['total']} ч. ➔ стало {new_l['total']} ч.")
                if not added and not updated:
                    st.info("Нагрузка из файла полностью совпадает с активной базой данных.")
                    
                st.write("---")
                col_btn1, col_btn2 = st.columns(2)
                with col_btn1:
                    if st.button("✅ Принять и обновить профиль в БД", use_container_width=True):
                        # Сохраняем псевдоним (alias) в БД, если ФИО было заменено/сопоставлено
                        if resolved_plan_fio != plan_teacher_fio:
                            similar_emp = find_similar_teacher(plan_teacher_fio, db)
                            if similar_emp:
                                aliases = similar_emp.get("aliases", [])
                                if plan_teacher_fio not in aliases:
                                    aliases.append(plan_teacher_fio)
                                    db.employees.update_one({"_id": similar_emp["_id"]}, {"$set": {"aliases": aliases}})
                        active_loads = list(db_loads)
                        for nl in new_loads:
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
                                
                        # Обновляем профиль во второй БД (workload_db)
                        if workload_db is not None:
                            similar_wl = find_similar_teacher(parsed_plan["fio"], workload_db)
                            profile_fields = {
                                "fio": parsed_plan["fio"],
                                "position": parsed_plan["position"],
                                "employment_conditions": parsed_plan["employment_conditions"],
                                "degree": parsed_plan["degree"],
                                "title": parsed_plan["title"],
                                "contract": parsed_plan.get("contract", {}),
                                "education": parsed_plan.get("education", {}),
                                "subjects": list(set([l["subject"] for l in active_loads if l.get("subject")]))
                            }
                            if similar_wl:
                                workload_db.employees.update_one({"_id": similar_wl["_id"]}, {"$set": profile_fields})
                            else:
                                existing_wl = workload_db.employees.find_one({"fio": parsed_plan["fio"]})
                                if existing_wl:
                                    workload_db.employees.update_one({"_id": existing_wl["_id"]}, {"$set": profile_fields})
                                else:
                                    workload_db.employees.insert_one(profile_fields)

                        db.employees.update_one(
                            {"_id": teacher_doc["_id"]},
                            {"$set": {
                                "fio": parsed_plan["fio"],
                                "position": parsed_plan["position"],
                                "employment_conditions": parsed_plan["employment_conditions"],
                                "degree": parsed_plan["degree"],
                                "title": parsed_plan["title"],
                                "contract": parsed_plan.get("contract", {}),
                                "education": parsed_plan.get("education", {}),
                                "department_name": parsed_plan.get("department_name", ""),
                                "institute_name": parsed_plan.get("institute_name", ""),
                                "loads": active_loads
                            }}
                        )
                        
                        db.individual_plans.insert_one({
                            "employee_id": str(teacher_doc["_id"]),
                            "year": parsed_year,
                            "position": parsed_plan["position"],
                            "employment_conditions": parsed_plan["employment_conditions"],
                            "degree": parsed_plan["degree"],
                            "title": parsed_plan["title"],
                            "contract": parsed_plan.get("contract", {}),
                            "education": parsed_plan.get("education", {}),
                            "department_name": parsed_plan.get("department_name", ""),
                            "institute_name": parsed_plan.get("institute_name", ""),
                            "loads": parsed_plan["loads"],
                            "timestamp": datetime.datetime.now().isoformat()
                        })
                        
                        st.success("Профиль преподавателя успешно обновлен в БД! История сохранена.")
                        del st.session_state["parsed_individual_plan"]
                        del st.session_state["parsed_individual_year"]
                        if "uploaded_plan_name" in st.session_state:
                            del st.session_state["uploaded_plan_name"]
                        st.rerun()
                with col_btn2:
                    if st.button("❌ Отклонить изменения", use_container_width=True):
                        del st.session_state["parsed_individual_plan"]
                        del st.session_state["parsed_individual_year"]
                        if "uploaded_plan_name" in st.session_state:
                            del st.session_state["uploaded_plan_name"]
                        st.rerun()
                        
        with col3_2:
            st.markdown("### 📄 Заполнение индивидуального плана по шаблону")
            
            data_sources = ["Текущий активный профиль из БД"]
            if history:
                for plan in history:
                    data_sources.append(f"План из истории за {plan.get('year')} год")
                    
            selected_source = st.selectbox("Использовать данные нагрузки из:", data_sources)
            
            if selected_source == "Текущий активный профиль из БД":
                source_loads = teacher_doc.get("loads", [])
                source_pos = position
                source_cond = conditions
                source_degree = degree
                source_title = title
            else:
                target_year = selected_source.replace("План из истории за ", "").replace(" год", "")
                target_plan = next(p for p in history if p.get("year") == target_year)
                source_loads = target_plan.get("loads", [])
                source_pos = target_plan.get("position", "")
                source_cond = target_plan.get("employment_conditions", "")
                source_degree = target_plan.get("degree", "")
                source_title = target_plan.get("title", "")
                
            default_template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_uploaded_p3_Копия individ_plrprGOD Шаблон для заполнения с переменными.xlsx")
            if not os.path.exists(default_template_path):
                default_template_path = "C:/Users/vrlab/Desktop/Копия individ_plrprGOD Шаблон для заполнения с переменными.xlsx"
            if not os.path.exists(default_template_path) and os.path.exists("individ_plan_Архипов.xlsx"):
                default_template_path = "individ_plan_Архипов.xlsx"
            
            uploaded_template = st.file_uploader(
                "📂 Загрузите файл-шаблон индивидуального плана (.xlsx):",
                type=["xlsx"],
                key="tab3_template_uploader_new"
            )
            
            if uploaded_template is not None:
                template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"_uploaded_template_{uploaded_template.name}")
                with open(template_path, "wb") as _f:
                    _f.write(uploaded_template.getbuffer())
                st.success(f"✅ Шаблон `{uploaded_template.name}` готов.")
            elif os.path.exists(default_template_path):
                template_path = default_template_path
                st.info(f"Используется шаблон: `{os.path.basename(template_path)}`")
            else:
                template_path = None
                st.warning("⚠️ Загрузите файл-шаблон (.xlsx) через поле выше.")
                
            search_name_in_template = st.text_input(
                "Искомая фамилия/ФИО в шаблоне для замены:",
                value=st.session_state.get("tab3_search_name_input_new", selected_fio),
                key="tab3_search_name_input_new"
            )
            
            final_fio_to_insert = st.text_input("ФИО для вставки в документ:", value=selected_fio)
            
            st.write("---")
            if st.button("🚀 Сгенерировать индивидуальный план", use_container_width=True, key="tab3_gen_btn"):
                if not template_path or not os.path.exists(template_path):
                    st.error("Пожалуйста, укажите корректный путь к шаблону.")
                elif not source_loads:
                    st.error("Выбранный источник не содержит учебной нагрузки.")
                else:
                    with st.spinner("Заполнение шаблона..."):
                        try:
                            import importlib
                            import excel_filler
                            importlib.reload(excel_filler)
                            from excel_filler import fill_teacher_plan
                            
                            output_path = f"individ_plan_{selected_fio}.xlsx"
                            
                            fill_teacher_plan(
                                template_path=template_path,
                                output_path=output_path,
                                new_fio=final_fio_to_insert,
                                loads=source_loads,
                                search_name=search_name_in_template,
                                position=source_pos,
                                employment_conditions=source_cond,
                                degree=source_degree,
                                title=source_title
                            )
                            
                            st.success(f"🎉 Документ индивидуального плана успешно сгенерирован: `{output_path}`!")
                            
                            with open(output_path, "rb") as f:
                                file_bytes = f.read()
                                
                            st.download_button(
                                label="📥 Скачать заполненный индивидуальный план",
                                data=file_bytes,
                                file_name=output_path,
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                use_container_width=True
                            )
                        except Exception as e_fill:
                            st.error("Ошибка при генерации документа:")
                            st.exception(e_fill)
    else:
        st.info("В базе данных нет преподавателей. Сначала импортируйте файл нагрузки на первой вкладке.")


# --- СТРАНИЦА 5: ДОКУМЕНТЫ ПО ШАБЛОНАМ WORD (.DOCX) ---
elif page == "📝 Документы по шаблонам Word (.docx)":
    st.subheader("📝 Генерация документов по шаблонам Word (.docx)")
    st.write("Создавайте и заполняйте приказы о направлении на практику и другие представления, объединяя данные о практиках студентов из БД САПР и анкетные данные преподавателей из БД Нагрузка.")
    
    practices = list(db.practices.find())
    
    # Выбор практики
    selected_practice = None
    if practices:
        practice_options = [f"📋 {p.get('employee_fio', '—')} — {p.get('type', '—')} (Гр. {p.get('group', '—')})" for p in practices]
        sel_idx = st.selectbox("Выберите запись о практике из БД САПР:", range(len(practice_options)), format_func=lambda x: practice_options[x])
        selected_practice = practices[sel_idx]
    else:
        st.warning("⚠️ В базе данных САПР (university_db) отсутствуют записи о практиках. Будет использован ручной ввод.")
        
    st.write("---")
    
    # 1. Поиск студента и группы
    student_options = ["-- Ввести вручную --"]
    selected_group_name = selected_practice.get("group", "") if selected_practice else ""
    if selected_group_name:
        students_in_group = list(db.students.find({"group": selected_group_name}))
        student_options.extend([s["fio"] for s in students_in_group if s.get("fio")])
        
    student_sel = st.selectbox("Выберите студента из группы (или введите вручную):", student_options)
    
    default_student_fio = ""
    if student_sel != "-- Ввести вручную --":
        default_student_fio = student_sel
        
    # 2. Поиск преподавателя во второй БД
    matched_teacher = None
    if selected_practice and workload_db is not None:
        matched_teacher = find_similar_teacher(selected_practice.get("employee_fio", ""), workload_db)
        
    # Список всех преподавателей из второй БД для ручного сопоставления
    all_workload_teachers = []
    if workload_db is not None:
        all_workload_teachers = list(workload_db.employees.find())
        
    st.markdown("### 📝 Редактирование полей шаблона")
    st.caption("Отредактируйте данные ниже. Шаблон Word (.docx) будет заполнен этими значениями.")
    
    p_col1, p_col2 = st.columns(2)
    
    with p_col1:
        st.markdown("**🎓 Данные студента и практики**")
        s_fio = st.text_input("ФИО Студента (student_fio):", value=default_student_fio, key="p5_student_fio")
        s_group = st.text_input("Группа (group_name):", value=selected_group_name, key="p5_group_name")
        
        default_direction = ""
        if selected_practice:
            # Пытаемся найти направление из групп
            grp_doc = db.groups.find_one({"name": selected_group_name})
            if grp_doc:
                default_direction = grp_doc.get("direction", "")
        s_dir = st.text_input("Направление подготовки (direction):", value=default_direction, key="p5_direction")
        
        p_type = st.text_input("Тип практики (practice_type):", value=selected_practice.get("type", "") if selected_practice else "", key="p5_practice_type")
        p_org = st.text_input("Организация практики (organization_name):", value=selected_practice.get("org", "") if selected_practice else "", key="p5_organization_name")
        p_order_date = st.text_input("Дата приказа (order_date):", value=selected_practice.get("order_date", "") if selected_practice else "", key="p5_order_date")
        p_order_signer = st.text_input("Подписавший приказ (order_signed_by):", value=selected_practice.get("order_signer", "") if selected_practice else "", key="p5_order_signed_by")
        
        default_dept = ""
        if matched_teacher:
            default_dept = matched_teacher.get("department_name", "")
        s_dept = st.text_input("Кафедра (department_name):", value=default_dept, key="p5_department_name")
        
    with p_col2:
        st.markdown("**👤 Данные преподавателя (руководителя)**")
        
        # Если преподаватель не совпал автоматически, дадим выбрать вручную
        if all_workload_teachers:
            manual_teacher_options = ["-- Выбрать из списка --"] + [t["fio"] for t in all_workload_teachers]
            # Ищем индекс текущего совпавшего, если он есть
            manual_teacher_index = 0
            if matched_teacher and matched_teacher.get("fio") in manual_teacher_options:
                manual_teacher_index = manual_teacher_options.index(matched_teacher["fio"])
                
            manual_teacher_sel = st.selectbox(
                "Сопоставить с преподавателем из БД Нагрузка:",
                manual_teacher_options,
                index=manual_teacher_index
            )
            if manual_teacher_sel != "-- Выбрать из списка --":
                matched_teacher = next(t for t in all_workload_teachers if t["fio"] == manual_teacher_sel)
                
        t_fio = st.text_input("ФИО Руководителя (teacher_fio):", value=matched_teacher.get("fio", selected_practice.get("employee_fio", "") if selected_practice else "") if matched_teacher else (selected_practice.get("employee_fio", "") if selected_practice else ""), key="p5_teacher_fio")
        t_pos = st.text_input("Должность (teacher_position):", value=matched_teacher.get("position", "") if matched_teacher else "", key="p5_teacher_position")
        t_degree = st.text_input("Ученая степень (teacher_degree):", value=matched_teacher.get("degree", "") if matched_teacher else "", key="p5_teacher_degree")
        t_title = st.text_input("Ученое звание (teacher_rank):", value=matched_teacher.get("title", "") if matched_teacher else "", key="p5_teacher_rank")
        
        t_contract_num = ""
        t_contract_date = ""
        if matched_teacher and matched_teacher.get("contract"):
            t_contract_num = matched_teacher["contract"].get("number", "")
            t_contract_date = matched_teacher["contract"].get("date", "")
        t_c_num = st.text_input("Номер контракта (contract_number):", value=t_contract_num, key="p5_contract_number")
        t_c_date = st.text_input("Дата контракта (contract_date):", value=t_contract_date, key="p5_contract_date")
        
        # Поля для таблицы
        st.markdown("**📖 Дополнительные поля по дисциплине**")
        sub_name = st.text_input("Дисциплина (subject_name):", value=selected_practice.get("type", "Производственная практика") if selected_practice else "Производственная практика", key="p5_subject_name")
        sub_lect = st.text_input("Часы лекций (subject_lectures):", value="0", key="p5_subject_lectures")
        sub_labs = st.text_input("Часы лабораторных (subject_labs):", value="0", key="p5_subject_labs")

    st.write("---")
    
    # Загрузка или выбор шаблона
    uploaded_docx_template = st.file_uploader(
        "📂 Загрузите свой файл-шаблон Word (.docx):",
        type=["docx"],
        key="p5_template_uploader"
    )
    
    default_docx_template_path = "default_template.docx"
    
    if uploaded_docx_template is not None:
        template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"_uploaded_docx_{uploaded_docx_template.name}")
        with open(template_path, "wb") as _f:
            _f.write(uploaded_docx_template.getbuffer())
        st.success(f"✅ Готов пользовательский шаблон: `{uploaded_docx_template.name}`")
    else:
        # Автоматически создаем демонстрационный шаблон, если его нет
        if not os.path.exists(default_docx_template_path):
            create_default_template(default_docx_template_path)
        template_path = default_docx_template_path
        st.info("Используется стандартный шаблон DOCX.")
        
    if st.button("📝 Сгенерировать представление в приказ (.docx)", use_container_width=True):
        if not s_fio:
            st.error("Пожалуйста, заполните ФИО студента.")
        else:
            with st.spinner("Заполнение шаблона Word..."):
                try:
                    output_docx_path = f"Представление_{s_fio.replace(' ', '_')}.docx"
                    
                    data_to_fill = {
                        "student_fio": s_fio,
                        "group_name": s_group,
                        "direction": s_dir,
                        "department_name": s_dept,
                        "practice_type": p_type,
                        "organization_name": p_org,
                        "order_date": p_order_date,
                        "order_signed_by": p_order_signer,
                        "teacher_fio": t_fio,
                        "teacher_position": t_pos,
                        "teacher_degree": t_degree,
                        "teacher_rank": t_title,
                        "contract_number": t_c_num,
                        "contract_date": t_c_date,
                        "subject_name": sub_name,
                        "subject_lectures": sub_lect,
                        "subject_labs": sub_labs
                    }
                    
                    fill_template(template_path, output_docx_path, data_to_fill)
                    st.success(f"🎉 Документ Word успешно сгенерирован: `{output_docx_path}`!")
                    
                    with open(output_docx_path, "rb") as f:
                        docx_bytes = f.read()
                        
                    st.download_button(
                        label="📥 Скачать готовое представление (.docx)",
                        data=docx_bytes,
                        file_name=output_docx_path,
                        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                        use_container_width=True
                    )
                except Exception as e_docx:
                    st.error(f"Не удалось заполнить шаблон Word: {e_docx}")


# --- СТРАНИЦА 6: ОБУЧЕНИЕ ИИ ---
elif page == "🎓 Обучение ИИ":
    st.subheader("🎓 Интерактивное обучение локальной ИИ-модели")
    st.write("Этот раздел позволяет обучить ИИ на реальных примерах файлов. Вы указываете правильное сопоставление колонок и извлекаемые данные для различных типов Excel-таблиц. Эти примеры сохраняются в файл `ai_training_examples.json` и автоматически подмешиваются в контекст ИИ при последующих разборах файлов (few-shot prompting).")

    # Выбор источника файлов
    train_source = st.radio("Источник файлов для обучения:", ["Сканировать папку", "Загрузить файл вручную"], horizontal=True)

    uploaded_train_file = None
    selected_train_filepath = None

    if train_source == "Сканировать папку":
        default_dir = r"c:\Users\vrlab\Desktop\Пример документов 2 типа"
        scan_path = st.text_input("Путь к папке для сканирования:", value=default_dir)
        
        excel_files = []
        if os.path.exists(scan_path) and os.path.isdir(scan_path):
            for f in os.listdir(scan_path):
                if f.endswith((".xlsx", ".xls")) and not f.startswith("~$"):
                    excel_files.append(f)
                    
        if excel_files:
            selected_filename = st.selectbox("Выберите файл для обучения:", excel_files)
            selected_train_filepath = os.path.join(scan_path, selected_filename)
            st.info(f"📁 Выбран файл: `{selected_train_filepath}`")
        else:
            st.warning("⚠️ В указанной папке не найдено Excel файлов или папка не существует.")
    else:
        uploaded_train_file = st.file_uploader("Загрузите Excel-файл (.xlsx, .xls) для обучения:", type=["xlsx", "xls"])

    # Читаем файл-поток
    file_stream = None
    filename_to_display = ""
    if selected_train_filepath:
        try:
            file_stream = open(selected_train_filepath, "rb")
            filename_to_display = os.path.basename(selected_train_filepath)
        except Exception as e_file:
            st.error(f"Не удалось открыть файл: {e_file}")
    elif uploaded_train_file:
        file_stream = uploaded_train_file
        filename_to_display = uploaded_train_file.name

    if file_stream is not None:
        try:
            # Чтение имен листов
            if hasattr(file_stream, "seek"):
                file_stream.seek(0)
            xls_train = pd.ExcelFile(file_stream)
            sheet_names = xls_train.sheet_names
            xls_train.close()
            
            # --- ИНИЦИАЛИЗАЦИЯ И СБРОС СЕССИИ ДЛЯ ТЕКУЩЕГО ФАЙЛА ---
            if "train_sheet_mappings" not in st.session_state:
                st.session_state["train_sheet_mappings"] = {}
                
            current_train_file_key = f"train_file_{filename_to_display}"
            if st.session_state.get("last_train_file_key") != current_train_file_key:
                st.session_state["last_train_file_key"] = current_train_file_key
                st.session_state["train_sheet_mappings"] = {}
                if "train_ai_raw_mapping" in st.session_state:
                    del st.session_state["train_ai_raw_mapping"]
            
            st.write("---")
            st.markdown("### 🔮 Шаг 1: Первоначальный разбор ИИ")
            st.write("Запустите разбор, чтобы увидеть, как ИИ справляется с этим файлом по умолчанию:")
            
            if st.button("🔮 Запустить анализ структуры ИИ", use_container_width=True):
                with st.spinner("ИИ анализирует все листы файла..."):
                    try:
                        if not connected:
                            st.error("Локальный ИИ (Ollama) недоступен. Подключение необходимо для обучения.")
                        else:
                            analyzed_sheets = []
                            for sheet_name in sheet_names:
                                if hasattr(file_stream, "seek"):
                                    file_stream.seek(0)
                                
                                is_sapr = "сапр" in filename_to_display.lower()
                                is_profile_sheet = any(kw in sheet_name.lower() for kw in ["титульн", "тит", "общие", "сведения", "profile", "анкета"])
                                
                                if is_profile_sheet and not is_sapr:
                                    profile_res = ai_parser_helper.analyze_teacher_profile_with_ai(
                                        file_stream, selected_model, ollama_url, target_sheet=sheet_name
                                    )
                                    mapping = {
                                        "sheet_name": sheet_name,
                                        "document_type": "individual_plan",
                                        "header_row_index": 0,
                                        "transpose": False,
                                        "column_mapping": {},
                                        "metadata": profile_res
                                    }
                                else:
                                    mapping = ai_parser_helper.analyze_excel_structure(
                                        file_stream, selected_model, ollama_url, target_sheet=sheet_name
                                    )
                                
                                st.session_state["train_sheet_mappings"][sheet_name] = mapping
                                analyzed_sheets.append(sheet_name)
                                
                            if sheet_names:
                                st.session_state["train_ai_raw_mapping"] = st.session_state["train_sheet_mappings"][sheet_names[0]]
                                
                            st.success(f"Разбор ИИ успешно завершен для всех листов: {', '.join(analyzed_sheets)}!")
                    except Exception as e_train_ai:
                        st.error(f"Ошибка при ИИ-анализе: {e_train_ai}")
                        
            # Если есть результаты ИИ или пользователь хочет настроить вручную
            st.write("---")
            st.markdown("### 🛠️ Шаг 2: Проверка и корректировка данных")
            st.write("Сравните распознанные данные с реальными и скорректируйте их. Это поможет обучить модель.")
            
            # Читаем превью колонок
            if hasattr(file_stream, "seek"):
                file_stream.seek(0)
            
            active_train_sheet = st.selectbox("Лист для анализа:", sheet_names, index=0)
            
            # Получаем или инициализируем raw_mapping для выбранного листа
            raw_mapping = st.session_state["train_sheet_mappings"].get(active_train_sheet, {})
            # Если для этого листа нет записи, пробуем найти наилучший сохраненный шаблон
            if not raw_mapping:
                best_ex = ai_parser_helper.find_best_training_example(filename_to_display, sheet_name=active_train_sheet)
                if best_ex:
                    raw_mapping = {
                        "sheet_name": best_ex.get("sheet_name"),
                        "document_type": best_ex.get("document_type"),
                        "header_row_index": best_ex.get("header_row_index"),
                        "transpose": best_ex.get("transpose", False),
                        "column_mapping": best_ex.get("correct_mapping"),
                        "metadata": best_ex.get("correct_metadata")
                    }
                    st.session_state["train_sheet_mappings"][active_train_sheet] = raw_mapping
                    st.info(f"🎯 Лист сопоставлен по шаблону: `{best_ex.get('filename')}`.")
            
            # Если все еще пусто, но ИИ-анализ вернул mapping именно для этого листа, подтягиваем его
            if not raw_mapping and st.session_state.get("train_ai_raw_mapping", {}).get("sheet_name") == active_train_sheet:
                raw_mapping = st.session_state["train_ai_raw_mapping"]
                st.session_state["train_sheet_mappings"][active_train_sheet] = raw_mapping
            
            # Документ-тип
            doc_types = ["department_load", "individual_plan", "practice_order", "student_list", "other"]
            default_doc_type = raw_mapping.get("document_type", "department_load" if "сапр" in filename_to_display.lower() else "individual_plan")
            if default_doc_type not in doc_types:
                default_doc_type = "department_load"
                
            selected_doc_type = st.selectbox(
                "Тип документа:", 
                doc_types, 
                index=doc_types.index(default_doc_type),
                key=f"train_doc_type_{active_train_sheet}"
            )
            
            # Строка заголовка
            default_header_row = int(raw_mapping.get("header_row_index", 2 if selected_doc_type == "department_load" else 0))
            selected_header_row = st.number_input(
                "Строка заголовков (0-based):", 
                min_value=0, 
                value=default_header_row,
                key=f"train_header_row_{active_train_sheet}"
            )
            
            # Выбор ориентации таблицы (по строкам или по столбцам)
            orientation_opts = ["По столбцам (вертикальная)", "По строкам (горизонтальная)"]
            default_orient = "По строкам (горизонтальная)" if raw_mapping.get("transpose") else "По столбцам (вертикальная)"
            table_orientation = st.radio(
                "Ориентация табличной части:",
                options=orientation_opts,
                index=orientation_opts.index(default_orient),
                key=f"train_orientation_{active_train_sheet}"
            )
            is_transposed = (table_orientation == "По строкам (горизонтальная)")

            # Читаем превью колонок, передавая строку заголовка для правильного поиска по строкам
            preview_text, col_names, active_train_sheet = ai_parser_helper.prepare_excel_preview(
                file_stream, target_sheet=active_train_sheet, header_row_idx=selected_header_row
            )
            
            if hasattr(file_stream, "seek"):
                file_stream.seek(0)
            df_train_sheet = pd.read_excel(file_stream, sheet_name=active_train_sheet, header=None)
            
            # Если горизонтальная ориентация, то опциями для сопоставления должны быть строки
            if is_transposed:
                col_options_labels = []
                for r in range(df_train_sheet.shape[0]):
                    row_vals = [str(df_train_sheet.iloc[r, c]).strip() for c in range(df_train_sheet.shape[1]) if pd.notna(df_train_sheet.iloc[r, c])]
                    row_desc = " | ".join(row_vals[:4])
                    if len(row_desc) > 40:
                        row_desc = row_desc[:37] + "..."
                    col_options_labels.append(f"Строка {r} - \"{row_desc}\"" if row_desc else f"Строка {r}")
            else:
                col_options_labels = col_names
            
            st.markdown("#### 🔗 Сопоставление колонок/строк (для табличной части)")
            
            # Базовые поля, которые есть почти везде
            base_fields = {
                "subject_name": "Дисциплина (subject_name)",
                "group_name": "Группа (group_name)",
                "semester_number": "Семестр (semester_number)"
            }
            
            # Поля часов (нагрузка)
            hours_fields = {
                "hours_lectures": "Часы лекций (hours_lectures)",
                "hours_practicals": "Часы практик (hours_practicals)",
                "hours_laboratories": "Часы лабораторий (hours_laboratories)",
                "hours_consultations": "Часы консультаций (hours_consultations)",
                "hours_exams": "Часы экзаменов (hours_exams)",
                "hours_zachets": "Часы зачетов (hours_zachets)",
                "hours_coursework": "Часы курсовых работ/проектов (hours_coursework)",
                "hours_practice": "Часы руководства практикой (hours_practice)",
                "hours_vkr": "Часы руководства ВКР (hours_vkr)",
                "hours_gek": "Часы участия в ГЭК (hours_gek)",
                "hours_additional": "Дополнительные часы (hours_additional)",
                "hours_total": "Итого учебных часов (hours_total)"
            }
            
            # Специфичные поля
            practice_fields = {
                "practice_type": "Тип практики (practice_type)",
                "practice_org": "Организация проведения практики (practice_org)",
                "practice_kind": "Вид практики (practice_kind)"
            }
            student_fields = {
                "student_fio": "ФИО Студента (student_fio)",
                "student_profile": "Профиль обучения студента (student_profile)"
            }
            teacher_field = {
                "teacher_fio": "ФИО Преподавателя (teacher_fio)"
            }
            
            schema_fields = {}
            if selected_doc_type == "individual_plan":
                schema_fields = {
                    "department_name": "Кафедры",
                    "institute_name": "наименование института/факультета",
                    "employee_fio": "фамилия, имя, отчество преподавателя",
                    "employee_position_name": "наименование должности",
                    "academic_year": "учебный год",
                    "department_head": "Заведующий кафедрой",
                    "employee_position": "Должность",
                    "employee_rate": "Размер ставки",
                    "employment_conditions": "Условия привлечения*",
                    "employee_degree": "Ученая степень",
                    "employee_title": "Ученое звание",
                    "employee_contract": "Сведения о контракте: дата, номер",
                    "employee_contract_duration": "срок действия",
                    "employee_edu_institution": "наименование образовательной организации, год окончания",
                    "employee_edu_specialty": "наименование образовательной программы",
                    "employee_edu_qualification": "квалификация",
                    "taught_subjects": "Перечень преподаваемых дисциплин."
                }
            elif selected_doc_type == "department_load":
                schema_fields.update(base_fields)
                schema_fields.update(teacher_field)
                schema_fields.update(hours_fields)
            elif selected_doc_type == "practice_order":
                schema_fields.update(student_fields)
                schema_fields.update(practice_fields)
            elif selected_doc_type == "student_list":
                schema_fields.update(student_fields)
                schema_fields.update(base_fields)
            else:
                # Fallback to all fields
                schema_fields.update(base_fields)
                schema_fields.update(teacher_field)
                schema_fields.update(hours_fields)
                schema_fields.update(student_fields)
                schema_fields.update(practice_fields)
            
            dropdown_options = ["-- Отсутствует --"] + col_options_labels
            user_column_mapping = {}
            
            map_col1, map_col2 = st.columns(2)
            for idx, (field_key, field_name) in enumerate(schema_fields.items()):
                target_col = map_col1 if idx % 2 == 0 else map_col2
                with target_col:
                    # Обработка сохраненных значений, которые могут быть списком (из-за multiselect)
                    current_idx = raw_mapping.get("column_mapping", {}).get(field_key)
                    default_indices = []
                    
                    if current_idx is not None:
                        if isinstance(current_idx, list):
                            for idx_val in current_idx:
                                try:
                                    if int(idx_val) < len(col_options_labels):
                                        default_indices.append(int(idx_val))
                                except:
                                    pass
                        else:
                            try:
                                if int(current_idx) < len(col_options_labels):
                                    default_indices.append(int(current_idx))
                            except:
                                pass
                                
                    default_options = [col_options_labels[i] for i in default_indices]
                        
                    sel_opts = st.multiselect(
                        f"📍 {field_name}:",
                        options=col_options_labels,
                        default=default_options,
                        key=f"train_col_{active_train_sheet}_{field_key}"
                    )
                    
                    if not sel_opts:
                        user_column_mapping[field_key] = None
                    else:
                        user_column_mapping[field_key] = [col_options_labels.index(opt) for opt in sel_opts]
                        
            # Метаданные
            st.markdown("#### 📝 Метаданные из ячеек (профиль, контракт, образование)")
            
            def parse_excel_coordinate(coord_str):
                if not coord_str:
                    return None
                import re
                match = re.match(r'^([A-Z]+)(\d+)$', coord_str.strip().upper())
                if not match:
                    return None
                col_str, row_str = match.groups()
                col_idx = 0
                for char in col_str:
                    col_idx = col_idx * 26 + (ord(char) - ord('A') + 1)
                col_idx -= 1
                row_idx = int(row_str) - 1
                return row_idx, col_idx

            meta_fields = {
                "employee_fio": "ФИО Преподавателя (employee_fio)",
                "employee_position": "Должность (employee_position)",
                "employee_rate": "Ставка / Условия (employee_rate)",
                "employee_degree": "Степень (employee_degree)",
                "employee_title": "Звание (employee_title)",
                "employee_contract_num": "Номер контракта (employee_contract_num)",
                "employee_contract_date": "Дата контракта (employee_contract_date)",
                "employee_contract_duration": "Срок контракта (employee_contract_duration)",
                "employee_edu_institution": "Образовательная организация (employee_edu_institution)",
                "employee_edu_year": "Год окончания (employee_edu_year)",
                "employee_edu_specialty": "Специальность (employee_edu_specialty)",
                "employee_edu_qualification": "Квалификация (employee_edu_qualification)",
                "department_name": "Кафедра (department_name)",
                "department_head": "Зав. кафедрой (department_head)",
                "department_direction": "Направление кафедры (department_direction)",
                "institute_name": "Название института (institute_name)",
                "academic_year": "Учебный год (academic_year)",
                "taught_subjects": "Перечень преподаваемых дисциплин (taught_subjects)"
            }

            saved_meta = raw_mapping.get("metadata", {}) or {}
            if not saved_meta and raw_mapping.get("teacher_profile"):
                saved_meta = {
                    "employee_fio": raw_mapping["teacher_profile"].get("fio"),
                    "employee_position": raw_mapping["teacher_profile"].get("position"),
                    "employee_rate": raw_mapping["teacher_profile"].get("employment_conditions"),
                    "employee_degree": raw_mapping["teacher_profile"].get("degree"),
                    "employee_title": raw_mapping["teacher_profile"].get("title"),
                }

            user_metadata = {}
            meta_fields_list = list(meta_fields.items())
            meta_col1, meta_col2 = st.columns(2)

            for idx, (field_key, field_name) in enumerate(meta_fields_list):
                target_col = meta_col1 if idx < 9 else meta_col2
                with target_col:
                    field_data = saved_meta.get(field_key) or {}
                    if isinstance(field_data, dict):
                        saved_val = ""  # Никогда не предзаполняем старым значением
                        saved_coord = field_data.get("coordinate") or ""
                        import re
                        # Защита от старых багов, где координата содержала ФИО
                        if saved_coord and not re.match(r'^[A-Z]+\d+$', str(saved_coord).strip().upper()):
                            saved_coord = ""
                    else:
                        saved_val = ""
                        saved_coord = ""

                    st.caption(f"**{field_name}**")
                    sc1, sc2 = st.columns([1, 2])
                    with sc1:
                        c_val = st.text_input(
                            "Ячейка (напр., B5):",
                            value=saved_coord,
                            key=f"train_coord_{active_train_sheet}_{field_key}",
                            label_visibility="collapsed"
                        )
                    with sc2:
                        pulled_val = ""
                        if c_val:
                            parsed = parse_excel_coordinate(c_val)
                            if parsed:
                                r_idx, col_idx = parsed
                                if r_idx < df_train_sheet.shape[0] and col_idx < df_train_sheet.shape[1]:
                                    cell_val = df_train_sheet.iloc[r_idx, col_idx]
                                    if pd.notna(cell_val):
                                        pulled_val = str(cell_val).strip()

                        default_txt = pulled_val if c_val else saved_val
                        t_val = st.text_input(
                            "Значение:",
                            value=default_txt,
                            key=f"train_val_{active_train_sheet}_{field_key}",
                            label_visibility="collapsed"
                        )

                    user_metadata[field_key] = {
                        "value": t_val,
                        "coordinate": c_val
                    }

            # Сохраняем текущие значения в сессию
            st.session_state["train_sheet_mappings"][active_train_sheet] = {
                "sheet_name": active_train_sheet,
                "document_type": selected_doc_type,
                "header_row_index": selected_header_row,
                "transpose": is_transposed,
                "column_mapping": user_column_mapping,
                "metadata": user_metadata
            }
            
            st.write("---")
            st.markdown("### 💾 Шаг 3: Сохранение обучающего примера")
            
            if st.button("💾 Сохранить этот пример в базу обучения ИИ", use_container_width=True):
                # Сохраняем пример
                new_example = {
                    "filename": filename_to_display,
                    "sheet_name": active_train_sheet,
                    "document_type": selected_doc_type,
                    "header_row_index": selected_header_row,
                    "transpose": is_transposed,
                    "columns_present": col_names,
                    "correct_mapping": user_column_mapping,
                    "correct_metadata": user_metadata
                }
                
                training_data = ai_parser_helper.load_training_examples()
                
                # Удаляем дубликаты по имени файла и листа
                training_data = [ex for ex in training_data if not (ex.get("filename") == filename_to_display and ex.get("sheet_name") == active_train_sheet)]
                training_data.append(new_example)
                
                try:
                    with open("ai_training_examples.json", "w", encoding="utf-8") as f_train:
                        json.dump(training_data, f_train, ensure_ascii=False, indent=2)
                    st.success("🎉 Обучающий пример успешно сохранен в `ai_training_examples.json`!")
                    st.rerun()
                except Exception as e_save_train:
                    st.error(f"Не удалось сохранить пример: {e_save_train}")
                    
        except Exception as e_sheet_train:
            st.error(f"Ошибка разбора структуры файла: {e_sheet_train}")
        finally:
            if selected_train_filepath and file_stream:
                file_stream.close()
                
    st.write("---")
    st.markdown("### 🗄️ База сохраненных обучающих примеров")
    
    saved_examples = ai_parser_helper.load_training_examples()
    if not saved_examples:
        st.info("База обучения пуста. Добавьте первый пример выше.")
    else:
        st.write(f"Всего сохранено обучающих примеров: **{len(saved_examples)}**")
        for idx, ex in enumerate(saved_examples):
            with st.expander(f"⭐ {ex.get('filename')} — {ex.get('sheet_name')} ({ex.get('document_type')})"):
                st.write(f"**Строка заголовков:** {ex.get('header_row_index')}")
                st.write(f"**Ориентация:** {'По строкам (горизонтальная)' if ex.get('transpose') else 'По столбцам (вертикальная)'}")
                st.write("**Колонки (маппинг):**", ex.get("correct_mapping"))
                st.write("**Метаданные:**", ex.get("correct_metadata"))
                
                if st.button("🗑️ Удалить этот обучающий пример", key=f"del_ex_{idx}"):
                    # Удаляем
                    saved_examples.pop(idx)
                    try:
                        with open("ai_training_examples.json", "w", encoding="utf-8") as f_train:
                            json.dump(saved_examples, f_train, ensure_ascii=False, indent=2)
                        st.success("Пример удален!")
                        st.rerun()
                    except Exception as e_del:
                        st.error(f"Не удалось удалить: {e_del}")

# Временный отладочный блок для чтения .doc файла (будет удален после получения текста)
try:
    import win32com.client
    import os
    st.sidebar.markdown("---")
    if st.sidebar.button("🔍 Прочитать .doc файл"):
        doc_path = r"C:\Users\ytud\Downloads\ЩученковОриг (1).doc"
        if os.path.exists(doc_path):
            word = win32com.client.Dispatch("Word.Application")
            word.Visible = False
            try:
                doc = word.Documents.Open(doc_path)
                paras = []
                for p in doc.Paragraphs:
                    paras.append(p.Range.Text)
                doc.Close()
                text_content = "\n".join(paras)
                st.session_state["temp_doc_content"] = text_content
            except Exception as ex_read:
                st.session_state["temp_doc_content"] = f"Error reading document: {ex_read}"
            finally:
                word.Quit()
        else:
            st.session_state["temp_doc_content"] = f"File not found: {doc_path}"
except Exception as ex_import:
    st.session_state["temp_doc_content"] = f"Error importing win32com: {ex_import}"

if "temp_doc_content" in st.session_state:
    st.markdown("### 📄 Содержимое .doc файла")
    st.text_area("Текст документа:", value=st.session_state["temp_doc_content"], height=600)
