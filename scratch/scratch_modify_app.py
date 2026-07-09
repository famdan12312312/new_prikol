filepath = "app.py"
with open(filepath, "r", encoding="utf-8") as f:
    content = f.read()

start_marker = 'elif page == "✏️ Редактирование нагрузки":'
end_marker = 'elif page == "📄 Индивидуальные планы и История":'

start_idx = content.find(start_marker)
if start_idx == -1:
    print("Error: start marker not found")
    exit(1)

end_idx = content.find(end_marker)
if end_idx == -1:
    print("Error: end marker not found")
    exit(1)

new_section_code = """elif page == "✏️ Редактирование нагрузки":
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
            
            df_sapr.insert(0, "Выбрать", False)
            
            edited_df = st.data_editor(
                df_sapr,
                column_config={
                    "Выбрать": st.column_config.CheckboxColumn(required=True),
                    "Index": None
                },
                disabled=[c for c in df_sapr.columns if c != "Выбрать"],
                use_container_width=True,
                hide_index=True,
                key="p3_sapr_editor"
            )
            
            chosen_rows = edited_df[edited_df["Выбрать"] == True]
            st.write(f"Выбрано дополнительных строк нагрузки из САПР: **{len(chosen_rows)}**")
            
            st.write("---")
            st.markdown("### Шаг 3: Шаблон индивидуального плана и генерация")
            
            uploaded_template = st.file_uploader(
                "📂 Загрузите файл-шаблон индивидуального плана (.xlsx):",
                type=["xlsx"],
                key="p3_template_uploader"
            )
            
            default_template_path = "c:/Users/vrlab/Desktop/Пример документов 2 типа/individ_plrpr2025 Обухов.xlsx"
            if not os.path.exists(default_template_path) and os.path.exists("individ_plan_Архипов.xlsx"):
                default_template_path = "individ_plan_Архипов.xlsx"
                
            if uploaded_template is not None:
                template_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), f"_uploaded_p3_{uploaded_template.name}")
                with open(template_path, "wb") as _f:
                    _f.write(uploaded_template.getbuffer())
                st.success(f"✅ Готов шаблон: `{uploaded_template.name}`")
            elif os.path.exists(default_template_path):
                template_path = default_template_path
                st.info(f"Используется стандартный шаблон: `{os.path.basename(template_path)}`")
            else:
                template_path = None
                st.warning("⚠️ Загрузите файл-шаблон (.xlsx) через поле выше.")
                
            search_name_in_template = st.text_input(
                "Искомая фамилия/ФИО в шаблоне для замены:",
                value=selected_fio,
                key="p3_search_name_input"
            )
            
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
                            
                            # Объединяем в памяти с текущими нагрузками преподавателя
                            combined_loads = list(current_loads)
                            for t_load in transferred_loads:
                                dup = False
                                for existing in combined_loads:
                                    if (str(existing.get("subject")).strip().lower() == str(t_load["subject"]).strip().lower() and
                                        str(existing.get("group")).strip().lower() == str(t_load["group"]).strip().lower() and
                                        str(existing.get("semester")).strip().lower() == str(t_load["semester"]).strip().lower()):
                                        dup = True
                                        break
                                if not dup:
                                    combined_loads.append(t_load)
                            
                            from excel_filler import fill_teacher_plan
                            output_path = f"individ_plan_{selected_fio.replace(' ', '_')}.xlsx"
                            
                            fill_teacher_plan(
                                template_path=template_path,
                                output_path=output_path,
                                new_fio=selected_fio,
                                loads=combined_loads,
                                search_name=search_name_in_template,
                                position=position,
                                employment_conditions=employment_conditions,
                                degree=degree,
                                title=title
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

"""

new_content = content[:start_idx] + new_section_code + "\n" + content[end_idx:]

with open(filepath, "w", encoding="utf-8") as f:
    f.write(new_content)

print("Modification of app.py successful")
