from local_db import MongoClient

def setup_database():
    """
    Инициализирует базу данных SQLite и очищает таблицы при запуске.
    """
    client = MongoClient()
    db_name = "university_db"
    db = client[db_name]
    
    # Очищаем коллекции при каждом запуске
    db.employees.drop()
    db.imported_files.drop()
    
    print(f"Подключение к локальной базе данных SQLite '{db_name}' выполнено успешно.")
    print("Таблицы инициализированы (очищены).")


def setup_workload_database():
    """
    Инициализирует Вторую базу данных SQLite (workload_db) для распределения нагрузки.
    """
    client = MongoClient(db_file="workload_db.sqlite")
    db_name = "workload_db"
    db = client[db_name]
    
    # Инициализируем коллекции (создаём таблицы если их нет)
    _ = db.employees
    _ = db.departments
    _ = db.institutes
    _ = db.groups
    _ = db.students
    _ = db.subjects
    _ = db.practices
    _ = db.unassigned_loads
    _ = db.individual_plans
    _ = db.imported_files
    
    print(f"Подключение к Второй базе данных SQLite '{db_name}' (workload_db.sqlite) выполнено успешно.")


if __name__ == "__main__":
    setup_database()
    setup_workload_database()
