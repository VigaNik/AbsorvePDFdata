import mysql.connector
from tabulate import tabulate

# Database connection parameters
db_config = {
    "host": "localhost",
    "user": "nikita",
    "password": "nikita!##123",
    "database": "academic_journals",
    "charset": "utf8mb4",
    "collation": "utf8mb4_unicode_ci"
}


def connect_to_db():
    """Establish connection to the database"""
    try:
        conn = mysql.connector.connect(**db_config)
        return conn
    except mysql.connector.Error as err:
        print(f"Connection error: {err}")
        return None


def check_all_files(conn):
    """Показать все файлы в базе данных"""
    cursor = conn.cursor()

    print("📁 ВСЕ ФАЙЛЫ В БАЗЕ ДАННЫХ")
    print("=" * 80)

    query = """
    SELECT 
        j.journal_name AS 'Журнал',
        i.issue_number AS 'Выпуск',
        COUNT(DISTINCT mt.page_name) AS 'Страниц',
        COUNT(DISTINCT dr.reference_id) AS 'Ссылок',
        COUNT(DISTINCT ft.footnote_id) AS 'Сносок',
        i.created_at AS 'Дата импорта'
    FROM journals j
    JOIN issues i ON j.journal_id = i.journal_id
    LEFT JOIN main_texts mt ON i.issue_id = mt.issue_id
    LEFT JOIN document_references dr ON i.issue_id = dr.issue_id
    LEFT JOIN footnotes_table ft ON i.issue_id = ft.issue_id
    GROUP BY j.journal_name, i.issue_number, i.created_at
    ORDER BY j.journal_name, CAST(i.issue_number AS UNSIGNED)
    """

    cursor.execute(query)
    results = cursor.fetchall()

    if results:
        headers = ['Журнал', 'Выпуск', 'Страниц', 'Ссылок', 'Сносок', 'Дата импорта']
        print(tabulate(results, headers=headers, tablefmt='grid'))
        print(f"\n📊 Всего найдено: {len(results)} файлов")
    else:
        print("❌ Файлы не найдены в базе данных")

    cursor.close()


def check_journal_files(conn, journal_name):
    """Проверить файлы конкретного журнала"""
    cursor = conn.cursor()

    print(f"📖 ФАЙЛЫ ЖУРНАЛА: {journal_name}")
    print("=" * 60)

    query = """
    SELECT 
        i.issue_number AS 'Выпуск',
        COUNT(DISTINCT mt.page_name) AS 'Страниц',
        COUNT(DISTINCT dr.reference_id) AS 'Ссылок',
        COUNT(DISTINCT ft.footnote_id) AS 'Сносок',
        i.created_at AS 'Дата импорта'
    FROM issues i
    JOIN journals j ON i.journal_id = j.journal_id
    LEFT JOIN main_texts mt ON i.issue_id = mt.issue_id
    LEFT JOIN document_references dr ON i.issue_id = dr.issue_id
    LEFT JOIN footnotes_table ft ON i.issue_id = ft.issue_id
    WHERE j.journal_name = %s
    GROUP BY i.issue_number, i.created_at
    ORDER BY CAST(i.issue_number AS UNSIGNED)
    """

    cursor.execute(query, (journal_name,))
    results = cursor.fetchall()

    if results:
        headers = ['Выпуск', 'Страниц', 'Ссылок', 'Сносок', 'Дата импорта']
        print(tabulate(results, headers=headers, tablefmt='grid'))
        print(f"\n📊 Найдено выпусков: {len(results)}")
    else:
        print(f"❌ Файлы журнала '{journal_name}' не найдены")

    cursor.close()


def check_missing_files(conn, expected_files):
    """Проверить какие файлы отсутствуют в базе данных"""
    cursor = conn.cursor()

    print("🔍 ПРОВЕРКА ОТСУТСТВУЮЩИХ ФАЙЛОВ")
    print("=" * 50)

    # Получить все существующие файлы
    query = """
    SELECT CONCAT(LOWER(j.journal_name), i.issue_number) as file_pattern
    FROM journals j
    JOIN issues i ON j.journal_id = i.journal_id
    """

    cursor.execute(query)
    existing_files = {row[0] for row in cursor.fetchall()}

    missing_files = []
    for expected_file in expected_files:
        # Преобразуем имя файла в паттерн (убираем расширения и префиксы)
        file_pattern = expected_file.lower()
        file_pattern = file_pattern.replace('_footnotes.xml', '')
        file_pattern = file_pattern.replace('_footnotes.csv', '')
        file_pattern = file_pattern.replace('.xml', '')
        file_pattern = file_pattern.replace('.csv', '')

        if file_pattern not in existing_files:
            missing_files.append(expected_file)

    if missing_files:
        print("❌ Отсутствующие файлы:")
        for file in missing_files:
            print(f"   - {file}")
    else:
        print("✅ Все ожидаемые файлы найдены в базе данных")

    cursor.close()
    return missing_files


def get_database_stats(conn):
    """Получить общую статистику базы данных"""
    cursor = conn.cursor()

    print("📊 СТАТИСТИКА БАЗЫ ДАННЫХ")
    print("=" * 40)

    stats_queries = [
        ("Журналов", "SELECT COUNT(*) FROM journals"),
        ("Выпусков", "SELECT COUNT(*) FROM issues"),
        ("Страниц с текстом", "SELECT COUNT(*) FROM main_texts"),
        ("Ссылок", "SELECT COUNT(*) FROM document_references"),
        ("Сносок", "SELECT COUNT(*) FROM footnotes_table")
    ]

    stats = []
    for name, query in stats_queries:
        cursor.execute(query)
        count = cursor.fetchone()[0]
        stats.append([name, count])

    print(tabulate(stats, headers=['Параметр', 'Количество'], tablefmt='grid'))
    cursor.close()


def search_file(conn, search_term):
    """Поиск файла по части названия"""
    cursor = conn.cursor()

    print(f"🔎 ПОИСК: '{search_term}'")
    print("=" * 40)

    query = """
    SELECT 
        j.journal_name AS 'Журнал',
        i.issue_number AS 'Выпуск',
        COUNT(DISTINCT mt.page_name) AS 'Страниц',
        COUNT(DISTINCT dr.reference_id) AS 'Ссылок',
        COUNT(DISTINCT ft.footnote_id) AS 'Сносок'
    FROM journals j
    JOIN issues i ON j.journal_id = i.journal_id
    LEFT JOIN main_texts mt ON i.issue_id = mt.issue_id
    LEFT JOIN document_references dr ON i.issue_id = dr.issue_id
    LEFT JOIN footnotes_table ft ON i.issue_id = ft.issue_id
    WHERE LOWER(CONCAT(j.journal_name, i.issue_number)) LIKE %s
    GROUP BY j.journal_name, i.issue_number
    ORDER BY j.journal_name, CAST(i.issue_number AS UNSIGNED)
    """

    cursor.execute(query, (f'%{search_term.lower()}%',))
    results = cursor.fetchall()

    if results:
        headers = ['Журнал', 'Выпуск', 'Страниц', 'Ссылок', 'Сносок']
        print(tabulate(results, headers=headers, tablefmt='grid'))
        print(f"\n📊 Найдено: {len(results)} совпадений")
    else:
        print(f"❌ Файлы содержащие '{search_term}' не найдены")

    cursor.close()


def main():
    """Главная функция для проверки файлов"""
    conn = connect_to_db()
    if not conn:
        return

    while True:
        print("\n" + "=" * 60)
        print("🗃️  ПРОВЕРКА ФАЙЛОВ В БАЗЕ ДАННЫХ")
        print("=" * 60)
        print("1. Показать все файлы")
        print("2. Проверить конкретный журнал")
        print("3. Общая статистика")
        print("4. Поиск файла")
        print("5. Проверить отсутствующие файлы")
        print("0. Выход")
        print("=" * 60)

        choice = input("Выберите действие (0-5): ").strip()

        if choice == '1':
            check_all_files(conn)

        elif choice == '2':
            journal_name = input("Введите название журнала (Tarbiz, Leshonenu, и т.д.): ").strip()
            check_journal_files(conn, journal_name)

        elif choice == '3':
            get_database_stats(conn)

        elif choice == '4':
            search_term = input("Введите часть названия файла для поиска: ").strip()
            search_file(conn, search_term)

        elif choice == '5':
            print("Введите ожидаемые файлы (по одному на строку, пустая строка для завершения):")
            expected_files = []
            while True:
                file_name = input("Файл: ").strip()
                if not file_name:
                    break
                expected_files.append(file_name)

            if expected_files:
                check_missing_files(conn, expected_files)
            else:
                print("Список файлов пуст")

        elif choice == '0':
            break

        else:
            print("❌ Неверный выбор. Попробуйте еще раз.")

        input("\nНажмите Enter для продолжения...")

    conn.close()
    print("Соединение с базой данных закрыто")


if __name__ == "__main__":
    main()