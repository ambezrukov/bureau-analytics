"""
Скрипт импорта данных из XLS-выгрузок ProjectMate в SQLite-базу.

Запуск: python scripts/import_data.py
Требования: pywin32 (win32com), pandas, openpyxl

Отступил от ТЗ: вместо xlrd используется win32com.client (Excel COM),
т.к. файлы >2MB и xlrd считает их повреждёнными.

Версия 2.0: добавлены все поля нового формата выгрузки (февраль 2026+).
Обратная совместимость со старым форматом сохранена.
"""

import sqlite3
import os
import sys
import pathlib
from datetime import datetime

# Устанавливаем UTF-8 для консоли Windows
if sys.platform == 'win32':
    sys.stdout.reconfigure(encoding='utf-8', errors='replace')
    sys.stderr.reconfigure(encoding='utf-8', errors='replace')

# --- Константы ---
BASE_DIR = pathlib.Path(__file__).parent.parent
DATA_DIR = BASE_DIR.parent / "data"  # Z:/Мой диск/ИИ/employees_KPI/data
DB_PATH  = DATA_DIR / "bureau_data.sqlite"

# Маппинг столбцов XLS → поля базы данных
# Поддерживаются оба формата: старый (до фев 2026) и новый (фев 2026+)
COLUMN_MAP = {
    # --- Идентификаторы ---
    'ID':               'id',
    'Период':           'period',
    'Начало':           'start_datetime',

    # --- Организационные ---
    'Категория':        'category',          # новый формат
    'Сотрудник':        'employee',
    'Квалификация':     'qualification',     # новый формат
    'Подразделение':    'department',        # новый формат (был закомментирован)
    'Исполнитель':      'contractor',        # новый формат
    'Заказчик':         'client',
    'Центр затрат':     'cost_center',

    # --- Проект ---
    'Код проекта':      'project_code',
    'Тип проекта':      'project_type',
    'Проект':           'project_name',

    # --- Описание работы (раньше не импортировалось по ТЗ) ---
    'Вид деятельности': 'activity_type',
    'Тема':             'topic',
    'Тип задания':      'task_type',
    'Задание':          'task',

    # --- Часы ---
    'Длительность':           'duration',
    'Округлено':              'rounded',
    'К оплате':               'billable_hours',
    'Сверхурочная':           'overtime',
    'В калькуляции, час':     'hours_calc',      # новый формат
    'Отклонение, час':        'hours_deviation', # новый формат

    # --- Ставка и суммы ---
    'Ставка':                 'rate',
    'Включать в счет':        'include_in_bill',
    'Сумма':                  'amount',           # старый формат
    'Сумма исходная':         'amount_original',  # новый формат
    'Сумма, учтено':          'amount_recorded',  # новый формат
    'Сумма в калькуляции':    'amount_calc',      # новый формат
    'Коэф-т':                 'coefficient',      # новый формат
    'Сумма, отклонение':      'amount_deviation', # новый формат

    # --- Валюта и учётная валюта ---
    'Валюта':                                     'currency',
    'Курс к учетной валюте (значения)':           'exchange_rate',          # новый формат
    'Сумма, учтено в учетной валюте':             'amount_recorded_rub',    # новый формат
    'Сумма в калькуляции в учетной валюте':       'amount_calc_rub',        # новый формат
    'Сумма, отклонение в учетной валюте':         'amount_deviation_rub',   # новый формат

    # --- Калькуляция ---
    'Номер калькуляции':   'calc_number',  # новый формат
    'Дата калькуляции':    'calc_date',    # новый формат
    'Состояние калькуляции': 'calc_state', # новый формат

    # --- Акт ---
    'Нома акта':    'act_number',  # новый формат (опечатка в ProjectMate: "Нома" вместо "Номер")
    'Дата акта':    'act_date',    # новый формат
    'Состояние акта': 'act_state', # новый формат

    # --- Статус записи ---
    'Состояние записи о времени': 'entry_state',  # новый формат
    'Состояние':                  'entry_state',  # старый формат (оба → одно поле)

    # --- Аудит ---
    'Создано':    'created_at',
    'Создал':     'created_by',
    'Изменено':   'modified_at',
    'Изменил':    'modified_by',
}

# Булевы поля (будут конвертированы в 0/1)
BOOL_FIELDS = {'overtime', 'include_in_bill'}

# Поля с датами/временем
DATETIME_FIELDS = {'start_datetime', 'created_at', 'modified_at'}

# Поля только с датой (без времени)
DATE_FIELDS = {'calc_date', 'act_date'}

# Числовые поля NOT NULL (заменяем None → 0.0)
NUMERIC_NOT_NULL = {'duration', 'rounded', 'billable_hours', 'rate', 'amount'}

# Строковые поля NOT NULL (заменяем None → '')
STRING_NOT_NULL = {'currency', 'created_by', 'modified_by'}


def init_database(db_path: pathlib.Path) -> sqlite3.Connection:
    """Создаёт базу данных и все таблицы если они не существуют."""
    conn = sqlite3.connect(str(db_path))
    c = conn.cursor()

    c.executescript('''
        CREATE TABLE IF NOT EXISTS time_entries (
            -- Идентификаторы
            id              INTEGER PRIMARY KEY,
            period          TEXT NOT NULL,
            start_datetime  DATETIME NOT NULL,

            -- Организационные
            category        TEXT,
            employee        TEXT NOT NULL,
            qualification   TEXT,
            department      TEXT,
            contractor      TEXT,
            client          TEXT,
            cost_center     TEXT,

            -- Проект
            project_code    TEXT,
            project_type    TEXT,
            project_name    TEXT,

            -- Описание работы
            activity_type   TEXT,
            topic           TEXT,
            task_type       TEXT,
            task            TEXT,

            -- Часы
            duration        REAL NOT NULL DEFAULT 0,
            rounded         REAL NOT NULL DEFAULT 0,
            billable_hours  REAL NOT NULL DEFAULT 0,
            overtime        BOOLEAN NOT NULL DEFAULT 0,
            hours_calc      REAL,
            hours_deviation REAL,

            -- Ставка и суммы
            rate              REAL NOT NULL DEFAULT 0,
            include_in_bill   BOOLEAN NOT NULL DEFAULT 0,
            amount            REAL NOT NULL DEFAULT 0,
            amount_original   REAL,
            amount_recorded   REAL,
            amount_calc       REAL,
            coefficient       REAL,
            amount_deviation  REAL,

            -- Валюта и учётная валюта
            currency              TEXT NOT NULL DEFAULT 'RUB',
            exchange_rate         REAL,
            amount_recorded_rub   REAL,
            amount_calc_rub       REAL,
            amount_deviation_rub  REAL,

            -- Калькуляция
            calc_number TEXT,
            calc_date   TEXT,
            calc_state  TEXT,

            -- Акт
            act_number  TEXT,
            act_date    TEXT,
            act_state   TEXT,

            -- Статус записи
            entry_state TEXT,

            -- Аудит
            created_at  DATETIME NOT NULL,
            created_by  TEXT NOT NULL DEFAULT '',
            modified_at DATETIME NOT NULL,
            modified_by TEXT NOT NULL DEFAULT '',

            -- Служебное
            import_date DATETIME NOT NULL
        );

        CREATE INDEX IF NOT EXISTS idx_te_period       ON time_entries(period);
        CREATE INDEX IF NOT EXISTS idx_te_employee     ON time_entries(employee);
        CREATE INDEX IF NOT EXISTS idx_te_project_type ON time_entries(project_type);
        CREATE INDEX IF NOT EXISTS idx_te_department   ON time_entries(department);

        CREATE TABLE IF NOT EXISTS staff (
            employee            TEXT PRIMARY KEY,
            work_group          TEXT,
            level               TEXT,
            salary              REAL,
            target_hours        REAL,
            target_realization  REAL,
            hire_date           DATE,
            dismissal_date      DATE
        );

        CREATE TABLE IF NOT EXISTS salary_data (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            employee    TEXT NOT NULL,
            period      TEXT NOT NULL,
            salary      REAL,
            bonus       REAL,
            total       REAL,
            comment     TEXT,
            import_date DATETIME NOT NULL,
            UNIQUE(employee, period)
        );

        CREATE TABLE IF NOT EXISTS exchange_rates (
            period   TEXT PRIMARY KEY,
            usd_rate REAL NOT NULL,
            eur_rate REAL NOT NULL
        );

        CREATE TABLE IF NOT EXISTS import_log (
            id            INTEGER PRIMARY KEY AUTOINCREMENT,
            filename      TEXT NOT NULL,
            period        TEXT NOT NULL,
            rows_imported INTEGER NOT NULL,
            import_date   DATETIME NOT NULL,
            data_type     TEXT NOT NULL DEFAULT 'time_entries'
        );
    ''')
    conn.commit()
    return conn


def migrate_database(conn: sqlite3.Connection):
    """
    Добавляет новые столбцы в существующую базу данных.
    Безопасно: пропускает уже существующие столбцы.
    """
    c = conn.cursor()

    # Получаем существующие столбцы time_entries
    existing = {row[1] for row in c.execute('PRAGMA table_info(time_entries)')}

    # Новые столбцы: (имя, тип, default)
    new_columns = [
        ('category',             'TEXT',     None),
        ('qualification',        'TEXT',     None),
        ('contractor',           'TEXT',     None),
        ('activity_type',        'TEXT',     None),
        ('topic',                'TEXT',     None),
        ('task_type',            'TEXT',     None),
        ('task',                 'TEXT',     None),
        ('hours_calc',           'REAL',     None),
        ('hours_deviation',      'REAL',     None),
        ('amount_original',      'REAL',     None),
        ('amount_recorded',      'REAL',     None),
        ('amount_calc',          'REAL',     None),
        ('coefficient',          'REAL',     None),
        ('amount_deviation',     'REAL',     None),
        ('exchange_rate',        'REAL',     None),
        ('amount_recorded_rub',  'REAL',     None),
        ('amount_calc_rub',      'REAL',     None),
        ('amount_deviation_rub', 'REAL',     None),
        ('calc_number',          'TEXT',     None),
        ('calc_date',            'TEXT',     None),
        ('calc_state',           'TEXT',     None),
        ('act_number',           'TEXT',     None),
        ('act_date',             'TEXT',     None),
        ('act_state',            'TEXT',     None),
        ('entry_state',          'TEXT',     None),
    ]

    added = []
    for col_name, col_type, col_default in new_columns:
        if col_name not in existing:
            if col_default is not None:
                sql = f'ALTER TABLE time_entries ADD COLUMN {col_name} {col_type} DEFAULT {col_default}'
            else:
                sql = f'ALTER TABLE time_entries ADD COLUMN {col_name} {col_type}'
            c.execute(sql)
            added.append(col_name)

    # Миграция таблицы staff — добавить dismissal_date если нет
    existing_staff = {row[1] for row in c.execute('PRAGMA table_info(staff)')}
    if 'dismissal_date' not in existing_staff:
        c.execute('ALTER TABLE staff ADD COLUMN dismissal_date DATE')
        added.append('staff.dismissal_date')

    conn.commit()

    # Создаём индексы для новых полей (после того как столбцы точно существуют)
    c.execute('CREATE INDEX IF NOT EXISTS idx_te_entry_state  ON time_entries(entry_state)')
    c.execute('CREATE INDEX IF NOT EXISTS idx_te_project_code ON time_entries(project_code)')
    conn.commit()

    if added:
        print(f"✅ Миграция БД: добавлено {len(added)} столбцов: {', '.join(added)}")
    else:
        print("✅ Миграция БД: схема актуальна, изменений не требуется")


def read_xls_via_com(filepath: pathlib.Path) -> list:
    """
    Читает XLS-файл ProjectMate через Excel COM (win32com).
    Возвращает список словарей с данными.

    Отступил от ТЗ: xlrd не работает с этими файлами (>2MB),
    поэтому используем win32com.client. Каждый файл открывается
    в отдельном экземпляре Excel (DispatchEx) чтобы избежать
    проблем с COM RPC.
    """
    import time

    try:
        import win32com.client
    except ImportError:
        print("ОШИБКА: pywin32 не установлен. Запустите: pip install pywin32")
        sys.exit(1)

    excel = win32com.client.DispatchEx('Excel.Application')
    excel.Visible = False
    excel.DisplayAlerts = False
    time.sleep(1)  # Даём Excel запуститься

    try:
        wb = excel.Workbooks.Open(str(filepath.absolute()))
        time.sleep(1)  # Даём файлу открыться

        # Ищем лист «Данные»
        ws = None
        for i in range(1, wb.Sheets.Count + 1):
            if wb.Sheets(i).Name == 'Данные':
                ws = wb.Sheets(i)
                break

        if ws is None:
            wb.Close(False)
            raise ValueError(f"Лист 'Данные' не найден в файле {filepath.name}")

        data = ws.UsedRange.Value
        try:
            wb.Close(SaveChanges=False)
        except Exception:
            pass

        if not data or len(data) < 2:
            return []

        headers = [str(h) if h is not None else '' for h in data[0]]
        rows = []
        for row in data[1:]:
            row_dict = {}
            for i, h in enumerate(headers):
                val = row[i] if i < len(row) else None
                row_dict[h] = val
            rows.append(row_dict)

        return rows

    finally:
        try:
            excel.Quit()
        except Exception:
            pass  # COM Quit() может упасть — не критично


def clean_value(value, field_name: str):
    """Очищает и конвертирует значение ячейки для записи в базу."""
    if value is None:
        return None

    # decimal.Decimal → float (новый формат выгрузки использует Decimal для чисел)
    import decimal
    if isinstance(value, decimal.Decimal):
        return float(value)

    # Булевы поля → 0/1
    if field_name in BOOL_FIELDS:
        if isinstance(value, bool):
            return 1 if value else 0
        if isinstance(value, str):
            return 1 if value.lower() in ('да', 'yes', 'true', '1') else 0
        return 1 if value else 0

    # Поля дат/времени → ISO-строка с временем
    if field_name in DATETIME_FIELDS:
        if hasattr(value, 'year'):  # pywintypes.datetime или datetime
            try:
                return (f"{value.year:04d}-{value.month:02d}-{value.day:02d} "
                        f"{value.hour:02d}:{value.minute:02d}:{value.second:02d}")
            except Exception:
                return str(value)
        return str(value) if value else None

    # Поля только с датой → ISO-строка без времени
    if field_name in DATE_FIELDS:
        if hasattr(value, 'year'):
            try:
                return f"{value.year:04d}-{value.month:02d}-{value.day:02d}"
            except Exception:
                return str(value)
        return str(value) if value else None

    # ID → int
    if field_name == 'id':
        try:
            return int(float(value))
        except (ValueError, TypeError):
            return None

    # Пустые строки → NULL
    if isinstance(value, str) and value.strip() == '':
        return None

    return value


def import_xls(filepath: pathlib.Path, conn: sqlite3.Connection) -> str:
    """
    Импортирует один XLS-файл в базу данных.
    Возвращает строку с результатом.
    """
    filename = filepath.name
    print(f"  Читаю {filename}...", end='', flush=True)

    rows = read_xls_via_com(filepath)
    if not rows:
        return f"⚠️  {filename}: файл пуст или не удалось прочитать"

    print(f" {len(rows)} строк...", end='', flush=True)

    import_date = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    records = []
    skipped = 0

    for raw_row in rows:
        record = {}

        # Маппинг всех столбцов из COLUMN_MAP
        for xls_col, db_field in COLUMN_MAP.items():
            raw_val = raw_row.get(xls_col)
            cleaned = clean_value(raw_val, db_field)
            # Не перезаписываем уже заполненное поле значением None
            # (актуально для entry_state, который маппится из двух столбцов)
            if cleaned is not None or db_field not in record:
                record[db_field] = cleaned

        # Обязательные поля — если нет id, пропускаем строку
        if record.get('id') is None:
            skipped += 1
            continue

        # Совместимость: старый формат имел "Сумма" → amount,
        # новый — "Сумма, учтено" → amount_recorded.
        # Если amount не заполнен, берём amount_recorded как основную сумму.
        if not record.get('amount') and record.get('amount_recorded'):
            record['amount'] = record['amount_recorded']

        # Служебное поле
        record['import_date'] = import_date

        # Числовые NOT NULL поля — заменяем None на 0
        for field in NUMERIC_NOT_NULL:
            if record.get(field) is None:
                record[field] = 0.0

        # Булевы NOT NULL поля
        for field in ('overtime', 'include_in_bill'):
            if record.get(field) is None:
                record[field] = 0

        # Строковые NOT NULL поля
        for field in STRING_NOT_NULL:
            if record.get(field) is None:
                record[field] = ''

        # Datetime NOT NULL поля — если пусты, ставим дату импорта
        for field in ('start_datetime', 'created_at', 'modified_at'):
            if record.get(field) is None:
                record[field] = import_date

        records.append(record)

    if not records:
        return f"⚠️  {filename}: нет корректных записей (пропущено {skipped})"

    # Вставляем в базу — динамически берём столбцы из первой записи
    cols = list(records[0].keys())
    placeholders = ','.join(['?'] * len(cols))
    sql = f'INSERT OR IGNORE INTO time_entries ({",".join(cols)}) VALUES ({placeholders})'

    inserted = 0
    for record in records:
        values = tuple(record.get(c) for c in cols)
        conn.execute(sql, values)
        inserted += conn.execute('SELECT changes()').fetchone()[0]

    period = records[0].get('period', '?') or '?'
    conn.execute(
        'INSERT INTO import_log (filename, period, rows_imported, import_date, data_type) VALUES (?,?,?,?,?)',
        (filename, period, inserted, import_date, 'time_entries')
    )
    conn.commit()

    already = len(records) - inserted
    if already > 0:
        return f"✅  {filename}: +{inserted} новых, {already} уже в базе. Период: {period}"
    return f"✅  {filename}: +{inserted} записей. Период: {period}"


def print_statistics(conn: sqlite3.Connection):
    """Выводит статистику по базе данных."""
    print("\n" + "=" * 70)
    print("📊 СТАТИСТИКА БАЗЫ ДАННЫХ")
    print("=" * 70)

    rows = conn.execute('''
        SELECT
            period,
            COUNT(*) as records,
            COUNT(DISTINCT employee) as employees,
            ROUND(SUM(duration), 1) as total_hours,
            ROUND(SUM(billable_hours), 1) as billable
        FROM time_entries
        GROUP BY period
        ORDER BY MIN(start_datetime)
    ''').fetchall()

    print(f"\n{'Период':<20} {'Записей':>8} {'Сотруд.':>8} {'Часов':>10} {'К оплате':>10}")
    print("-" * 60)
    for r in rows:
        print(f"{r[0]:<20} {r[1]:>8} {r[2]:>8} {r[3]:>10.1f} {r[4]:>10.1f}")

    total = conn.execute('SELECT COUNT(*), COUNT(DISTINCT employee) FROM time_entries').fetchone()
    print("-" * 60)
    print(f"ИТОГО: {total[0]} записей, {total[1]} уникальных сотрудников")

    # Типы проектов
    types = conn.execute('''
        SELECT project_type, COUNT(*), ROUND(SUM(duration),1)
        FROM time_entries
        WHERE project_type IS NOT NULL
        GROUP BY project_type
        ORDER BY SUM(duration) DESC
    ''').fetchall()
    print("\n📋 По типам проектов:")
    for t in types:
        print(f"  {t[0]}: {t[1]} записей, {t[2]} ч.")

    # Покрытие новых полей (проверяем качество данных)
    dept_filled = conn.execute(
        'SELECT COUNT(*) FROM time_entries WHERE department IS NOT NULL'
    ).fetchone()[0]
    total_rows = conn.execute('SELECT COUNT(*) FROM time_entries').fetchone()[0]
    if total_rows > 0:
        print(f"\n📐 Заполненность новых полей:")
        print(f"  Подразделение: {dept_filled}/{total_rows} ({100*dept_filled//total_rows}%)")

    # Статусы записей (новое поле)
    states = conn.execute('''
        SELECT entry_state, COUNT(*)
        FROM time_entries
        WHERE entry_state IS NOT NULL
        GROUP BY entry_state
        ORDER BY COUNT(*) DESC
    ''').fetchall()
    if states:
        print("\n📌 Статусы записей:")
        for s in states:
            print(f"  {s[0]}: {s[1]}")

    # Валюты
    currencies = conn.execute('''
        SELECT currency, COUNT(*), ROUND(SUM(amount),0)
        FROM time_entries
        GROUP BY currency
    ''').fetchall()
    print("\n💱 Валюты:")
    for c in currencies:
        print(f"  {c[0] or 'N/A'}: {c[1]} записей, сумма {c[2]:,.0f}")

    print("=" * 70)


def main():
    print("=" * 70)
    print("📥 ИМПОРТ ДАННЫХ ProjectMate → SQLite  (v2.0)")
    print("=" * 70)
    print(f"\nПапка данных: {DATA_DIR}")
    print(f"База данных:  {DB_PATH}\n")

    # Проверяем папку с данными
    if not DATA_DIR.exists():
        print(f"ОШИБКА: Папка данных не найдена: {DATA_DIR}")
        sys.exit(1)

    # Создаём/открываем базу
    conn = init_database(DB_PATH)
    print(f"✅ База данных инициализирована: {DB_PATH.name}")

    # Миграция существующей базы (добавляем новые столбцы)
    migrate_database(conn)
    print()

    # Находим все XLS-файлы
    xls_files = sorted(DATA_DIR.glob('*.xls')) + sorted(DATA_DIR.glob('*.xlsx'))
    # Исключаем справочники и шаблоны
    xls_files = [f for f in xls_files if not f.name.startswith('Справочник')]

    if not xls_files:
        print("ОШИБКА: XLS-файлы не найдены в папке данных")
        sys.exit(1)

    print(f"📂 Найдено файлов: {len(xls_files)}")
    print()

    # Импортируем каждый файл
    for xls_file in xls_files:
        result = import_xls(xls_file, conn)
        print(result)

    # Статистика
    print_statistics(conn)
    conn.close()
    print("\n✅ Импорт завершён!")


if __name__ == '__main__':
    main()
