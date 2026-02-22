"""
Утилиты загрузки данных из SQLite-базы.
Поиск базы: текущая папка → ../data/ → переменная окружения DB_PATH → st.secrets.
"""

import sqlite3
import pandas as pd
import pathlib
import os
from typing import List, Optional
import streamlit as st


def _find_db() -> str:
    """Ищет файл базы данных в нескольких местах."""
    candidates = [
        pathlib.Path('bureau_data.sqlite'),
        pathlib.Path(__file__).parent.parent / 'bureau_data.sqlite',
        pathlib.Path(__file__).parent.parent.parent / 'data' / 'bureau_data.sqlite',
    ]

    # Переменная окружения
    env_path = os.environ.get('DB_PATH')
    if env_path:
        candidates.insert(0, pathlib.Path(env_path))

    for p in candidates:
        if p.exists():
            return str(p)

    # st.secrets (для Streamlit Cloud)
    try:
        return st.secrets.get('DB_PATH', '')
    except Exception:
        pass

    raise FileNotFoundError(
        "Файл bureau_data.sqlite не найден. "
        "Поместите его в папку dashboard/ или задайте переменную DB_PATH."
    )


@st.cache_resource
def get_connection():
    """Возвращает подключение к базе данных (кешируется на весь сеанс)."""
    db_path = _find_db()
    conn = sqlite3.connect(db_path, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


@st.cache_data(ttl=3600)
def load_time_entries() -> pd.DataFrame:
    """Загружает все записи о рабочем времени."""
    conn = get_connection()
    df = pd.read_sql('''
        SELECT
            t.*,
            COALESCE(t.department, s.work_group) AS work_group_resolved,
            s.level,
            s.target_hours,
            s.target_realization
        FROM time_entries t
        LEFT JOIN staff s ON t.employee = s.employee
    ''', conn)
    return df


@st.cache_data(ttl=3600)
def load_staff() -> pd.DataFrame:
    """Загружает справочник сотрудников."""
    conn = get_connection()
    return pd.read_sql('SELECT * FROM staff ORDER BY employee', conn)


@st.cache_data(ttl=3600)
def load_exchange_rates() -> pd.DataFrame:
    """Загружает курсы валют."""
    conn = get_connection()
    return pd.read_sql('SELECT * FROM exchange_rates', conn)


@st.cache_data(ttl=3600)
def get_periods() -> List[str]:
    """Возвращает список всех периодов в хронологическом порядке."""
    conn = get_connection()
    rows = conn.execute('''
        SELECT DISTINCT period, MIN(start_datetime) as dt
        FROM time_entries
        GROUP BY period
        ORDER BY dt
    ''').fetchall()
    return [r[0] for r in rows]


@st.cache_data(ttl=3600)
def get_work_groups() -> List[str]:
    """Возвращает список всех рабочих групп из справочника."""
    conn = get_connection()
    rows = conn.execute('''
        SELECT DISTINCT work_group
        FROM staff
        WHERE work_group IS NOT NULL AND work_group != ''
        ORDER BY work_group
    ''').fetchall()
    return [r[0] for r in rows]


def render_sidebar(key_prefix: str = '') -> dict:
    """
    Отрисовывает единую боковую панель фильтров.
    Возвращает словарь с выбранными значениями фильтров.
    """
    with st.sidebar:
        st.image("https://img.icons8.com/color/96/scales.png", width=60)
        st.title("Аналитика бюро")
        st.divider()

        periods = get_periods()
        work_groups = get_work_groups()
        has_staff = len(work_groups) > 0

        # Уровень агрегации
        level = st.radio(
            "Уровень",
            ["Бюро", "По группам", "По сотрудникам"],
            key=f"{key_prefix}level"
        )

        # Рабочая группа (если есть справочник)
        selected_groups = []
        if has_staff and level in ("По группам", "По сотрудникам"):
            selected_groups = st.multiselect(
                "Рабочая группа",
                work_groups,
                default=work_groups,
                key=f"{key_prefix}groups"
            )

        # Сотрудник
        selected_employee = None
        if level == "По сотрудникам":
            # Фильтруем сотрудников по выбранным группам
            conn = get_connection()
            if has_staff and selected_groups:
                employees = conn.execute('''
                    SELECT DISTINCT t.employee
                    FROM time_entries t
                    JOIN staff s ON t.employee = s.employee
                    WHERE s.work_group IN ({})
                    ORDER BY t.employee
                '''.format(','.join('?' * len(selected_groups))),
                    selected_groups
                ).fetchall()
            else:
                employees = conn.execute(
                    'SELECT DISTINCT employee FROM time_entries ORDER BY employee'
                ).fetchall()

            employee_list = [r[0] for r in employees]
            if employee_list:
                selected_employee = st.selectbox(
                    "Сотрудник",
                    employee_list,
                    key=f"{key_prefix}employee"
                )

        # Периоды
        default_periods = periods[-3:] if len(periods) >= 3 else periods
        selected_periods = st.multiselect(
            "Периоды",
            periods,
            default=default_periods,
            key=f"{key_prefix}periods"
        )

        # Типы проектов
        project_types = [
            'Работа по договорам',
            'Внутренний проект',
            'Обучение',
            'Отпуск',
            'Больничный лист',
            'Развитие бизнеса',
        ]
        selected_types = st.multiselect(
            "Типы проектов",
            project_types,
            default=project_types,
            key=f"{key_prefix}types"
        )

        st.divider()
        st.caption("Данные: ProjectMate")

    return {
        'level': level,
        'work_groups': selected_groups,
        'employee': selected_employee,
        'periods': selected_periods,
        'project_types': selected_types,
        'has_staff': has_staff,
    }


def get_filtered_data(filters: dict) -> pd.DataFrame:
    """
    Возвращает отфильтрованный DataFrame согласно выбранным фильтрам.
    """
    df = load_time_entries()

    if not df.empty:
        # Фильтр по периодам
        if filters.get('periods'):
            df = df[df['period'].isin(filters['periods'])]

        # Фильтр по типам проектов
        if filters.get('project_types'):
            df = df[df['project_type'].isin(filters['project_types']) |
                    df['project_type'].isna()]

        # Фильтр по рабочей группе
        if filters.get('work_groups') and filters['has_staff']:
            df = df[df['work_group_resolved'].isin(filters['work_groups'])]

        # Фильтр по сотруднику
        if filters.get('employee'):
            df = df[df['employee'] == filters['employee']]

    return df
