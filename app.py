"""
Главная страница Streamlit-дашборда «Аналитика бюро».
"""

import streamlit as st
import sys
import pathlib

# Добавляем папку dashboard в sys.path
sys.path.insert(0, str(pathlib.Path(__file__).parent))

st.set_page_config(
    page_title="Аналитика бюро",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

from utils.data import render_sidebar, get_filtered_data, get_periods, load_time_entries

# Боковая панель
filters = render_sidebar(key_prefix='main_')

st.title("⚖️ Аналитика юридического бюро")
st.caption("Интерактивный дашборд на основе данных ProjectMate")

st.divider()

# Краткая сводка
try:
    df = load_time_entries()
    periods = get_periods()

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("Периодов в базе", len(periods))
    with col2:
        st.metric("Всего записей", f"{len(df):,}")
    with col3:
        st.metric("Сотрудников", df['employee'].nunique())
    with col4:
        total_hours = df['duration'].sum()
        st.metric("Всего часов", f"{total_hours:,.0f}")

    st.divider()

    st.markdown("""
    ### Навигация

    Используйте меню слева для перехода между разделами:

    | Раздел | Содержание |
    |--------|-----------|
    | 📊 Обзор | KPI-карточки, динамика по месяцам, распределение по типам проектов |
    | 👥 Сотрудники | Таблица с показателями сотрудников за 3 периода, реализация |
    | 📈 Динамика | Конструктор графиков: выберите показатель и разрез |
    | 📁 Проекты | Топ-10 проектов по часам и суммам |
    | 💰 ФОТ | В разработке |

    ### Текущие данные

    """)

    # Таблица периодов
    import pandas as pd
    summary = df.groupby('period').agg(
        Записей=('id', 'count'),
        Сотрудников=('employee', 'nunique'),
        Часов=('duration', 'sum'),
        К_оплате=('billable_hours', 'sum'),
    ).reset_index()
    summary['Часов'] = summary['Часов'].round(1)
    summary['К_оплате'] = summary['К_оплате'].round(1)
    summary.columns = ['Период', 'Записей', 'Сотрудников', 'Часов', 'К оплате']

    st.dataframe(
        summary,
        use_container_width=True,
        hide_index=True,
    )

except Exception as e:
    st.error(f"Ошибка подключения к базе данных: {e}")
    st.info("Убедитесь, что файл bureau_data.sqlite находится в папке dashboard/")
