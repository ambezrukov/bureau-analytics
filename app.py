"""
Главная страница Streamlit-дашборда «Аналитика бюро».
"""

import streamlit as st
import sys
import pathlib
import traceback

# Добавляем папку dashboard в sys.path чтобы utils/ был виден
sys.path.insert(0, str(pathlib.Path(__file__).parent))

st.set_page_config(
    page_title="Аналитика бюро",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

# Импортируем утилиты — оборачиваем в try чтобы видеть ошибку
try:
    from utils.data import render_sidebar, get_periods, load_time_entries
    _import_ok = True
except Exception as _e:
    _import_ok = False
    _import_error = traceback.format_exc()

if not _import_ok:
    st.error("❌ Ошибка импорта модулей")
    st.code(_import_error)
    st.stop()

# Боковая панель
try:
    filters = render_sidebar(key_prefix='main_')
except Exception as e:
    st.error(f"❌ Ошибка боковой панели: {e}")
    st.code(traceback.format_exc())
    st.stop()

st.title("⚖️ Аналитика юридического бюро")
st.caption("Интерактивный дашборд на основе данных ProjectMate")
st.divider()

# Краткая сводка
try:
    import pandas as pd

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

    summary = df.groupby('period').agg(
        Записей=('id', 'count'),
        Сотрудников=('employee', 'nunique'),
        Часов=('duration', 'sum'),
        К_оплате=('billable_hours', 'sum'),
    ).reset_index()
    summary['Часов'] = summary['Часов'].round(1)
    summary['К_оплате'] = summary['К_оплате'].round(1)
    summary.columns = ['Период', 'Записей', 'Сотрудников', 'Часов', 'К оплате']

    st.dataframe(summary, use_container_width=True, hide_index=True)

except Exception as e:
    st.error(f"❌ Ошибка загрузки данных: {e}")
    st.code(traceback.format_exc())
    st.info("Убедитесь, что файл bureau_data.sqlite находится в папке dashboard/")
