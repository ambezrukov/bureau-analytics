"""
Страница «Сотрудники» — таблица KPI по каждому сотруднику за 3 периода.
"""

import sys
import pathlib
sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

import streamlit as st
import pandas as pd
from utils.data import render_sidebar, get_filtered_data, get_periods, load_time_entries
from utils.charts import horizontal_bar_realization

filters = render_sidebar(key_prefix='staff_')

st.title("👥 Сотрудники")

if not filters['periods']:
    st.warning("Выберите хотя бы один период в боковой панели.")
    st.stop()

# Загружаем все данные (без фильтра по типу проектов — нужны все для реализации)
filters_all_types = filters.copy()
filters_all_types['project_types'] = [
    'Работа по договорам', 'Внутренний проект', 'Обучение',
    'Отпуск', 'Больничный лист', 'Развитие бизнеса',
]
df = get_filtered_data(filters_all_types)

if df.empty:
    st.warning("Нет данных для выбранных фильтров.")
    st.stop()

all_periods = get_periods()

# Порог для подсветки
col_thr1, col_thr2, _ = st.columns([1, 1, 4])
with col_thr1:
    threshold_low = st.number_input("Порог (красный) %", value=70, min_value=0, max_value=100, step=5)
with col_thr2:
    threshold_high = st.number_input("Порог (зелёный) %", value=85, min_value=0, max_value=100, step=5)

st.divider()

# ── Таблица по сотрудникам ────────────────────────────────────────────────────
def calc_employee_stats(data: pd.DataFrame) -> pd.DataFrame:
    """Считает показатели по сотрудникам для одного периода."""
    total = data.groupby('employee')['duration'].sum().rename('total_hours')
    commercial = (
        data[data['project_type'] == 'Работа по договорам']
        .groupby('employee')['duration'].sum()
        .rename('commercial_hours')
    )
    billable = data.groupby('employee')['billable_hours'].sum()

    result = pd.concat([total, commercial, billable], axis=1).reset_index()
    result.columns = ['employee', 'total_hours', 'commercial_hours', 'billable_hours']
    result['realization_pct'] = (
        result['billable_hours'] / result['commercial_hours'].replace(0, float('nan')) * 100
    ).round(1)
    return result


# Берём последние 3 из выбранных периодов
selected_sorted = sorted(filters['periods'], key=lambda p: all_periods.index(p) if p in all_periods else 0)
show_periods = selected_sorted[-3:]  # максимум 3

dfs = {}
for p in show_periods:
    df_p = df[df['period'] == p]
    if not df_p.empty:
        dfs[p] = calc_employee_stats(df_p)

if not dfs:
    st.warning("Нет данных по сотрудникам.")
    st.stop()

# Объединяем в одну таблицу
all_employees = sorted(set(e for d in dfs.values() for e in d['employee']))
table_rows = []

for emp in all_employees:
    row = {'Сотрудник': emp}
    for p in show_periods:
        if p in dfs:
            emp_data = dfs[p][dfs[p]['employee'] == emp]
            if not emp_data.empty:
                r = emp_data.iloc[0]
                row[f'{p} | Часов'] = round(r['total_hours'], 1)
                row[f'{p} | К оплате'] = round(r['billable_hours'], 1)
                row[f'{p} | Реализация %'] = r['realization_pct']
            else:
                row[f'{p} | Часов'] = None
                row[f'{p} | К оплате'] = None
                row[f'{p} | Реализация %'] = None
    table_rows.append(row)

table_df = pd.DataFrame(table_rows)


def style_realization(val):
    """Подсветка ячеек реализации."""
    if pd.isna(val):
        return ''
    if val < threshold_low:
        return 'background-color: #FFD7D7; color: #C00000; font-weight: bold'
    elif val > threshold_high:
        return 'background-color: #D7F5D7; color: #375623; font-weight: bold'
    return ''


# Применяем стиль только к столбцам реализации
real_cols = [c for c in table_df.columns if 'Реализация' in c]
styled = table_df.style.applymap(style_realization, subset=real_cols)

st.subheader(f"Показатели сотрудников ({', '.join(show_periods)})")
st.dataframe(styled, use_container_width=True, hide_index=True, height=600)

# Кнопка скачать
import io
buf = io.BytesIO()
table_df.to_excel(buf, index=False, sheet_name='Сотрудники')
buf.seek(0)
st.download_button(
    "⬇️ Скачать в Excel",
    buf,
    file_name="сотрудники_kpi.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)

st.divider()

# ── График: реализация по сотрудникам (последний период) ──────────────────────
if show_periods:
    last_period = show_periods[-1]
    if last_period in dfs:
        df_real = dfs[last_period][['employee', 'realization_pct']].dropna()
        if not df_real.empty:
            st.subheader(f"Реализация по сотрудникам — {last_period}")
            fig = horizontal_bar_realization(df_real, threshold_low, threshold_high)
            st.plotly_chart(fig, use_container_width=True)
