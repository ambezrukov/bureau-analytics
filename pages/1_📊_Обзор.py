"""
Страница «Обзор» — KPI-карточки и базовые графики.
"""

import sys
import pathlib
sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

import streamlit as st
import pandas as pd
from utils.data import render_sidebar, get_filtered_data, get_periods
from utils.charts import line_chart_hours_realization, bar_chart_project_types

st.set_page_config(page_title="Обзор | Аналитика бюро", page_icon="📊", layout="wide")

filters = render_sidebar(key_prefix='obzor_')

st.title("📊 Обзор")

if not filters['periods']:
    st.warning("Выберите хотя бы один период в боковой панели.")
    st.stop()

df = get_filtered_data(filters)

if df.empty:
    st.warning("Нет данных для выбранных фильтров.")
    st.stop()

# ── KPI-карточки ──────────────────────────────────────────────────────────────
periods = get_periods()

# Текущий период = последний из выбранных
selected_sorted = sorted(filters['periods'], key=lambda p: periods.index(p) if p in periods else 0)
current_period = selected_sorted[-1] if selected_sorted else None
prev_periods = selected_sorted[:-1]

df_cur = df[df['period'] == current_period] if current_period else df
df_prev = df[df['period'].isin(prev_periods)] if prev_periods else pd.DataFrame()

def calc_kpis(data: pd.DataFrame) -> dict:
    """Вычисляет ключевые показатели."""
    total_hours = data['duration'].sum()
    commercial = data[data['project_type'] == 'Работа по договорам']['duration'].sum()
    billable = data['billable_hours'].sum()
    realization = (billable / commercial * 100) if commercial > 0 else 0

    # Сумма счетов в рублях (с конвертацией по умолч. курсам)
    USD_RATE = 88.0
    EUR_RATE = 95.0
    bills = data[data['include_in_bill'] == 1]
    amount_rub = (
        bills[bills['currency'] == 'RUB']['amount'].sum() +
        bills[bills['currency'] == 'USD']['amount'].sum() * USD_RATE +
        bills[bills['currency'] == 'EUR']['amount'].sum() * EUR_RATE
    )
    return {
        'total_hours': total_hours,
        'commercial': commercial,
        'billable': billable,
        'realization': realization,
        'amount_rub': amount_rub,
    }

kpi_cur = calc_kpis(df_cur)
kpi_prev = calc_kpis(df_prev) if not df_prev.empty else None

def delta_str(cur, prev, fmt='.1f', suffix=''):
    if prev is None or prev == 0:
        return None
    d = cur - prev
    sign = '+' if d >= 0 else ''
    return f"{sign}{d:{fmt}}{suffix}"

st.subheader(f"Период: {current_period or 'все'}")

col1, col2, col3, col4, col5 = st.columns(5)

with col1:
    st.metric(
        "Всего часов",
        f"{kpi_cur['total_hours']:,.1f}",
        delta=delta_str(kpi_cur['total_hours'], kpi_prev['total_hours'] if kpi_prev else None, '.0f'),
    )
with col2:
    st.metric(
        "Коммерческие ч.",
        f"{kpi_cur['commercial']:,.1f}",
        delta=delta_str(kpi_cur['commercial'], kpi_prev['commercial'] if kpi_prev else None, '.0f'),
    )
with col3:
    st.metric(
        "К оплате ч.",
        f"{kpi_cur['billable']:,.1f}",
        delta=delta_str(kpi_cur['billable'], kpi_prev['billable'] if kpi_prev else None, '.0f'),
    )
with col4:
    st.metric(
        "Реализация %",
        f"{kpi_cur['realization']:.1f}%",
        delta=delta_str(kpi_cur['realization'], kpi_prev['realization'] if kpi_prev else None, '.1f', '%'),
    )
with col5:
    st.metric(
        "Сумма счетов ₽",
        f"{kpi_cur['amount_rub']:,.0f}",
        delta=delta_str(kpi_cur['amount_rub'], kpi_prev['amount_rub'] if kpi_prev else None, '.0f'),
    )

st.divider()

# ── Линейный график: динамика по месяцам ──────────────────────────────────────
st.subheader("Динамика по месяцам")

df_by_period = df.groupby('period').agg(
    total_hours=('duration', 'sum'),
    commercial_hours=('duration', lambda x: x[df.loc[x.index, 'project_type'] == 'Работа по договорам'].sum()),
    billable_hours=('billable_hours', 'sum'),
).reset_index()

# Реализация
df_by_period['realization_pct'] = (
    df_by_period['billable_hours'] / df_by_period['commercial_hours'].replace(0, float('nan')) * 100
).round(1)

# Сортируем по порядку периодов
all_periods = get_periods()
df_by_period['_order'] = df_by_period['period'].apply(
    lambda p: all_periods.index(p) if p in all_periods else 0
)
df_by_period = df_by_period.sort_values('_order').drop(columns='_order')

st.plotly_chart(line_chart_hours_realization(df_by_period), use_container_width=True)

# ── Распределение по типам проектов ───────────────────────────────────────────
col_left, col_right = st.columns(2)

with col_left:
    st.subheader("По типам проектов (часы)")
    st.plotly_chart(bar_chart_project_types(df), use_container_width=True)

with col_right:
    st.subheader("Топ-10 сотрудников по часам")
    top_empl = df.groupby('employee')['duration'].sum().nlargest(10).reset_index()
    top_empl.columns = ['Сотрудник', 'Часов']
    top_empl['Часов'] = top_empl['Часов'].round(1)
    st.dataframe(top_empl, use_container_width=True, hide_index=True)

# ── Сводная таблица ───────────────────────────────────────────────────────────
st.divider()
st.subheader("Сводная таблица по периодам")

pivot = df.groupby('period').agg(
    Записей=('id', 'count'),
    Сотрудников=('employee', 'nunique'),
).reset_index()

df_commercial = df[df['project_type'] == 'Работа по договорам'].groupby('period').agg(
    Коммерч_ч=('duration', 'sum'),
).reset_index()

df_bill = df.groupby('period').agg(
    К_оплате_ч=('billable_hours', 'sum'),
    Всего_ч=('duration', 'sum'),
).reset_index()

merged = pivot.merge(df_bill, on='period').merge(df_commercial, on='period', how='left')
merged['Реализация_%'] = (merged['К_оплате_ч'] / merged['Коммерч_ч'].replace(0, float('nan')) * 100).round(1)

merged.columns = ['Период', 'Записей', 'Сотрудников', 'К оплате ч.', 'Всего ч.', 'Коммерч. ч.', 'Реализация %']
merged['Всего ч.'] = merged['Всего ч.'].round(1)
merged['К оплате ч.'] = merged['К оплате ч.'].round(1)
merged['Коммерч. ч.'] = merged['Коммерч. ч.'].round(1)

# Сортируем по порядку
merged['_order'] = merged['Период'].apply(lambda p: all_periods.index(p) if p in all_periods else 0)
merged = merged.sort_values('_order').drop(columns='_order')

st.dataframe(merged, use_container_width=True, hide_index=True)
