"""
Страница «Проекты» — топ проектов по часам и суммам.
"""

import sys
import pathlib
sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

import streamlit as st
import pandas as pd
from utils.data import render_sidebar, get_filtered_data
from utils.charts import bar_top_projects

filters = render_sidebar(key_prefix='proj_')

st.title("📁 Проекты")

if not filters['periods']:
    st.warning("Выберите хотя бы один период в боковой панели.")
    st.stop()

df = get_filtered_data(filters)

if df.empty:
    st.warning("Нет данных для выбранных фильтров.")
    st.stop()

USD_RATE = 88.0
EUR_RATE = 95.0

# Считаем суммы с конвертацией
df = df.copy()
df['amount_rub'] = 0.0
df.loc[df['currency'] == 'RUB', 'amount_rub'] = df.loc[df['currency'] == 'RUB', 'amount']
df.loc[df['currency'] == 'USD', 'amount_rub'] = df.loc[df['currency'] == 'USD', 'amount'] * USD_RATE
df.loc[df['currency'] == 'EUR', 'amount_rub'] = df.loc[df['currency'] == 'EUR', 'amount'] * EUR_RATE

# Агрегация по проектам
proj_stats = df.groupby(['project_name', 'client', 'project_type']).agg(
    hours=('duration', 'sum'),
    employees=('employee', 'nunique'),
    amount_rub=('amount_rub', 'sum'),
).reset_index()

proj_stats['hours'] = proj_stats['hours'].round(1)
proj_stats['amount_rub'] = proj_stats['amount_rub'].round(0)

# ── Топ по часам ──────────────────────────────────────────────────────────────
st.subheader("Топ-10 проектов по часам")

top_hours = proj_stats.nlargest(10, 'hours')[['project_name', 'client', 'hours', 'employees', 'amount_rub']]
top_hours.columns = ['Проект', 'Клиент', 'Часов', 'Сотрудников', 'Сумма (руб.)']

col_chart, col_table = st.columns([1, 1])

with col_chart:
    top_for_chart = proj_stats.nlargest(10, 'hours')[['project_name', 'hours']]
    top_for_chart.columns = ['project_name', 'hours']
    fig = bar_top_projects(top_for_chart, 'hours', 'project_name', 'Топ-10 по часам')
    st.plotly_chart(fig, use_container_width=True)

with col_table:
    st.dataframe(top_hours, use_container_width=True, hide_index=True)

st.divider()

# ── Топ по суммам ─────────────────────────────────────────────────────────────
st.subheader("Топ-10 проектов по сумме")

top_amount = proj_stats[proj_stats['amount_rub'] > 0].nlargest(10, 'amount_rub')[
    ['project_name', 'client', 'hours', 'employees', 'amount_rub']
]
top_amount.columns = ['Проект', 'Клиент', 'Часов', 'Сотрудников', 'Сумма (руб.)']

col_chart2, col_table2 = st.columns([1, 1])

with col_chart2:
    top_for_chart2 = proj_stats[proj_stats['amount_rub'] > 0].nlargest(10, 'amount_rub')[
        ['project_name', 'amount_rub']
    ]
    top_for_chart2.columns = ['project_name', 'amount_rub']
    top_for_chart2['amount_rub'] = (top_for_chart2['amount_rub'] / 1_000_000).round(2)
    fig2 = bar_top_projects(top_for_chart2, 'amount_rub', 'project_name', 'Топ-10 по сумме (млн руб.)')
    st.plotly_chart(fig2, use_container_width=True)

with col_table2:
    top_amount_disp = top_amount.copy()
    top_amount_disp['Сумма (руб.)'] = top_amount_disp['Сумма (руб.)'].apply(lambda x: f"{x:,.0f}")
    st.dataframe(top_amount_disp, use_container_width=True, hide_index=True)

st.divider()

# ── Полная таблица проектов ───────────────────────────────────────────────────
st.subheader("Все проекты")

all_projects = proj_stats.sort_values('hours', ascending=False).copy()
all_projects.columns = ['Проект', 'Клиент', 'Тип', 'Часов', 'Сотрудников', 'Сумма (руб.)']
all_projects['Сумма (руб.)'] = all_projects['Сумма (руб.)'].apply(lambda x: f"{x:,.0f}")

st.dataframe(all_projects, use_container_width=True, hide_index=True, height=400)

# Скачать
import io
buf = io.BytesIO()
proj_stats.to_excel(buf, index=False, sheet_name='Проекты')
buf.seek(0)
st.download_button(
    "⬇️ Скачать в Excel",
    buf,
    file_name="проекты.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
)
