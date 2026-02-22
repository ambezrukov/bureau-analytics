"""
Страница «Динамика» — конструктор произвольных графиков.
"""

import sys
import pathlib
sys.path.insert(0, str(pathlib.Path(__file__).parent.parent))

import streamlit as st
import pandas as pd
from utils.data import render_sidebar, get_filtered_data, get_periods
from utils.charts import line_chart_dynamic

st.set_page_config(page_title="Динамика | Аналитика бюро", page_icon="📈", layout="wide")

filters = render_sidebar(key_prefix='dyn_')

st.title("📈 Динамика")

if not filters['periods']:
    st.warning("Выберите хотя бы один период в боковой панели.")
    st.stop()

# Настройки графика
col1, col2, col3 = st.columns(3)

with col1:
    y_metric = st.selectbox(
        "Показатель (ось Y)",
        [
            ("Всего часов", "total_hours"),
            ("Коммерческие часы", "commercial_hours"),
            ("К оплате (часы)", "billable_hours"),
            ("Реализация %", "realization_pct"),
            ("Сумма счетов (руб.)", "amount_rub"),
        ],
        format_func=lambda x: x[0],
    )

with col2:
    breakdown = st.selectbox(
        "Разрез",
        ["Бюро (итог)", "По сотрудникам", "По типам проектов"],
    )

with col3:
    show_all_periods = st.checkbox("Все периоды (игнорировать выбор)", value=False)

st.divider()

# Загружаем данные
filters_wide = filters.copy()
if show_all_periods:
    filters_wide['periods'] = get_periods()

df = get_filtered_data(filters_wide)

if df.empty:
    st.warning("Нет данных для выбранных фильтров.")
    st.stop()

all_periods = get_periods()
USD_RATE = 88.0
EUR_RATE = 95.0


def compute_metrics(data: pd.DataFrame, group_cols: list) -> pd.DataFrame:
    """Вычисляет метрики по группировке."""
    total = data.groupby(group_cols)['duration'].sum().rename('total_hours')
    commercial = (
        data[data['project_type'] == 'Работа по договорам']
        .groupby(group_cols)['duration'].sum()
        .rename('commercial_hours')
    )
    billable = data.groupby(group_cols)['billable_hours'].sum().rename('billable_hours')

    bills = data[data['include_in_bill'] == 1]
    amount_rub = (
        bills[bills['currency'] == 'RUB'].groupby(group_cols)['amount'].sum() +
        bills[bills['currency'] == 'USD'].groupby(group_cols)['amount'].sum() * USD_RATE +
        bills[bills['currency'] == 'EUR'].groupby(group_cols)['amount'].sum() * EUR_RATE
    ).rename('amount_rub')

    result = pd.concat([total, commercial, billable, amount_rub], axis=1).reset_index()

    # Реализация
    result['realization_pct'] = (
        result.get('billable_hours', pd.Series(dtype=float)) /
        result.get('commercial_hours', pd.Series(dtype=float)).replace(0, float('nan')) * 100
    ).round(1)

    # Сортируем по периодам
    result['_order'] = result['period'].apply(
        lambda p: all_periods.index(p) if p in all_periods else 0
    )
    result = result.sort_values('_order').drop(columns='_order')
    return result


metric_col, metric_label = y_metric

if breakdown == "Бюро (итог)":
    df_grouped = compute_metrics(df, ['period'])
    color_col = None
    title = f"{metric_label} — бюро"

elif breakdown == "По сотрудникам":
    # Ограничиваем до топ-10 по метрике
    top_employees = df.groupby('employee')['duration'].sum().nlargest(10).index.tolist()
    df_top = df[df['employee'].isin(top_employees)]
    df_grouped = compute_metrics(df_top, ['period', 'employee'])
    color_col = 'employee'
    title = f"{metric_label} — по сотрудникам (топ-10)"

else:  # По типам проектов
    df_grouped = compute_metrics(df, ['period', 'project_type'])
    df_grouped['project_type'] = df_grouped['project_type'].fillna('Не указан')
    color_col = 'project_type'
    title = f"{metric_label} — по типам проектов"

# Строим график
if metric_col in df_grouped.columns:
    fig = line_chart_dynamic(df_grouped, metric_col, metric_label, color_col, title)
    st.plotly_chart(fig, use_container_width=True)

    # Таблица под графиком
    with st.expander("Данные таблицы"):
        display_cols = ['period']
        if color_col:
            display_cols.append(color_col)
        display_cols.append(metric_col)
        st.dataframe(
            df_grouped[display_cols].rename(columns={
                'period': 'Период',
                'employee': 'Сотрудник',
                'project_type': 'Тип проекта',
                metric_col: metric_label,
            }),
            use_container_width=True,
            hide_index=True,
        )
else:
    st.warning(f"Метрика '{metric_col}' недоступна для текущего разреза.")
