"""
Функции построения графиков через Plotly.
Все подписи на русском, цветовая схема — синяя гамма.
"""

import plotly.graph_objects as go
import plotly.express as px
import pandas as pd
import streamlit as st


# Цветовая палитра
COLORS = {
    'primary':   '#1F4E79',
    'secondary': '#2E75B6',
    'accent':    '#70AD47',
    'warning':   '#FF9900',
    'danger':    '#C00000',
    'light':     '#BDD7EE',
    'sequence':  px.colors.sequential.Blues[3:],
}

COLOR_SEQUENCE = [
    '#1F4E79', '#2E75B6', '#4472C4', '#70AD47',
    '#ED7D31', '#A9D18E', '#9DC3E6', '#FF9900',
]


def metric_card(label: str, value: str, delta: str = None, delta_positive: bool = True):
    """Карточка KPI через st.metric."""
    st.metric(
        label=label,
        value=value,
        delta=delta,
        delta_color="normal" if delta_positive else "inverse"
    )


def line_chart_hours_realization(df_grouped: pd.DataFrame) -> go.Figure:
    """
    Линейный график: часы и реализация по периодам (две оси Y).
    df_grouped должен содержать: period, total_hours, billable_hours, realization_pct
    """
    fig = go.Figure()

    fig.add_trace(go.Scatter(
        x=df_grouped['period'],
        y=df_grouped['total_hours'],
        name='Всего часов',
        mode='lines+markers',
        line=dict(color=COLORS['primary'], width=2),
        marker=dict(size=6),
        yaxis='y1',
    ))

    fig.add_trace(go.Scatter(
        x=df_grouped['period'],
        y=df_grouped['billable_hours'],
        name='К оплате',
        mode='lines+markers',
        line=dict(color=COLORS['secondary'], width=2, dash='dash'),
        marker=dict(size=6),
        yaxis='y1',
    ))

    fig.add_trace(go.Scatter(
        x=df_grouped['period'],
        y=df_grouped['realization_pct'],
        name='Реализация %',
        mode='lines+markers',
        line=dict(color=COLORS['accent'], width=2),
        marker=dict(size=6),
        yaxis='y2',
    ))

    fig.update_layout(
        title='Динамика часов и реализации',
        xaxis=dict(title='Период'),
        yaxis=dict(title='Часы', side='left'),
        yaxis2=dict(
            title='Реализация %',
            side='right',
            overlaying='y',
            showgrid=False,
            range=[0, 120],
        ),
        legend=dict(x=0, y=1.1, orientation='h'),
        hovermode='x unified',
        height=400,
        margin=dict(l=0, r=0, t=40, b=0),
        plot_bgcolor='white',
        paper_bgcolor='white',
    )
    fig.update_xaxes(showgrid=True, gridcolor='#f0f0f0')
    fig.update_yaxes(showgrid=True, gridcolor='#f0f0f0')

    return fig


def bar_chart_project_types(df: pd.DataFrame) -> go.Figure:
    """Столбчатая диаграмма: распределение часов по типам проектов."""
    grouped = df.groupby('project_type', dropna=False)['duration'].sum().reset_index()
    grouped = grouped.sort_values('duration', ascending=False)
    grouped['project_type'] = grouped['project_type'].fillna('Не указан')

    fig = px.bar(
        grouped,
        x='project_type',
        y='duration',
        title='Распределение часов по типам проектов',
        labels={'project_type': 'Тип проекта', 'duration': 'Часов'},
        color='project_type',
        color_discrete_sequence=COLOR_SEQUENCE,
    )
    fig.update_layout(
        showlegend=False,
        height=350,
        margin=dict(l=0, r=0, t=40, b=0),
        plot_bgcolor='white',
        paper_bgcolor='white',
    )
    return fig


def horizontal_bar_realization(df_employees: pd.DataFrame,
                                threshold_low: float = 70,
                                threshold_high: float = 85) -> go.Figure:
    """
    Горизонтальная столбчатая диаграмма: реализация по сотрудникам.
    Цвет: красный < threshold_low, зелёный > threshold_high, синий — остальное.
    """
    df = df_employees.copy().sort_values('realization_pct', ascending=True)

    colors = []
    for v in df['realization_pct']:
        if pd.isna(v):
            colors.append(COLORS['light'])
        elif v < threshold_low:
            colors.append(COLORS['danger'])
        elif v > threshold_high:
            colors.append(COLORS['accent'])
        else:
            colors.append(COLORS['secondary'])

    fig = go.Figure(go.Bar(
        x=df['realization_pct'],
        y=df['employee'],
        orientation='h',
        marker_color=colors,
        text=df['realization_pct'].apply(lambda x: f'{x:.1f}%' if pd.notna(x) else '—'),
        textposition='outside',
    ))

    # Линии порогов
    fig.add_vline(x=threshold_low, line_dash='dash', line_color=COLORS['danger'],
                  annotation_text=f'{threshold_low}%', annotation_position='top')
    fig.add_vline(x=threshold_high, line_dash='dash', line_color=COLORS['accent'],
                  annotation_text=f'{threshold_high}%', annotation_position='top')

    fig.update_layout(
        title='Реализация по сотрудникам (%)',
        xaxis=dict(title='Реализация %', range=[0, max(110, df['realization_pct'].max() * 1.1)]),
        yaxis=dict(title=''),
        height=max(300, len(df) * 25 + 80),
        margin=dict(l=0, r=60, t=40, b=0),
        plot_bgcolor='white',
        paper_bgcolor='white',
    )
    return fig


def line_chart_dynamic(df_grouped: pd.DataFrame,
                        y_col: str,
                        y_label: str,
                        color_col: str = None,
                        title: str = '') -> go.Figure:
    """
    Универсальный линейный график для страницы «Динамика».
    df_grouped: period, [color_col], y_col
    """
    if color_col and color_col in df_grouped.columns:
        fig = px.line(
            df_grouped, x='period', y=y_col, color=color_col,
            title=title,
            labels={'period': 'Период', y_col: y_label, color_col: 'Сотрудник/группа'},
            markers=True,
            color_discrete_sequence=COLOR_SEQUENCE,
        )
    else:
        fig = px.line(
            df_grouped, x='period', y=y_col,
            title=title,
            labels={'period': 'Период', y_col: y_label},
            markers=True,
            color_discrete_sequence=[COLORS['primary']],
        )

    fig.update_layout(
        height=420,
        hovermode='x unified',
        margin=dict(l=0, r=0, t=40, b=0),
        plot_bgcolor='white',
        paper_bgcolor='white',
        legend=dict(x=1.01, y=1),
    )
    fig.update_xaxes(showgrid=True, gridcolor='#f0f0f0')
    fig.update_yaxes(showgrid=True, gridcolor='#f0f0f0')
    return fig


def bar_top_projects(df_top: pd.DataFrame,
                      x_col: str,
                      y_col: str,
                      title: str) -> go.Figure:
    """Горизонтальная столбчатая диаграмма топ-проектов."""
    fig = px.bar(
        df_top,
        x=x_col,
        y=y_col,
        orientation='h',
        title=title,
        color_discrete_sequence=[COLORS['secondary']],
        text=x_col,
    )
    fig.update_traces(textposition='outside', texttemplate='%{text:.1f}')
    fig.update_layout(
        yaxis=dict(autorange='reversed'),
        height=max(300, len(df_top) * 30 + 80),
        margin=dict(l=0, r=60, t=40, b=0),
        plot_bgcolor='white',
        paper_bgcolor='white',
        showlegend=False,
    )
    return fig
