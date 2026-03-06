"""
Точка входа. Настройка навигации через st.navigation.
"""
import streamlit as st
import sys
import pathlib

sys.path.insert(0, str(pathlib.Path(__file__).parent))

st.set_page_config(
    page_title="Аналитика бюро",
    page_icon="⚖️",
    layout="wide",
    initial_sidebar_state="expanded",
)

pg = st.navigation([
    st.Page("pages/0_Главная.py",       title="🏠 Главная",    default=True),
    st.Page("pages/1_📊_Обзор.py",      title="📊 Обзор"),
    st.Page("pages/2_👥_Сотрудники.py", title="👥 Сотрудники"),
    st.Page("pages/3_📈_Динамика.py",   title="📈 Динамика"),
    st.Page("pages/4_📁_Проекты.py",    title="📁 Проекты"),
    st.Page("pages/5_💰_ФОТ.py",        title="💰 ФОТ"),
])
pg.run()
