import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide")

st.title("Dashboard")

# Wczytaj Excel
df = pd.read_excel("data_man/dane.xlsx")

df["Date"] = pd.to_datetime(df["Date"]).dt.date

col1, col2 = st.columns(2)

# Wykres 1 - lewa kolumna
with col1:
    st.subheader("Weight")
    fig1 = px.line(df, x="Date", y="Weight")  # <-- ZMIEŃ KOLUMNY
    fig1.update_xaxes(tickformat="%Y-%m-%d")
    st.plotly_chart(fig1, width="stretch")

# Wykres 2 - prawa kolumna
with col2:
    st.subheader("Volume")
    fig2 = px.bar(df, x="Date", y="Volumen")  # <-- ZMIEŃ KOLUMNY
    fig2.update_xaxes(tickformat="%Y-%m-%d")
    st.plotly_chart(fig2, width="stretch")

# Tabela pod wykresami (pełna szerokość)
st.dataframe(df, width="stretch")