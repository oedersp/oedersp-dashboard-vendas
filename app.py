import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(page_title="Dashboard Executivo", layout="wide")

try:
    arquivo = "Rela_televendas.xlsx"
    vendas = pd.read_excel(arquivo, sheet_name="racional de vendas")
    atendimento = pd.read_excel(arquivo, sheet_name="racional de atendimento")
except Exception as e:
    st.error(f"Erro ao carregar arquivo: {e}")
    st.stop()

vendas.columns = vendas.columns.str.strip()
atendimento.columns = atendimento.columns.str.strip()

vendas["MESES"] = pd.to_datetime(vendas["MESES"], errors='coerce')
vendas["MES"] = vendas["MESES"].dt.strftime("%b/%Y")
vendas["META DE VENDAS"] = vendas["META DE VENDAS"].ffill()

st.title("📊 Dashboard Executivo — Victor Hugo")

col1, col2, col3, col4, col5 = st.columns(5)
with col1:
    st.metric("Total Vendido", f"R$ {vendas['R$ VENDAS'].sum():,.0f}")
with col2:
    st.metric("Peças", f"{vendas['PEÇAS VENDIDAS'].sum():,.0f}")
with col3:
    st.metric("Convertidos", f"{vendas['CLIENTES CONVERTIDOS'].sum():,.0f}")
with col4:
    st.metric("P.A", f"{vendas['P.A'].mean():.2f}")
with col5:
    st.metric("Meta", f"R$ {vendas['META DE VENDAS'].sum():,.0f}")

st.divider()

st.subheader("Evolução de Vendas")
fig = px.line(vendas, x="MES", y="R$ VENDAS", markers=True)
st.plotly_chart(fig, use_container_width=True)

st.success("✅ Dashboard carregado com sucesso!")
