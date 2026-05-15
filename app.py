import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

st.set_page_config(
    page_title="Dashboard Executivo Victor Hugo",
    layout="wide"
)

# =====================================================
# CARREGAMENTO
# =====================================================

arquivo = "Rela_televendas.xlsx"

vendas = pd.read_excel(arquivo, sheet_name="racional de vendas")
atendimento = pd.read_excel(arquivo, sheet_name="racional de atendimento")
horarios = pd.read_excel(arquivo, sheet_name="quantidade por horário")

# =====================================================
# LIMPEZA DOS DADOS
# =====================================================

vendas.columns = vendas.columns.str.strip()
atendimento.columns = atendimento.columns.str.strip()

vendas["MESES"] = pd.to_datetime(vendas["MESES"])
vendas["MES"] = vendas["MESES"].dt.strftime("%b/%Y")

vendas["META DE VENDAS"] = vendas["META DE VENDAS"].fillna(method="ffill")

atendimento["Meses"] = pd.to_datetime(atendimento["Meses"])
atendimento["Mes"] = atendimento["Meses"].dt.strftime("%b/%Y")

atendimento["VENDEDORA"] = atendimento["VENDEDORA"].fillna(method="ffill")

# =====================================================
# SIDEBAR
# =====================================================

st.sidebar.title("Filtros")

vendedoras = vendas["VENDODORA"].unique()

filtro_vendedora = st.sidebar.multiselect(
    "Selecione a Vendedora",
    vendedoras,
    default=vendedoras
)

vendas_filtrado = vendas[
    vendas["VENDODORA"].isin(filtro_vendedora)
]

# =====================================================
# KPIS
# =====================================================

st.title("Dashboard Executivo — Victor Hugo")

col1, col2, col3, col4, col5 = st.columns(5)

valor_total = vendas_filtrado["R$ VENDAS"].sum()
pecas = vendas_filtrado["PEÇAS VENDIDAS"].sum()
convertidos = vendas_filtrado["CLIENTES CONVERTIDOS"].sum()
pa_medio = vendas_filtrado["P.A"].mean()
meta_total = vendas_filtrado["META DE VENDAS"].sum()

with col1:
    st.metric("Total Vendido", f"R$ {valor_total:,.0f}")

with col2:
    st.metric("Peças Vendidas", f"{pecas:,.0f}")

with col3:
    st.metric("Clientes Convertidos", f"{convertidos:,.0f}")

with col4:
    st.metric("P.A Médio", f"{pa_medio:.2f}")

with col5:
    st.metric("Meta Total", f"R$ {meta_total:,.0f}")

st.divider()

# =====================================================
# 1 - EVOLUÇÃO MENSAL POR VENDEDORA
# =====================================================

fig1 = px.line(
    vendas_filtrado,
    x="MES",
    y="R$ VENDAS",
    color="VENDODORA",
    markers=True,
    text="R$ VENDAS",
    title="Evolução Mensal de Vendas por Vendedora"
)

fig1.update_traces(textposition="top center")
fig1.update_layout(
    height=500,
    font=dict(size=14)
)

st.plotly_chart(fig1, use_container_width=True)

# =====================================================
# 2 - PEÇAS X CONVERTIDOS X PA
# =====================================================

fig2 = go.Figure()

fig2.add_trace(go.Bar(
    x=vendas_filtrado["MES"],
    y=vendas_filtrado["PEÇAS VENDIDAS"],
    name="Peças Vendidas"
))

fig2.add_trace(go.Bar(
    x=vendas_filtrado["MES"],
    y=vendas_filtrado["CLIENTES CONVERTIDOS"],
    name="Clientes Convertidos"
))

fig2.add_trace(go.Scatter(
    x=vendas_filtrado["MES"],
    y=vendas_filtrado["P.A"],
    name="P.A",
    mode="lines+markers"
))

fig2.update_layout(
    barmode="group",
    title="Peças Vendidas x Convertidos x P.A",
    height=500,
    font=dict(size=14)
)

st.plotly_chart(fig2, use_container_width=True)

# =====================================================
# 3 - VENDAS VS META POR VENDEDORA
# =====================================================

vendedora_meta = vendas_filtrado.groupby("VENDODORA", as_index=False).agg({
    "R$ VENDAS": "sum",
    "META DE VENDAS": "sum"
})

fig3 = go.Figure()

fig3.add_trace(go.Bar(
    x=vendedora_meta["VENDODORA"],
    y=vendedora_meta["R$ VENDAS"],
    name="Vendas"
))

fig3.add_trace(go.Bar(
    x=vendedora_meta["VENDODORA"],
    y=vendedora_meta["META DE VENDAS"],
    name="Meta"
))

fig3.update_layout(
    title="Vendas x Meta por Vendedora",
    barmode="group",
    height=500,
    font=dict(size=14)
)

st.plotly_chart(fig3, use_container_width=True)

# =====================================================
# 4 - COMPARATIVO COM LINHA DE TENDÊNCIA
# =====================================================

fig4 = px.scatter(
    vendas_filtrado,
    x="META DE VENDAS",
    y="R$ VENDAS",
    trendline="ols",
    color="VENDODORA",
    title="Comparativo Vendas vs Meta"
)

fig4.update_layout(
    height=500,
    font=dict(size=14)
)

st.plotly_chart(fig4, use_container_width=True)

# =====================================================
# 5 - HEATMAP
# =====================================================

st.subheader("Mapa de Calor — Quantidade por Horário")

horarios = horarios.fillna(0)

horarios.columns = [
    "Mes",
    "Dia",
    *list(range(24))
]

heatmap_df = horarios.drop(columns=["Mes"])
heatmap_df = heatmap_df.groupby("Dia").sum()

fig5 = px.imshow(
    heatmap_df,
    color_continuous_scale="Reds",
    aspect="auto",
    title="Quantidade por Horário e Dia"
)

fig5.update_layout(
    height=500,
    font=dict(size=14)
)

st.plotly_chart(fig5, use_container_width=True)

# =====================================================
# 6 - TABELA EXECUTIVA POR VENDEDOR
# =====================================================

st.subheader("Resumo Executivo por Vendedora")

resumo = vendas_filtrado.groupby("VENDODORA", as_index=False).agg({
    "R$ VENDAS": "sum",
    "PEÇAS VENDIDAS": "sum",
    "CLIENTES CONVERTIDOS": "sum",
    "P.A": "mean"
})

st.dataframe(resumo, use_container_width=True)

# =====================================================
# 7 - KPIS DE ATENDIMENTO
# =====================================================

st.subheader("Indicadores de Atendimento")

col1, col2, col3, col4 = st.columns(4)

with col1:
    st.metric(
        "Contatos",
        f"{atendimento['CONTATOS'].sum():,.0f}"
    )

with col2:
    st.metric(
        "Interações",
        f"{atendimento['INTERAÇÕES'].sum():,.0f}"
    )

with col3:
    st.metric(
        "Tempo Médio 1ª Resposta",
        atendimento['TEMPOR DE 1º RESPOSTA'].dropna().iloc[0]
    )

with col4:
    st.metric(
        "TMA Médio",
        atendimento['TMA'].dropna().iloc[0]
    )

# =====================================================
# 8 - EVOLUÇÃO DE ATENDIMENTO
# =====================================================

fig6 = px.line(
    atendimento,
    x="Mes",
    y="CONTATOS",
    color="VENDEDORA",
    markers=True,
    title="Evolução Mensal de Atendimento"
)

fig6.update_layout(
    height=500,
    font=dict(size=14)
)

st.plotly_chart(fig6, use_container_width=True)

# =====================================================
# 9 - VENDAS VS META / PEÇAS VS CONVERTIDOS
# =====================================================

fig7 = make_subplots(specs=[[{"secondary_y": True}]])

fig7.add_trace(
    go.Bar(
        x=vendas_filtrado["MES"],
        y=vendas_filtrado["R$ VENDAS"],
        name="Vendas"
    ),
    secondary_y=False
)

fig7.add_trace(
    go.Scatter(
        x=vendas_filtrado["MES"],
        y=vendas_filtrado["META DE VENDAS"],
        name="Meta",
        mode="lines+markers"
    ),
    secondary_y=False
)

fig7.add_trace(
    go.Scatter(
        x=vendas_filtrado["MES"],
        y=vendas_filtrado["PEÇAS VENDIDAS"],
        name="Peças Vendidas",
        mode="lines+markers"
    ),
    secondary_y=True
)

fig7.add_trace(
    go.Scatter(
        x=vendas_filtrado["MES"],
        y=vendas_filtrado["CLIENTES CONVERTIDOS"],
        name="Convertidos",
        mode="lines+markers"
    ),
    secondary_y=True
)

fig7.update_layout(
    title="Vendas vs Meta | Peças Vendidas vs Convertidos",
    height=600,
    font=dict(size=14)
)

st.plotly_chart(fig7, use_container_width=True)
