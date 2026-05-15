import streamlit as st
import pandas as pd

st.set_page_config(page_title="Dashboard", layout="wide")
st.title("06 - Dashboard Operacional")

if "processado" not in st.session_state:
    st.warning("Execute o processamento na etapa 04.")
    st.stop()

df = st.session_state["processado"]

c1, c2, c3, c4 = st.columns(4)
c1.metric("Linhas Originais", len(st.session_state.get("cji3", [])))
c2.metric("Linhas Geradas", len(df))
c3.metric("Valor Total", f"{pd.to_numeric(df.get('Valor Capitalizado', 0), errors='coerce').sum():,.2f}")
c4.metric("Empresas", df.get("Empresa", pd.Series(dtype=str)).nunique())

st.subheader("Valor por tratamento")
if "Tratamento" in df.columns:
    resumo = df.groupby("Tratamento", dropna=False)["Valor Capitalizado"].sum().reset_index()
    st.bar_chart(resumo, x="Tratamento", y="Valor Capitalizado")
    st.dataframe(resumo, use_container_width=True)

st.subheader("Pendências")
pend = df[df.get("Status", "").astype(str).str.contains("Pendência|Aguardando", case=False, na=False)]
st.dataframe(pend, use_container_width=True)
