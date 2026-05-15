import streamlit as st

st.set_page_config(page_title="Motor de Imobilização SAP", layout="wide")

st.title("Motor de Imobilização e Rateio SAP - MVP")
st.markdown("""
Esta aplicação piloto substitui a lógica pesada do Excel por um fluxo parametrizável em Python.

**Fluxo:** upload das bases → validação → central de decisão → rateio → template SAP AS01.
""")

st.info("Use o menu lateral para seguir as etapas. Comece em **01 Upload Bases**.")

st.subheader("Princípio do sistema")
st.write("A CJI3 é a matéria-prima. O produto final é o template SAP e os relatórios de decisão/rateio.")

st.code("""
CJI3 + Chamados + Classes + Vida Útil + Localização
        ↓
Validação e padronização
        ↓
Central de decisão de ativos
        ↓
Motor de rateio
        ↓
Template SAP AS01 + Logs
""")
