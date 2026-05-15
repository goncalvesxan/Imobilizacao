import streamlit as st
import pandas as pd
from core.importador import padronizar_cji3
from core.validador import validar_base_cji3

st.set_page_config(page_title="Upload das Bases", layout="wide")
st.title("01 - Upload das Bases")

st.markdown("Carregue inicialmente a base **CJI3**. As demais bases serão usadas nas próximas versões para enriquecimento e validação.")

arquivo_cji3 = st.file_uploader("Base CJI3 Excel", type=["xlsx", "xlsm", "xls"])

if arquivo_cji3:
    xls = pd.ExcelFile(arquivo_cji3)
    aba = st.selectbox("Selecione a aba da CJI3", xls.sheet_names)
    df_raw = pd.read_excel(arquivo_cji3, sheet_name=aba)
    df_cji3, log_colunas = padronizar_cji3(df_raw)
    log_validacao = validar_base_cji3(df_cji3)

    st.session_state["cji3_raw"] = df_raw
    st.session_state["cji3"] = df_cji3
    st.session_state["log_colunas"] = log_colunas
    st.session_state["log_validacao"] = log_validacao

    st.success(f"Base carregada: {len(df_cji3):,} linhas")
    c1, c2 = st.columns(2)
    with c1:
        st.subheader("Mapeamento de colunas")
        st.dataframe(log_colunas, use_container_width=True)
    with c2:
        st.subheader("Validações iniciais")
        if log_validacao.empty:
            st.success("Nenhuma inconsistência crítica encontrada nas colunas mínimas.")
        else:
            st.dataframe(log_validacao, use_container_width=True)

    st.subheader("Prévia da base padronizada")
    st.dataframe(df_cji3.head(100), use_container_width=True)
else:
    st.warning("Aguardando upload da base CJI3.")
