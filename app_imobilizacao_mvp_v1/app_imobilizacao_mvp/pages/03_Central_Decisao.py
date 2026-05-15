import streamlit as st
import pandas as pd
from core.regras import carregar_regras, sugerir_decisoes
from core.schema import TIPOS_DECISAO, TRATAMENTOS, CRITERIOS_RATEIO

st.set_page_config(page_title="Central de Decisão", layout="wide")
st.title("03 - Central de Decisão")

if "cji3" not in st.session_state:
    st.warning("Carregue a base CJI3 primeiro na etapa 01.")
    st.stop()

cji3 = st.session_state["cji3"]
regras = st.session_state.get("regras", carregar_regras())

if st.button("Gerar sugestões automáticas") or "decisoes" not in st.session_state:
    st.session_state["decisoes"] = sugerir_decisoes(cji3, regras)

decisoes = st.session_state["decisoes"]

st.markdown("Revise as decisões sugeridas. Esta tela é onde o conhecimento operacional começa a virar parâmetro.")

cols_chave = [c for c in ["ID_Linha", "Empresa", "Centro", "PEP", "Descricao", "Classe Custo", "Quantidade", "Valor", "Fornecedor Nome", "Pedido", "Nota Fiscal", "Tipo Decisão", "Tratamento", "Critério Rateio", "Regra Aplicada"] if c in decisoes.columns]

editado = st.data_editor(
    decisoes[cols_chave],
    use_container_width=True,
    height=620,
    num_rows="fixed",
    column_config={
        "Tipo Decisão": st.column_config.SelectboxColumn("Tipo Decisão", options=TIPOS_DECISAO),
        "Tratamento": st.column_config.SelectboxColumn("Tratamento", options=TRATAMENTOS),
        "Critério Rateio": st.column_config.SelectboxColumn("Critério Rateio", options=CRITERIOS_RATEIO),
    }
)

if st.button("Confirmar decisões"):
    base = decisoes.drop(columns=[c for c in ["Tipo Decisão", "Tratamento", "Critério Rateio", "Regra Aplicada"] if c in decisoes.columns])
    confirmado = base.merge(editado[["ID_Linha", "Tipo Decisão", "Tratamento", "Critério Rateio", "Regra Aplicada"]], on="ID_Linha", how="left")
    st.session_state["decisoes_confirmadas"] = confirmado
    st.success("Decisões confirmadas para processamento.")

st.subheader("Resumo das decisões")
st.dataframe(editado.groupby(["Tipo Decisão", "Tratamento"], dropna=False).size().reset_index(name="Qtd Linhas"), use_container_width=True)
