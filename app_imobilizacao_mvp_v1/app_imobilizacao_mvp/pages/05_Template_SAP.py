import streamlit as st
import pandas as pd
from core.exportador_sap import montar_template_sap
from core.utils import excel_bytes_abas

st.set_page_config(page_title="Template SAP", layout="wide")
st.title("05 - Template SAP AS01")

if "processado" not in st.session_state:
    st.warning("Execute o processamento na etapa 04.")
    st.stop()

template = montar_template_sap(st.session_state["processado"])
st.session_state["template_sap"] = template

st.subheader("Prévia do Template SAP")
st.dataframe(template, use_container_width=True, height=620)

pendencias = template[template["Status"].astype(str).str.contains("Pendência|Aguardando", case=False, na=False)].copy()
logs = st.session_state.get("log_validacao", pd.DataFrame())

arquivo = excel_bytes_abas({
    "Template_AS01": template,
    "Pendencias": pendencias,
    "Log_Validacoes": logs,
    "Base_Processada": st.session_state["processado"]
})

st.download_button(
    "Baixar pacote Excel final",
    data=arquivo,
    file_name="Template_AS01_MVP.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
)
