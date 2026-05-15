import streamlit as st
import pandas as pd
from pathlib import Path
from core.regras import carregar_regras

st.set_page_config(page_title="Regras Parametrizadas", layout="wide")
st.title("02 - Regras Parametrizadas")

st.markdown("Aqui ficam as regras iniciais de sugestão. No MVP, elas podem ser editadas na tela e salvas em CSV.")

regras = carregar_regras()

editado = st.data_editor(
    regras,
    use_container_width=True,
    num_rows="dynamic",
    column_config={
        "prioridade": st.column_config.NumberColumn("Prioridade", min_value=1, step=1),
        "operador": st.column_config.SelectboxColumn("Operador", options=["contem", "igual", "comeca_com"]),
        "tipo_sugerido": st.column_config.SelectboxColumn("Tipo sugerido", options=["Produto", "Servico", "Componente", "Principal", "Complemento Ativo Existente", "Nao Capitalizavel", "Pendente"]),
        "tratamento_sugerido": st.column_config.TextColumn("Tratamento sugerido"),
        "criterio_rateio": st.column_config.TextColumn("Critério rateio"),
    }
)

if st.button("Salvar regras parametrizadas"):
    Path("config").mkdir(exist_ok=True)
    editado.to_csv("config/regras_padrao.csv", index=False)
    st.session_state["regras"] = editado
    st.success("Regras salvas em config/regras_padrao.csv")

st.session_state["regras"] = editado
