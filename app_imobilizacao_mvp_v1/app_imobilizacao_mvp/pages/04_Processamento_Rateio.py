import streamlit as st
from core.rateio import processar_rateio

st.set_page_config(page_title="Processamento e Rateio", layout="wide")
st.title("04 - Processamento e Rateio")

if "decisoes_confirmadas" not in st.session_state:
    st.warning("Confirme as decisões na etapa 03 antes de processar.")
    st.stop()

if st.button("Processar rateio e formação de ativos"):
    processado = processar_rateio(st.session_state["decisoes_confirmadas"])
    st.session_state["processado"] = processado
    st.success(f"Processamento concluído: {len(processado):,} linhas geradas.")

if "processado" in st.session_state:
    df = st.session_state["processado"]
    c1, c2, c3 = st.columns(3)
    c1.metric("Linhas geradas", len(df))
    c2.metric("Valor capitalizado", f"{df.get('Valor Capitalizado', 0).sum():,.2f}")
    c3.metric("Pendências", int((df.get("Status", "") == "Pendência de análise").sum()))

    st.subheader("Resultado do processamento")
    st.dataframe(df, use_container_width=True, height=620)
else:
    st.info("Clique em processar para gerar o resultado.")
