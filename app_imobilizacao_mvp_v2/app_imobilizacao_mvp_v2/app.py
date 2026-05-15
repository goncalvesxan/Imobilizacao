from __future__ import annotations
import pandas as pd
import streamlit as st
from core.loaders import (
    read_excel_any, standardize_cji3, standardize_classes, standardize_localizacao,
    standardize_empresas, standardize_chamados, standardize_razao, standardize_vida_util
)
from core.enrichment import enrich_cji3, build_prefill, decision_to_template
from core.schema import TRATAMENTOS, CRITERIOS_RATEIO
from core.utils import excel_bytes, norm_text, to_number

st.set_page_config(page_title="Motor de Imobilização SAP - MVP V2", layout="wide")

st.markdown("""
<style>
.block-container {padding-top: 1.2rem;}
.asset-card {border:1px solid #9E9E9E; padding:14px; border-radius:6px; background:#FAFAFA;}
.small-label {font-size: 0.82rem; color:#333; margin-bottom: -12px;}
.stTabs [data-baseweb="tab-list"] {gap: 8px;}
.stTabs [data-baseweb="tab"] {border:1px solid #BDBDBD; border-radius:6px 6px 0 0; padding:8px 14px;}
</style>
""", unsafe_allow_html=True)

st.title("Cadastro de Ativo - Motor de Imobilização SAP")
st.caption("MVP V2: CJI3 + tabelas auxiliares + formulário de decisão no formato operacional do Excel.")

for k, default in {
    "bases": {}, "logs": {}, "cji3": pd.DataFrame(), "cji3_enriquecida": pd.DataFrame(), "decisoes": []
}.items():
    if k not in st.session_state:
        st.session_state[k] = default

tab_upload, tab_cadastro, tab_template = st.tabs(["1. Carregamento das Bases", "2. Cadastro / Decisão", "3. Template AS01 / Exportação"])

with tab_upload:
    st.subheader("Carregar bases necessárias")
    st.write("Carregue a CJI3 e as tabelas auxiliares. Cada base carregada será padronizada para alimentar o formulário e o template AS01.")
    c1, c2, c3 = st.columns(3)
    with c1:
        arq_cji3 = st.file_uploader("Base CJI3", type=["xlsx", "xlsm"], key="up_cji3")
        arq_chamados = st.file_uploader("Chamados / Status", type=["xlsx", "xlsm"], key="up_chamados")
    with c2:
        arq_classes = st.file_uploader("Tabela Classes", type=["xlsx", "xlsm"], key="up_classes")
        arq_vida = st.file_uploader("Tabela Vida Útil", type=["xlsx", "xlsm"], key="up_vida")
    with c3:
        arq_loc = st.file_uploader("Tabela Empresa x Centro x Local", type=["xlsx", "xlsm"], key="up_loc")
        arq_emp = st.file_uploader("Tabela Empresas", type=["xlsx", "xlsm"], key="up_emp")
        arq_razao = st.file_uploader("Razão / Contas Imob. Andamento", type=["xlsx", "xlsm"], key="up_razao")

    if st.button("Processar bases carregadas", type="primary"):
        bases, logs = {}, {}
        try:
            if arq_cji3:
                raw = read_excel_any(arq_cji3)
                cji3, log = standardize_cji3(raw)
                bases["cji3"] = cji3
                logs["CJI3"] = log
            if arq_chamados:
                bases["chamados"] = standardize_chamados(read_excel_any(arq_chamados))
            if arq_classes:
                bases["classes"] = standardize_classes(read_excel_any(arq_classes))
            if arq_vida:
                bases["vida_util"] = standardize_vida_util(arq_vida)
            if arq_loc:
                bases["localizacao"] = standardize_localizacao(read_excel_any(arq_loc))
            if arq_emp:
                bases["empresas"] = standardize_empresas(read_excel_any(arq_emp))
            if arq_razao:
                bases["razao"] = standardize_razao(read_excel_any(arq_razao))

            if "cji3" in bases:
                st.session_state.bases = bases
                st.session_state.logs = logs
                st.session_state.cji3 = bases["cji3"]
                st.session_state.cji3_enriquecida = enrich_cji3(bases["cji3"], bases)
                st.success("Bases processadas com sucesso.")
            else:
                st.warning("Carregue ao menos a CJI3 para iniciar o cadastro.")
        except Exception as e:
            st.error(f"Erro ao processar bases: {e}")

    st.divider()
    st.subheader("Diagnóstico das bases em memória")
    if st.session_state.bases:
        resumo = [{"Base": nome, "Linhas": len(df), "Colunas": len(df.columns)} for nome, df in st.session_state.bases.items()]
        st.dataframe(pd.DataFrame(resumo), use_container_width=True, hide_index=True)
        if "CJI3" in st.session_state.logs:
            st.write("Mapeamento de colunas da CJI3")
            st.dataframe(st.session_state.logs["CJI3"], use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma base carregada ainda.")

with tab_cadastro:
    df = st.session_state.cji3_enriquecida.copy()
    if df.empty:
        st.warning("Carregue e processe a CJI3 na primeira aba.")
    else:
        st.subheader("Cadastro / Central de Decisão")
        top1, top2, top3 = st.columns([1.2, 1.2, 2])
        with top1:
            peps = [""] + sorted([str(x) for x in df["PEP"].dropna().unique()])[:3000]
            filtro_pep = st.selectbox("Fabricante / PEP", peps)
        with top2:
            filtro_texto = st.text_input("Filtrar", placeholder="descrição, NF, pedido, fornecedor...")
        with top3:
            st.caption("A lista abaixo representa as linhas CJI3. Selecione um ID para preencher o formulário.")

        view = df.copy()
        if filtro_pep:
            view = view[view["PEP"].astype(str) == filtro_pep]
        if filtro_texto:
            nt = norm_text(filtro_texto)
            mask = view.astype(str).apply(lambda col: col.map(norm_text).str.contains(nt, na=False)).any(axis=1)
            view = view[mask]

        left, right = st.columns([0.95, 1.55], gap="large")
        with right:
            cols_show = [c for c in ["ID_CJI3", "Empresa", "Centro", "PEP", "Classe Custo", "Texto Pedido", "Texto Material", "Quantidade", "Valor", "NFe", "Pedido", "Classe", "Inventário", "Série"] if c in view.columns]
            st.dataframe(view[cols_show].head(500), use_container_width=True, height=230, hide_index=True)
            ids = view["ID_CJI3"].tolist()
            selected_id = st.selectbox("Selecionar ID_CJI3", ids, index=0 if ids else None)
        with left:
            row = df[df["ID_CJI3"] == selected_id].iloc[0] if selected_id else df.iloc[0]
            pref = build_prefill(row, st.session_state.bases.get("vida_util"))
            st.markdown('<div class="asset-card">', unsafe_allow_html=True)
            st.write("**Cadastro**")
            cA, cB, cC = st.columns([1.1, 0.55, 1])
            with cA:
                n_imob = st.text_input("N IMOB", value="")
                classe = st.text_input("Classe", value=pref["Classe"])
            with cB:
                sn_imob = st.text_input("Sn IMOB", value="")
            with cC:
                componente = st.checkbox("Componente")
                empresa = st.text_input("Empresa", value=pref["Empresa"])
            denom = st.text_input("Descr_01 / Denominação", value=pref["Denominação"])
            descr2 = st.text_input("Descr_02", value=pref["Descr_02"])
            descr3 = st.text_input("Descr_03", value=pref["Descr_03"])
            c1, c2 = st.columns(2)
            with c1:
                serie = st.text_input("Série", value=pref["Série"])
            with c2:
                inventario = st.text_input("N Plaqueta / Inventário", value=pref["Inventário"])
            c1, c2, c3 = st.columns([0.9, 0.7, 1])
            with c1:
                qtd = st.number_input("QTD", value=float(pref["Qtd"]), min_value=0.0, step=1.0)
            with c2:
                un_med = st.text_input("UN MED", value=pref["Un_Med"])
            with c3:
                cc_sap = st.text_input("C.C SAP", value=pref["C.C_SAP"])
            c1, c2, c3 = st.columns([1, 0.8, 1])
            with c1:
                fornec = st.text_input("FORNEC", value=pref["Código_FORNC"])
            with c2:
                centro = st.text_input("CENTRO", value=pref["CENTRO"])
            with c3:
                nf = st.text_input("NF / Pedido", value=pref["Nota Fiscal / Pedido"])
            c1, c2 = st.columns([1, 1.4])
            with c1:
                dt_deprec = st.text_input("Dt.Inic. Deprec.", value=pref["Início Depreciação"])
            with c2:
                pep = st.text_input("PEP", value=pref["Fabricante / Pep"])
            vida = st.text_input("Vida Útil", value=str(pref["Vida Útil"]))
            st.markdown('</div>', unsafe_allow_html=True)

        st.divider()
        b1, b2, b3, b4 = st.columns([1.2, 1.2, 1.2, 2.6])
        with b1:
            tratamento = st.selectbox("Tratamento", TRATAMENTOS)
        with b2:
            criterio = st.selectbox("Critério Rateio", CRITERIOS_RATEIO)
        with b3:
            valor_cap = st.number_input("Valor Capitalizado", value=float(pref["Valor Capitalizado"]), step=100.0)
        with b4:
            obs = st.text_input("Observação / regra aplicada", value="")

        ac1, ac2, ac3 = st.columns([1, 1, 4])
        with ac1:
            if st.button("SALVAR", type="primary"):
                decisao = {
                    "Classe": classe, "Empresa": empresa, "Denominação": denom, "Descr_02": descr2, "Descr_03": descr3,
                    "Série": serie, "Inventário": inventario, "Qtd": qtd, "Un_Med": un_med, "C.C_SAP": cc_sap,
                    "CENTRO": centro, "Segmento": row.get("Segmento Chamado", row.get("Segmento", "")),
                    "Código_FORNC": fornec, "NOME_Fornecedor": pref["NOME_Fornecedor"], "Fabricante / Pep": pep,
                    "Vida Útil": vida, "Início Depreciação": dt_deprec, "Nota Fiscal / Pedido": nf,
                    "Valor Capitalizado": valor_cap, "Origem_ID": selected_id, "Tratamento": tratamento,
                    "Critério Rateio": criterio, "Status": "Salvo", "Observação": obs,
                    "N IMOB": n_imob, "SN IMOB": sn_imob, "Componente": "Sim" if componente else "Não"
                }
                st.session_state.decisoes = [d for d in st.session_state.decisoes if d.get("Origem_ID") != selected_id]
                st.session_state.decisoes.append(decisao)
                st.success(f"Decisão salva para ID_CJI3 {selected_id}.")
        with ac2:
            if st.button("EXCLUIR"):
                antes = len(st.session_state.decisoes)
                st.session_state.decisoes = [d for d in st.session_state.decisoes if d.get("Origem_ID") != selected_id]
                st.warning("Decisão excluída." if len(st.session_state.decisoes) < antes else "Não havia decisão salva para este ID.")

        st.subheader("Itens cadastrados / decisões salvas")
        if st.session_state.decisoes:
            st.dataframe(pd.DataFrame(st.session_state.decisoes), use_container_width=True, height=260, hide_index=True)
        else:
            st.info("Nenhuma decisão salva ainda.")

with tab_template:
    st.subheader("Template AS01 e exportação")
    if not st.session_state.decisoes:
        st.warning("Nenhum item salvo na Central de Decisão.")
    else:
        template = decision_to_template(st.session_state.decisoes)
        st.dataframe(template, use_container_width=True, height=360, hide_index=True)
        pendencias = template[template["Tratamento"].astype(str).str.contains("Pendência|Não capitalizável", case=False, na=False)].copy()
        pacote = excel_bytes({
            "Template_AS01": template,
            "Decisoes_Completas": pd.DataFrame(st.session_state.decisoes),
            "Pendencias": pendencias,
            "CJI3_Enriquecida": st.session_state.cji3_enriquecida.head(100000),
        })
        st.download_button(
            "Baixar pacote Excel final",
            data=pacote,
            file_name="Template_AS01_MVP_V2.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary"
        )

    st.divider()
    st.write("**Bases carregadas disponíveis para auditoria:**")
    if st.session_state.bases:
        for nome, bdf in st.session_state.bases.items():
            with st.expander(f"{nome} - {len(bdf)} linhas"):
                st.dataframe(bdf.head(200), use_container_width=True, hide_index=True)
