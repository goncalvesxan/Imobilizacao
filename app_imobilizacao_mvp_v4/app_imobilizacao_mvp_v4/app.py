from __future__ import annotations

import pandas as pd
import streamlit as st

from core.loaders import (
    read_excel_any,
    standardize_cji3,
    standardize_classes,
    standardize_localizacao,
    standardize_empresas,
    standardize_chamados,
    standardize_razao,
    standardize_vida_util,
)
from core.enrichment import enrich_cji3, build_prefill, decision_to_template, get_vida_util_val
from core.schema import TRATAMENTOS, CRITERIOS_RATEIO
from core.utils import excel_bytes, norm_text, to_number

st.set_page_config(page_title="Motor de Imobilização SAP - MVP V4", layout="wide")

st.markdown(
    """
<style>
.block-container {padding-top: 1.2rem; max-width: 100%;}
.asset-card {border:1px solid #9E9E9E; padding:14px; border-radius:6px; background:#FAFAFA;}
.stTabs [data-baseweb="tab-list"] {gap: 8px;}
.stTabs [data-baseweb="tab"] {border:1px solid #BDBDBD; border-radius:6px 6px 0 0; padding:8px 14px;}
.metric-box {border:1px solid #DDD; border-radius:8px; padding:10px; background:#FFF;}
</style>
""",
    unsafe_allow_html=True,
)

st.title("Cadastro de Ativo - Motor de Imobilização SAP")
st.caption("MVP V4: seleção acumulada de múltiplos itens CJI3 + tratamento em lote + template AS01.")

for k, default in {
    "bases": {},
    "logs": {},
    "cji3": pd.DataFrame(),
    "cji3_enriquecida": pd.DataFrame(),
    "decisoes": [],
    "ids_selecionados": [],
}.items():
    if k not in st.session_state:
        st.session_state[k] = default


def _str_clean(v) -> str:
    if pd.isna(v):
        return ""
    return str(v).strip()


def _unique_options(df: pd.DataFrame | None, col: str) -> list[str]:
    if df is None or df.empty or col not in df.columns:
        return [""]
    vals = [_str_clean(x).replace(".0", "") for x in df[col].dropna().unique()]
    vals = sorted({v for v in vals if v})
    return [""] + vals


def _select_index(options: list[str], value: str) -> int:
    value = _str_clean(value).replace(".0", "")
    if value and value not in options:
        options.append(value)
    try:
        return options.index(value)
    except ValueError:
        return 0


def _vida_options(vida_df: pd.DataFrame | None) -> list[str]:
    if vida_df is None or vida_df.empty:
        return [""]
    vals: set[str] = set()
    for c in vida_df.columns:
        if c in ["Classe", "Conta SAP", "Conta SAP Deprec"]:
            continue
        for v in vida_df[c].dropna().unique():
            sv = _str_clean(v).replace(".0", "")
            if sv:
                vals.add(sv)
    return [""] + sorted(vals)


def _same_or_first(rows: pd.DataFrame, col: str) -> str:
    if col not in rows.columns or rows.empty:
        return ""
    vals = [_str_clean(v).replace(".0", "") for v in rows[col].dropna().unique() if _str_clean(v)]
    vals = list(dict.fromkeys(vals))
    if not vals:
        return ""
    return vals[0] if len(vals) == 1 else vals[0]


def _join_unique(rows: pd.DataFrame, col: str, limit: int = 8) -> str:
    if col not in rows.columns or rows.empty:
        return ""
    vals = [_str_clean(v).replace(".0", "") for v in rows[col].dropna().unique() if _str_clean(v)]
    vals = list(dict.fromkeys(vals))[:limit]
    return " | ".join(vals)


def _build_prefill_lote(rows: pd.DataFrame, vida_df: pd.DataFrame | None) -> dict:
    """Cria preenchimento inicial para tratamento aplicado a 1 ou vários IDs CJI3."""
    if rows.empty:
        return {}
    base = build_prefill(rows.iloc[0], vida_df)
    qtd_total = sum(to_number(v, 0) for v in rows.get("Quantidade", pd.Series(dtype=float)))
    valor_total = sum(to_number(v, 0) for v in rows.get("Valor", pd.Series(dtype=float)))
    n = len(rows)

    # Campos comuns do lote. Quando houver divergência, mantém o primeiro valor e deixa rastreabilidade no quadro selecionado.
    base["Empresa"] = _same_or_first(rows, "Empresa") or base.get("Empresa", "")
    base["C.C_SAP"] = _same_or_first(rows, "Centro") or base.get("C.C_SAP", "")
    base["CENTRO"] = _same_or_first(rows, "Centro") or base.get("CENTRO", "")
    base["Código_FORNC"] = _same_or_first(rows, "Conta Contrapartida") or base.get("Código_FORNC", "")
    base["NOME_Fornecedor"] = _same_or_first(rows, "Fornecedor Nome") or base.get("NOME_Fornecedor", "")
    base["Fabricante / Pep"] = _same_or_first(rows, "PEP") or base.get("Fabricante / Pep", "")
    base["Qtd"] = qtd_total if qtd_total > 0 else n
    base["Valor Capitalizado"] = valor_total

    nf = _join_unique(rows, "NFe") or _join_unique(rows, "NF Remessa")
    pedido = _join_unique(rows, "Pedido")
    base["Nota Fiscal / Pedido"] = f"{nf} / {pedido}".strip(" /")

    if n > 1:
        texto = _join_unique(rows, "Texto Pedido", limit=2) or _join_unique(rows, "Texto Material", limit=2)
        base["Denominação"] = f"AGRUPAMENTO CJI3 - {n} ITENS"[:50]
        base["Descr_02"] = texto[:50]
        base["Descr_03"] = f"IDs: {', '.join(map(str, rows['ID_CJI3'].tolist()[:8]))}"[:50]
    return base


def _ids_key(ids: list[int]) -> str:
    return ",".join(map(str, sorted(set(int(x) for x in ids))))


tab_upload, tab_cadastro, tab_template = st.tabs(
    ["1. Carregamento das Bases", "2. Cadastro / Decisão", "3. Template AS01 / Exportação"]
)

with tab_upload:
    st.subheader("Carregar bases necessárias")
    st.write("Carregue a CJI3 e as tabelas auxiliares necessárias para montar o cadastro AS01.")
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
                st.session_state.ids_selecionados = []
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
            st.caption("Selecione um ou vários ID_CJI3, adicione ao quadro de tratamento e aplique uma única regra ao conjunto.")

        view = df.copy()
        if filtro_pep:
            view = view[view["PEP"].astype(str) == filtro_pep]
        if filtro_texto:
            nt = norm_text(filtro_texto)
            mask = view.astype(str).apply(lambda col: col.map(norm_text).str.contains(nt, na=False)).any(axis=1)
            view = view[mask]

        if view.empty:
            st.warning("Nenhum registro encontrado com os filtros informados.")
            st.stop()

        # Quadro CJI3 completo, com todas as colunas e seleção múltipla por ID.
        st.write("**Quadro CJI3 - todas as colunas do arquivo de origem**")
        cji3_cols = [c for c in st.session_state.cji3.columns if c in view.columns]
        if "ID_CJI3" in cji3_cols:
            cji3_cols = ["ID_CJI3"] + [c for c in cji3_cols if c != "ID_CJI3"]
        st.dataframe(view[cji3_cols].head(1000), use_container_width=True, height=280, hide_index=True)

        ids = [int(x) for x in view["ID_CJI3"].dropna().tolist()]
        ids_para_adicionar = st.multiselect("Selecionar ID_CJI3 para acumular no tratamento", options=ids, format_func=lambda x: f"ID {x}")
        bt1, bt2, bt3, bt4 = st.columns([1.2, 1.4, 1.4, 3])
        with bt1:
            if st.button("Adicionar selecionados", type="primary"):
                atual = set(st.session_state.ids_selecionados)
                atual.update(ids_para_adicionar)
                st.session_state.ids_selecionados = sorted(atual)
                st.rerun()
        with bt2:
            if st.button("Adicionar todos filtrados"):
                atual = set(st.session_state.ids_selecionados)
                atual.update(ids)
                st.session_state.ids_selecionados = sorted(atual)
                st.rerun()
        with bt3:
            if st.button("Limpar quadro selecionado"):
                st.session_state.ids_selecionados = []
                st.rerun()

        selected_ids = [int(x) for x in st.session_state.ids_selecionados]
        selected_rows = df[df["ID_CJI3"].isin(selected_ids)].copy()

        st.write("**Quadro de itens selecionados para receber o tratamento**")
        if selected_rows.empty:
            st.info("Nenhum item acumulado ainda. Selecione IDs no quadro CJI3 acima e clique em 'Adicionar selecionados'.")
            st.stop()
        else:
            st.dataframe(selected_rows[cji3_cols], use_container_width=True, height=220, hide_index=True)
            total_qtd = sum(to_number(v, 0) for v in selected_rows.get("Quantidade", pd.Series(dtype=float)))
            total_valor = sum(to_number(v, 0) for v in selected_rows.get("Valor", pd.Series(dtype=float)))
            m1, m2, m3, m4 = st.columns(4)
            m1.metric("Itens selecionados", len(selected_rows))
            m2.metric("Quantidade total", f"{total_qtd:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            m3.metric("Valor total", f"{total_valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
            m4.metric("PEPs distintos", selected_rows["PEP"].astype(str).nunique() if "PEP" in selected_rows else 0)

        pref = _build_prefill_lote(selected_rows, st.session_state.bases.get("vida_util"))
        classes_df = st.session_state.bases.get("classes")
        vida_df = st.session_state.bases.get("vida_util")
        lote_key = _ids_key(selected_ids)

        st.divider()
        left, right = st.columns([0.9, 1.5], gap="large")
        with left:
            st.markdown('<div class="asset-card">', unsafe_allow_html=True)
            st.write("**Cadastro do ativo / tratamento do lote selecionado**")
            cA, cB, cC = st.columns([1.1, 0.55, 1])
            with cA:
                n_imob = st.text_input("N IMOB", value="", key=f"n_imob_{lote_key}")
                classe_opts = _unique_options(classes_df, "Classe")
                classe_default = pref.get("Classe", "")
                classe = st.selectbox(
                    "Classe",
                    classe_opts,
                    index=_select_index(classe_opts, classe_default),
                    key=f"classe_{lote_key}",
                    help="Campo alimentado pela Tabela Classes. A sugestão inicial vem da conta/classe de custo da CJI3.",
                )
            with cB:
                sn_imob = st.text_input("Sn IMOB", value="", key=f"sn_{lote_key}")
            with cC:
                componente = st.checkbox("Componente", key=f"comp_{lote_key}")
                empresa = st.text_input("Empresa", value=pref.get("Empresa", ""), key=f"empresa_{lote_key}")

            denom = st.text_input("Descr_01 / Denominação", value=pref.get("Denominação", ""), key=f"denom_{lote_key}")
            descr2 = st.text_input("Descr_02", value=pref.get("Descr_02", ""), key=f"descr2_{lote_key}")
            descr3 = st.text_input("Descr_03", value=pref.get("Descr_03", ""), key=f"descr3_{lote_key}")
            c1, c2 = st.columns(2)
            with c1:
                serie = st.text_input("Série", value=pref.get("Série", ""), key=f"serie_{lote_key}")
            with c2:
                inventario = st.text_input("N Plaqueta / Inventário", value=pref.get("Inventário", ""), key=f"inv_{lote_key}")

            c1, c2, c3 = st.columns([0.9, 0.7, 1])
            with c1:
                qtd = st.number_input("QTD", value=float(pref.get("Qtd", 0) or 0), min_value=0.0, step=1.0, key=f"qtd_{lote_key}")
            with c2:
                un_med = st.text_input("UN MED", value=pref.get("Un_Med", "UN"), key=f"un_{lote_key}")
            with c3:
                cc_sap = st.text_input("C.C SAP", value=pref.get("C.C_SAP", ""), key=f"cc_{lote_key}")

            c1, c2, c3 = st.columns([1, 0.8, 1])
            with c1:
                fornec = st.text_input("FORNEC", value=pref.get("Código_FORNC", ""), key=f"fornec_{lote_key}")
            with c2:
                centro = st.text_input("CENTRO", value=pref.get("CENTRO", ""), key=f"centro_{lote_key}")
            with c3:
                nf = st.text_input("NF / Pedido", value=pref.get("Nota Fiscal / Pedido", ""), key=f"nf_{lote_key}")

            c1, c2 = st.columns([1, 1.4])
            with c1:
                dt_deprec = st.text_input("Dt.Inic. Deprec.", value=pref.get("Início Depreciação", ""), key=f"dep_{lote_key}")
            with c2:
                pep = st.text_input("PEP", value=pref.get("Fabricante / Pep", ""), key=f"pep_{lote_key}")

            vida_sugerida = get_vida_util_val(classe, empresa, vida_df) or pref.get("Vida Útil", "")
            vida_opts = _vida_options(vida_df)
            vida = st.selectbox(
                "Vida Útil",
                vida_opts,
                index=_select_index(vida_opts, str(vida_sugerida)),
                key=f"vida_{lote_key}_{classe}_{empresa}",
                help="Vida útil sugerida pela tabela de depreciação conforme Classe e Empresa. Pode ser alterada manualmente.",
            )
            st.markdown("</div>", unsafe_allow_html=True)

        with right:
            st.write("**Resumo técnico do lote selecionado**")
            resumo_cols = [c for c in ["ID_CJI3", "Empresa", "Centro", "PEP", "Classe Custo", "Texto Pedido", "Texto Material", "Quantidade", "Valor", "Conta Contrapartida", "Fornecedor Nome", "Pedido", "NFe"] if c in selected_rows.columns]
            st.dataframe(selected_rows[resumo_cols], use_container_width=True, height=420, hide_index=True)

        st.divider()
        b1, b2, b3, b4 = st.columns([1.2, 1.2, 1.2, 2.6])
        with b1:
            tratamento = st.selectbox("Tratamento", TRATAMENTOS, key=f"trat_{lote_key}")
        with b2:
            criterio = st.selectbox("Critério Rateio", CRITERIOS_RATEIO, key=f"crit_{lote_key}")
        with b3:
            valor_cap = st.number_input("Valor Capitalizado", value=float(pref.get("Valor Capitalizado", 0) or 0), step=100.0, key=f"valor_{lote_key}")
        with b4:
            obs = st.text_input("Observação / regra aplicada", value="", key=f"obs_{lote_key}")

        ac1, ac2, ac3, ac4 = st.columns([1, 1, 1.2, 4])
        with ac1:
            if st.button("SALVAR LOTE", type="primary"):
                decisao = {
                    "Classe": classe,
                    "Empresa": empresa,
                    "Denominação": denom,
                    "Descr_02": descr2,
                    "Descr_03": descr3,
                    "Série": serie,
                    "Inventário": inventario,
                    "Qtd": qtd,
                    "Un_Med": un_med,
                    "C.C_SAP": cc_sap,
                    "CENTRO": centro,
                    "Segmento": _same_or_first(selected_rows, "Segmento Chamado") or _same_or_first(selected_rows, "Segmento"),
                    "Código_FORNC": fornec,
                    "NOME_Fornecedor": pref.get("NOME_Fornecedor", ""),
                    "Fabricante / Pep": pep,
                    "Vida Útil": vida,
                    "Início Depreciação": dt_deprec,
                    "Nota Fiscal / Pedido": nf,
                    "Valor Capitalizado": valor_cap,
                    "Origem_ID": lote_key,
                    "Origem_IDs": lote_key,
                    "Qtd_Itens_CJI3": len(selected_rows),
                    "Tratamento": tratamento,
                    "Critério Rateio": criterio,
                    "Status": "Salvo",
                    "Observação": obs,
                    "N IMOB": n_imob,
                    "SN IMOB": sn_imob,
                    "Componente": "Sim" if componente else "Não",
                }
                st.session_state.decisoes = [d for d in st.session_state.decisoes if d.get("Origem_ID") != lote_key]
                st.session_state.decisoes.append(decisao)
                st.success(f"Decisão salva para o lote de IDs: {lote_key}.")
        with ac2:
            if st.button("EXCLUIR LOTE"):
                antes = len(st.session_state.decisoes)
                st.session_state.decisoes = [d for d in st.session_state.decisoes if d.get("Origem_ID") != lote_key]
                st.warning("Decisão do lote excluída." if len(st.session_state.decisoes) < antes else "Não havia decisão salva para este lote.")
        with ac3:
            if st.button("SALVAR E LIMPAR"):
                decisao = {
                    "Classe": classe, "Empresa": empresa, "Denominação": denom, "Descr_02": descr2, "Descr_03": descr3,
                    "Série": serie, "Inventário": inventario, "Qtd": qtd, "Un_Med": un_med, "C.C_SAP": cc_sap,
                    "CENTRO": centro, "Segmento": _same_or_first(selected_rows, "Segmento Chamado") or _same_or_first(selected_rows, "Segmento"),
                    "Código_FORNC": fornec, "NOME_Fornecedor": pref.get("NOME_Fornecedor", ""), "Fabricante / Pep": pep,
                    "Vida Útil": vida, "Início Depreciação": dt_deprec, "Nota Fiscal / Pedido": nf, "Valor Capitalizado": valor_cap,
                    "Origem_ID": lote_key, "Origem_IDs": lote_key, "Qtd_Itens_CJI3": len(selected_rows),
                    "Tratamento": tratamento, "Critério Rateio": criterio, "Status": "Salvo", "Observação": obs,
                    "N IMOB": n_imob, "SN IMOB": sn_imob, "Componente": "Sim" if componente else "Não",
                }
                st.session_state.decisoes = [d for d in st.session_state.decisoes if d.get("Origem_ID") != lote_key]
                st.session_state.decisoes.append(decisao)
                st.session_state.ids_selecionados = []
                st.success(f"Decisão salva para o lote de IDs: {lote_key}.")
                st.rerun()

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
        pacote = excel_bytes(
            {
                "Template_AS01": template,
                "Decisoes_Completas": pd.DataFrame(st.session_state.decisoes),
                "Pendencias": pendencias,
                "CJI3_Enriquecida": st.session_state.cji3_enriquecida.head(100000),
                "CJI3_Original_Padronizada": st.session_state.cji3.head(100000),
            }
        )
        st.download_button(
            "Baixar pacote Excel final",
            data=pacote,
            file_name="Template_AS01_MVP_V4.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
        )

    st.divider()
    st.write("**Bases carregadas disponíveis para auditoria:**")
    if st.session_state.bases:
        for nome, bdf in st.session_state.bases.items():
            with st.expander(f"{nome} - {len(bdf)} linhas"):
                st.dataframe(bdf.head(200), use_container_width=True, hide_index=True)
