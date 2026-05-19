from __future__ import annotations

from datetime import datetime
import html

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
from core.schema import TRATAMENTOS, CRITERIOS_RATEIO, TEMPLATE_AS01
from core.utils import excel_bytes, norm_text, to_number

st.set_page_config(page_title="Capitalização de Ativos - Imobilizações Adições", layout="wide")

st.markdown(
    """
<style>
:root {--navy:#0F172A; --blue:#2563EB; --soft:#F8FAFC; --line:#E2E8F0; --muted:#64748B;}
.block-container {padding-top: 0.75rem; padding-left: 1.2rem; padding-right: 1.2rem; max-width: 100%;}
.main .block-container {background: linear-gradient(180deg, #F8FAFC 0%, #FFFFFF 38%, #F8FAFC 100%);}
h1, h2, h3 {letter-spacing: -0.02em;}
.hero {border:1px solid #CBD5E1; border-radius:22px; padding:22px 26px; margin-bottom:12px; background:linear-gradient(135deg,#0F172A 0%,#1D4ED8 58%,#38BDF8 100%); color:white; box-shadow:0 12px 36px rgba(15,23,42,.18);}
.hero-title {font-size:2.15rem; font-weight:800; line-height:1.1; margin-bottom:6px;}
.hero-subtitle {font-size:.95rem; opacity:.92;}
.hero-pill {display:inline-block; margin-top:12px; padding:5px 11px; border-radius:999px; background:rgba(255,255,255,.15); border:1px solid rgba(255,255,255,.25); font-size:.78rem; font-weight:700;}
.asset-card {border:1px solid #D7DEE8; padding:18px; border-radius:16px; background:#FFFFFF; box-shadow:0 8px 24px rgba(31,41,55,0.08);}
.stTabs [data-baseweb="tab-list"] {gap: 8px; border-bottom:1px solid #CBD5E1; background:#F8FAFC; padding-top:4px; position: sticky; top: 0; z-index: 99;}
.stTabs [data-baseweb="tab"] {border:1px solid #CBD5E1; border-radius:12px 12px 0 0; padding:10px 18px; background:#F1F5F9; font-weight:700;}
.stTabs [aria-selected="true"] {background:#FFFFFF; border-top:4px solid #2563EB; color:#1E3A8A;}
.selection-help {border-left:5px solid #2563EB; background:linear-gradient(90deg,#EFF6FF,#FFFFFF); padding:13px 17px; border-radius:14px; margin-bottom:12px; color:#1F2937; border-top:1px solid #DBEAFE; border-right:1px solid #DBEAFE; border-bottom:1px solid #DBEAFE;}
.small-note {color:#64748B; font-size:0.88rem;}
.status-card {border:1px solid #E2E8F0; border-radius:16px; padding:14px 16px; background:#FFFFFF; box-shadow:0 6px 18px rgba(15,23,42,0.06);}
.status-title {font-size:0.72rem; color:#64748B; text-transform:uppercase; letter-spacing:.06em; font-weight:700;}
.status-value {font-size:1.65rem; font-weight:800; color:#0F172A; margin-top:2px;}
.badge-ok {display:inline-block; padding:3px 9px; border-radius:999px; background:#DCFCE7; color:#166534; font-weight:700; font-size:.78rem;}
.badge-warn {display:inline-block; padding:3px 9px; border-radius:999px; background:#FEF3C7; color:#92400E; font-weight:700; font-size:.78rem;}
.badge-stop {display:inline-block; padding:3px 9px; border-radius:999px; background:#FEE2E2; color:#991B1B; font-weight:700; font-size:.78rem;}
div[data-testid="stDataFrame"], div[data-testid="stDataEditor"] {border-radius:14px; overflow:hidden; border:1px solid #E2E8F0; box-shadow:0 4px 16px rgba(15,23,42,.045);}
[data-testid="stMetric"] {background:#FFFFFF; border:1px solid #E2E8F0; padding:12px 14px; border-radius:15px; box-shadow:0 5px 14px rgba(15,23,42,.05);}
[data-testid="stMetricLabel"] {font-weight:700; color:#64748B;}
[data-testid="stMetricValue"] {font-weight:800; color:#0F172A;}
.stButton > button, .stDownloadButton > button {border-radius:11px; font-weight:700; border:1px solid #CBD5E1;}
.lote-panel {border:1px solid #CBD5E1; border-radius:18px; padding:17px; background:linear-gradient(180deg,#FFFFFF 0%,#F8FAFC 100%); box-shadow:0 10px 28px rgba(15,23,42,.08); margin-bottom:16px;}
.panel-title {font-size:1.05rem; font-weight:800; color:#0F172A; margin-bottom:4px;}
.panel-subtitle {font-size:.84rem; color:#64748B; margin-bottom:10px;}
.legend {display:flex; gap:10px; align-items:center; flex-wrap:wrap; margin: 6px 0 10px 0;}
.legend-item {border:1px solid #CBD5E1; border-radius:999px; padding:5px 11px; background:#fff; font-size:.82rem; color:#334155; font-weight:700;}
.prof-divider {height:1px; background:#E2E8F0; margin:14px 0;}
.section-label {font-size:.78rem; color:#2563EB; font-weight:800; letter-spacing:.08em; text-transform:uppercase; margin-top:8px;}
.compact-note {font-size:.82rem; color:#64748B;}

.hover-scroll {max-height: 310px; overflow:auto; border:1px solid #D8E0EA; border-radius:16px; background:#FFFFFF; box-shadow:0 8px 24px rgba(15,23,42,.06);}
.hover-table {border-collapse:separate; border-spacing:0; width:100%; min-width:980px; font-size:.82rem;}
.hover-table th {position:sticky; top:0; background:#F8FAFC; color:#475569; font-weight:800; text-align:left; border-bottom:1px solid #D8E0EA; padding:10px 12px; z-index:3;}
.hover-table td {border-bottom:1px solid #E5EAF1; padding:10px 12px; color:#0F172A; vertical-align:top; white-space:nowrap;}
.hover-table tr:hover {background:#EFF6FF; cursor:pointer;}
.hover-table tr:hover td {box-shadow: inset 0 1px 0 #BFDBFE, inset 0 -1px 0 #BFDBFE;}
.tooltip-anchor {position:relative; font-weight:800; color:#1D4ED8;}
.row-tooltip {display:none; position:fixed; left:clamp(20px, 22vw, 340px); top:120px; width:min(820px, 72vw); max-height:70vh; overflow:auto; z-index:999999; background:#FFFFFF; color:#0F172A; border:1px solid #CBD5E1; border-radius:18px; padding:0; box-shadow:0 24px 64px rgba(15,23,42,.24);}
.hover-table tr:hover .row-tooltip {display:block;}
.tooltip-header {position:sticky; top:0; background:linear-gradient(135deg,#0F172A,#1E40AF); color:#FFFFFF; padding:14px 18px; font-weight:900; border-radius:18px 18px 0 0; letter-spacing:.2px;}
.tooltip-grid {display:grid; grid-template-columns:180px 1fr; gap:0; padding:12px 16px 16px 16px;}
.tooltip-key {font-size:.76rem; color:#64748B; font-weight:800; padding:8px 10px; border-bottom:1px solid #EEF2F7; background:#F8FAFC;}
.tooltip-val {font-size:.80rem; color:#111827; padding:8px 10px; border-bottom:1px solid #EEF2F7; line-height:1.35; white-space:normal; word-break:break-word;}
.hover-hint {font-size:.78rem; color:#64748B; margin:4px 0 8px 0;}

</style>
""",
    unsafe_allow_html=True,
)

st.markdown(
    """
    <div class="hero">
      <div class="hero-title">Capitalização de Ativos - Imobilizações Adições</div>
      <div class="hero-subtitle">Motor operacional para seleção CJI3, decisões de imobilização, rastreabilidade e geração do template SAP AS01.</div>
      <div class="hero-pill">MVP V10.6 • Tooltip de chamados • Visual profissional</div>
    </div>
    """,
    unsafe_allow_html=True,
)

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


# Cockpit superior sempre ativo: ajuda o usuário a se localizar sem abrir abas.
_cji3_count = len(st.session_state.cji3) if isinstance(st.session_state.cji3, pd.DataFrame) else 0
_dec_count = len(st.session_state.decisoes)
_lote_count = len(st.session_state.ids_selecionados)
_valor_cap_cockpit = 0.0
if st.session_state.decisoes:
    _df_cockpit = pd.DataFrame(st.session_state.decisoes)
    if "Valor Capitalizado" in _df_cockpit.columns:
        _valor_cap_cockpit = pd.to_numeric(_df_cockpit["Valor Capitalizado"], errors="coerce").fillna(0).sum()
_c0, _c1, _c2, _c3, _c4 = st.columns([1.1, 1, 1, 1, 1.2])
_c0.metric("Bases carregadas", len(st.session_state.bases))
_c1.metric("Linhas CJI3", _cji3_count)
_c2.metric("IDs no lote", _lote_count)
_c3.metric("Decisões salvas", _dec_count)
_c4.metric("Valor capitalizado", f"{_valor_cap_cockpit:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
st.markdown('<div class="compact-note">Fluxo recomendado: carregar bases → revisar chamados → selecionar CJI3 → salvar decisões → exportar AS01.</div>', unsafe_allow_html=True)


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

def _format_hover_value(v) -> str:
    if pd.isna(v):
        return ""
    if isinstance(v, (pd.Timestamp, datetime)):
        return v.strftime("%d/%m/%Y %H:%M")
    return str(v).replace("nan", "").strip()


def _chamado_hover_table(df: pd.DataFrame, visible_cols: list[str], max_rows: int = 80) -> str:
    """Tabela HTML com tooltip organizado por linha para visualizar todos os campos do chamado."""
    if df is None or df.empty:
        return '<div class="selection-help">Nenhum chamado encontrado para o PEP selecionado.</div>'

    view = df.head(max_rows).copy()
    visible_cols = [c for c in visible_cols if c in view.columns]
    if not visible_cols:
        visible_cols = list(view.columns[:8])

    headers = "".join(f"<th>{html.escape(str(c))}</th>" for c in visible_cols)
    rows_html = []
    for _, row in view.iterrows():
        titulo = _format_hover_value(row.get("N_CHAMADO", "")) or _format_hover_value(row.get(visible_cols[0], ""))
        subtitulo = _format_hover_value(row.get("PEP", ""))
        header = f"Chamado {html.escape(titulo)}"
        if subtitulo:
            header += f" · PEP {html.escape(subtitulo)}"

        detalhe_html = []
        for col in view.columns:
            val = _format_hover_value(row.get(col, ""))
            if val != "":
                detalhe_html.append(
                    f'<div class="tooltip-key">{html.escape(str(col))}</div>'
                    f'<div class="tooltip-val">{html.escape(val)}</div>'
                )
        tooltip_box = (
            '<div class="row-tooltip">'
            f'<div class="tooltip-header">{header}</div>'
            f'<div class="tooltip-grid">{"".join(detalhe_html)}</div>'
            '</div>'
        )

        cells = []
        for i, col in enumerate(visible_cols):
            val = html.escape(_format_hover_value(row.get(col, "")))
            if i == 0:
                cells.append(f'<td class="tooltip-anchor">{val}{tooltip_box}</td>')
            else:
                cells.append(f"<td>{val}</td>")
        rows_html.append('<tr>' + ''.join(cells) + '</tr>')

    aviso = ""
    if len(df) > max_rows:
        aviso = f'<div class="hover-hint">Exibindo {max_rows} de {len(df)} registros. Use os filtros para refinar a consulta.</div>'

    return (
        '<div class="hover-hint">Passe o mouse sobre qualquer linha para abrir uma caixa organizada com todos os campos do chamado.</div>'
        + aviso
        + '<div class="hover-scroll">'
        + '<table class="hover-table">'
        + f'<thead><tr>{headers}</tr></thead>'
        + f'<tbody>{"".join(rows_html)}</tbody>'
        + '</table></div>'
    )

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



def _ids_from_saved_decisions(decisoes: list[dict]) -> set[int]:
    """Retorna todos os ID_CJI3 que já foram usados em decisões salvas."""
    ids: set[int] = set()
    for d in decisoes or []:
        raw = str(d.get("Origem_IDs") or d.get("Origem_ID") or "")
        raw = raw.split("#R")[0]
        for part in raw.replace(";", ",").replace("|", ",").split(","):
            part = part.strip()
            if part.isdigit():
                ids.add(int(part))
    return ids


def _status_item_cji3(id_cji3, ids_salvos: set[int], ids_lote: set[int]) -> str:
    try:
        i = int(id_cji3)
    except Exception:
        return "Indefinido"
    if i in ids_salvos:
        return "Salvo / bloqueado"
    if i in ids_lote:
        return "No lote atual"
    return "Disponível"


def _render_status_cards(total: int, disponiveis: int, no_lote: int, salvos: int, valor_disp: float, valor_salvo: float) -> None:
    c1, c2, c3, c4, c5 = st.columns([1, 1, 1, 1, 1.2])
    cards = [
        (c1, "Registros na visão", total, ""),
        (c2, "Disponíveis", disponiveis, "badge-ok"),
        (c3, "No lote atual", no_lote, "badge-warn"),
        (c4, "Já salvos", salvos, "badge-stop"),
        (c5, "Valor disponível", f"{valor_disp:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."), ""),
    ]
    for col, title, value, badge in cards:
        with col:
            st.markdown(
                f'<div class="status-card"><div class="status-title">{title}</div><div class="status-value">{value}</div></div>',
                unsafe_allow_html=True,
            )




def _status_badge_value(status: str) -> str:
    if status == "Salvo / bloqueado":
        return "🔴 Já utilizado"
    if status == "No lote atual":
        return "🟡 No lote"
    if status == "Disponível":
        return "🟢 Disponível"
    return "⚪ Indefinido"


def _style_status_rows(row):
    status = str(row.get("Situação", ""))
    if status == "Salvo / bloqueado":
        return ["background-color: #FEE2E2; color: #7F1D1D; font-weight: 600" for _ in row]
    if status == "No lote atual":
        return ["background-color: #FEF3C7; color: #78350F; font-weight: 600" for _ in row]
    if status == "Disponível":
        return ["background-color: #ECFDF5; color: #064E3B" for _ in row]
    return ["" for _ in row]


def _render_lote_decisoes_panel(df_base: pd.DataFrame, cji3_cols: list[str], titulo: str = "Painel permanente do lote") -> None:
    """Mostra sempre o lote atual e as decisões salvas, sem depender de seleção ativa."""
    ids_salvos = _ids_from_saved_decisions(st.session_state.decisoes)
    ids_lote = [int(x) for x in st.session_state.ids_selecionados if int(x) not in ids_salvos]
    st.session_state.ids_selecionados = ids_lote
    selected_rows_panel = df_base[df_base["ID_CJI3"].isin(ids_lote)].copy() if not df_base.empty and "ID_CJI3" in df_base.columns else pd.DataFrame()
    df_dec_panel = pd.DataFrame(st.session_state.decisoes) if st.session_state.decisoes else pd.DataFrame()

    qtd_lote = sum(to_number(v, 0) for v in selected_rows_panel.get("Quantidade", pd.Series(dtype=float))) if not selected_rows_panel.empty else 0
    val_lote = sum(to_number(v, 0) for v in selected_rows_panel.get("Valor", pd.Series(dtype=float))) if not selected_rows_panel.empty else 0
    val_salvo = sum(to_number(v, 0) for v in df_dec_panel.get("Valor Capitalizado", pd.Series(dtype=float))) if not df_dec_panel.empty else 0

    st.markdown('<div class="lote-panel">', unsafe_allow_html=True)
    st.markdown(f'<div class="panel-title">{titulo}</div><div class="panel-subtitle">Acompanhamento sempre visível dos itens acumulados e decisões já salvas.</div>', unsafe_allow_html=True)
    k1, k2, k3, k4, k5 = st.columns(5)
    k1.metric("IDs no lote atual", len(ids_lote))
    k2.metric("Qtd lote", f"{qtd_lote:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    k3.metric("Valor lote", f"{val_lote:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
    k4.metric("IDs já utilizados", len(ids_salvos))
    k5.metric("Valor salvo", f"{val_salvo:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    t_lote, t_dec = st.tabs(["Lote atual acumulado", "Itens cadastrados / decisões salvas"])
    with t_lote:
        if selected_rows_panel.empty:
            st.info("Nenhum ID acumulado no lote atual.")
        else:
            lote_cols = [c for c in ["ID_CJI3", "Empresa", "Centro", "PEP", "Classe Custo", "Texto Pedido", "Texto Material", "Quantidade", "Valor", "Conta Contrapartida", "Fornecedor Nome", "Pedido", "NFe"] if c in selected_rows_panel.columns]
            if not lote_cols:
                lote_cols = cji3_cols[:18]
            lote_panel = selected_rows_panel[lote_cols].copy()
            lote_panel.insert(1, "Situação", "No lote atual")
            st.dataframe(lote_panel.style.apply(_style_status_rows, axis=1), use_container_width=True, height=220, hide_index=True)
    with t_dec:
        if df_dec_panel.empty:
            st.info("Nenhuma decisão salva ainda.")
        else:
            st.caption("As decisões salvas bloqueiam automaticamente os respectivos ID_CJI3 na seleção para evitar duplicidade.")
            st.dataframe(df_dec_panel, use_container_width=True, height=220, hide_index=True)
    st.markdown('</div>', unsafe_allow_html=True)


def _gerar_decisoes_replicadas(decisao_base: dict, replicar: bool, qtd_replicas: float, valor_total: float) -> list[dict]:
    """Gera uma ou várias linhas cadastradas a partir de uma decisão.

    Quando Replicar=True, o cadastro é repetido conforme a QTD informada e
    o Valor Capitalizado é dividido igualmente entre as linhas criadas.
    """
    try:
        n = int(round(float(qtd_replicas or 0)))
    except Exception:
        n = 0
    if not replicar or n <= 1:
        d = decisao_base.copy()
        d["Replicar"] = "Não"
        d["Nº Réplica"] = ""
        d["Total Réplicas"] = 1
        return [d]

    valor_unit = float(valor_total or 0) / n if n else float(valor_total or 0)
    saida = []
    for i in range(1, n + 1):
        d = decisao_base.copy()
        d["Origem_ID"] = f"{decisao_base.get('Origem_ID', '')}#R{i:03d}"
        d["Replicar"] = "Sim"
        d["Nº Réplica"] = i
        d["Total Réplicas"] = n
        d["Qtd"] = 1
        d["Valor Capitalizado"] = round(valor_unit, 2)
        # Ajuda a rastrear que o cadastro foi desmembrado em partes iguais.
        obs = str(d.get("Observação", "") or "").strip()
        extra = f"Replicado por QTD: {i}/{n}; valor unitário: {valor_unit:.2f}"
        d["Observação"] = f"{obs} | {extra}" if obs else extra
        saida.append(d)
    return saida




def _build_template_as01_padronizado(decisoes: list[dict], somente_salvos: bool = True, incluir_pendencias: bool = True) -> pd.DataFrame:
    """Gera a visão padronizada para carga/apoio AS01 a partir das decisões salvas."""
    if not decisoes:
        return pd.DataFrame(columns=TEMPLATE_AS01)
    df = pd.DataFrame(decisoes).copy()
    if somente_salvos and "Status" in df.columns:
        df = df[df["Status"].astype(str).str.upper().eq("SALVO")].copy()
    if not incluir_pendencias and "Tratamento" in df.columns:
        mask_pend = df["Tratamento"].astype(str).str.contains("Pendência|Não capitalizável", case=False, na=False)
        df = df[~mask_pend].copy()
    template = decision_to_template(df.to_dict("records"))
    for col in TEMPLATE_AS01:
        if col not in template.columns:
            template[col] = ""
    for col in ["Classe", "Empresa", "C.C_SAP", "CENTRO", "Código_FORNC", "Fabricante / Pep", "Vida Útil"]:
        if col in template.columns:
            template[col] = template[col].astype(str).str.replace(".0", "", regex=False).str.strip()
    if "Qtd" in template.columns:
        template["Qtd"] = pd.to_numeric(template["Qtd"], errors="coerce").fillna(0)
    if "Valor Capitalizado" in template.columns:
        template["Valor Capitalizado"] = pd.to_numeric(template["Valor Capitalizado"], errors="coerce").fillna(0)
    return template[TEMPLATE_AS01]


def _validar_template_as01(template: pd.DataFrame) -> pd.DataFrame:
    obrigatorios = ["Classe", "Empresa", "Denominação", "Qtd", "Un_Med", "C.C_SAP", "CENTRO", "Fabricante / Pep", "Vida Útil", "Nota Fiscal / Pedido"]
    linhas = []
    total = len(template)
    for col in obrigatorios:
        if col not in template.columns:
            vazios = total
        elif col == "Qtd":
            vazios = int((pd.to_numeric(template[col], errors="coerce").fillna(0) <= 0).sum())
        else:
            serie = template[col]
            vazios = int(serie.isna().sum() + serie.astype(str).str.strip().isin(["", "nan", "None", "0", "0.0"]).sum())
        linhas.append({
            "Campo obrigatório": col,
            "Registros vazios/inválidos": vazios,
            "Total de registros": total,
            "Status": "OK" if vazios == 0 else "Revisar",
        })
    return pd.DataFrame(linhas)


def _parametros_exportacao(nome_arquivo: str, somente_salvos: bool, incluir_pendencias: bool, template: pd.DataFrame) -> pd.DataFrame:
    return pd.DataFrame([
        {"Parâmetro": "Versão do aplicativo", "Valor": "MVP V10.6"},
        {"Parâmetro": "Data/hora de geração", "Valor": datetime.now().strftime("%d/%m/%Y %H:%M:%S")},
        {"Parâmetro": "Nome sugerido do arquivo", "Valor": nome_arquivo},
        {"Parâmetro": "Somente status Salvo", "Valor": "Sim" if somente_salvos else "Não"},
        {"Parâmetro": "Incluir pendências/não capitalizáveis", "Valor": "Sim" if incluir_pendencias else "Não"},
        {"Parâmetro": "Quantidade de linhas no template", "Valor": len(template)},
        {"Parâmetro": "Valor capitalizado total", "Valor": float(template["Valor Capitalizado"].sum()) if "Valor Capitalizado" in template.columns and not template.empty else 0},
    ])


def _readme_exportacao() -> pd.DataFrame:
    return pd.DataFrame([
        {"Etapa": 1, "Descrição": "Conferir a aba Template_AS01_Padronizado."},
        {"Etapa": 2, "Descrição": "Validar campos obrigatórios na aba Validacoes_AS01."},
        {"Etapa": 3, "Descrição": "Usar a aba Decisoes_Completas como rastreabilidade do lote acumulado."},
        {"Etapa": 4, "Descrição": "Usar a aba CJI3_IDs_Utilizados para auditar quais lançamentos CJI3 já foram tratados."},
        {"Etapa": 5, "Descrição": "Enviar ao SAP somente após revisão dos campos obrigatórios e regras de capitalização."},
    ])


tab_upload, tab_chamados, tab_cadastro, tab_template = st.tabs(
    ["1. Carregamento das Bases", "2. Gestão de Chamados", "3. Cadastro / Decisão", "4. Template AS01 / Exportação"]
)

with tab_upload:
    st.markdown('<div class="section-label">Entrada de dados</div>', unsafe_allow_html=True)
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



with tab_chamados:
    st.markdown('<div class="section-label">Gestão operacional dos chamados</div>', unsafe_allow_html=True)
    st.subheader("Gestão do arquivo de Chamados")
    chamados = st.session_state.bases.get("chamados", pd.DataFrame()).copy()
    cji3_base = st.session_state.cji3_enriquecida.copy()

    if chamados.empty:
        st.warning("Carregue o arquivo de Chamados na aba 1 para habilitar esta gestão.")
        st.info("Esta aba foi criada para consultar chamados por PEP, etapa da solicitação, classe, filial, centro, NF, pedido e cruzar com os lançamentos CJI3 já carregados.")
    else:
        for col in chamados.columns:
            if chamados[col].dtype == "object":
                chamados[col] = chamados[col].astype(str).replace("nan", "")

        # Cria aliases para preservar sua nomenclatura operacional na tela, mesmo quando a base veio com nomes diferentes.
        if "ETAPA SOLICITAÇÃO" not in chamados.columns:
            chamados["ETAPA SOLICITAÇÃO"] = chamados["Status Chamado"] if "Status Chamado" in chamados.columns else ""
        if "CLASSE DO CHAMADO" not in chamados.columns:
            chamados["CLASSE DO CHAMADO"] = chamados["Classe Chamado"] if "Classe Chamado" in chamados.columns else ""
        if "FILIAL" not in chamados.columns:
            chamados["FILIAL"] = chamados["Centro"] if "Centro" in chamados.columns else ""

        st.markdown(
            '<div class="selection-help"><b>Objetivo desta aba</b><br>'
            'Gerenciar a carteira de chamados antes da decisão do ativo, com indicadores por etapa, classe e filial, além do cruzamento por PEP com a CJI3.</div>',
            unsafe_allow_html=True,
        )

        f1, f2, f3, f4, f5, f6 = st.columns([0.9, 0.9, 1.1, 1.1, 1, 1.6])
        with f1:
            ch_emp = st.selectbox("Empresa", _unique_options(chamados, "Empresa"), key="ch_emp")
        with f2:
            ch_filial = st.selectbox("FILIAL", _unique_options(chamados, "FILIAL"), key="ch_filial")
        with f3:
            ch_etapa = st.selectbox("ETAPA SOLICITAÇÃO", _unique_options(chamados, "ETAPA SOLICITAÇÃO"), key="ch_etapa")
        with f4:
            ch_classe = st.selectbox("CLASSE DO CHAMADO", _unique_options(chamados, "CLASSE DO CHAMADO"), key="ch_classe")
        with f5:
            ch_centro = st.selectbox("Centro", _unique_options(chamados, "Centro"), key="ch_centro")
        with f6:
            ch_texto = st.text_input("Filtro geral", placeholder="PEP, chamado, NF, pedido, plaqueta, série...", key="ch_texto")

        ch_view = chamados.copy()
        if ch_emp and "Empresa" in ch_view.columns:
            ch_view = ch_view[ch_view["Empresa"].astype(str) == ch_emp]
        if ch_filial and "FILIAL" in ch_view.columns:
            ch_view = ch_view[ch_view["FILIAL"].astype(str) == ch_filial]
        if ch_etapa and "ETAPA SOLICITAÇÃO" in ch_view.columns:
            ch_view = ch_view[ch_view["ETAPA SOLICITAÇÃO"].astype(str) == ch_etapa]
        if ch_classe and "CLASSE DO CHAMADO" in ch_view.columns:
            ch_view = ch_view[ch_view["CLASSE DO CHAMADO"].astype(str) == ch_classe]
        if ch_centro and "Centro" in ch_view.columns:
            ch_view = ch_view[ch_view["Centro"].astype(str) == ch_centro]
        if ch_texto:
            nt = norm_text(ch_texto)
            mask = ch_view.astype(str).apply(lambda col: col.map(norm_text).str.contains(nt, na=False)).any(axis=1)
            ch_view = ch_view[mask]

        peps_cji3 = set()
        if not cji3_base.empty and "PEP" in cji3_base.columns:
            peps_cji3 = set(cji3_base["PEP"].dropna().astype(str).str.strip())
        if "PEP" in ch_view.columns:
            ch_view["Existe na CJI3"] = ch_view["PEP"].astype(str).str.strip().apply(lambda x: "Sim" if x in peps_cji3 else "Não")
        else:
            ch_view["Existe na CJI3"] = "Não avaliado"

        # Cards executivos
        m1, m2, m3, m4, m5, m6 = st.columns(6)
        m1.metric("Qtde de Registros", len(ch_view))
        m2.metric("PEPs distintos", ch_view["PEP"].astype(str).nunique() if "PEP" in ch_view else 0)
        m3.metric("Com CJI3", int((ch_view["Existe na CJI3"] == "Sim").sum()) if "Existe na CJI3" in ch_view else 0)
        m4.metric("Sem CJI3", int((ch_view["Existe na CJI3"] == "Não").sum()) if "Existe na CJI3" in ch_view else 0)
        m5.metric("Etapas", ch_view["ETAPA SOLICITAÇÃO"].astype(str).nunique() if "ETAPA SOLICITAÇÃO" in ch_view else 0)
        m6.metric("Filiais", ch_view["FILIAL"].astype(str).nunique() if "FILIAL" in ch_view else 0)

        st.markdown('<div class="prof-divider"></div>', unsafe_allow_html=True)
        st.markdown('<div class="section-label">Distribuição dos chamados</div>', unsafe_allow_html=True)

        def _ranking_chamados(df: pd.DataFrame, col: str, titulo: str) -> pd.DataFrame:
            if col not in df.columns or df.empty:
                return pd.DataFrame({titulo: [], "Qtde": []})
            out = (
                df[col].fillna("Não informado").astype(str).str.strip().replace("", "Não informado")
                .value_counts(dropna=False).reset_index()
            )
            out.columns = [titulo, "Qtde"]
            out["%"] = (out["Qtde"] / max(out["Qtde"].sum(), 1) * 100).round(2)
            return out

        rank_etapa = _ranking_chamados(ch_view, "ETAPA SOLICITAÇÃO", "ETAPA SOLICITAÇÃO")
        rank_classe = _ranking_chamados(ch_view, "CLASSE DO CHAMADO", "CLASSE DO CHAMADO")
        rank_filial = _ranking_chamados(ch_view, "FILIAL", "FILIAL")

        r1, r2, r3 = st.columns(3)
        with r1:
            st.markdown('<div class="asset-card"><b>Quantidade por ETAPA SOLICITAÇÃO</b>', unsafe_allow_html=True)
            st.dataframe(rank_etapa, use_container_width=True, height=260, hide_index=True)
            if not rank_etapa.empty:
                st.bar_chart(rank_etapa.set_index("ETAPA SOLICITAÇÃO")["Qtde"])
            st.markdown('</div>', unsafe_allow_html=True)
        with r2:
            st.markdown('<div class="asset-card"><b>Quantidade por CLASSE DO CHAMADO</b>', unsafe_allow_html=True)
            st.dataframe(rank_classe, use_container_width=True, height=260, hide_index=True)
            if not rank_classe.empty:
                st.bar_chart(rank_classe.set_index("CLASSE DO CHAMADO")["Qtde"])
            st.markdown('</div>', unsafe_allow_html=True)
        with r3:
            st.markdown('<div class="asset-card"><b>Quantidade por FILIAL</b>', unsafe_allow_html=True)
            st.dataframe(rank_filial, use_container_width=True, height=260, hide_index=True)
            if not rank_filial.empty:
                st.bar_chart(rank_filial.set_index("FILIAL")["Qtde"])
            st.markdown('</div>', unsafe_allow_html=True)

        st.markdown('<div class="prof-divider"></div>', unsafe_allow_html=True)
        col_prioridade = [
            "N_CHAMADO", "ETAPA SOLICITAÇÃO", "CLASSE DO CHAMADO", "FILIAL", "Empresa", "Centro", "PEP",
            "Segmento Chamado", "Tipo Chamado", "Inventário", "Série", "NF Chamado",
            "Pedido Chamado", "Existe na CJI3"
        ]
        col_prioridade = [c for c in col_prioridade if c in ch_view.columns]
        demais = [c for c in ch_view.columns if c not in col_prioridade]
        ch_cols = col_prioridade + demais

        st.markdown('<div class="section-label">Base analítica de chamados</div>', unsafe_allow_html=True)
        st.dataframe(ch_view[ch_cols], use_container_width=True, height=390, hide_index=True)

        st.divider()
        c1, c2 = st.columns([1, 2])
        with c1:
            pep_opts = _unique_options(ch_view, "PEP")
            pep_detalhe = st.selectbox("Analisar PEP do chamado", pep_opts, key="ch_pep_detalhe")
        with c2:
            st.caption("Ao escolher um PEP, a tela mostra o chamado e os respectivos lançamentos CJI3 encontrados para apoiar a decisão na próxima aba.")

        if pep_detalhe:
            detalhe_ch = ch_view[ch_view["PEP"].astype(str).str.strip() == pep_detalhe].copy() if "PEP" in ch_view.columns else pd.DataFrame()
            detalhe_cji3 = cji3_base[cji3_base["PEP"].astype(str).str.strip() == pep_detalhe].copy() if not cji3_base.empty and "PEP" in cji3_base.columns else pd.DataFrame()

            d1, d2 = st.columns([1, 1.4])
            with d1:
                st.write("**Dados do chamado selecionado**")
                cols_visiveis_tooltip = [c for c in ["N_CHAMADO", "ETAPA SOLICITAÇÃO", "CLASSE DO CHAMADO", "FILIAL", "Empresa", "Centro", "PEP", "Inventário", "Série", "NF Chamado", "Pedido Chamado", "Existe na CJI3"] if c in detalhe_ch.columns]
                st.markdown(_chamado_hover_table(detalhe_ch[ch_cols], cols_visiveis_tooltip, max_rows=80), unsafe_allow_html=True)
            with d2:
                st.write("**Lançamentos CJI3 vinculados ao PEP**")
                if detalhe_cji3.empty:
                    st.info("Nenhum lançamento CJI3 encontrado para este PEP.")
                else:
                    resumo_cji3 = [c for c in ["ID_CJI3", "Empresa", "Centro", "PEP", "Classe Custo", "Texto Pedido", "Texto Material", "Quantidade", "Valor", "Conta Contrapartida", "Fornecedor Nome", "Pedido", "NFe"] if c in detalhe_cji3.columns]
                    st.dataframe(detalhe_cji3[resumo_cji3], use_container_width=True, height=260, hide_index=True)
                    q = sum(to_number(v, 0) for v in detalhe_cji3.get("Quantidade", pd.Series(dtype=float)))
                    val = sum(to_number(v, 0) for v in detalhe_cji3.get("Valor", pd.Series(dtype=float)))
                    st.caption(f"Resumo CJI3 do PEP: {len(detalhe_cji3)} lançamento(s) | Quantidade {q:,.2f} | Valor {val:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

        st.divider()
        pacote_chamados = excel_bytes({
            "Chamados_Filtrados": ch_view,
            "Qtd_por_Etapa": rank_etapa,
            "Qtd_por_Classe": rank_classe,
            "Qtd_por_Filial": rank_filial,
            "Chamados_Original": chamados,
        })
        st.download_button(
            "Baixar gestão de chamados em Excel",
            data=pacote_chamados,
            file_name="Gestao_Chamados_MVP_V10_5.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

with tab_template:
    st.markdown('<div class="section-label">Saída SAP</div>', unsafe_allow_html=True)
    st.subheader("Template AS01 / Exportação padronizada")
    st.markdown(
        """
        <div class="selection-help">
        <b>Exportação do lote acumulado</b><br>
        Esta aba gera o relatório Excel padronizado para conferência e apoio à carga SAP AS01, usando as decisões salvas no lote acumulado.
        </div>
        """,
        unsafe_allow_html=True,
    )

    total_decisoes = len(st.session_state.decisoes)
    df_decisoes = pd.DataFrame(st.session_state.decisoes) if st.session_state.decisoes else pd.DataFrame()
    ids_utilizados = sorted(_ids_from_saved_decisions(st.session_state.decisoes)) if st.session_state.decisoes else []

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Decisões salvas", total_decisoes)
    c2.metric("IDs CJI3 tratados", len(ids_utilizados))
    c3.metric("Linhas CJI3 carregadas", len(st.session_state.cji3) if isinstance(st.session_state.cji3, pd.DataFrame) else 0)
    valor_dec = 0.0
    if not df_decisoes.empty and "Valor Capitalizado" in df_decisoes.columns:
        valor_dec = pd.to_numeric(df_decisoes["Valor Capitalizado"], errors="coerce").fillna(0).sum()
    c4.metric("Valor capitalizado", f"{valor_dec:,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))

    st.divider()
    p1, p2, p3 = st.columns([1.5, 1, 1])
    with p1:
        nome_arquivo = st.text_input("Nome do arquivo Excel", value="Template_AS01_MVP_V10_4.xlsx")
    with p2:
        somente_salvos = st.checkbox("Exportar somente status Salvo", value=True)
    with p3:
        incluir_pendencias = st.checkbox("Incluir pendências/não capitalizáveis", value=True)

    b1, b2 = st.columns([1, 3])
    with b1:
        if st.button("Atualizar relatório AS01", type="primary", use_container_width=True):
            st.rerun()
    with b2:
        st.caption("A prévia abaixo fica sempre ativa e é recalculada conforme os lotes/decisões salvas são atualizados.")

    if not st.session_state.decisoes:
        template = pd.DataFrame(columns=TEMPLATE_AS01)
        validacoes = _validar_template_as01(template)
        pendencias = pd.DataFrame()
        st.info("Nenhum item salvo ainda. O relatório AS01 permanecerá visível e será atualizado automaticamente quando você salvar decisões na Central de Decisão.")
    else:
        template = _build_template_as01_padronizado(st.session_state.decisoes, somente_salvos=somente_salvos, incluir_pendencias=incluir_pendencias)
        validacoes = _validar_template_as01(template)
        pendencias = template[template["Tratamento"].astype(str).str.contains("Pendência|Não capitalizável", case=False, na=False)].copy() if "Tratamento" in template.columns else pd.DataFrame()

    st.markdown("### Prévia do Template AS01 padronizado")
    st.caption("Esta visão fica sempre ativa para conferência. Quando não houver decisões salvas, apenas o cabeçalho padrão será exibido.")
    st.dataframe(template, use_container_width=True, height=380, hide_index=True)

    st.markdown("### Validações obrigatórias")
    st.dataframe(validacoes, use_container_width=True, height=220, hide_index=True)

    ids_df = pd.DataFrame({"ID_CJI3_Utilizado": ids_utilizados})
    param_df = _parametros_exportacao(nome_arquivo, somente_salvos, incluir_pendencias, template)
    pacote = excel_bytes(
        {
            "Template_AS01_Padronizado": template,
            "Validacoes_AS01": validacoes,
            "Decisoes_Completas": df_decisoes,
            "Pendencias": pendencias,
            "CJI3_IDs_Utilizados": ids_df,
            "Parametros_Exportacao": param_df,
            "Leia_me": _readme_exportacao(),
            "CJI3_Enriquecida": st.session_state.cji3_enriquecida.head(100000) if isinstance(st.session_state.cji3_enriquecida, pd.DataFrame) else pd.DataFrame(),
            "CJI3_Original_Padronizada": st.session_state.cji3.head(100000) if isinstance(st.session_state.cji3, pd.DataFrame) else pd.DataFrame(),
        }
    )
    st.download_button(
        "Baixar pacote Excel final padronizado",
        data=pacote,
        file_name=nome_arquivo if nome_arquivo.lower().endswith(".xlsx") else f"{nome_arquivo}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

    st.divider()
    st.write("**Bases carregadas disponíveis para auditoria:**")
    if st.session_state.bases:
        for nome, bdf in st.session_state.bases.items():
            with st.expander(f"{nome} - {len(bdf)} linhas"):
                st.dataframe(bdf.head(200), use_container_width=True, hide_index=True)
    else:
        st.info("Nenhuma base auxiliar carregada nesta sessão.")



with tab_cadastro:
    df = st.session_state.cji3_enriquecida.copy()
    if df.empty:
        st.warning("Carregue e processe a CJI3 na primeira aba.")
    else:
        st.markdown('<div class="section-label">Motor de decisão</div>', unsafe_allow_html=True)
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

        # Visão intuitiva para seleção: mostra primeiro as colunas que ajudam a decidir,
        # mas mantém uma aba com todas as colunas da CJI3 para auditoria.
        st.markdown(
            '<div class="selection-help"><b>Seleção dos itens CJI3</b><br>'
            'Use a visão resumida para selecionar rapidamente os lançamentos que receberão o mesmo tratamento. '
            'A visão completa continua disponível no modo “Todas as colunas”. Use filtros, status visual e painel permanente para trabalhar com grande volume sem perder rastreabilidade.</div>',
            unsafe_allow_html=True,
        )

        cji3_cols = [c for c in st.session_state.cji3.columns if c in view.columns]
        if "ID_CJI3" in cji3_cols:
            cji3_cols = ["ID_CJI3"] + [c for c in cji3_cols if c != "ID_CJI3"]

        # Colunas prioritárias para tomada de decisão operacional.
        cols_resumo = [
            "ID_CJI3", "Situação", "Empresa", "Centro", "FILIAL", "Segmento", "PEP",
            "Denominação objeto", "Classe Custo", "Descrição Classe Custo",
            "Material", "Texto Pedido", "Texto Material", "Quantidade", "Valor",
            "Conta Contrapartida", "Fornecedor Nome", "Pedido", "NFe", "Data de lançamento",
        ]
        cols_resumo = [c for c in cols_resumo if c in view.columns]
        if "ID_CJI3" not in cols_resumo:
            cols_resumo = ["ID_CJI3"] + cols_resumo

        # Filtros visuais adicionais para reduzir ruído antes da seleção.
        f1, f2, f3, f4, f5 = st.columns([1, 1, 1.1, 1.1, 1.5])
        with f1:
            emp_opts = [""] + sorted([str(x) for x in view.get("Empresa", pd.Series(dtype=str)).dropna().astype(str).unique()])
            filtro_emp = st.selectbox("Empresa", emp_opts, key="flt_emp_visual")
        with f2:
            centro_opts = [""] + sorted([str(x) for x in view.get("Centro", pd.Series(dtype=str)).dropna().astype(str).unique()])
            filtro_centro = st.selectbox("Centro", centro_opts, key="flt_centro_visual")
        with f3:
            classe_opts_filtro = [""] + sorted([str(x) for x in view.get("Classe Custo", pd.Series(dtype=str)).dropna().astype(str).unique()])
            filtro_classe_custo = st.selectbox("Classe custo", classe_opts_filtro, key="flt_classe_custo_visual")
        with f4:
            fornecedor_opts = [""] + sorted([str(x) for x in view.get("Conta Contrapartida", pd.Series(dtype=str)).dropna().astype(str).unique()])[:2000]
            filtro_fornec = st.selectbox("Fornecedor / contrap.", fornecedor_opts, key="flt_fornec_visual")
        with f5:
            modo_visao = st.radio("Modo", ["Seleção resumida", "Todas as colunas"], horizontal=True, key="modo_visao_cji3")

        view_sel = view.copy()
        if filtro_emp and "Empresa" in view_sel.columns:
            view_sel = view_sel[view_sel["Empresa"].astype(str) == filtro_emp]
        if filtro_centro and "Centro" in view_sel.columns:
            view_sel = view_sel[view_sel["Centro"].astype(str) == filtro_centro]
        if filtro_classe_custo and "Classe Custo" in view_sel.columns:
            view_sel = view_sel[view_sel["Classe Custo"].astype(str) == filtro_classe_custo]
        if filtro_fornec and "Conta Contrapartida" in view_sel.columns:
            view_sel = view_sel[view_sel["Conta Contrapartida"].astype(str) == filtro_fornec]

        if view_sel.empty:
            st.warning("Nenhum registro encontrado com os filtros visuais informados.")
            st.stop()

        ids_salvos = _ids_from_saved_decisions(st.session_state.decisoes)
        ids_lote_atual = set(int(x) for x in st.session_state.ids_selecionados)
        view_sel = view_sel.copy()
        view_sel["Situação"] = view_sel["ID_CJI3"].apply(lambda x: _status_item_cji3(x, ids_salvos, ids_lote_atual))

        cdisp1, cdisp2, cdisp3 = st.columns([1.2, 1.2, 3])
        with cdisp1:
            mostrar_salvos = st.checkbox("Mostrar itens já salvos", value=True, help="Itens salvos aparecem marcados como bloqueados para evitar duplicidade.")
        with cdisp2:
            somente_disp = st.checkbox("Ver somente disponíveis", value=False)
        with cdisp3:
            st.caption("A coluna Situação mostra se o ID_CJI3 está disponível, acumulado no lote atual ou já salvo/bloqueado.")

        if not mostrar_salvos:
            view_sel = view_sel[view_sel["Situação"] != "Salvo / bloqueado"]
        if somente_disp:
            view_sel = view_sel[view_sel["Situação"] == "Disponível"]

        if view_sel.empty:
            st.warning("Nenhum registro disponível após aplicar os filtros de situação.")
            st.stop()

        disp_mask = view_sel["Situação"] == "Disponível"
        lote_mask = view_sel["Situação"] == "No lote atual"
        salvo_mask = view_sel["Situação"] == "Salvo / bloqueado"
        valor_disp = sum(to_number(v, 0) for v in view_sel.loc[disp_mask, "Valor"]) if "Valor" in view_sel.columns else 0
        valor_salvo = sum(to_number(v, 0) for v in view_sel.loc[salvo_mask, "Valor"]) if "Valor" in view_sel.columns else 0
        _render_status_cards(len(view_sel), int(disp_mask.sum()), int(lote_mask.sum()), int(salvo_mask.sum()), valor_disp, valor_salvo)

        _render_lote_decisoes_panel(df, cji3_cols, "Painel permanente do lote e decisões")

        st.markdown(
            '<div class="legend">'
            '<span class="legend-item">🟢 Disponível</span>'
            '<span class="legend-item">🟡 No lote atual</span>'
            '<span class="legend-item">🔴 Já utilizado / bloqueado</span>'
            '</div>',
            unsafe_allow_html=True,
        )
        st.write("**Quadro CJI3 para seleção**")
        vlim1, vlim2 = st.columns([1, 4])
        with vlim1:
            limite_linhas_grade = st.selectbox("Linhas na grade", [500, 1000, 1500, 3000], index=2, help="Limite visual para manter a tela leve. Os filtros continuam atuando sobre a base carregada.")
        with vlim2:
            st.caption("Para bases muito grandes, filtre por PEP/empresa/centro antes de selecionar. Isso mantém a operação rápida e evita seleção indevida.")
        if modo_visao == "Seleção resumida":
            grid = view_sel[cols_resumo].copy()
        else:
            grid = view_sel[cji3_cols].copy()

        ids_atuais = set(int(x) for x in st.session_state.ids_selecionados)
        ids_salvos = _ids_from_saved_decisions(st.session_state.decisoes)
        if "Situação" not in grid.columns:
            grid.insert(1, "Situação", grid["ID_CJI3"].apply(lambda x: _status_item_cji3(x, ids_salvos, ids_atuais)))
        if "Status Visual" not in grid.columns:
            grid.insert(1, "Status Visual", grid["Situação"].apply(_status_badge_value))
        grid.insert(0, "Selecionar", grid.apply(lambda r: (int(r["ID_CJI3"]) in ids_atuais) and (r.get("Situação") != "Salvo / bloqueado") if pd.notna(r.get("ID_CJI3")) else False, axis=1))

        edited = st.data_editor(
            grid.head(int(limite_linhas_grade)),
            use_container_width=True,
            height=420,
            hide_index=True,
            disabled=[c for c in grid.columns if c != "Selecionar"],
            column_config={
                "Selecionar": st.column_config.CheckboxColumn("Sel.", help="Marque para adicionar ao tratamento", width="small"),
                "ID_CJI3": st.column_config.NumberColumn("ID", width="small"),
                "Status Visual": st.column_config.TextColumn("Status", help="Marcação visual da linha", width="medium"),
                "Situação": st.column_config.TextColumn("Situação", help="Disponível, no lote atual ou já salvo/bloqueado", width="medium"),
                "Empresa": st.column_config.TextColumn("Empresa", width="small"),
                "Centro": st.column_config.TextColumn("Centro", width="small"),
                "PEP": st.column_config.TextColumn("PEP", width="medium"),
                "Classe Custo": st.column_config.TextColumn("Classe Custo", width="medium"),
                "Texto Pedido": st.column_config.TextColumn("Texto Pedido", width="large"),
                "Texto Material": st.column_config.TextColumn("Texto Material", width="large"),
                "Quantidade": st.column_config.NumberColumn("Qtd", format="%.2f", width="small"),
                "Valor": st.column_config.NumberColumn("Valor", format="%.2f", width="medium"),
                "Conta Contrapartida": st.column_config.TextColumn("FORNEC", width="medium"),
                "Fornecedor Nome": st.column_config.TextColumn("Nome fornecedor", width="large"),
            },
            key="editor_cji3_selecao",
        )

        with st.expander("Mapa visual da CJI3 com marcação por cor", expanded=True):
            st.caption("Linhas verdes estão disponíveis; amarelas estão no lote atual; vermelhas já foram utilizadas e ficam bloqueadas para evitar duplicidade.")
            mapa_cols = [c for c in ["Status Visual", "Situação", "ID_CJI3", "Empresa", "Centro", "PEP", "Classe Custo", "Texto Pedido", "Texto Material", "Quantidade", "Valor", "Conta Contrapartida", "Fornecedor Nome", "Pedido", "NFe"] if c in grid.columns]
            mapa = grid[mapa_cols].copy().head(500)
            st.dataframe(mapa.style.apply(_style_status_rows, axis=1), use_container_width=True, height=260, hide_index=True)

        ids_marcados = []
        ids_bloqueados_marcados = []
        ids_salvos = _ids_from_saved_decisions(st.session_state.decisoes)
        if "Selecionar" in edited.columns and "ID_CJI3" in edited.columns:
            ids_marcados_raw = [int(x) for x in edited.loc[edited["Selecionar"] == True, "ID_CJI3"].dropna().tolist()]
            ids_marcados = [x for x in ids_marcados_raw if x not in ids_salvos]
            ids_bloqueados_marcados = [x for x in ids_marcados_raw if x in ids_salvos]

        a1, a2, a3, a4, a5 = st.columns([1.2, 1.3, 1.2, 1.4, 3])
        with a1:
            if st.button("Adicionar marcados", type="primary"):
                atual = set(st.session_state.ids_selecionados)
                atual.update(ids_marcados)
                st.session_state.ids_selecionados = sorted(atual)
                if ids_bloqueados_marcados:
                    st.warning(f"{len(ids_bloqueados_marcados)} ID(s) já estavam salvos e foram ignorados para evitar duplicidade.")
                st.rerun()
        with a2:
            if st.button("Adicionar todos disponíveis"):
                atual = set(st.session_state.ids_selecionados)
                ids_disponiveis_visao = [int(x) for x in view_sel.loc[view_sel["Situação"] == "Disponível", "ID_CJI3"].dropna().tolist()]
                atual.update(ids_disponiveis_visao)
                st.session_state.ids_selecionados = sorted(atual)
                st.rerun()
        with a3:
            if st.button("Remover marcados"):
                remover = set(ids_marcados)
                st.session_state.ids_selecionados = [x for x in st.session_state.ids_selecionados if int(x) not in remover]
                st.rerun()
        with a4:
            if st.button("Limpar selecionados"):
                st.session_state.ids_selecionados = []
                st.rerun()
        with a5:
            st.caption("Dica: marque vários itens com o checkbox, depois clique em Adicionar marcados. O quadro abaixo acumula o lote que receberá o tratamento.")

        selected_ids = [int(x) for x in st.session_state.ids_selecionados]
        ids_salvos_geral = _ids_from_saved_decisions(st.session_state.decisoes)
        # Segurança adicional: se um ID já salvo ficou acumulado no lote por uma ação anterior,
        # ele é removido antes de novo tratamento para evitar duplicidade.
        selected_ids = [x for x in selected_ids if x not in ids_salvos_geral]
        st.session_state.ids_selecionados = selected_ids
        selected_rows = df[df["ID_CJI3"].isin(selected_ids)].copy()

        st.write("**Quadro de itens selecionados para receber o tratamento**")
        if selected_rows.empty:
            st.info("Nenhum item acumulado ainda. Selecione IDs no quadro CJI3 acima e clique em 'Adicionar marcados'.")
            st.stop()
        else:
            lote_view = selected_rows[cji3_cols].copy()
            if "Situação" not in lote_view.columns:
                lote_view.insert(1, "Situação", lote_view["ID_CJI3"].apply(lambda x: _status_item_cji3(x, ids_salvos_geral, set(selected_ids))))
            if "Status Visual" not in lote_view.columns:
                lote_view.insert(1, "Status Visual", lote_view["Situação"].apply(_status_badge_value))
            st.dataframe(lote_view.style.apply(_style_status_rows, axis=1), use_container_width=True, height=220, hide_index=True)
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
            replicar = st.checkbox(
                "Replicar cadastro pela QTD",
                key=f"replicar_{lote_key}",
                help="Quando ativado, o cadastro salvo será repetido conforme o campo QTD. O valor capitalizado será dividido igualmente entre as linhas geradas.",
            )
            if replicar:
                n_rep = int(round(float(qtd or 0))) if qtd else 0
                valor_unitario_prev = (float(valor_cap or 0) / n_rep) if n_rep > 0 else 0
                st.caption(f"Serão geradas {n_rep} linha(s) cadastrada(s), com valor unitário aproximado de {valor_unitario_prev:,.2f}.".replace(",", "X").replace(".", ",").replace("X", "."))
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
                    "Replicar": "Sim" if replicar else "Não",
                }
                novas_decisoes = _gerar_decisoes_replicadas(decisao, replicar, qtd, valor_cap)
                st.session_state.decisoes = [
                    d for d in st.session_state.decisoes
                    if str(d.get("Origem_ID", "")).split("#R")[0] != lote_key
                ]
                st.session_state.decisoes.extend(novas_decisoes)
                st.session_state.ids_selecionados = [x for x in st.session_state.ids_selecionados if int(x) not in set(selected_ids)]
                st.success(f"Decisão salva para o lote de IDs: {lote_key}. Linhas geradas: {len(novas_decisoes)}. IDs marcados como salvos/bloqueados.")
                st.rerun()
        with ac2:
            if st.button("EXCLUIR LOTE"):
                antes = len(st.session_state.decisoes)
                st.session_state.decisoes = [
                    d for d in st.session_state.decisoes
                    if str(d.get("Origem_ID", "")).split("#R")[0] != lote_key
                ]
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
                    "Replicar": "Sim" if replicar else "Não",
                }
                novas_decisoes = _gerar_decisoes_replicadas(decisao, replicar, qtd, valor_cap)
                st.session_state.decisoes = [
                    d for d in st.session_state.decisoes
                    if str(d.get("Origem_ID", "")).split("#R")[0] != lote_key
                ]
                st.session_state.decisoes.extend(novas_decisoes)
                st.session_state.ids_selecionados = []
                st.success(f"Decisão salva para o lote de IDs: {lote_key}. Linhas geradas: {len(novas_decisoes)}.")
                st.rerun()

        st.subheader("Itens cadastrados / decisões salvas")
        st.caption("Esta é a mesma base do painel permanente acima. Mantida aqui para conferência ao final do cadastro do lote.")
        if st.session_state.decisoes:
            df_dec = pd.DataFrame(st.session_state.decisoes)
            ids_tratados = _ids_from_saved_decisions(st.session_state.decisoes)
            st.caption(f"Itens CJI3 já marcados como tratados: {len(ids_tratados)}. Esses IDs ficam bloqueados na seleção para evitar duplicidade.")
            st.dataframe(df_dec, use_container_width=True, height=260, hide_index=True)
        else:
            st.info("Nenhuma decisão salva ainda.")
