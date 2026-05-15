from __future__ import annotations
import pandas as pd
from .utils import norm_text, to_number


def _key(s):
    return s.astype(str).str.strip()


def enrich_cji3(cji3: pd.DataFrame, bases: dict[str, pd.DataFrame]) -> pd.DataFrame:
    df = cji3.copy()

    # Localização: Empresa + Centro
    loc = bases.get("localizacao")
    if loc is not None and not loc.empty:
        loc2 = loc.copy()
        df["_Empresa"] = _key(df["Empresa"])
        df["_Centro"] = _key(df["Centro"])
        loc2["_Empresa"] = _key(loc2["Empresa"])
        loc2["_Centro"] = _key(loc2["Centro"])
        cols = ["_Empresa", "_Centro", "Local de negócios", "Nome da Filial", "Centro de Lucro"]
        df = df.merge(loc2[[c for c in cols if c in loc2.columns]].drop_duplicates(["_Empresa", "_Centro"]), how="left", on=["_Empresa", "_Centro"])
        df.drop(columns=["_Empresa", "_Centro"], inplace=True, errors="ignore")

    # Empresas: Empresa
    emp = bases.get("empresas")
    if emp is not None and not emp.empty:
        emp2 = emp.copy()
        df["_Empresa"] = _key(df["Empresa"])
        emp2["_Empresa"] = _key(emp2["Empresa"])
        df = df.merge(emp2[["_Empresa", "Nome Empresa"]].drop_duplicates("_Empresa"), how="left", on="_Empresa")
        df.drop(columns=["_Empresa"], inplace=True, errors="ignore")

    # Chamados: PEP, com fallback primeira ocorrência
    chamados = bases.get("chamados")
    if chamados is not None and not chamados.empty:
        ch = chamados.copy()
        df["_PEP"] = _key(df["PEP"])
        ch["_PEP"] = _key(ch["PEP"])
        keep = ["_PEP", "N_CHAMADO", "Segmento Chamado", "Classe Chamado", "Tipo Chamado", "Inventário", "Série", "NF Chamado", "Pedido Chamado", "Status Chamado"]
        ch = ch[[c for c in keep if c in ch.columns]].dropna(subset=["_PEP"]).drop_duplicates("_PEP")
        df = df.merge(ch, how="left", on="_PEP")
        df.drop(columns=["_PEP"], inplace=True, errors="ignore")

    # Classes: Classe de custo -> Classe SAP
    classes = bases.get("classes")
    if classes is not None and not classes.empty:
        cl = classes.copy()
        df["_ClasseCusto"] = _key(df["Classe Custo"])
        # Tenta CAP Custo e Conta SAP como chave
        frames = []
        for chave in ["CAP Custo", "Conta SAP"]:
            if chave in cl.columns:
                tmp = cl.copy()
                tmp["_ClasseCusto"] = _key(tmp[chave])
                frames.append(tmp[["_ClasseCusto", "Classe", "CAP Custo", "CAP Deprec", "Conta SAP"]])
        if frames:
            mapa = pd.concat(frames, ignore_index=True).dropna(subset=["_ClasseCusto"]).drop_duplicates("_ClasseCusto")
            df = df.merge(mapa, how="left", on="_ClasseCusto")
        df.drop(columns=["_ClasseCusto"], inplace=True, errors="ignore")

    return df


def get_vida_util(row: pd.Series, vida: pd.DataFrame | None):
    if vida is None or vida.empty:
        return ""
    classe = str(row.get("Classe") or row.get("Classe Chamado") or "").strip()
    empresa = str(row.get("Empresa") or "").strip()
    if not classe:
        return ""
    v = vida.copy()
    v["_Classe"] = v["Classe"].astype(str).str.strip()
    hit = v[v["_Classe"] == classe]
    if hit.empty:
        return ""
    r = hit.iloc[0]
    # Empresas podem estar como int, float ou string nas colunas
    possible = [empresa, empresa.replace(".0", "")]
    for col in vida.columns:
        if str(col).strip() in possible or str(col).strip().replace(".0", "") in possible:
            val = r.get(col)
            if pd.notna(val): return val
    for col in ["Total Geral", "Total"]:
        if col in vida.columns and pd.notna(r.get(col)):
            return r.get(col)
    return ""


def build_prefill(row: pd.Series, vida: pd.DataFrame | None = None) -> dict:
    classe = row.get("Classe") if pd.notna(row.get("Classe", pd.NA)) else row.get("Classe Chamado", "")
    nf = row.get("NFe") if pd.notna(row.get("NFe", pd.NA)) else row.get("NF Remessa", "")
    pedido = row.get("Pedido") if pd.notna(row.get("Pedido", pd.NA)) else row.get("Pedido Chamado", "")
    denom = row.get("Texto Pedido") if pd.notna(row.get("Texto Pedido", pd.NA)) else row.get("Denominação objeto", "")
    descr2 = row.get("Texto Material") if pd.notna(row.get("Texto Material", pd.NA)) else row.get("Descrição Classe Custo", "")
    return {
        "Classe": "" if pd.isna(classe) else str(classe),
        "Empresa": "" if pd.isna(row.get("Empresa")) else str(row.get("Empresa")),
        "Denominação": "" if pd.isna(denom) else str(denom)[:50],
        "Descr_02": "" if pd.isna(descr2) else str(descr2)[:50],
        "Descr_03": "" if pd.isna(row.get("Denominação objeto")) else str(row.get("Denominação objeto"))[:50],
        "Série": "" if pd.isna(row.get("Série")) else str(row.get("Série")),
        "Inventário": "" if pd.isna(row.get("Inventário")) else str(row.get("Inventário")),
        "Qtd": to_number(row.get("Quantidade"), 1) or 1,
        "Un_Med": "UN",
        "C.C_SAP": "" if pd.isna(row.get("Centro de Lucro")) else str(row.get("Centro de Lucro")),
        "CENTRO": "" if pd.isna(row.get("Centro")) else str(row.get("Centro")),
        "Segmento": "" if pd.isna(row.get("Segmento Chamado", pd.NA)) else str(row.get("Segmento Chamado")),
        "Código_FORNC": "" if pd.isna(row.get("Fornecedor Cod", pd.NA)) else str(row.get("Fornecedor Cod")),
        "NOME_Fornecedor": "" if pd.isna(row.get("Fornecedor Nome", pd.NA)) else str(row.get("Fornecedor Nome")),
        "Fabricante / Pep": "" if pd.isna(row.get("PEP")) else str(row.get("PEP")),
        "Vida Útil": get_vida_util(row, vida),
        "Início Depreciação": "",
        "Nota Fiscal / Pedido": f"{'' if pd.isna(nf) else nf} / {'' if pd.isna(pedido) else pedido}".strip(" /"),
        "Valor Capitalizado": to_number(row.get("Valor"), 0),
    }


def decision_to_template(decisions: list[dict]) -> pd.DataFrame:
    from .schema import TEMPLATE_AS01
    df = pd.DataFrame(decisions)
    for c in TEMPLATE_AS01:
        if c not in df.columns:
            df[c] = ""
    return df[TEMPLATE_AS01]
