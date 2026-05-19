from __future__ import annotations
import pandas as pd
from .schema import CJI3_MAP
from .utils import clean_columns, find_col, norm_text


def read_excel_any(file, header=0):
    return clean_columns(pd.read_excel(file, header=header))


def standardize_cji3(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    df = clean_columns(df)
    rename, log = {}, []
    for std, candidates in CJI3_MAP.items():
        col = find_col(df, candidates)
        if col:
            rename[col] = std
            log.append({"Base": "CJI3", "Campo padrão": std, "Coluna encontrada": col, "Status": "OK"})
        else:
            log.append({"Base": "CJI3", "Campo padrão": std, "Coluna encontrada": "", "Status": "Não encontrada"})
    out = df.rename(columns=rename).copy()
    for c in CJI3_MAP:
        if c not in out.columns:
            out[c] = pd.NA
    out.insert(0, "ID_CJI3", range(1, len(out) + 1))
    return out, pd.DataFrame(log)


def standardize_classes(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(df)
    rename = {}
    for std, cands in {
        "Classe": ["Classe"], "CAP Custo": ["CAP Custo"], "CAP Deprec": ["CAP Deprec"], "Conta SAP": ["Conta SAP"]
    }.items():
        col = find_col(df, cands)
        if col: rename[col] = std
    out = df.rename(columns=rename).copy()
    for c in ["Classe", "CAP Custo", "CAP Deprec", "Conta SAP"]:
        if c not in out.columns: out[c] = pd.NA
    return out


def standardize_localizacao(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(df)
    rename = {}
    mapping = {
        "Empresa": ["Empr", "Empresa"], "Centro": ["Cen.", "Centro"], "Local de negócios": ["Local de negócios", "Local de negocios"],
        "Nome da Filial": ["Nome da Filial"], "Centro de Lucro": ["Centro de Lucro", "Centro lucro"]
    }
    for std, cands in mapping.items():
        col = find_col(df, cands)
        if col: rename[col] = std
    out = df.rename(columns=rename).copy()
    for c in mapping:
        if c not in out.columns: out[c] = pd.NA
    return out


def standardize_empresas(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(df)
    rename = {}
    if find_col(df, ["Empr", "Empresa"]): rename[find_col(df, ["Empr", "Empresa"])] = "Empresa"
    if find_col(df, ["Nome", "Empresa Nome"]): rename[find_col(df, ["Nome", "Empresa Nome"])] = "Nome Empresa"
    out = df.rename(columns=rename).copy()
    for c in ["Empresa", "Nome Empresa"]:
        if c not in out.columns: out[c] = pd.NA
    return out


def standardize_chamados(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(df)
    rename = {}
    mapping = {
        "N_CHAMADO": ["Solicitação", "Solicitacao"], "Empresa": ["Empresa"], "Centro": ["Centro"], "PEP": ["Elemento PEP", "PEP"],
        "FILIAL": ["FILIAL", "Filial"],
        "Segmento Chamado": ["Segmento", "Segmento1"],
        "Classe Chamado": ["CLASSE DO CHAMADO", "Classe do chamado", "Classe Chamado", "Classificação do Imobilizado", "Classificacao do Imobilizado"],
        "ETAPA SOLICITAÇÃO": ["ETAPA SOLICITAÇÃO", "Etapa Solicitação", "Etapa Solicitacao", "Etapa", "Status", "Situação", "Situacao"],
        "Tipo Chamado": ["Tipo"], "Inventário": ["Nº Plaqueta", "N Plaqueta"], "Série": ["Nº Série", "N Serie"],
        "NF Chamado": ["Nota Fiscal", "NF", "Nº NF", "N. Nf-e"], "Pedido Chamado": ["Pedido", "Documento de compras"],
        "Status Chamado": ["Status", "Situação", "Situacao", "ETAPA SOLICITAÇÃO", "Etapa Solicitação", "Etapa Solicitacao"]
    }
    for std, cands in mapping.items():
        col = find_col(df, cands)
        if col: rename[col] = std
    out = df.rename(columns=rename).copy()
    for c in mapping:
        if c not in out.columns: out[c] = pd.NA
    return out


def standardize_razao(df: pd.DataFrame) -> pd.DataFrame:
    df = clean_columns(df)
    rename = {}
    mapping = {
        "Documento": ["Nº documento", "Documento"], "Pedido": ["Documento de compras", "Pedido"], "PEP": ["Elemento PEP", "PEP"],
        "Centro": ["Centro"], "Conta": ["Conta", "Conta do Razão", "Conta do Razao"], "Fornecedor Cod": ["Código Fornecedor", "Codigo Fornecedor"],
        "Fornecedor Nome": ["N_Fornec", "Fornecedor"], "Valor Razão": ["Montante em moeda interna", "Valor"], "Referência": ["Referência", "Referencia"]
    }
    for std, cands in mapping.items():
        col = find_col(df, cands)
        if col: rename[col] = std
    out = df.rename(columns=rename).copy()
    for c in mapping:
        if c not in out.columns: out[c] = pd.NA
    return out


def standardize_vida_util(file) -> pd.DataFrame:
    raw = pd.read_excel(file, header=None)
    # Procura linha onde começam os cabeçalhos: Classe / Conta SAP / Conta SAP Deprec
    header_idx = None
    for i in range(min(10, len(raw))):
        vals = [norm_text(x) for x in raw.iloc[i].tolist()]
        if "classe" in vals and any("conta sap" in v for v in vals):
            header_idx = i
            break
    if header_idx is None:
        df = clean_columns(pd.read_excel(file))
    else:
        header = raw.iloc[header_idx].tolist()
        df = raw.iloc[header_idx + 1:].copy()
        df.columns = [str(c).strip() for c in header]
        df = clean_columns(df)
    # Normaliza nomes básicos e remove vazios
    rename = {}
    if find_col(df, ["Classe"]): rename[find_col(df, ["Classe"])] = "Classe"
    if find_col(df, ["Conta SAP"]): rename[find_col(df, ["Conta SAP"])] = "Conta SAP"
    if find_col(df, ["Conta SAP Deprec", "Conta SAP Deprec."]): rename[find_col(df, ["Conta SAP Deprec", "Conta SAP Deprec."])] = "Conta SAP Deprec"
    out = df.rename(columns=rename).copy()
    out = out[out.get("Classe").notna()] if "Classe" in out.columns else out
    if "Classe" not in out.columns: out["Classe"] = pd.NA
    return out
