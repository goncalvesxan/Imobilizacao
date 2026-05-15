from __future__ import annotations
import pandas as pd

COLUNAS_MINIMAS = ["Empresa", "Centro", "PEP", "Descricao", "Classe Custo", "Quantidade", "Valor"]


def validar_base_cji3(df: pd.DataFrame) -> pd.DataFrame:
    logs = []
    for col in COLUNAS_MINIMAS:
        if col not in df.columns:
            logs.append({"Tipo": "Estrutura", "Campo": col, "Problema": "Coluna ausente", "Severidade": "Alta"})
        elif df[col].isna().all():
            logs.append({"Tipo": "Conteúdo", "Campo": col, "Problema": "Coluna sem dados", "Severidade": "Alta"})
        elif df[col].isna().any():
            logs.append({"Tipo": "Conteúdo", "Campo": col, "Problema": f"{int(df[col].isna().sum())} linhas vazias", "Severidade": "Média"})
    if "Valor" in df.columns:
        valores = pd.to_numeric(df["Valor"], errors="coerce")
        if valores.isna().any():
            logs.append({"Tipo": "Conteúdo", "Campo": "Valor", "Problema": "Há valores não numéricos", "Severidade": "Alta"})
    if "Quantidade" in df.columns:
        qtd = pd.to_numeric(df["Quantidade"], errors="coerce")
        if qtd.isna().any():
            logs.append({"Tipo": "Conteúdo", "Campo": "Quantidade", "Problema": "Há quantidades não numéricas", "Severidade": "Média"})
    return pd.DataFrame(logs)
