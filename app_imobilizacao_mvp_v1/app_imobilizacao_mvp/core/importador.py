from __future__ import annotations
import pandas as pd
from .schema import COLUNAS_CJI3_CANDIDATAS
from .utils import localizar_coluna


def padronizar_cji3(df: pd.DataFrame) -> tuple[pd.DataFrame, pd.DataFrame]:
    """Renomeia colunas conhecidas da CJI3 para um padrão interno."""
    renomear = {}
    log = []
    for campo_padrao, candidatos in COLUNAS_CJI3_CANDIDATAS.items():
        col = localizar_coluna(df, candidatos)
        if col:
            renomear[col] = campo_padrao
            log.append({"Campo Padrão": campo_padrao, "Coluna Encontrada": col, "Status": "OK"})
        else:
            log.append({"Campo Padrão": campo_padrao, "Coluna Encontrada": "", "Status": "Não encontrada"})
    out = df.rename(columns=renomear).copy()
    for col in COLUNAS_CJI3_CANDIDATAS:
        if col not in out.columns:
            out[col] = pd.NA
    out["ID_Linha"] = range(1, len(out) + 1)
    return out, pd.DataFrame(log)
