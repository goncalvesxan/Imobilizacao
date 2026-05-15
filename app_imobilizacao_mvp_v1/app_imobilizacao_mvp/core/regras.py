from __future__ import annotations
from pathlib import Path
import pandas as pd
from .utils import normalizar_texto


def carregar_regras(caminho: str | Path = "config/regras_padrao.csv") -> pd.DataFrame:
    p = Path(caminho)
    if p.exists():
        return pd.read_csv(p)
    return pd.DataFrame(columns=["prioridade", "campo", "operador", "valor", "tipo_sugerido", "tratamento_sugerido", "criterio_rateio", "observacao"])


def aplicar_regras_linha(row: pd.Series, regras: pd.DataFrame) -> dict:
    regras_ordenadas = regras.sort_values("prioridade") if "prioridade" in regras.columns else regras
    for _, regra in regras_ordenadas.iterrows():
        campo = str(regra.get("campo", ""))
        operador = normalizar_texto(regra.get("operador", ""))
        valor_regra = normalizar_texto(regra.get("valor", ""))
        valor_campo = normalizar_texto(row.get(campo, ""))
        match = False
        if operador == "contem" and valor_regra in valor_campo:
            match = True
        elif operador == "igual" and valor_regra == valor_campo:
            match = True
        elif operador == "comeca_com" and valor_campo.startswith(valor_regra):
            match = True
        if match:
            return {
                "Tipo Decisão": regra.get("tipo_sugerido", "Pendente"),
                "Tratamento": regra.get("tratamento_sugerido", "Pendência de análise"),
                "Critério Rateio": regra.get("criterio_rateio", "Sem rateio"),
                "Regra Aplicada": regra.get("observacao", "Regra parametrizada")
            }
    return {"Tipo Decisão": "Pendente", "Tratamento": "Pendência de análise", "Critério Rateio": "Sem rateio", "Regra Aplicada": "Sem regra encontrada"}


def sugerir_decisoes(df: pd.DataFrame, regras: pd.DataFrame) -> pd.DataFrame:
    sugestoes = df.apply(lambda r: aplicar_regras_linha(r, regras), axis=1, result_type="expand")
    return pd.concat([df.reset_index(drop=True), sugestoes.reset_index(drop=True)], axis=1)
