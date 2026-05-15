from __future__ import annotations
import math
import pandas as pd


def _num(v, default=0.0):
    try:
        if pd.isna(v):
            return default
        return float(v)
    except Exception:
        return default


def processar_rateio(df: pd.DataFrame) -> pd.DataFrame:
    """Gera linhas de ativo a partir das decisões.

    MVP: implementa tratamentos principais e mantém pendência quando falta decisão.
    """
    saidas = []
    for _, row in df.iterrows():
        tratamento = str(row.get("Tratamento", "Pendência de análise"))
        criterio = str(row.get("Critério Rateio", "Sem rateio"))
        qtd = _num(row.get("Quantidade"), 1)
        valor = _num(row.get("Valor"), 0)
        base = row.to_dict()

        if tratamento == "Novo ativo múltiplo por quantidade" and qtd > 1:
            n = int(math.floor(qtd))
            valor_unit = valor / n if n else valor
            for i in range(1, n + 1):
                item = base.copy()
                item["Subitem"] = i
                item["Qtd Capitalizada"] = 1
                item["Valor Capitalizado"] = round(valor_unit, 2)
                item["Status"] = "Processado"
                item["Observação Rateio"] = f"Linha dividida em {n} ativos por quantidade"
                saidas.append(item)
        elif tratamento in ["Servico incorporado ao ativo", "Ativo composto por agrupamento", "Ativo principal com componentes", "Complemento de ativo existente"]:
            item = base.copy()
            item["Subitem"] = 1
            item["Qtd Capitalizada"] = qtd if qtd else 1
            item["Valor Capitalizado"] = round(valor, 2)
            item["Status"] = "Aguardando vínculo/agrupamento" if tratamento != "Ativo composto por agrupamento" else "Processado"
            item["Observação Rateio"] = f"Tratamento conceitual: {tratamento} / Critério: {criterio}"
            saidas.append(item)
        elif tratamento == "Não capitalizável":
            item = base.copy()
            item["Subitem"] = 0
            item["Qtd Capitalizada"] = 0
            item["Valor Capitalizado"] = 0
            item["Status"] = "Excluído da capitalização"
            item["Observação Rateio"] = "Marcado como não capitalizável"
            saidas.append(item)
        elif tratamento == "Novo ativo unitário":
            item = base.copy()
            item["Subitem"] = 1
            item["Qtd Capitalizada"] = qtd if qtd else 1
            item["Valor Capitalizado"] = round(valor, 2)
            item["Status"] = "Processado"
            item["Observação Rateio"] = "Ativo unitário"
            saidas.append(item)
        else:
            item = base.copy()
            item["Subitem"] = 0
            item["Qtd Capitalizada"] = qtd
            item["Valor Capitalizado"] = valor
            item["Status"] = "Pendência de análise"
            item["Observação Rateio"] = "Tratamento não definido"
            saidas.append(item)
    return pd.DataFrame(saidas)
