from __future__ import annotations
import pandas as pd
from .schema import CAMPOS_TEMPLATE_SAP


def montar_template_sap(df: pd.DataFrame) -> pd.DataFrame:
    out = pd.DataFrame()
    out["Classe"] = df.get("Classe", df.get("Classe Custo", ""))
    out["Empresa"] = df.get("Empresa", "")
    out["Denominação"] = df.get("Descricao", "")
    out["Descr_02"] = df.get("PEP", "")
    out["Descr_03"] = df.get("Observação Rateio", "")
    out["Série"] = df.get("Série", "")
    out["Inventário"] = df.get("Inventário", "")
    out["Qtd"] = df.get("Qtd Capitalizada", df.get("Quantidade", ""))
    out["Un_Med"] = df.get("Unidade", "")
    out["C.C_SAP"] = df.get("Centro Custo", "")
    out["CENTRO"] = df.get("Centro", "")
    out["Segmento"] = df.get("Segmento", "")
    out["Código_FORNC"] = df.get("Fornecedor Cod", "")
    out["NOME_Fornecedor"] = df.get("Fornecedor Nome", "")
    out["Fabricante / Pep"] = df.get("PEP", "")
    out["Vida Útil"] = df.get("Vida Útil", "")
    out["Início Depreciação"] = df.get("Início Depreciação", "")
    pedido = df.get("Pedido", "").astype(str) if "Pedido" in df else ""
    nf = df.get("Nota Fiscal", "").astype(str) if "Nota Fiscal" in df else ""
    out["Nota Fiscal / Pedido"] = (nf + " / " + pedido) if hasattr(nf, "astype") else ""
    out["Valor Capitalizado"] = df.get("Valor Capitalizado", "")
    out["Tipo Decisão"] = df.get("Tipo Decisão", "")
    out["Tratamento"] = df.get("Tratamento", "")
    out["Critério Rateio"] = df.get("Critério Rateio", "")
    out["Status"] = df.get("Status", "")
    for col in CAMPOS_TEMPLATE_SAP:
        if col not in out.columns:
            out[col] = ""
    return out[CAMPOS_TEMPLATE_SAP]
