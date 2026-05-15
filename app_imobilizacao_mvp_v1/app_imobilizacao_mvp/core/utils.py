from __future__ import annotations
import re
import unicodedata
from io import BytesIO
from pathlib import Path
import pandas as pd


def normalizar_texto(txt: object) -> str:
    if pd.isna(txt):
        return ""
    s = str(txt).strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s)
    return s


def localizar_coluna(df: pd.DataFrame, candidatos: list[str]) -> str | None:
    mapa = {normalizar_texto(c): c for c in df.columns}
    for cand in candidatos:
        n = normalizar_texto(cand)
        if n in mapa:
            return mapa[n]
    for cand in candidatos:
        n = normalizar_texto(cand)
        for col_norm, col_real in mapa.items():
            if n in col_norm or col_norm in n:
                return col_real
    return None


def excel_bytes_abas(abas: dict[str, pd.DataFrame]) -> bytes:
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for nome, df in abas.items():
            safe = nome[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
            ws = writer.sheets[safe]
            for idx, col in enumerate(df.columns):
                largura = min(max(len(str(col)) + 2, 12), 38)
                if not df.empty:
                    largura = min(max(largura, int(df[col].astype(str).str.len().quantile(0.90)) + 2), 45)
                ws.set_column(idx, idx, largura)
    return output.getvalue()


def carregar_excel(uploaded_file, sheet_name=0) -> pd.DataFrame:
    return pd.read_excel(uploaded_file, sheet_name=sheet_name)


def garantir_pasta(caminho: str | Path) -> Path:
    p = Path(caminho)
    p.mkdir(parents=True, exist_ok=True)
    return p
