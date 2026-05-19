from __future__ import annotations
import re
import unicodedata
from io import BytesIO
import pandas as pd


def norm_text(value) -> str:
    if pd.isna(value):
        return ""
    s = str(value).strip().lower()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s)
    return s


def clean_columns(df: pd.DataFrame) -> pd.DataFrame:
    out = df.copy()
    out.columns = [str(c).strip() for c in out.columns]
    return out


def find_col(df: pd.DataFrame, candidates: list[str]) -> str | None:
    cols = {norm_text(c): c for c in df.columns}
    for cand in candidates:
        n = norm_text(cand)
        if n in cols:
            return cols[n]
    for cand in candidates:
        n = norm_text(cand)
        for k, real in cols.items():
            if n and (n in k or k in n):
                return real
    return None


def first_value(row, cols: list[str], default=""):
    for c in cols:
        if c in row.index and not pd.isna(row.get(c)) and str(row.get(c)).strip() != "":
            return row.get(c)
    return default


def to_number(v, default=0.0):
    try:
        if pd.isna(v):
            return default
        if isinstance(v, str):
            v = v.replace(".", "").replace(",", ".") if v.count(",") == 1 else v
        return float(v)
    except Exception:
        return default


def excel_bytes(sheets: dict[str, pd.DataFrame]) -> bytes:
    bio = BytesIO()
    with pd.ExcelWriter(bio, engine="xlsxwriter") as writer:
        for name, df in sheets.items():
            safe = name[:31]
            df.to_excel(writer, sheet_name=safe, index=False)
            wb = writer.book
            ws = writer.sheets[safe]
            header_fmt = wb.add_format({"bold": True, "bg_color": "#D9EAF7", "border": 1})
            money_fmt = wb.add_format({"num_format": "#,##0.00"})
            for i, col in enumerate(df.columns):
                ws.write(0, i, col, header_fmt)
                width = min(max(len(str(col)) + 2, 12), 38)
                if not df.empty:
                    try:
                        width = min(max(width, int(df[col].astype(str).str.len().quantile(0.90)) + 2), 45)
                    except Exception:
                        pass
                ws.set_column(i, i, width)
                if "valor" in norm_text(col) or "montante" in norm_text(col):
                    ws.set_column(i, i, width, money_fmt)
            ws.freeze_panes(1, 0)
            ws.autofilter(0, 0, max(len(df), 1), max(len(df.columns) - 1, 0))
    return bio.getvalue()
