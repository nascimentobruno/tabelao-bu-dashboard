from __future__ import annotations

import json
import math
import unicodedata
from pathlib import Path
from datetime import datetime, date

import pandas as pd
import numpy as np

# =========================
# CONFIG
# =========================
EXCEL_FILE = Path("TABELAO_v1.0.xlsx")

DOCS_DIR = Path("docs")
DATA_DIR = DOCS_DIR / "data"
ASSETS_DIR = DOCS_DIR / "assets"

# abas azuis
SHEETS = {
    "imoveis": "Imoveis",
    "carbuy": "Carbuy",
    "veiculos": "Veiculos",
}

# leitura
HEADER_ROW = 1  # no seu excel os títulos estão na linha 2 (header=1)
MAX_ROWS_PER_JSON_PART = 2000  # como você disse: não passa de 2 mil por BU (por mês)

# se a coluna for "cot" estilo 7-jan, a gente precisa de um ano default
DEFAULT_YEAR_FOR_COT = 2026  # ajuste se mudar o arquivo para outro ano

# =========================
# HELPERS
# =========================
def norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c))


def fmt_date_ddmmyyyy(x) -> str:
    """Retorna dd/mm/aaaa ou ''."""
    if x is None:
        return ""
    try:
        if pd.isna(x):
            return ""
    except Exception:
        pass
    dt = pd.to_datetime(x, dayfirst=True, errors="coerce")
    if pd.isna(dt):
        return ""
    return dt.strftime("%d/%m/%Y")


def fmt_money(x) -> str:
    """R$ 1.234.567,89 | vazio/NaN -> R$ 0,00"""
    if x is None:
        return "R$ 0,00"
    try:
        if pd.isna(x):
            return "R$ 0,00"
    except Exception:
        pass

    s = str(x).strip()
    if s in ["", "-", "nan", "NaN", "None"]:
        return "R$ 0,00"

    try:
        # entende BR e EN
        if "," in s and "." in s:
            v = float(s.replace(".", "").replace(",", "."))
        elif "," in s:
            v = float(s.replace(",", "."))
        else:
            v = float(s)

        out = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {out}"
    except Exception:
        return "R$ 0,00"


def fmt_efficiency(x) -> str:
    """0.3333 -> 33,33% | 65,22 -> 65,22% | '-' -> '-' """
    if x is None:
        return "-"
    try:
        if pd.isna(x):
            return "-"
    except Exception:
        pass

    s = str(x).strip()
    if s in ["", "-", "nan", "NaN", "None"]:
        return "-"

    s = s.replace("%", "").replace(",", ".").strip()

    try:
        v = float(s)
        if 0 <= v <= 1:
            v *= 100
        return f"{v:.2f}".replace(".", ",") + "%"
    except Exception:
        return "-"


def make_json_safe_df(df: pd.DataFrame) -> pd.DataFrame:
    """
    Converte qualquer coisa não serializável em JSON para tipos simples:
    - datetime/Timestamp -> string ISO
    - NaN/NaT -> ""
    - numpy types -> python native
    """
    def conv(v):
        if v is None:
            return ""
        try:
            if pd.isna(v):
                return ""
        except Exception:
            pass

        if isinstance(v, (pd.Timestamp, datetime, date)):
            return v.isoformat()
        if isinstance(v, (np.integer,)):
            return int(v)
        if isinstance(v, (np.floating,)):
            fv = float(v)
            if math.isnan(fv):
                return ""
            return fv
        if isinstance(v, (np.bool_,)):
            return bool(v)
        return v

    return df.map(conv)


def split_into_parts(rows: list[dict], max_rows: int) -> list[list[dict]]:
    if max_rows <= 0:
        return [rows]
    return [rows[i:i + max_rows] for i in range(0, len(rows), max_rows)]


def detect_date_column(df: pd.DataFrame) -> str | None:
    """
    Detecta uma coluna de data de forma robusta:
    - tenta nomes conhecidos: Data / Data e Hora / cot etc
    - se não achar, tenta por "parece data" em amostra
    """
    if df is None or df.empty:
        return None

    cols_norm = {norm(c): c for c in df.columns}

    # prioridades
    preferred = ["data", "dataehora", "datahora", "dt", "cot"]
    candidates = [cols_norm[k] for k in preferred if k in cols_norm]

    if candidates:
        return candidates[0]

    # fallback: tenta achar coluna que converte bem em datetime
    for c in df.columns:
        s = df[c].astype(str).str.strip()
        sample = s.head(50)
        dt_try = pd.to_datetime(sample, dayfirst=True, errors="coerce")
        if dt_try.notna().sum() >= max(5, int(len(sample) * 0.3)):
            return c

    return None


def parse_datetime_series(df: pd.DataFrame, date_col: str) -> pd.Series:
    """
    Converte a coluna detectada em datetime.
    Regras:
    - Se a coluna for 'cot' (ex: '7-jan'), concatena '-ANO' antes de converter.
    - Caso contrário, to_datetime normal.
    """
    if not date_col:
        return pd.Series([pd.NaT] * len(df))

    if norm(date_col) == "cot":
        s = df[date_col].astype(str).str.strip()
        # exemplo: 7-jan -> 7-jan-2026
        s2 = s + f"-{DEFAULT_YEAR_FOR_COT}"
        return pd.to_datetime(s2, dayfirst=True, errors="coerce")

    return pd.to_datetime(df[date_col], dayfirst=True, errors="coerce")


# =========================
# MAIN
# =========================
def main():
    DOCS_DIR.mkdir(exist_ok=True)
    DATA_DIR.mkdir(exist_ok=True)
    ASSETS_DIR.mkdir(exist_ok=True)

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Não achei o Excel em: {EXCEL_FILE.resolve()}")

    # manifest (mês -> partes por BU)
    manifest = {
        "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "source": EXCEL_FILE.name,
        "months": [],
        "files": {bu: {} for bu in SHEETS.keys()},
    }

    months_set: set[str] = set()

    # limpa data antiga (opcional)
    (DATA_DIR / "manifest.json").unlink(missing_ok=True)

    for bu_key, sheet_name in SHEETS.items():
        # lê aba
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=HEADER_ROW)

        # remove Unnamed e colunas totalmente vazias
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        df = df.dropna(axis=1, how="all")

        # detecta coluna de data
        date_col = detect_date_column(df)

        if not date_col:
            # sem data: joga tudo em SEM_MES (mas ainda assim, remove linhas vazias completas)
            df = df.dropna(how="all").copy()
            if df.empty:
                continue
            df["Data"] = ""
            df["__month"] = "SEM_MES"
        else:
            dt = parse_datetime_series(df, date_col)

            # >>> ESTA LINHA É O QUE EVITA “1 MILHÃO DE LINHAS”
            # mantém SOMENTE as linhas que possuem data válida
            df = df[dt.notna()].copy()
            dt = dt[dt.notna()]

            if df.empty:
                continue

            # garante coluna Data dd/mm/aaaa
            df["Data"] = dt.dt.strftime("%d/%m/%Y")
            df["__month"] = dt.dt.strftime("%Y-%m")

        # formata por nome de coluna (normalizado)
        for col in list(df.columns):
            c = norm(col)
            if "eficiencia" in c:
                df[col] = df[col].map(fmt_efficiency)
            elif c in ("faturamento", "r$ estoque", "estoque r$", "rs estoque") or ("faturamento" in c) or ("r$" in c) or ("rs" in c):
                # cuidado: não queremos formatar qualquer coisa errada como dinheiro,
                # mas na sua planilha esses campos vêm com "R$" no nome ou são Faturamento
                df[col] = df[col].map(fmt_money)

        # garante json serializável
        df = make_json_safe_df(df)

        # separa por mês e grava em partes
        for month in sorted(df["__month"].dropna().unique().tolist()):
            month_df = df[df["__month"] == month].copy()
            if month_df.empty:
                continue

            months_set.add(str(month))

            rows = month_df.to_dict(orient="records")
            parts = split_into_parts(rows, MAX_ROWS_PER_JSON_PART)

            manifest["files"][bu_key].setdefault(str(month), [])

            for idx, part_rows in enumerate(parts, start=1):
                out_name = f"{bu_key}_{month}_part{idx}.json"
                out_path = DATA_DIR / out_name
                with open(out_path, "w", encoding="utf-8") as f:
                    json.dump(part_rows, f, ensure_ascii=False)

                manifest["files"][bu_key][str(month)].append({
                    "file": out_name,
                    "rows": len(part_rows),
                })

    # months no topo
    manifest["months"] = sorted(months_set)

    # escreve manifest
    with open(DATA_DIR / "manifest.json", "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print("OK - dashboard gerado (manifest + JSONs por mês/partes)")


if __name__ == "__main__":
    main()
