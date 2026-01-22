from __future__ import annotations

import json
import unicodedata
from pathlib import Path
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP

import pandas as pd
import numpy as np

# =========================
# CONFIG
# =========================
EXCEL_FILE = Path("TABELAO_v1.0.xlsx")

DOCS_DIR = Path("docs")
DATA_DIR = DOCS_DIR / "data"
ASSETS_DIR = DOCS_DIR / "assets"

SHEETS = {
    "imoveis": "Imoveis",
    "carbuy": "Carbuy",
    "veiculos": "Veiculos",
}

ROWS_PER_PART = 2000  # margem segura


# =========================
# HELPERS
# =========================
def norm(s: str) -> str:
    s = "" if s is None else str(s)
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s)
    return "".join(c for c in s if not unicodedata.combining(c))


def fmt_date(x) -> str:
    """Normaliza datas para dd/mm/yyyy."""
    if pd.isna(x):
        return ""
    dt = pd.to_datetime(x, dayfirst=True, errors="coerce")
    if pd.isna(dt):
        return ""
    return dt.strftime("%d/%m/%Y")


def ym_from_date_str(date_str: str) -> str:
    """Extrai YYYY-MM de uma data dd/mm/yyyy."""
    if not date_str:
        return ""
    dt = pd.to_datetime(date_str, dayfirst=True, errors="coerce")
    if pd.isna(dt):
        return ""
    return f"{dt.year}-{str(dt.month).zfill(2)}"


def as_decimal_money(value):
    """
    Converte valor vindo do Excel (float/int/str) para Decimal com 2 casas.
    Evita o bug do 5182575.239999999 virar 5.182.575.239.999.999,00.
    Retorna Decimal ou None.
    """
    if value is None or value == "" or value == "-":
        return None

    # Numérico direto
    if isinstance(value, (int, float, Decimal, np.integer, np.floating)):
        return Decimal(str(value)).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)

    # Texto "R$ 5.182.575,24" ou "5.182.575,24"
    if isinstance(value, str):
        s = value.strip().replace("R$", "").strip()
        if not s:
            return None
        s = s.replace(".", "").replace(",", ".")
        try:
            return Decimal(s).quantize(Decimal("0.01"), rounding=ROUND_HALF_UP)
        except Exception:
            return None

    return None


def fmt_efficiency(x):
    """
    Mantém eficiência como string legível:
    - se vier 0-1 => vira 0-100%
    - se vier 0-100 => mantém
    """
    if pd.isna(x) or x == "":
        return "-"
    try:
        v = float(str(x).replace("%", "").replace(",", "."))
        if 0 <= v <= 1:
            v *= 100
        return f"{v:.2f}".replace(".", ",") + "%"
    except Exception:
        return "-"


def json_safe(df: pd.DataFrame) -> pd.DataFrame:
    """Converte tipos pandas/numpy para JSON-friendly."""
    def conv(v):
        if pd.isna(v):
            return ""
        if isinstance(v, (pd.Timestamp, datetime)):
            return v.isoformat()
        if isinstance(v, np.integer):
            return int(v)
        if isinstance(v, np.floating):
            return float(v)
        return v

    return df.applymap(conv)


def is_money_col(col_name: str) -> bool:
    c = norm(col_name)
    # cobre seus casos: "R$ Estoque" e "Faturamento"
    return ("r$" in c) or ("faturamento" in c)


# =========================
# MAIN
# =========================
def main():
    DOCS_DIR.mkdir(exist_ok=True)
    DATA_DIR.mkdir(exist_ok=True)
    ASSETS_DIR.mkdir(exist_ok=True)

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(EXCEL_FILE)

    manifest = {
        "generated_at": datetime.now().strftime("%d/%m/%Y %H:%M"),
        "source": EXCEL_FILE.name,
        "months": [],
        "files": {}
    }

    all_months = set()

    for bu, sheet in SHEETS.items():
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet, header=1)

        # remove colunas Unnamed / vazias
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        df = df.dropna(axis=1, how="all")

        if "Data" not in df.columns:
            continue

        # data + mês
        df["Data"] = df["Data"].map(fmt_date)
        df["__month"] = df["Data"].map(ym_from_date_str)

        # 🚫 remove linhas SEM MÊS
        df = df[df["__month"] != ""]

        if df.empty:
            continue

        # formata colunas (CORREÇÃO: moeda vira NÚMERO no JSON, não string "R$ ...")
        for col in df.columns:
            if col == "__month":
                continue

            c = norm(col)

            if "eficiencia" in c:
                df[col] = df[col].map(fmt_efficiency)

            elif is_money_col(col):
                # transforma em float com 2 casas (via Decimal), sem explodir
                def _money_to_float(v):
                    d = as_decimal_money(v)
                    return float(d) if d is not None else 0.0

                df[col] = df[col].map(_money_to_float)

        df = json_safe(df)

        manifest["files"][bu] = {}

        for month, g in df.groupby("__month"):
            all_months.add(month)

            out_dir = DATA_DIR / bu
            out_dir.mkdir(exist_ok=True)

            rows = g.drop(columns="__month").to_dict(orient="records")

            parts = [
                rows[i:i + ROWS_PER_PART]
                for i in range(0, len(rows), ROWS_PER_PART)
            ]

            manifest["files"][bu][month] = []

            for i, part in enumerate(parts, start=1):
                fname = f"{month}_part{i}.json"

                with open(out_dir / fname, "w", encoding="utf-8") as f:
                    json.dump(part, f, ensure_ascii=False)

                # ✅ melhor: manifest guarda só o nome do arquivo
                # (seu index já monta data/<BU>/<file>)
                manifest["files"][bu][month].append({
                    "file": fname,
                    "rows": len(part)
                })

    manifest["months"] = sorted(all_months)

    with open(DATA_DIR / "manifest.json", "w", encoding="utf-8") as f:
        json.dump(manifest, f, ensure_ascii=False, indent=2)

    print("OK — dashboard gerado (SEM SEM_MES) e moeda numérica (sem bug de float)")


if __name__ == "__main__":
    main()
