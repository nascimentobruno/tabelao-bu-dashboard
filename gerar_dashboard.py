from __future__ import annotations

import html
import unicodedata
from pathlib import Path
from datetime import datetime
import pandas as pd

# =========================
# CONFIG
# =========================
EXCEL_FILE = Path("TABELAO_v1.0.xlsx")
DOCS_DIR = Path("docs")
ASSETS_CSS = "assets/style.css"

SHEETS = {
    "Imoveis": "Imoveis",
    "Carbuy": "Carbuy",
    "Veiculos": "Veiculos",
}

ROW_LIMIT = None  # None = todas as linhas


# =========================
# NORMALIZAÇÃO DE TEXTO
# =========================
def norm(s: str) -> str:
    """remove acentos + lowercase + trim"""
    s = "" if s is None else str(s)
    s = s.strip().lower()
    s = unicodedata.normalize("NFKD", s)
    s = "".join(ch for ch in s if not unicodedata.combining(ch))
    return s


# =========================
# FORMATADORES
# =========================
def fmt_date(x) -> str:
    if pd.isna(x) or str(x).strip() in ["", "-", "nan", "NaN"]:
        return ""
    dt = pd.to_datetime(x, dayfirst=True, errors="coerce")
    if pd.isna(dt):
        return ""
    return dt.strftime("%d/%m/%Y")


def fmt_money(x) -> str:
    """R$ 1.234.567,89 | vazio/NaN -> R$ 0,00"""
    if pd.isna(x):
        return "R$ 0,00"
    s = str(x).strip()
    if s in ["", "-", "nan", "NaN"]:
        return "R$ 0,00"
    try:
        # tenta entender BR e EN
        if "," in s and "." in s:
            # pode ser 1.234.567,89
            v = float(s.replace(".", "").replace(",", "."))
        elif "," in s:
            v = float(s.replace(",", "."))
        else:
            v = float(s)

        out = f"{v:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        return f"R$ {out}"
    except Exception:
        return s


def fmt_efficiency(x) -> str:
    """
    Eficiência SEMPRE em percentual:
    0.3333 -> 33,33%
    65.3 -> 65,30%
    '-' / vazio / NaN -> '-'
    """
    if pd.isna(x):
        return "-"
    s = str(x).strip()
    if s in ["", "-", "nan", "NaN"]:
        return "-"

    s2 = s.replace("%", "").strip().replace(",", ".")
    try:
        v = float(s2)
    except Exception:
        return "-"

    if 0 <= v <= 1:
        v *= 100

    return f"{v:.2f}".replace(".", ",") + "%"


def ym_from_date_str(date_str: str) -> str:
    """dd/mm/aaaa -> YYYY-MM (para filtro)"""
    if not date_str:
        return ""
    dt = pd.to_datetime(date_str, dayfirst=True, errors="coerce")
    if pd.isna(dt):
        return ""
    return f"{dt.year}-{str(dt.month).zfill(2)}"


def month_label_from_ym(ym: str) -> str:
    """YYYY-MM -> MM/YYYY"""
    try:
        yyyy, mm = ym.split("-")
        return f"{mm}/{yyyy}"
    except Exception:
        return ym


def safe_str(x) -> str:
    if pd.isna(x):
        return ""
    return str(x)


# =========================
# HTML TABLE
# =========================
def df_to_html_table(df: pd.DataFrame, months_set: set[str]) -> str:
    """
    Renderiza tabela e marca cada linha com data-month="YYYY-MM"
    """
    df = df.copy()

    if ROW_LIMIT is not None:
        df = df.head(int(ROW_LIMIT))

    df.columns = [safe_str(c).strip() for c in df.columns]
    has_data = "Data" in df.columns

    thead = "<thead><tr>" + "".join(
        f"<th>{html.escape(c)}</th>" for c in df.columns
    ) + "</tr></thead>"

    rows = []
    for _, row in df.iterrows():
        month_key = ""
        if has_data:
            month_key = ym_from_date_str(safe_str(row["Data"]))
            if month_key:
                months_set.add(month_key)

        tds = "".join(
            f"<td>{html.escape(safe_str(v))}</td>" for v in row.tolist()
        )

        if month_key:
            rows.append(f'<tr data-month="{html.escape(month_key)}">{tds}</tr>')
        else:
            rows.append(f"<tr>{tds}</tr>")

    tbody = "<tbody>" + "".join(rows) + "</tbody>"
    return f'<table class="data-table">{thead}{tbody}</table>'


# =========================
# HTML PAGE
# =========================
def build_html(excel_name: str, tables: dict[str, str], months_sorted: list[str], updated_str: str) -> str:
    tab_ids = {
        "Imoveis": "tab-imoveis",
        "Carbuy": "tab-carbuy",
        "Veiculos": "tab-veiculos",
    }

    # options do select
    options = ['<option value="ALL">Todos</option>']
    for ym in months_sorted:
        options.append(
            f'<option value="{html.escape(ym)}">{html.escape(month_label_from_ym(ym))}</option>'
        )
    options_html = "\n".join(options)

    return f"""<!doctype html>
<html lang="pt-br">
<head>
  <meta charset="utf-8" />
  <meta name="viewport" content="width=device-width,initial-scale=1" />
  <title>Tabelão • BU</title>
  <link rel="stylesheet" href="{ASSETS_CSS}">
</head>

<body>
  <div class="layout">
    <aside class="sidebar">
      <div class="brand">Tabelão • BU</div>

      <div class="month-filter">
        <label for="monthFilter">Mês</label>
        <select id="monthFilter" aria-label="Filtro de mês">
          {options_html}
        </select>
      </div>

      <button class="nav-btn active" type="button" data-tab="{tab_ids["Imoveis"]}">Imoveis</button>
      <button class="nav-btn" type="button" data-tab="{tab_ids["Carbuy"]}">Carbuy</button>
      <button class="nav-btn" type="button" data-tab="{tab_ids["Veiculos"]}">Veiculos</button>

      <div class="updated">
        Atualizado:
        <span>{updated_str}</span>
      </div>
    </aside>

    <main class="content">
      <header class="topbar">
        <div>
          <div class="title">Dashboard por BU</div>
          <div class="subtitle">Fonte: {html.escape(excel_name)}</div>
        </div>
      </header>

      <section class="card">
        <div class="table-wrap" id="{tab_ids["Imoveis"]}">
          {tables["Imoveis"]}
        </div>

        <div class="table-wrap hidden" id="{tab_ids["Carbuy"]}">
          {tables["Carbuy"]}
        </div>

        <div class="table-wrap hidden" id="{tab_ids["Veiculos"]}">
          {tables["Veiculos"]}
        </div>
      </section>

      <footer class="footer">
        <span>Gerado automaticamente via Python • pronto para GitHub Pages</span>
      </footer>
    </main>
  </div>

<script>
  const buttons = Array.from(document.querySelectorAll('.nav-btn'));
  const tabs = Array.from(document.querySelectorAll('.table-wrap'));
  const monthSelect = document.getElementById('monthFilter');

  function showTab(tabId) {{
    tabs.forEach(t => t.classList.add('hidden'));
    const el = document.getElementById(tabId);
    if (el) el.classList.remove('hidden');

    buttons.forEach(b => b.classList.remove('active'));
    const activeBtn = buttons.find(b => b.dataset.tab === tabId);
    if (activeBtn) activeBtn.classList.add('active');

    applyMonthFilter(monthSelect.value);
  }}

  function applyMonthFilter(monthValue) {{
    const allRows = document.querySelectorAll('table.data-table tbody tr');
    allRows.forEach(tr => {{
      const m = tr.getAttribute('data-month') || '';
      if (monthValue === 'ALL') {{
        tr.style.display = '';
      }} else {{
        tr.style.display = (m === monthValue) ? '' : 'none';
      }}
    }});
  }}

  buttons.forEach(btn => {{
    btn.addEventListener('click', (e) => {{
      e.preventDefault();
      e.stopPropagation();
      showTab(btn.dataset.tab);
    }});
  }});

  monthSelect.addEventListener('change', () => {{
    applyMonthFilter(monthSelect.value);
  }});

  // default
  showTab("{tab_ids["Imoveis"]}");
</script>
</body>
</html>
"""


def main() -> None:
    DOCS_DIR.mkdir(parents=True, exist_ok=True)
    (DOCS_DIR / "assets").mkdir(parents=True, exist_ok=True)

    if not EXCEL_FILE.exists():
        raise FileNotFoundError(f"Não achei o Excel em: {EXCEL_FILE.resolve()}")

    tables: dict[str, str] = {}
    months_found: set[str] = set()

    for label, sheet_name in SHEETS.items():
        # header=1: ignora a 1ª linha agrupadora (Negócio/CRM/Mídia)
        df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, header=1)

        # remove Unnamed e colunas vazias
        df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
        df = df.dropna(axis=1, how="all")

        # Data
        if "Data" in df.columns:
            df["Data"] = df["Data"].map(fmt_date)

        # Formatação por coluna (normalizando acentos)
        for col in df.columns:
            c_norm = norm(col)

            # eficiência (pega "Eficiência" também)
            if "eficiencia" in c_norm:
                df[col] = df[col].map(fmt_efficiency)

            # moeda
            elif ("r$" in c_norm) or (c_norm == "faturamento"):
                df[col] = df[col].map(fmt_money)

        tables[label] = df_to_html_table(df, months_found)

    months_sorted = sorted(months_found)

    updated_str = datetime.now().strftime("%d/%m/%Y %H:%M")
    html_out = build_html(EXCEL_FILE.name, tables, months_sorted, updated_str)

    out_path = DOCS_DIR / "index.html"
    out_path.write_text(html_out, encoding="utf-8")

    print("[OK] dashboard gerado (filtro de mês na sidebar)")


if __name__ == "__main__":
    main()
