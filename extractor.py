import re
import unicodedata
from typing import Dict, Optional

import pandas as pd


SEMAFORO_RULES = {
    "Liquidez Corrente": {"green_min": 1.00, "yellow_min": 0.90, "direction": "high"},
    "Liquidez Reduzida": {"green_min": 0.80, "yellow_min": 0.60, "direction": "high"},
    "Autonomia Financeira": {"green_min": 0.30, "yellow_min": 0.20, "direction": "high"},
    "Endividamento": {"green_max": 0.70, "yellow_max": 0.80, "direction": "low"},
    "ROA": {"green_min": 0.02, "yellow_min": 0.00, "direction": "high"},
    "ROE": {"green_min": 0.05, "yellow_min": 0.00, "direction": "high"},
    "Margem Líquida": {"green_min": 0.03, "yellow_min": 0.00, "direction": "high"},
    "Margem EBITDA": {"green_min": 0.10, "yellow_min": 0.05, "direction": "high"},
}

TARGET_ROWS = {
    "numero_funcionarios": 189,
    "ebit": 278,
    "ebitda": 281,
    "ativo_corrente": 320,
    "inventarios": 329,
    "passivo_corrente": 361,
    "passivo_total": 363,
    "liquidez_corrente": 411,
    "liquidez_reduzida": 412,
    "autonomia_financeira": 417,
    "endividamento": 419,
}

FIELD_LABELS = {
    "numero_funcionarios": {
        "strict": [
            "numero de empregados",
            "número de empregados",
            "numero empregados",
            "número empregados",
            "numero de funcionarios",
            "número de funcionários",
            "employees",
            "average number of employees",
        ],
        "fallback": [
            "empregados",
            "funcionarios",
            "funcionários",
            "staff",
            "nofunc",
            "nºfunc",
        ],
    },
    "ebitda": {
        "strict": [
            "ebitda",
            "ebitda c/mep",
            "ebitda c mep",
        ],
        "fallback": [],
    },
    "ebit": {
        "strict": [
            "resultado operacional",
            "ebit",
            "operating result",
            "operating profit",
            "res.op",
        ],
        "fallback": [
            "resultados correntes",
        ],
    },
    "ativo_corrente": {
        "strict": [
            "total do activo corrente",
            "total do ativo corrente",
            "activo corrente",
            "ativo corrente",
            "current assets",
        ],
        "fallback": [],
    },
    "inventarios": {
        "strict": [
            "inventarios",
            "inventários",
            "existencias",
            "existências",
            "stocks",
            "inventories",
            "inventarios e ativos biologicos",
            "inventários e ativos biológicos",
        ],
        "fallback": [],
    },
    "passivo_corrente": {
        "strict": [
            "total do passivo corrente",
            "passivo corrente",
            "current liabilities",
            "debitos correntes",
            "débitos correntes",
        ],
        "fallback": [],
    },
    "passivo_total": {
        "strict": [
            "total do passivo",
            "passivo total",
            "total liabilities",
        ],
        "fallback": [
            "passivo",
            "liabilities",
        ],
    },
    "liquidez_corrente": {
        "strict": [
            "liquidez geral",
            "liquidez corrente",
            "current ratio",
        ],
        "fallback": [],
    },
    "liquidez_reduzida": {
        "strict": [
            "liquidez reduzida",
            "quick ratio",
        ],
        "fallback": [],
    },
    "autonomia_financeira": {
        "strict": [
            "autonomia financeira",
            "capital proprio / ativo",
            "capital próprio / ativo",
            "capital proprio / activo",
            "capital próprio / activo",
        ],
        "fallback": [],
    },
    "endividamento": {
        "strict": [
            "endividamento",
            "indebtedness",
            "passivo / ativo",
            "passivo / activo",
            "passivo / capital proprio",
            "passivo / capital próprio",
        ],
        "fallback": [],
    },
}


def normalize_text(value) -> str:
    if value is None:
        return ""
    text = str(value).strip().lower()
    text = unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("utf-8")
    text = text.replace("\n", " ")
    text = re.sub(r"\s+", " ", text)
    return text


def clean_number(value) -> Optional[float]:
    if value is None:
        return None

    if isinstance(value, (int, float)):
        if pd.isna(value):
            return None
        return float(value)

    text = str(value).strip()
    if not text:
        return None

    text = (
        text.replace("\xa0", " ")
        .replace("€", "")
        .replace("eur", "")
        .replace("EUR", "")
        .replace("%", "")
        .replace("m€", "")
        .replace("meur", "")
        .strip()
    )

    if text.startswith("(") and text.endswith(")"):
        text = "-" + text[1:-1]

    text = text.replace(" ", "")

    try:
        return float(text.replace(".", "").replace(",", "."))
    except Exception:
        pass

    try:
        return float(text.replace(",", ""))
    except Exception:
        return None


def row_to_text(row: pd.Series) -> str:
    vals = [str(v) for v in row.tolist() if pd.notna(v)]
    return " | ".join(vals)


def row_label_text(row: pd.Series) -> str:
    vals = []
    for v in row.tolist():
        if pd.isna(v):
            continue
        txt = normalize_text(v)
        if txt and clean_number(v) is None:
            vals.append(txt)
    return " | ".join(vals[:10])


def row_has_numeric_values(row: pd.Series, minimum: int = 1) -> bool:
    count = 0
    for v in row.tolist():
        if clean_number(v) is not None:
            count += 1
    return count >= minimum


def contains_term(label_text: str, term: str) -> bool:
    return term in label_text


def is_ratio_label(label: str) -> bool:
    ratio_markers = [
        "/",
        "%",
        "ratio",
        "margem",
        "liquidez",
        "autonomia financeira",
        "endividamento",
        "roe",
        "roa",
        "debitos correntes /",
        "débitos correntes /",
        "vendas /",
        "capital proprio /",
        "capital próprio /",
        "passivo /",
    ]
    return any(marker in label for marker in ratio_markers)


def find_company_info_xlsx(df: pd.DataFrame) -> Dict[str, Optional[str]]:
    nome = None
    nif = None
    cae = None
    cae_descricao = None
    ultimo_ano = None

    top_vals = [v for v in df.iloc[0].tolist() if pd.notna(v)]
    if top_vals:
        nome = str(top_vals[0]).strip()

    for i in range(min(120, len(df))):
        joined = row_to_text(df.iloc[i])
        norm = normalize_text(joined)

        if nif is None and "contribuinte" in norm:
            m = re.search(r"\b(\d{9})\b", joined)
            if m:
                nif = m.group(1)

        if ultimo_ano is None and ("ultimo ano disponivel" in norm or "ultima data de conta" in norm):
            years = [int(y) for y in re.findall(r"20\d{2}", joined)]
            years = [y for y in years if 2000 <= y <= 2100]
            if years:
                ultimo_ano = max(years)

    for i in range(min(200, len(df))):
        joined = row_to_text(df.iloc[i])
        norm = normalize_text(joined)

        if "codigo(s) cae rev.4" in norm or "codigo(s) cae rev.3" in norm:
            for j in range(i + 1, min(i + 15, len(df))):
                vals = [v for v in df.iloc[j].tolist() if pd.notna(v)]
                if len(vals) >= 2:
                    first = str(vals[0]).strip()
                    if re.fullmatch(r"\d{5}", first):
                        cae = first
                        cae_descricao = str(vals[-1]).strip()
                        break
            if cae:
                break

    return {
        "nome": nome,
        "nif": nif,
        "cae": cae,
        "cae_descricao": cae_descricao,
        "ultimo_ano_disponivel": ultimo_ano,
    }


def find_main_year_header_row(df: pd.DataFrame) -> int:
    best_row = None
    best_hits = 0

    for i in range(len(df)):
        joined = row_to_text(df.iloc[i])
        hits = len(re.findall(r"31/12/20\d{2}", joined))
        if hits > best_hits:
            best_hits = hits
            best_row = i

    if best_row is None or best_hits < 3:
        raise ValueError("Não foi possível localizar o cabeçalho principal dos anos no Excel.")

    return best_row


def build_year_map(df: pd.DataFrame, header_row: int) -> Dict[int, int]:
    year_map = {}
    row = df.iloc[header_row]

    for col_idx, value in enumerate(row.tolist()):
        if pd.isna(value):
            continue
        txt = normalize_text(value)
        m = re.search(r"(20\d{2})", txt)
        if m:
            year = int(m.group(1))
            if 2000 <= year <= 2100:
                year_map[col_idx] = year

    return year_map


def extract_values_around_year_columns(row: pd.Series, year_map: Dict[int, int]) -> Dict[int, float]:
    row_values = row.tolist()
    result = {}

    for year_col, year in year_map.items():
        found = None
        for offset in [0, 1, 2, 3, 4, 5]:
            idx = year_col - offset
            if 0 <= idx < len(row_values):
                value = clean_number(row_values[idx])
                if value is not None:
                    found = value
                    break
        if found is not None:
            result[year] = found

    return result


def find_row_flexible(
    df: pd.DataFrame,
    target_row: int,
    strict_labels,
    fallback_labels,
    tolerance: int = 12,
    allow_ratio: bool = True,
) -> Optional[int]:
    start = max(0, target_row - tolerance)
    end = min(len(df), target_row + tolerance + 1)

    def valid(i: int, labels) -> bool:
        label = row_label_text(df.iloc[i])
        if not allow_ratio and is_ratio_label(label):
            return False
        return any(contains_term(label, t) for t in labels) and row_has_numeric_values(df.iloc[i], 1)

    for i in range(start, end):
        if valid(i, strict_labels):
            return i

    for i in range(len(df)):
        if valid(i, strict_labels):
            return i

    for i in range(start, end):
        if valid(i, fallback_labels):
            return i

    for i in range(len(df)):
        if valid(i, fallback_labels):
            return i

    return None


def find_best_labeled_series(
    df: pd.DataFrame,
    year_map: Dict[int, int],
    patterns,
    prefer_lower=True,
    allow_ratio=False,
) -> Dict[int, float]:
    candidates = []
    for i in range(len(df)):
        label = row_label_text(df.iloc[i])
        if any(p in label for p in patterns) and row_has_numeric_values(df.iloc[i], 1):
            if not allow_ratio and is_ratio_label(label):
                continue
            series = extract_values_around_year_columns(df.iloc[i], year_map)
            if series:
                magnitude = max(abs(v) for v in series.values())
                candidates.append((i, magnitude))

    if not candidates:
        return {}

    if prefer_lower:
        candidates = sorted(candidates, key=lambda x: (x[0], x[1]), reverse=True)
    else:
        candidates = sorted(candidates, key=lambda x: (x[0], x[1]))

    best_row = candidates[0][0]
    return extract_values_around_year_columns(df.iloc[best_row], year_map)


def extract_vendas_series(df: pd.DataFrame, year_map: Dict[int, int]) -> Dict[int, float]:
    # 1) Volume de Negócios
    for i in range(len(df)):
        label = row_label_text(df.iloc[i])
        if ("volume de negocios" in label or "volume de negócios" in label) and row_has_numeric_values(df.iloc[i], 1):
            return extract_values_around_year_columns(df.iloc[i], year_map)

    # 2) Vendas e serviços prestados
    for i in range(len(df)):
        label = row_label_text(df.iloc[i])
        if "vendas e servicos prestados" in label or "vendas e serviços prestados" in label:
            if row_has_numeric_values(df.iloc[i], 1) and not is_ratio_label(label):
                return extract_values_around_year_columns(df.iloc[i], year_map)

            for j in range(i + 1, min(i + 12, len(df))):
                sublabel = row_label_text(df.iloc[j])
                if "vendas total" in sublabel and row_has_numeric_values(df.iloc[j], 1):
                    return extract_values_around_year_columns(df.iloc[j], year_map)
                if sublabel.strip() == "vendas" and row_has_numeric_values(df.iloc[j], 1):
                    return extract_values_around_year_columns(df.iloc[j], year_map)

    # 3) Vendas total
    for i in range(len(df)):
        label = row_label_text(df.iloc[i])
        if "vendas total" in label and row_has_numeric_values(df.iloc[i], 1):
            return extract_values_around_year_columns(df.iloc[i], year_map)

    # 4) Vendas
    for i in range(len(df)):
        label = row_label_text(df.iloc[i])
        if label.strip() == "vendas" and row_has_numeric_values(df.iloc[i], 1):
            return extract_values_around_year_columns(df.iloc[i], year_map)

    return {}


def extract_resultado_liquido_series(df: pd.DataFrame, year_map: Dict[int, int]) -> Dict[int, float]:
    """
    Prioridade:
    1. Resultado líquido do exercício / do período
    2. Resultado líquido
    Evita linhas de rácios, margens e percentagens.
    """
    strict_patterns = [
        "resultado liquido do exercicio",
        "resultado líquido do exercício",
        "resultado liquido do periodo",
        "resultado líquido do período",
        "resultado consolidado liquido",
        "resultado consolidado líquido",
    ]

    fallback_patterns = [
        "resultado liquido",
        "resultado líquido",
        "net income",
        "profit for the year",
    ]

    for i in range(len(df)):
        label = row_label_text(df.iloc[i])
        if any(p in label for p in strict_patterns):
            if row_has_numeric_values(df.iloc[i], 1) and not is_ratio_label(label):
                return extract_values_around_year_columns(df.iloc[i], year_map)

    for i in range(len(df)):
        label = row_label_text(df.iloc[i])
        if any(p in label for p in fallback_patterns):
            if row_has_numeric_values(df.iloc[i], 1) and not is_ratio_label(label):
                return extract_values_around_year_columns(df.iloc[i], year_map)

    return {}


def build_history_df(historical: Dict[str, Dict[int, float]]) -> pd.DataFrame:
    all_years = sorted({year for series in historical.values() for year in series.keys() if not str(year).startswith("_")})
    rows = []

    for year in all_years:
        rows.append(
            {
                "Ano": year,
                "ativo_total": historical.get("ativo_total", {}).get(year),
                "ativo_corrente": historical.get("ativo_corrente", {}).get(year),
                "inventarios": historical.get("inventarios", {}).get(year),
                "passivo_corrente": historical.get("passivo_corrente", {}).get(year),
                "passivo_total": historical.get("passivo_total", {}).get(year),
                "capital_proprio": historical.get("capital_proprio", {}).get(year),
                "vendas": historical.get("vendas", {}).get(year),
                "ebitda": historical.get("ebitda", {}).get(year),
                "ebit": historical.get("ebit", {}).get(year),
                "resultado_liquido": historical.get("resultado_liquido", {}).get(year),
                "numero_funcionarios": historical.get("numero_funcionarios", {}).get(year),
                "liquidez_corrente": historical.get("liquidez_corrente", {}).get(year),
                "liquidez_reduzida": historical.get("liquidez_reduzida", {}).get(year),
                "autonomia_financeira": historical.get("autonomia_financeira", {}).get(year),
                "endividamento": historical.get("endividamento", {}).get(year),
            }
        )

    return pd.DataFrame(rows)


def derive_missing_fields(latest_values: Dict[str, Optional[float]]) -> Dict[str, Optional[float]]:
    ativo_total = latest_values.get("ativo_total")
    capital_proprio = latest_values.get("capital_proprio")
    vendas = latest_values.get("vendas")

    autonomia = latest_values.get("autonomia_financeira")
    liquidez_corrente = latest_values.get("liquidez_corrente")
    liquidez_reduzida = latest_values.get("liquidez_reduzida")

    ratio_debitos_cp = latest_values.get("_ratio_debitos_correntes_capital")
    ratio_debitos_invent = latest_values.get("_ratio_debitos_correntes_inventarios")
    ratio_vendas_ativo_corrente = latest_values.get("_ratio_vendas_ativo_corrente")

    # 1) Ativo Total via autonomia financeira
    if ativo_total is None and capital_proprio is not None and autonomia not in (None, 0):
        autonomia_dec = autonomia / 100 if autonomia > 1 else autonomia
        if autonomia_dec not in (None, 0):
            latest_values["ativo_total"] = capital_proprio / autonomia_dec
            ativo_total = latest_values["ativo_total"]

    # 2) Passivo Total
    if latest_values.get("passivo_total") is None and ativo_total is not None and capital_proprio is not None:
        latest_values["passivo_total"] = ativo_total - capital_proprio

    # 3) Passivo Corrente via Débitos Correntes / Capital Próprio (%)
    if latest_values.get("passivo_corrente") is None and capital_proprio is not None and ratio_debitos_cp not in (None, 0):
        latest_values["passivo_corrente"] = capital_proprio * (ratio_debitos_cp / 100)

    # 4) Ativo Corrente via Liquidez Geral
    if latest_values.get("ativo_corrente") is None and latest_values.get("passivo_corrente") is not None and liquidez_corrente not in (None, 0):
        latest_values["ativo_corrente"] = latest_values["passivo_corrente"] * liquidez_corrente

    # 5) Ativo Corrente via Vendas / Ativo Corrente (%)
    if latest_values.get("ativo_corrente") is None and vendas is not None and ratio_vendas_ativo_corrente not in (None, 0):
        latest_values["ativo_corrente"] = vendas / (ratio_vendas_ativo_corrente / 100)

    # 6) Inventários via liquidez reduzida
    if latest_values.get("inventarios") is None:
        ac = latest_values.get("ativo_corrente")
        pc = latest_values.get("passivo_corrente")
        lr = latest_values.get("liquidez_reduzida")
        if ac is not None and pc is not None and lr is not None:
            latest_values["inventarios"] = ac - (lr * pc)

    # 7) Inventários via Débitos Correntes / Inventários (%)
    if latest_values.get("inventarios") is None and latest_values.get("passivo_corrente") is not None and ratio_debitos_invent not in (None, 0):
        latest_values["inventarios"] = latest_values["passivo_corrente"] / (ratio_debitos_invent / 100)

    # 8) Autonomia financeira
    if latest_values.get("autonomia_financeira") is None and latest_values.get("ativo_total") not in (None, 0) and capital_proprio is not None:
        latest_values["autonomia_financeira"] = capital_proprio / latest_values["ativo_total"]

    # 9) Endividamento
    if latest_values.get("endividamento") is None and latest_values.get("passivo_total") is not None and latest_values.get("ativo_total") not in (None, 0):
        latest_values["endividamento"] = latest_values["passivo_total"] / latest_values["ativo_total"]

    return latest_values


def extract_document(uploaded_file, filename: str) -> Dict:
    filename = filename.lower()

    if not filename.endswith(".xlsx"):
        raise ValueError("Esta versão da app aceita apenas ficheiros XLSX.")

    df = pd.read_excel(uploaded_file, sheet_name=0, header=None)
    company_info = find_company_info_xlsx(df)

    latest_year = company_info.get("ultimo_ano_disponivel")
    header_row = find_main_year_header_row(df)
    year_map = build_year_map(df, header_row)

    if not year_map:
        raise ValueError("Não foi possível construir o mapa de anos.")

    historical = {}
    debug_rows = {}

    # Campos principais com extração dedicada
    historical["vendas"] = extract_vendas_series(df, year_map)
    historical["ativo_total"] = find_best_labeled_series(
        df, year_map,
        ["total ativo", "total activo", "ativo total", "activo total", "total do ativo", "total do activo", "total assets"],
        prefer_lower=True,
        allow_ratio=False,
    )
    historical["capital_proprio"] = find_best_labeled_series(
        df, year_map,
        ["total capital proprio", "total capital próprio", "capital proprio", "capital próprio", "total do capital proprio", "total do capital próprio", "equity"],
        prefer_lower=True,
        allow_ratio=False,
    )
    historical["numero_funcionarios"] = find_best_labeled_series(
        df, year_map,
        ["numero de empregados", "número de empregados", "numero empregados", "número empregados", "numero de funcionarios", "número de funcionários", "employees"],
        prefer_lower=True,
        allow_ratio=False,
    )
    historical["ebit"] = find_best_labeled_series(
        df, year_map,
        ["ebit", "resultado operacional", "operating result", "operating profit", "res.op"],
        prefer_lower=True,
        allow_ratio=False,
    )
    historical["ebitda"] = find_best_labeled_series(
        df, year_map,
        ["ebitda c/mep", "ebitda c mep", "ebitda"],
        prefer_lower=True,
        allow_ratio=False,
    )
    historical["resultado_liquido"] = extract_resultado_liquido_series(df, year_map)

    debug_rows["vendas"] = "custom_search"
    debug_rows["ativo_total"] = "custom_search"
    debug_rows["capital_proprio"] = "custom_search"
    debug_rows["numero_funcionarios"] = "custom_search"
    debug_rows["ebit"] = "custom_search"
    debug_rows["ebitda"] = "custom_search"
    debug_rows["resultado_liquido"] = "custom_search"

    # Restantes campos
    for field, target_row in TARGET_ROWS.items():
        if field in historical:
            continue

        labels_cfg = FIELD_LABELS[field]
        matched_row = find_row_flexible(
            df,
            target_row,
            labels_cfg["strict"],
            labels_cfg["fallback"],
            tolerance=12,
            allow_ratio=field in {"liquidez_corrente", "liquidez_reduzida", "autonomia_financeira", "endividamento"},
        )
        debug_rows[field] = matched_row

        if matched_row is not None:
            historical[field] = extract_values_around_year_columns(df.iloc[matched_row], year_map)
        else:
            historical[field] = {}

    # Séries auxiliares para derivação
    historical["_ratio_vendas_ativo_corrente"] = find_best_labeled_series(
        df, year_map,
        ["vendas / activo corrente", "vendas / ativo corrente"],
        prefer_lower=True,
        allow_ratio=True,
    )
    historical["_ratio_debitos_correntes_capital"] = find_best_labeled_series(
        df, year_map,
        ["debitos correntes / capital proprio", "débitos correntes / capital próprio"],
        prefer_lower=True,
        allow_ratio=True,
    )
    historical["_ratio_debitos_correntes_inventarios"] = find_best_labeled_series(
        df, year_map,
        ["debitos correntes / inventarios", "débitos correntes / inventários", "debitos correntes / inventarios e ativos biologicos", "débitos correntes / inventários e ativos biológicos"],
        prefer_lower=True,
        allow_ratio=True,
    )

    all_years = sorted({
        year for series in historical.values()
        if isinstance(series, dict)
        for year in series.keys()
        if isinstance(year, int)
    })

    if latest_year is None and all_years:
        latest_year = max(all_years)

    if latest_year is None:
        raise ValueError("Não foi possível identificar o último ano disponível.")

    latest_values = {}
    for field, series in historical.items():
        if isinstance(series, dict):
            latest_values[field] = series.get(latest_year)

    latest_values = derive_missing_fields(latest_values)

    if latest_values.get("numero_funcionarios") is not None:
        company_info["numero_funcionarios"] = int(round(latest_values["numero_funcionarios"]))
    else:
        company_info["numero_funcionarios"] = None

    history_df = build_history_df({
        k: v for k, v in historical.items() if not k.startswith("_")
    })

    return {
        "source_type": "xlsx",
        "company_info": company_info,
        "latest_year": latest_year,
        "latest_values": latest_values,
        "historical": historical,
        "history_df": history_df,
        "debug": {
            "header_row": header_row,
            "year_map": year_map,
            "matched_rows": debug_rows,
        },
    }