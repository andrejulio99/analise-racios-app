import math
from pathlib import Path

import pandas as pd
import streamlit as st

from extractor import extract_document, SEMAFORO_RULES

st.set_page_config(
    page_title="CFA | Análise de Rácios Financeiros",
    page_icon="📊",
    layout="wide",
)

# =========================
# Helpers
# =========================
def fmt_number(value):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    try:
        return f"{float(value):,.0f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(value)


def fmt_decimal(value, digits=3):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    return f"{value:.{digits}f}"


def fmt_percent(value, digits=2):
    if value is None or (isinstance(value, float) and math.isnan(value)):
        return "—"
    return f"{value * 100:.{digits}f}%"


def safe_div(a, b):
    if a is None or b in (None, 0):
        return None
    return a / b


def normalize_percent_ratio(x):
    if x is None:
        return None
    if x > 1.5:
        return x / 100
    return x


def calculate_ratios(values: dict) -> dict:
    ativo_total = values.get("ativo_total")
    ativo_corrente = values.get("ativo_corrente")
    inventarios = values.get("inventarios")
    passivo_corrente = values.get("passivo_corrente")
    passivo_total = values.get("passivo_total")
    capital_proprio = values.get("capital_proprio")
    vendas = values.get("vendas")
    ebitda = values.get("ebitda")
    resultado_liquido = values.get("resultado_liquido")

    liquidez_corrente = values.get("liquidez_corrente")
    liquidez_reduzida = values.get("liquidez_reduzida")
    autonomia_financeira = values.get("autonomia_financeira")
    endividamento = values.get("endividamento")

    if liquidez_corrente is None:
        liquidez_corrente = safe_div(ativo_corrente, passivo_corrente)

    if liquidez_reduzida is None:
        if ativo_corrente is not None and inventarios is not None and passivo_corrente not in (None, 0):
            liquidez_reduzida = (ativo_corrente - inventarios) / passivo_corrente

    if autonomia_financeira is None:
        autonomia_financeira = safe_div(capital_proprio, ativo_total)
    else:
        autonomia_financeira = normalize_percent_ratio(autonomia_financeira)

    if endividamento is None:
        endividamento = safe_div(passivo_total, ativo_total)
    else:
        endividamento = normalize_percent_ratio(endividamento)

    return {
        "Liquidez Corrente": liquidez_corrente,
        "Liquidez Reduzida": liquidez_reduzida,
        "Autonomia Financeira": autonomia_financeira,
        "Endividamento": endividamento,
        "ROA": safe_div(resultado_liquido, ativo_total),
        "ROE": safe_div(resultado_liquido, capital_proprio),
        "Margem Líquida": safe_div(resultado_liquido, vendas),
        "Margem EBITDA": safe_div(ebitda, vendas),
    }


def interpret_ratio(name, value):
    if value is None:
        return "Dados insuficientes."

    if name == "Liquidez Corrente":
        if value < 1:
            return "Capacidade curta de cumprir obrigações de curto prazo."
        elif value <= 2:
            return "Situação equilibrada."
        return "Liquidez confortável."

    if name == "Liquidez Reduzida":
        if value < 0.6:
            return "Cobertura fraca sem inventários."
        elif value < 0.8:
            return "Situação intermédia."
        return "Boa cobertura sem inventários."

    if name == "Autonomia Financeira":
        if value < 0.20:
            return "Dependência elevada de capitais alheios."
        elif value < 0.30:
            return "Solidez intermédia."
        return "Boa solidez financeira."

    if name == "Endividamento":
        if value > 0.80:
            return "Nível de endividamento elevado."
        elif value > 0.70:
            return "Nível intermédio."
        return "Nível controlado."

    if name in ["ROA", "ROE", "Margem Líquida", "Margem EBITDA"]:
        if value < 0:
            return "Rentabilidade negativa."
        elif value < 0.02:
            return "Rentabilidade baixa."
        elif value < 0.10:
            return "Rentabilidade aceitável."
        return "Rentabilidade forte."

    return "Sem interpretação definida."


def semaforo(name, value):
    if value is None:
        return "⚪", "Sem dados"

    rule = SEMAFORO_RULES.get(name)
    if not rule:
        return "⚪", "Sem regra"

    if rule["direction"] == "high":
        if value >= rule["green_min"]:
            return "🟢", "Bom"
        if value >= rule["yellow_min"]:
            return "🟡", "Atenção"
        return "🔴", "Crítico"

    if rule["direction"] == "low":
        if value <= rule["green_max"]:
            return "🟢", "Bom"
        if value <= rule["yellow_max"]:
            return "🟡", "Atenção"
        return "🔴", "Crítico"

    return "⚪", "Sem regra"


def build_ratio_table(ratios: dict) -> pd.DataFrame:
    rows = []
    for name, value in ratios.items():
        icon, estado = semaforo(name, value)

        if name in [
            "Autonomia Financeira",
            "Endividamento",
            "ROA",
            "ROE",
            "Margem Líquida",
            "Margem EBITDA",
        ]:
            valor_txt = fmt_percent(value)
        else:
            valor_txt = fmt_decimal(value) if value is not None else "—"

        rows.append(
            {
                "Semáforo": icon,
                "Estado": estado,
                "Rácio": name,
                "Valor": valor_txt,
                "Interpretação": interpret_ratio(name, value),
            }
        )
    return pd.DataFrame(rows)


def editable_number_input(label, value, key):
    default = 0.0 if value is None else float(value)
    return st.number_input(label, value=default, step=1.0, format="%.2f", key=key)


# =========================
# Branding / Visual
# =========================
logo_path = Path("cfa-vp-fundo-820.png")

st.markdown(
    """
<style>
:root {
    --cfa-bg: #f5f3ef;
    --cfa-surface: #ffffff;
    --cfa-border: #d8dcd8;
    --cfa-primary: #004851;
    --cfa-primary-2: #0b5a63;
    --cfa-text: #16353a;
    --cfa-muted: #748385;
    --cfa-accent: #fa9370;
    --cfa-shadow: 0 10px 28px rgba(0,0,0,0.05);
}

html, body, [class*="css"] {
    font-family: "Segoe UI", Arial, sans-serif;
}

.stApp {
    background: linear-gradient(180deg, #f5f3ef 0%, #f7f6f3 100%) !important;
    color: var(--cfa-text);
}

.block-container {
    max-width: 1320px;
    padding-top: 1.2rem;
    padding-bottom: 2rem;
}

section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0b525b 0%, #073f45 100%);
    border-right: 1px solid rgba(255,255,255,0.08);
}

section[data-testid="stSidebar"] * {
    color: white !important;
}

.sidebar-copy {
    background: rgba(255,255,255,0.08);
    border: 1px solid rgba(255,255,255,0.10);
    border-radius: 18px;
    padding: 18px 16px;
    line-height: 1.65;
    font-size: 0.98rem;
    color: rgba(255,255,255,0.92);
    margin-top: 10px;
}

.top-wrap {
    max-width: 1100px;
    margin: 0 auto 18px auto;
}

.title-center {
    text-align: center;
    color: var(--cfa-primary);
    font-size: 3.1rem;
    line-height: 1.05;
    font-weight: 800;
    margin-top: 4px;
    margin-bottom: 10px;
}

.subtitle-center {
    text-align: center;
    color: var(--cfa-muted);
    font-size: 1rem;
    line-height: 1.6;
    max-width: 860px;
    margin: 0 auto 18px auto;
}

.upload-wrap {
    background: rgba(255,255,255,0.96);
    border: 1px solid var(--cfa-border);
    border-radius: 22px;
    padding: 14px 16px 6px 16px;
    box-shadow: var(--cfa-shadow);
    margin-bottom: 20px;
}

.kpi {
    background: rgba(255,255,255,0.96);
    border: 1px solid var(--cfa-border);
    border-radius: 20px;
    padding: 16px 18px;
    min-height: 112px;
    box-shadow: var(--cfa-shadow);
}

.kpi-title {
    color: var(--cfa-muted);
    font-size: 0.78rem;
    margin-bottom: 10px;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    font-weight: 700;
}

.kpi-value {
    color: var(--cfa-text);
    font-size: 1.45rem;
    font-weight: 800;
}

.value-card {
    background: linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(255,255,255,0.94) 100%);
    border: 1px solid var(--cfa-border);
    border-radius: 20px;
    padding: 18px 20px;
    min-height: 120px;
    box-shadow: var(--cfa-shadow);
}

.value-label {
    color: var(--cfa-muted);
    font-size: 0.78rem;
    text-transform: uppercase;
    letter-spacing: 0.05em;
    font-weight: 700;
    margin-bottom: 10px;
}

.value-number {
    color: var(--cfa-text);
    font-size: 1.95rem;
    font-weight: 800;
    line-height: 1.1;
}

.section-title {
    color: var(--cfa-primary);
    font-size: 2rem;
    font-weight: 800;
    margin-top: 0.6rem;
    margin-bottom: 1rem;
}

div[data-testid="stFileUploader"] {
    background: transparent;
    border: 0;
    padding: 0;
    box-shadow: none;
}

div[data-testid="stDataFrame"] {
    border-radius: 16px;
    overflow: hidden;
    border: 1px solid var(--cfa-border);
    background: white;
}

div[data-testid="stExpander"] {
    border: 1px solid var(--cfa-border) !important;
    border-radius: 16px !important;
    background: rgba(255,255,255,0.96);
}

.stTabs [data-baseweb="tab-list"] {
    gap: 14px;
}

.stTabs [data-baseweb="tab"] {
    border-radius: 999px;
    background: rgba(255,255,255,0.92);
    border: 1px solid var(--cfa-border);
    padding: 8px 18px;
    color: var(--cfa-primary);
}

.stTabs [aria-selected="true"] {
    background: var(--cfa-primary) !important;
    color: white !important;
    border-color: var(--cfa-primary) !important;
}

@media (max-width: 900px) {
    .title-center {
        font-size: 2.35rem;
    }
}
</style>
""",
    unsafe_allow_html=True,
)

# =========================
# Sidebar
# =========================
with st.sidebar:
    if logo_path.exists():
        st.image(str(logo_path), use_container_width=True)

    st.markdown(
        """
<div class="sidebar-copy">
Versão otimizada para ficheiros Excel (.xlsx), com leitura do último ano disponível, indicadores de elegibilidade, semáforos, evolução histórica e apresentação executiva.
</div>
""",
        unsafe_allow_html=True,
    )

# =========================
# Header
# =========================
st.markdown('<div class="top-wrap">', unsafe_allow_html=True)

st.markdown(
    '<div class="title-center">Análise de Rácios Financeiros</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="subtitle-center">Ferramenta de leitura financeira com apresentação executiva, pensada para análise rápida, validação de elegibilidade e visualização da evolução histórica.</div>',
    unsafe_allow_html=True,
)

st.markdown("</div>", unsafe_allow_html=True)

# =========================
# Upload
# =========================
st.markdown('<div class="upload-wrap">', unsafe_allow_html=True)
uploaded = st.file_uploader(
    "Upload de ficheiro XLSX",
    type=["xlsx"],
    key="upload_financeiro",
)
st.markdown("</div>", unsafe_allow_html=True)

if uploaded:
    try:
        parsed = extract_document(uploaded, uploaded.name)

        company = parsed["company_info"]
        latest_year = parsed["latest_year"]
        latest_values = parsed["latest_values"]
        history_df = parsed["history_df"]

        st.markdown('<div class="section-title">Informações da Empresa</div>', unsafe_allow_html=True)
        c1, c2, c3, c4 = st.columns(4)

        with c1:
            st.markdown(
                f"""
<div class='kpi'>
    <div class='kpi-title'>Nome</div>
    <div class='kpi-value' style='font-size:1.02rem'>{company.get('nome') or '—'}</div>
</div>
""",
                unsafe_allow_html=True,
            )
        with c2:
            cae_label = "—"
            if company.get("cae"):
                cae_label = f"{company.get('cae')} - {company.get('cae_descricao') or ''}".strip(" -")
            st.markdown(
                f"""
<div class='kpi'>
    <div class='kpi-title'>CAE</div>
    <div class='kpi-value' style='font-size:0.98rem'>{cae_label}</div>
</div>
""",
                unsafe_allow_html=True,
            )
        with c3:
            st.markdown(
                f"""
<div class='kpi'>
    <div class='kpi-title'>NIF</div>
    <div class='kpi-value'>{company.get('nif') or '—'}</div>
</div>
""",
                unsafe_allow_html=True,
            )
        with c4:
            st.markdown(
                f"""
<div class='kpi'>
    <div class='kpi-title'>Nº Funcionários ({latest_year})</div>
    <div class='kpi-value'>{fmt_number(company.get('numero_funcionarios'))}</div>
</div>
""",
                unsafe_allow_html=True,
            )

        st.markdown(
            f'<div class="section-title">Último ano disponível: {latest_year}</div>',
            unsafe_allow_html=True,
        )

        k1, k2, k3, k4 = st.columns(4)
        with k1:
            st.markdown(
                f"""
<div class='value-card'>
    <div class='value-label'>Ativo Total</div>
    <div class='value-number'>{fmt_number(latest_values.get('ativo_total'))}</div>
</div>
""",
                unsafe_allow_html=True,
            )
        with k2:
            st.markdown(
                f"""
<div class='value-card'>
    <div class='value-label'>Capital Próprio</div>
    <div class='value-number'>{fmt_number(latest_values.get('capital_proprio'))}</div>
</div>
""",
                unsafe_allow_html=True,
            )
        with k3:
            st.markdown(
                f"""
<div class='value-card'>
    <div class='value-label'>Volume de Negócios</div>
    <div class='value-number'>{fmt_number(latest_values.get('vendas'))}</div>
</div>
""",
                unsafe_allow_html=True,
            )
        with k4:
            st.markdown(
                f"""
<div class='value-card'>
    <div class='value-label'>Resultado Líquido</div>
    <div class='value-number'>{fmt_number(latest_values.get('resultado_liquido'))}</div>
</div>
""",
                unsafe_allow_html=True,
            )

        with st.expander("Ver dados extraídos do último ano", expanded=False):
            latest_table = pd.DataFrame(
                {
                    "Campo": [
                        "Ativo Total",
                        "Ativo Corrente",
                        "Inventários",
                        "Passivo Corrente",
                        "Passivo Total",
                        "Capital Próprio",
                        "Volume de Negócios",
                        "EBITDA",
                        "EBIT",
                        "Resultado Líquido",
                        "Nº Funcionários",
                    ],
                    "Valor": [
                        fmt_number(latest_values.get("ativo_total")),
                        fmt_number(latest_values.get("ativo_corrente")),
                        fmt_number(latest_values.get("inventarios")),
                        fmt_number(latest_values.get("passivo_corrente")),
                        fmt_number(latest_values.get("passivo_total")),
                        fmt_number(latest_values.get("capital_proprio")),
                        fmt_number(latest_values.get("vendas")),
                        fmt_number(latest_values.get("ebitda")),
                        fmt_number(latest_values.get("ebit")),
                        fmt_number(latest_values.get("resultado_liquido")),
                        fmt_number(latest_values.get("numero_funcionarios")),
                    ],
                }
            )
            st.dataframe(latest_table, use_container_width=True, hide_index=True)

        with st.expander("Corrigir valores manualmente", expanded=False):
            left, right = st.columns(2)
            edited = {}

            with left:
                edited["ativo_total"] = editable_number_input("Ativo Total", latest_values.get("ativo_total"), "ativo_total")
                edited["inventarios"] = editable_number_input("Inventários", latest_values.get("inventarios"), "inventarios")
                edited["passivo_total"] = editable_number_input("Passivo Total", latest_values.get("passivo_total"), "passivo_total")
                edited["vendas"] = editable_number_input("Volume de Negócios", latest_values.get("vendas"), "vendas")
                edited["ebit"] = editable_number_input("EBIT", latest_values.get("ebit"), "ebit")

            with right:
                edited["ativo_corrente"] = editable_number_input("Ativo Corrente", latest_values.get("ativo_corrente"), "ativo_corrente")
                edited["passivo_corrente"] = editable_number_input("Passivo Corrente", latest_values.get("passivo_corrente"), "passivo_corrente")
                edited["capital_proprio"] = editable_number_input("Capital Próprio", latest_values.get("capital_proprio"), "capital_proprio")
                edited["ebitda"] = editable_number_input("EBITDA", latest_values.get("ebitda"), "ebitda")
                edited["resultado_liquido"] = editable_number_input("Resultado Líquido", latest_values.get("resultado_liquido"), "resultado_liquido")

            edited["liquidez_corrente"] = latest_values.get("liquidez_corrente")
            edited["liquidez_reduzida"] = latest_values.get("liquidez_reduzida")
            edited["autonomia_financeira"] = latest_values.get("autonomia_financeira")
            edited["endividamento"] = latest_values.get("endividamento")

        if "edited" not in locals():
            edited = latest_values.copy()

        ratios = calculate_ratios(edited)

        st.markdown('<div class="section-title">Rácios e Elegibilidade</div>', unsafe_allow_html=True)
        ratio_table = build_ratio_table(ratios)
        st.dataframe(ratio_table, use_container_width=True, hide_index=True)

        with st.expander("ℹ️ Critérios dos semáforos", expanded=False):
            st.markdown(
                """
**Critérios usados na app**

- **Liquidez Corrente**
  - 🟢 >= 1,00
  - 🟡 entre 0,90 e 0,99
  - 🔴 < 0,90

- **Liquidez Reduzida**
  - 🟢 >= 0,80
  - 🟡 entre 0,60 e 0,79
  - 🔴 < 0,60

- **Autonomia Financeira**
  - 🟢 >= 30%
  - 🟡 entre 20% e 29,99%
  - 🔴 < 20%

- **Endividamento**
  - 🟢 <= 70%
  - 🟡 entre 70,01% e 80%
  - 🔴 > 80%

- **ROA**
  - 🟢 >= 2%
  - 🟡 entre 0% e 1,99%
  - 🔴 < 0%

- **ROE**
  - 🟢 >= 5%
  - 🟡 entre 0% e 4,99%
  - 🔴 < 0%

- **Margem Líquida**
  - 🟢 >= 3%
  - 🟡 entre 0% e 2,99%
  - 🔴 < 0%

- **Margem EBITDA**
  - 🟢 >= 10%
  - 🟡 entre 5% e 9,99%
  - 🔴 < 5%
"""
            )

        st.markdown('<div class="section-title">Evolução da Empresa</div>', unsafe_allow_html=True)

        if not history_df.empty and "Ano" in history_df.columns:
            chart_df = history_df.copy().sort_values("Ano")
            chart_df = chart_df.tail(5).copy()
            chart_df["Ano"] = chart_df["Ano"].astype(int).astype(str)
            chart_df = chart_df.set_index("Ano")

            tab1, tab2, tab3 = st.tabs(["Dimensão", "Rentabilidade", "Pessoas"])

            with tab1:
                cols = [c for c in ["ativo_total", "capital_proprio", "vendas"] if c in chart_df.columns]
                if cols:
                    st.markdown("**Ativo Total, Capital Próprio e Volume de Negócios**")
                    st.line_chart(chart_df[cols])

            with tab2:
                cols = [c for c in ["ebitda", "resultado_liquido"] if c in chart_df.columns]
                if cols:
                    st.markdown("**EBITDA e Resultado Líquido**")
                    st.line_chart(chart_df[cols])

            with tab3:
                if "numero_funcionarios" in chart_df.columns:
                    st.markdown("**Número de Funcionários**")
                    st.bar_chart(chart_df[["numero_funcionarios"]])

            with st.expander("Ver tabela histórica", expanded=False):
                show_hist = history_df.copy().sort_values("Ano")
                show_hist = show_hist.tail(5).copy()
                for c in show_hist.columns:
                    if c != "Ano":
                        show_hist[c] = show_hist[c].apply(fmt_number)
                show_hist["Ano"] = show_hist["Ano"].astype(int).astype(str)
                st.dataframe(show_hist, use_container_width=True, hide_index=True)
        else:
            st.info("Não foi possível construir a evolução histórica a partir deste ficheiro.")

    except Exception as e:
        st.error(f"Erro ao processar o ficheiro: {e}")

else:
    st.info("Faz upload de um ficheiro XLSX para começar.")