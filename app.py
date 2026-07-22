# app.py
import re
import io
import base64
import unicodedata
from datetime import datetime

import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak
from reportlab.lib.enums import TA_CENTER, TA_LEFT, TA_RIGHT, TA_JUSTIFY
from reportlab.graphics.shapes import Drawing, String
from reportlab.graphics.charts.linecharts import HorizontalLineChart
from reportlab.graphics.widgets.markers import makeMarker
from reportlab.lib.colors import HexColor

st.set_page_config(page_title="Dashboard Financeiro Premium", page_icon="📊", layout="wide", initial_sidebar_state="expanded")

# ---------------------------------------------------------------------------
# Paleta — Azul Maldivas
# ---------------------------------------------------------------------------
C_NAVY, C_NAVY2 = "#0a1628", "#1a3a5c"
C_TEAL, C_TEAL_LIGHT = "#2e86ab", "#7dd3fc"
C_SUCCESS, C_DANGER, C_GOLD = "#00d4aa", "#ff6b6b", "#feca57"
C_MEDAL_GOLD, C_MEDAL_SILVER, C_MEDAL_BRONZE = "#ffd700", "#c0c0c0", "#cd7f32"
GRADIENTE = [C_NAVY, C_NAVY2, C_TEAL, "#4ea8de", C_SUCCESS, C_TEAL_LIGHT]

MESES_ORDEM = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MESES_FULL = {"Jan": "Janeiro", "Fev": "Fevereiro", "Mar": "Março", "Abr": "Abril", "Mai": "Maio", "Jun": "Junho",
              "Jul": "Julho", "Ago": "Agosto", "Set": "Setembro", "Out": "Outubro", "Nov": "Novembro", "Dez": "Dezembro"}
MES_PREFIX = {"JAN": "Jan", "FEV": "Fev", "MAR": "Mar", "ABR": "Abr", "MAI": "Mai", "JUN": "Jun",
              "JUL": "Jul", "AGO": "Ago", "SET": "Set", "OUT": "Out", "NOV": "Nov", "DEZ": "Dez"}
QUARTER_MONTHS = {"1º Tri": ["Jan", "Fev", "Mar"], "2º Tri": ["Abr", "Mai", "Jun"],
                   "3º Tri": ["Jul", "Ago", "Set"], "4º Tri": ["Out", "Nov", "Dez"]}
QUARTER_OF = {m: q for q, ms in QUARTER_MONTHS.items() for m in ms}

# ---------------------------------------------------------------------------
# CSS — visual idêntico ao modelo Azul Maldivas
# ---------------------------------------------------------------------------
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700;800;900&display=swap');
* {{ font-family: 'Inter', -apple-system, BlinkMacSystemFont, sans-serif; }}
.main .block-container {{ padding: 2rem 3rem; max-width: 1400px; }}
.main-header {{
    background: linear-gradient(135deg, {C_NAVY} 0%, {C_NAVY2} 25%, {C_TEAL} 50%, {C_NAVY2} 75%, {C_NAVY} 100%);
    background-size: 400% 400%; animation: gradientShift 8s ease infinite;
    padding: 50px 40px; border-radius: 28px; text-align: center; margin-bottom: 40px;
    box-shadow: 0 25px 70px rgba(10,22,40,.45); border: 2px solid rgba(255,255,255,.2); }}
@keyframes gradientShift {{ 0%{{background-position:0% 50%}} 50%{{background-position:100% 50%}} 100%{{background-position:0% 50%}} }}
.main-header h1 {{ color:white; font-size:3rem; font-weight:900; text-shadow:3px 3px 10px rgba(0,0,0,.4); margin-bottom:10px; }}
.main-header h2 {{ color:white; font-size:1.3rem; font-weight:500; opacity:.95; }}
.main-header .badge {{ display:inline-block; background:rgba(255,255,255,.2); padding:8px 22px; border-radius:50px;
    margin-top:16px; font-size:.95rem; font-weight:600; color:white; border:1px solid rgba(255,255,255,.3); }}
.section-header {{ padding:22px 34px; border-radius:18px; margin:32px 0 20px 0; box-shadow:0 12px 35px rgba(0,0,0,.15); }}
.section-header h2 {{ color:white; font-size:1.5rem; font-weight:800; margin:0; }}
.kpi-card {{ border-radius:20px; padding:26px 18px; color:white; text-align:center; min-height:150px;
    box-shadow:0 15px 40px rgba(0,0,0,.25); display:flex; flex-direction:column; justify-content:center; }}
.kpi-label {{ font-size:.85rem; font-weight:700; opacity:.9; text-transform:uppercase; letter-spacing:1px; }}
.kpi-value {{ font-size:1.7rem; font-weight:900; margin:10px 0 4px 0; }}
.kpi-delta-up {{ font-size:.8rem; font-weight:700; color:#baffef; }}
.kpi-delta-down {{ font-size:.8rem; font-weight:700; color:#ffd6d6; }}
.legenda-box {{ background:linear-gradient(135deg,#fff 0%,#f0f8ff 100%); border:2px solid {C_TEAL}; border-left:8px solid {C_TEAL};
    border-radius:18px; padding:26px 30px; margin:25px 0; box-shadow:0 12px 30px rgba(46,134,171,.15); }}
.legenda-item {{ margin:8px 0; padding:12px 16px; border-radius:12px; border-left:4px solid; font-size:.95rem; }}
.ranking-card {{ background:white; border-radius:18px; padding:20px; margin:12px 0; box-shadow:0 8px 24px rgba(0,0,0,.1); border:3px solid; }}
.filtro-box {{ background:linear-gradient(135deg,{C_NAVY} 0%,{C_NAVY2} 100%); padding:22px 30px; border-radius:18px;
    margin-bottom:28px; box-shadow:0 12px 35px rgba(10,22,40,.3); }}
.filtro-box h3 {{ color:white; margin:0 0 12px 0; font-size:1.2rem; font-weight:700; }}
.stDataFrame {{ border-radius:16px; overflow:hidden; }}
</style>
""", unsafe_allow_html=True)

PLOTLY_CONFIG = {"displaylogo": False}

# ---------------------------------------------------------------------------
# Helpers de texto/número
# ---------------------------------------------------------------------------
def strip_accents(s):
    return "".join(c for c in unicodedata.normalize("NFKD", str(s)) if not unicodedata.combining(c))

def norm(s):
    return strip_accents(str(s)).upper().strip() if s is not None else ""

def parse_number(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return 0.0 if pd.isna(v) else float(v)
    s = str(v).strip()
    if s in ("", "-"):
        return 0.0
    neg = s.startswith("(") and s.endswith(")")
    if neg:
        s = s[1:-1]
    s = s.replace("R$", "").replace(" ", "")
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        val = float(s)
    except ValueError:
        return 0.0
    return -val if neg else val

def find_col(df, *keywords, exclude=None):
    for col in df.columns:
        h = norm(col)
        if all(k in h for k in keywords) and (not exclude or exclude not in h):
            return col
    return None

def find_sheet(sheetnames, *contains):
    for name in sheetnames:
        h = norm(name)
        if any(c in h for c in contains):
            return name
    return None

def month_from_value(v):
    if v is None or (isinstance(v, float) and pd.isna(v)):
        return None
    if isinstance(v, (pd.Timestamp, datetime)):
        return MESES_ORDEM[v.month - 1]
    s = str(v).strip()
    m = re.match(r"^(\d{1,2})[/-](\d{4})$", s)
    if m:
        return MESES_ORDEM[int(m.group(1)) - 1]
    try:
        dt = pd.to_datetime(v, dayfirst=True, errors="raise")
        return MESES_ORDEM[dt.month - 1]
    except Exception:
        return None

def fmt_r(v, mil=False):
    v = 0 if v is None or (isinstance(v, float) and np.isnan(v)) else v
    if mil:
        v = v / 1000
    sign = "-" if v < 0 else ""
    txt = f"{abs(v):,.0f}".replace(",", ".")
    return f"{sign}R$ {txt}{' mil' if mil else ''}"

def fmt_pct(v):
    return "-" if v is None or (isinstance(v, float) and np.isnan(v)) else f"{v:.1%}"

# ---------------------------------------------------------------------------
# Parsing "DRE 2026" (robusto por rótulo, não por índice fixo)
# ---------------------------------------------------------------------------
def find_row(ws, label, start=1, end=None):
    end = end or ws.max_row
    target = norm(label)
    for r in range(start, end + 1):
        if target in norm(ws.cell(row=r, column=2).value):
            return r
    return None

def find_header_row(ws):
    for r in range(1, ws.max_row + 1):
        vals = [norm(ws.cell(row=r, column=c).value)[:3] for c in range(3, ws.max_column + 1)]
        if sum(1 for v in vals if v in MES_PREFIX) >= 6:
            return r
    return None

def row_series(ws, row, cols):
    return [ws.cell(row=row, column=c).value or 0 for c in cols]

def parse_dre(ws):
    header_row = find_header_row(ws)
    if header_row is None:
        return None
    month_cols = [c for c in range(3, ws.max_column + 1) if norm(ws.cell(row=header_row, column=c).value)[:3] in MES_PREFIX]
    if not month_cols:
        return None

    direta_sec = find_row(ws, "PRODUÇÃO DIRETA")
    portal_sec = find_row(ws, "PORTAL MAAS")
    resultado_row = find_row(ws, "RESULTADO OPERACIONAL", start=portal_sec)
    while resultado_row and "DISTRIBUI" in norm(ws.cell(row=resultado_row, column=2).value):
        resultado_row = find_row(ws, "RESULTADO OPERACIONAL", start=resultado_row + 1)
    d_end, p_end = portal_sec, resultado_row

    rows = {
        "d_receita": find_row(ws, "RECEITA BRUTA", direta_sec, d_end),
        "d_impostos": find_row(ws, "IMPOSTOS DIRETOS", direta_sec, d_end),
        "d_custo": find_row(ws, "CUSTO OPERACIONAL", direta_sec, d_end),
        "d_cocorretagem": find_row(ws, "CO-CORRETAGEM", direta_sec, d_end),
        "d_rebate": find_row(ws, "REBATE AAI", direta_sec, d_end),
        "d_margem": find_row(ws, "MARGEM DE CONTRIBUIÇÃO", direta_sec, d_end),
        "d_despesas": find_row(ws, "DESPESAS", direta_sec, d_end),
        "d_folha": find_row(ws, "FOLHA+TERCEIROS", direta_sec, d_end),
        "p_receita": find_row(ws, "RECEITA BRUTA", portal_sec, p_end),
        "p_impostos": find_row(ws, "IMPOSTOS DIRETOS", portal_sec, p_end),
        "p_custo": find_row(ws, "CUSTO OPERACIONAL", portal_sec, p_end),
        "p_margem": find_row(ws, "MARGEM DE CONTRIBUIÇÃO", portal_sec, p_end),
        "p_despesas": find_row(ws, "DESPESAS", portal_sec, p_end),
        "p_folha": find_row(ws, "FOLHA+TERCEIROS", portal_sec, p_end),
    }
    socio_partner_row = find_row(ws, "Sócio Partner", start=resultado_row)
    socio_maldivas_row = find_row(ws, "Sócio Maldivas", start=resultado_row)
    valor_pagar_maldivas_row = find_row(ws, "Valor a pagar", start=resultado_row)

    def s(key):
        r = rows.get(key)
        return row_series(ws, r, month_cols) if r else [0] * len(month_cols)

    def rs(row):
        return row_series(ws, row, month_cols) if row else [0] * len(month_cols)

    meses = [MES_PREFIX[norm(ws.cell(row=header_row, column=c).value)[:3]] for c in month_cols]
    receita_direta, receita_portal = s("d_receita"), s("p_receita")

    df = pd.DataFrame({
        "ReceitaDireta": receita_direta, "ReceitaPortal": receita_portal,
        "Impostos": [a + b for a, b in zip(s("d_impostos"), s("p_impostos"))],
        "Custo": [a + b for a, b in zip(s("d_custo"), s("p_custo"))],
        "CoCorretagem": s("d_cocorretagem"), "RebateAAI": s("d_rebate"),
        "MargemContribuicao": [a + b for a, b in zip(s("d_margem"), s("p_margem"))],
        "Despesas": [a + b for a, b in zip(s("d_despesas"), s("p_despesas"))],
        "Folha": [a + b for a, b in zip(s("d_folha"), s("p_folha"))],
        "ResultadoOperacional": rs(resultado_row),
        "SocioPartner": rs(socio_partner_row), "SocioMaldivas": rs(socio_maldivas_row),
        "ValorPagarMaldivas": rs(valor_pagar_maldivas_row),
    }, index=meses)
    df = df.groupby(df.index).sum()
    df["TemDados"] = (df["ReceitaDireta"] + df["ReceitaPortal"]) > 0
    return df.reindex([m for m in MESES_ORDEM if m in df.index])

def parse_shares(wb):
    if "INPUTS" not in wb.sheetnames:
        return 0.7, 0.3
    ws = wb["INPUTS"]
    rp, rm = find_row(ws, "Sócio Partner"), find_row(ws, "Sócio Maldivas")
    return (float(ws.cell(row=rp, column=2).value or 0.7) if rp else 0.7,
            float(ws.cell(row=rm, column=2).value or 0.3) if rm else 0.3)

# ---------------------------------------------------------------------------
# Parsing "ASSERTIF DIRETO" (transações) e "DESPESAS"
# ---------------------------------------------------------------------------
def parse_transacoes(file_bytes, sheet_name):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
    except Exception:
        return None
    if df is None or df.empty:
        return None
    col_data = find_col(df, "DATA") or find_col(df, "MES")
    col_cliente = find_col(df, "CLIENTE")
    col_seguradora = find_col(df, "SEGURADORA")
    col_produto = find_col(df, "PRODUTO")
    col_originador = find_col(df, "ORIGINADOR")
    col_comissao = find_col(df, "COMISS") or find_col(df, "VALOR")
    if col_comissao is None:
        return None
    out = pd.DataFrame()
    out["Mes"] = df[col_data].apply(month_from_value) if col_data is not None else None
    out["Cliente"] = df[col_cliente] if col_cliente is not None else "N/D"
    out["Seguradora"] = df[col_seguradora] if col_seguradora is not None else "N/D"
    out["Produto"] = df[col_produto] if col_produto is not None else "N/D"
    out["Originador"] = df[col_originador] if col_originador is not None else "N/D"
    out["Valor"] = df[col_comissao].apply(parse_number)
    out = out[out["Valor"] != 0]
    return out if not out.empty else None

def parse_despesas(file_bytes, sheet_name):
    try:
        df = pd.read_excel(io.BytesIO(file_bytes), sheet_name=sheet_name)
    except Exception:
        return None
    if df is None or df.empty:
        return None
    col_data = find_col(df, "DATA") or find_col(df, "MES")
    col_categoria = find_col(df, "CATEGORIA")
    col_valor = find_col(df, "VALOR")
    if col_valor is None or col_categoria is None:
        return None
    out = pd.DataFrame()
    out["Mes"] = df[col_data].apply(month_from_value) if col_data is not None else None
    out["Categoria"] = df[col_categoria]
    out["Valor"] = df[col_valor].apply(parse_number).abs()
    out = out[out["Valor"] != 0]
    return out if not out.empty else None

@st.cache_data(show_spinner="Lendo planilha...")
def parse_workbook(file_bytes, file_name):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if "DRE 2026" not in wb.sheetnames:
        return None
    df = parse_dre(wb["DRE 2026"])
    if df is None:
        return None
    shares = parse_shares(wb)
    sheet_receitas = find_sheet(wb.sheetnames, "DIRETO") or find_sheet(wb.sheetnames, "RECEITA")
    sheet_despesas = find_sheet(wb.sheetnames, "DESPES")
    tx = parse_transacoes(file_bytes, sheet_receitas) if sheet_receitas else None
    desp = parse_despesas(file_bytes, sheet_despesas) if sheet_despesas else None
    return {"file_name": file_name, "df": df, "shares": shares, "tx": tx, "desp": desp}

def months_with_data(parsed_list, selected_files):
    s = set()
    for p in parsed_list:
        if p["file_name"] in selected_files:
            s.update(p["df"].index[p["df"]["TemDados"]])
    return sorted(s, key=lambda m: MESES_ORDEM.index(m))

def aggregate(parsed_list, selected_files, selected_months):
    sel = [p for p in parsed_list if p["file_name"] in selected_files]
    if not sel:
        return None
    parts = [p["df"].loc[p["df"].index.intersection(selected_months)] for p in sel]
    full = pd.concat(parts).groupby(level=0).sum()
    full["TemDados"] = full["TemDados"] > 0
    full = full.reindex([m for m in MESES_ORDEM if m in full.index])
    combined = full[full["TemDados"]].copy()

    tx_parts = [p["tx"] for p in sel if p["tx"] is not None]
    tx = pd.concat(tx_parts, ignore_index=True) if tx_parts else None
    if tx is not None and tx["Mes"].notna().any():
        tx = tx[tx["Mes"].isna() | tx["Mes"].isin(selected_months)]

    desp_parts = [p["desp"] for p in sel if p["desp"] is not None]
    desp = pd.concat(desp_parts, ignore_index=True) if desp_parts else None
    if desp is not None and desp["Mes"].notna().any():
        desp = desp[desp["Mes"].isna() | desp["Mes"].isin(selected_months)]

    partner_share = float(np.mean([p["shares"][0] for p in sel]))
    maldivas_share = float(np.mean([p["shares"][1] for p in sel]))
    return combined, tx, desp, (partner_share, maldivas_share)

def month_delta(series):
    if len(series) < 2:
        return None
    prev, curr = series.iloc[-2], series.iloc[-1]
    return None if prev == 0 else (curr - prev) / abs(prev)

def delta_html(delta):
    if delta is None:
        return '<span style="font-size:.8rem;color:#cfd8e3;">sem comparação</span>'
    cls = "kpi-delta-up" if delta >= 0 else "kpi-delta-down"
    arrow = "▲" if delta >= 0 else "▼"
    return f'<span class="{cls}">{arrow} {delta:+.1%} vs mês anterior</span>'

# ---------------------------------------------------------------------------
# Upload
# ---------------------------------------------------------------------------
with st.sidebar:
    st.markdown(f"""<div style="text-align:center;padding:15px 0;">
    <span style="font-size:2.6rem;">📊</span><h2 style="color:{C_NAVY};margin:10px 0 2px 0;">DASHBOARD</h2>
    <p style="color:#6c757d;font-size:.85rem;">Financeiro Premium</p></div>""", unsafe_allow_html=True)
    st.markdown("---")
    st.markdown("### 📁 Upload de Dados")
    uploads = st.file_uploader("Envie uma ou mais planilhas (.xlsx)", type=["xlsx"], accept_multiple_files=True)
    logo_file = st.file_uploader("Logo da empresa (opcional)", type=["png", "jpg", "jpeg"])
    if logo_file is not None:
        st.session_state["logo_bytes"] = logo_file.getvalue()
    st.markdown("---")
    show_tables = st.checkbox("📋 Mostrar tabelas detalhadas", value=True)
    show_charts = st.checkbox("📈 Mostrar gráficos", value=True)
    mil_mode = st.checkbox("Exibir valores em R$ mil")

st.markdown(f"""
<div class="main-header"><h1>📊 DASHBOARD FINANCEIRO</h1><h2>Painel Premium Interativo</h2>
<div class="badge">YTD 2026 · Multi-planilha</div></div>""", unsafe_allow_html=True)

if not uploads:
    st.info("Envie uma ou mais planilhas contendo as abas 'DRE 2026', 'ASSERTIF DIRETO' e/ou 'DESPESAS'.")
    st.stop()

parsed_list = []
for f in uploads:
    data = parse_workbook(f.getvalue(), f.name)
    if data is None:
        st.sidebar.warning(f"'{f.name}' não contém uma aba 'DRE 2026' válida — ignorado.")
        continue
    parsed_list.append(data)

if not parsed_list:
    st.error("Nenhum arquivo válido foi encontrado.")
    st.stop()

all_files = [p["file_name"] for p in parsed_list]

# ---------------------------------------------------------------------------
# Filtro — Planilhas
# ---------------------------------------------------------------------------
if "sel_files" not in st.session_state:
    st.session_state.sel_files = all_files
st.markdown('<div class="filtro-box"><h3>📁 Planilhas incluídas</h3></div>', unsafe_allow_html=True)
fb1, fb2 = st.columns([5, 1])
selected_files = fb1.multiselect("Planilhas", all_files, key="sel_files", label_visibility="collapsed")
if fb2.button("Todas", key="btn_all_files"):
    st.session_state.sel_files = all_files
    st.rerun()
if not selected_files:
    st.warning("Selecione ao menos uma planilha.")
    st.stop()

available_months = months_with_data(parsed_list, selected_files)
if not available_months:
    st.warning("Nenhum dos arquivos selecionados possui meses com dados lançados.")
    st.stop()
quarters_available = [q for q in QUARTER_MONTHS if any(m in available_months for m in QUARTER_MONTHS[q])]

# ---------------------------------------------------------------------------
# Filtro — Período (Trimestre + Mês, interativo e sincronizado)
# ---------------------------------------------------------------------------
if "sel_months" not in st.session_state or not set(st.session_state.sel_months) <= set(available_months):
    st.session_state.sel_months = available_months
if "sel_quarters" not in st.session_state:
    st.session_state.sel_quarters = quarters_available

def sync_quarter_to_months():
    qs = st.session_state.sel_quarters
    months = [m for q in qs for m in QUARTER_MONTHS[q] if m in available_months]
    st.session_state.sel_months = sorted(set(months), key=lambda m: MESES_ORDEM.index(m)) or available_months

st.markdown('<div class="filtro-box"><h3>🗓️ Período de Análise</h3></div>', unsafe_allow_html=True)
qc1, qc2, qc3, qc4 = st.columns(4)
if qc1.button("📅 Todos os meses"):
    st.session_state.sel_months = available_months
    st.session_state.sel_quarters = quarters_available
if qc2.button("📆 Último trimestre") and quarters_available:
    last_q = quarters_available[-1]
    st.session_state.sel_quarters = [last_q]
    st.session_state.sel_months = [m for m in QUARTER_MONTHS[last_q] if m in available_months]
if qc3.button("🗓️ Último mês"):
    st.session_state.sel_months = [available_months[-1]]
if qc4.button("🧹 Limpar"):
    st.session_state.sel_months, st.session_state.sel_quarters = [], []

colt, colm = st.columns(2)
with colt:
    st.multiselect("Trimestre", quarters_available, key="sel_quarters", on_change=sync_quarter_to_months)
with colm:
    selected_months = st.multiselect("Meses (apenas com dados lançados)", available_months, key="sel_months")

st.caption(f"Exibindo apenas meses com dados reais: {', '.join(available_months)}. Meses futuros sem lançamento não entram no dashboard nem nos cálculos de variação.")

if not selected_months:
    st.warning("Selecione ao menos um mês.")
    st.stop()

combined, combined_tx, combined_desp, (partner_share, maldivas_share) = aggregate(parsed_list, selected_files, selected_months)
if combined.empty:
    st.warning("Sem dados para a seleção atual.")
    st.stop()

periodo_label = f"{combined.index[0]} a {combined.index[-1]} 2026" if len(combined.index) > 1 else f"{combined.index[0]} 2026"

# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------
receita_mensal = combined["ReceitaDireta"] + combined["ReceitaPortal"]
receita_total = receita_mensal.sum()
custos_totais = (combined["Impostos"] + combined["Custo"] + combined["RebateAAI"] - combined["CoCorretagem"]).sum()
margem_contrib = combined["MargemContribuicao"].sum()
despesas_total = combined["Despesas"].sum()
resultado_operacional = combined["ResultadoOperacional"].sum()
margem_lucro = (resultado_operacional / receita_total) if receita_total else 0
status = "LUCRO" if resultado_operacional >= 0 else "PREJUÍZO"

delta_receita = month_delta(receita_mensal)
delta_custos = month_delta(combined["Impostos"] + combined["Custo"] + combined["RebateAAI"] - combined["CoCorretagem"])
delta_margem = month_delta(combined["MargemContribuicao"])
delta_despesas = month_delta(combined["Despesas"])
delta_resultado = month_delta(combined["ResultadoOperacional"])

st.markdown(f'<div class="section-header" style="background:linear-gradient(135deg,{C_NAVY},{C_NAVY2});"><h2>💰 Indicadores Principais (KPIs)</h2></div>', unsafe_allow_html=True)
kpi_defs = [
    ("FATURAMENTO", fmt_r(receita_total, mil_mode), delta_receita, C_NAVY),
    ("CUSTOS TOTAIS", fmt_r(custos_totais, mil_mode), delta_custos, C_DANGER),
    ("MARGEM CONTRIB.", fmt_r(margem_contrib, mil_mode), delta_margem, C_TEAL),
    ("DESPESAS TOTAIS", fmt_r(despesas_total, mil_mode), delta_despesas, C_GOLD),
    ("RESULTADO OPER.", fmt_r(resultado_operacional, mil_mode), delta_resultado, C_SUCCESS if resultado_operacional >= 0 else C_DANGER),
]
kcols = st.columns(5)
for col, (label, value, delta, cor) in zip(kcols, kpi_defs):
    col.markdown(f"""<div class="kpi-card" style="background:linear-gradient(145deg,{cor},{cor}cc);">
    <div class="kpi-label">{label}</div><div class="kpi-value">{value}</div>{delta_html(delta)}</div>""", unsafe_allow_html=True)

st.markdown(f"""
<div class="legenda-box"><h3 style="color:{C_NAVY};">📌 Legenda dos Indicadores</h3>
<div class="legenda-item" style="background:rgba(10,22,40,.06);border-color:{C_NAVY};"><b>Faturamento:</b> Receita Bruta (Produção Direta + Portal MAAS)</div>
<div class="legenda-item" style="background:rgba(255,107,107,.08);border-color:{C_DANGER};"><b>Custos Totais:</b> Impostos + Custo Operacional (D.A) + Rebate AAI − Co-Corretagem</div>
<div class="legenda-item" style="background:rgba(46,134,171,.08);border-color:{C_TEAL};"><b>Margem de Contribuição:</b> Faturamento − Custos Totais</div>
<div class="legenda-item" style="background:rgba(254,202,87,.1);border-color:{C_GOLD};"><b>Despesas Totais:</b> Despesas Operacionais + Folha/Terceiros</div>
<div class="legenda-item" style="background:rgba(0,212,170,.08);border-color:{C_SUCCESS};"><b>Resultado Operacional:</b> Margem de Contribuição − Despesas · base da distribuição entre sócios</div>
</div>""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Evolução Mensal
# ---------------------------------------------------------------------------
if show_charts:
    st.markdown(f'<div class="section-header" style="background:linear-gradient(135deg,{C_NAVY2},{C_TEAL});"><h2>📈 Evolução Mensal — Receita vs Resultado</h2></div>', unsafe_allow_html=True)
    meses_full = [MESES_FULL[m] for m in combined.index]
    crescimento = receita_mensal.pct_change() * 100

    fig = make_subplots(rows=1, cols=3, subplot_titles=("<b>Receita Bruta</b>", "<b>Crescimento Mensal (%)</b>", "<b>Resultado Operacional</b>"),
                         horizontal_spacing=0.08, column_widths=[0.35, 0.3, 0.35])
    fig.add_trace(go.Bar(x=meses_full, y=receita_mensal, marker_color=C_TEAL_LIGHT,
                          text=[f"R$ {v/1000:.1f}K" for v in receita_mensal], textposition="outside",
                          hovertemplate="%{x}: R$ %{y:,.0f}<extra></extra>"), row=1, col=1)
    fig.add_trace(go.Scatter(x=meses_full, y=crescimento, mode="lines+markers+text",
                              line=dict(color=C_TEAL_LIGHT, width=4, shape="spline"),
                              marker=dict(size=16, color=C_TEAL_LIGHT, line=dict(width=3, color="white")),
                              text=[("-" if pd.isna(v) else f"{v:+.1f}%") for v in crescimento], textposition="top center",
                              hovertemplate="%{x}: %{y:+.1f}%<extra></extra>"), row=1, col=2)
    fig.add_hline(y=0, line_dash="dash", line_color=C_DANGER, row=1, col=2)
    fig.add_trace(go.Bar(x=meses_full, y=combined["ResultadoOperacional"], marker_color=C_TEAL_LIGHT,
                          text=[f"R$ {v/1000:.1f}K" for v in combined["ResultadoOperacional"]], textposition="outside",
                          hovertemplate="%{x}: R$ %{y:,.0f}<extra></extra>"), row=1, col=3)
    fig.add_hline(y=0, line_color=C_DANGER, row=1, col=3)
    fig.update_layout(height=480, showlegend=False, paper_bgcolor="white", plot_bgcolor="rgba(0,0,0,0)",
                       font=dict(family="Inter"), hovermode="x unified")
    st.plotly_chart(fig, use_container_width=True, config=PLOTLY_CONFIG)

    if show_tables:
        tab = pd.DataFrame({"Mês": meses_full, "Receita Bruta": receita_mensal.values,
                             "Crescimento": ["-"] + [f"{v:+.1f}%" for v in crescimento.values[1:]],
                             "Resultado Operacional": combined["ResultadoOperacional"].values,
                             "Margem": [f"{(r/rb):.1%}" if rb else "-" for r, rb in zip(combined["ResultadoOperacional"], receita_mensal)]})
        st.dataframe(tab.style.format({"Receita Bruta": lambda v: fmt_r(v, mil_mode), "Resultado Operacional": lambda v: fmt_r(v, mil_mode)}),
                     hide_index=True, use_container_width=True)

# ---------------------------------------------------------------------------
# Ranking Seguradoras
# ---------------------------------------------------------------------------
if show_charts and combined_tx is not None:
    st.markdown(f'<div class="section-header" style="background:linear-gradient(135deg,{C_NAVY2},{C_NAVY});"><h2>🏆 Ranking — Maiores Comissões por Seguradora</h2></div>', unsafe_allow_html=True)
    df_seg = combined_tx.groupby("Seguradora")["Valor"].sum().reset_index().sort_values("Valor", ascending=False)
    df_seg = df_seg[df_seg["Valor"] > 0]
    if not df_seg.empty:
        df_seg["Pct"] = df_seg["Valor"] / df_seg["Valor"].sum() * 100
        top = df_seg.head(15)
        fig_seg = go.Figure(go.Bar(y=top["Seguradora"], x=top["Valor"], orientation="h",
                                    marker=dict(color=top["Valor"], colorscale=[[0, "#4ea8de"], [0.5, C_TEAL], [1, C_NAVY]], showscale=True),
                                    text=[f"{fmt_r(v, mil_mode)} ({p:.1f}%)" for v, p in zip(top["Valor"], top["Pct"])], textposition="outside",
                                    hovertemplate="%{y}: R$ %{x:,.0f}<extra></extra>"))
        fig_seg.update_layout(height=600, paper_bgcolor="white", plot_bgcolor="rgba(0,0,0,0)",
                               yaxis=dict(categoryorder="total ascending"), margin=dict(l=200))
        st.plotly_chart(fig_seg, use_container_width=True, config=PLOTLY_CONFIG)
    else:
        st.info("Sem dados de seguradora para o período selecionado.")

# ---------------------------------------------------------------------------
# Distribuição de Resultados — Sócios
# ---------------------------------------------------------------------------
if show_charts:
    st.markdown(f'<div class="section-header" style="background:linear-gradient(135deg,{C_TEAL},#4ea8de);"><h2>🤝 Distribuição de Resultados — Sócios</h2></div>', unsafe_allow_html=True)
    socio_partner_total = combined["SocioPartner"].sum()
    socio_maldivas_total = combined["SocioMaldivas"].sum()
    denom = socio_partner_total + socio_maldivas_total
    partner_pct = (socio_partner_total / denom) if denom else partner_share
    maldivas_pct = 1 - partner_pct

    fig_dist = go.Figure()
    fig_dist.add_bar(name=f"Partner ({partner_pct:.0%})", x=meses_full, y=combined["SocioPartner"], marker_color=C_NAVY,
                      text=[fmt_r(v, mil_mode) for v in combined["SocioPartner"]], textposition="outside")
    fig_dist.add_bar(name=f"Maldivas ({maldivas_pct:.0%})", x=meses_full, y=combined["SocioMaldivas"], marker_color=C_TEAL_LIGHT,
                      text=[fmt_r(v, mil_mode) for v in combined["SocioMaldivas"]], textposition="outside")
    fig_dist.add_trace(go.Scatter(name="Resultado Total", x=meses_full, y=combined["ResultadoOperacional"], mode="lines+markers",
                                   line=dict(color=C_SUCCESS, width=3, dash="dot")))
    fig_dist.add_hline(y=0, line_color=C_DANGER)
    fig_dist.update_layout(height=480, barmode="group", paper_bgcolor="white", plot_bgcolor="rgba(0,0,0,0)",
                            legend=dict(orientation="h", y=-0.2), hovermode="x unified")
    st.plotly_chart(fig_dist, use_container_width=True, config=PLOTLY_CONFIG)

    box_color = C_SUCCESS if resultado_operacional >= 0 else C_DANGER
    st.markdown(f"""<div style="background:linear-gradient(135deg,{box_color},{box_color}cc);padding:22px 30px;border-radius:16px;text-align:center;">
    <span style="color:white;font-size:1.1rem;font-weight:700;">TOTAIS DO PERÍODO — Resultado: {fmt_r(resultado_operacional, mil_mode)}</span>
    <div style="display:flex;justify-content:center;gap:30px;margin-top:14px;flex-wrap:wrap;">
    <div style="background:rgba(255,255,255,.2);padding:12px 22px;border-radius:10px;color:white;"><div style="font-size:.85rem;">Partner ({partner_pct:.0%})</div><div style="font-size:1.3rem;font-weight:900;">{fmt_r(socio_partner_total, mil_mode)}</div></div>
    <div style="background:rgba(255,255,255,.2);padding:12px 22px;border-radius:10px;color:white;"><div style="font-size:.85rem;">Maldivas ({maldivas_pct:.0%})</div><div style="font-size:1.3rem;font-weight:900;">{fmt_r(socio_maldivas_total, mil_mode)}</div></div>
    <div style="background:rgba(255,255,255,.3);padding:12px 22px;border-radius:10px;color:white;"><div style="font-size:.85rem;">Status</div><div style="font-size:1.3rem;font-weight:900;">{status}</div></div>
    </div></div>""", unsafe_allow_html=True)

    quarter_series = combined.groupby([QUARTER_OF[m] for m in combined.index])["ValorPagarMaldivas"].sum()
    quarter_series = quarter_series.reindex([q for q in quarters_available if q in quarter_series.index])
    if not quarter_series.empty:
        st.markdown("##### 📅 Resultado Trimestral — Valor a Receber (Maldivas)")
        qcols2 = st.columns(len(quarter_series))
        for col, q in zip(qcols2, quarter_series.index):
            col.metric(q, fmt_r(quarter_series[q], mil_mode))

# ---------------------------------------------------------------------------
# Análise por Produto
# ---------------------------------------------------------------------------
if show_charts and combined_tx is not None:
    df_prod = combined_tx.groupby("Produto")["Valor"].sum().reset_index()
    df_prod = df_prod[df_prod["Valor"] > 0].sort_values("Valor", ascending=False)
    if not df_prod.empty:
        st.markdown(f'<div class="section-header" style="background:linear-gradient(135deg,{C_GOLD},#ff9f43);"><h2 style="color:{C_NAVY};">📦 Análise por Tipo de Produto</h2></div>', unsafe_allow_html=True)
        c1, c2 = st.columns(2)
        with c1:
            fig_p = px.sunburst(df_prod, path=["Produto"], values="Valor", color="Valor", color_continuous_scale="YlOrRd")
            fig_p.update_layout(height=520, paper_bgcolor="white")
            st.plotly_chart(fig_p, use_container_width=True, config=PLOTLY_CONFIG)
        with c2:
            fig_pb = go.Figure(go.Bar(y=df_prod["Produto"], x=df_prod["Valor"], orientation="h",
                                       marker=dict(color=df_prod["Valor"], colorscale="YlOrRd"),
                                       text=[fmt_r(v, mil_mode) for v in df_prod["Valor"]], textposition="outside"))
            fig_pb.update_layout(height=520, paper_bgcolor="white", plot_bgcolor="rgba(0,0,0,0)", yaxis=dict(categoryorder="total ascending"))
            st.plotly_chart(fig_pb, use_container_width=True, config=PLOTLY_CONFIG)

# ---------------------------------------------------------------------------
# Ranking Originadores
# ---------------------------------------------------------------------------
ranking_df = None
if show_charts and combined_tx is not None:
    df_orig = (combined_tx.groupby("Originador").agg(Valor=("Valor", "sum"), Operacoes=("Valor", "count")).reset_index())
    df_orig = df_orig[df_orig["Valor"] > 0].sort_values("Valor", ascending=False)
    df_orig["TicketMedio"] = df_orig["Valor"] / df_orig["Operacoes"]
    if not df_orig.empty:
        ranking_df = df_orig.head(3)
        st.markdown(f'<div class="section-header" style="background:linear-gradient(135deg,{C_TEAL},{C_SUCCESS});"><h2>👥 Ranking — Top Originadores</h2></div>', unsafe_allow_html=True)
        c1, c2 = st.columns([1, 1])
        with c1:
            top5 = df_orig.head(5)
            outros = df_orig.iloc[5:]["Valor"].sum()
            labels = list(top5["Originador"]) + (["Outros"] if outros > 0 else [])
            values = list(top5["Valor"]) + ([outros] if outros > 0 else [])
            fig_o = go.Figure(go.Pie(labels=labels, values=values, hole=.55, marker=dict(colors=GRADIENTE)))
            fig_o.update_layout(height=520, paper_bgcolor="white")
            st.plotly_chart(fig_o, use_container_width=True, config=PLOTLY_CONFIG)
        with c2:
            st.markdown("##### 🏆 Top 3 Originadores")
            medals, mcolors = ["🥇", "🥈", "🥉"], [C_MEDAL_GOLD, C_MEDAL_SILVER, C_MEDAL_BRONZE]
            for i, row in df_orig.head(3).reset_index(drop=True).iterrows():
                st.markdown(f"""<div class="ranking-card" style="border-color:{mcolors[i]};">
                <div style="display:flex;align-items:center;gap:16px;"><span style="font-size:2.4rem;">{medals[i]}</span>
                <div><b>{row['Originador']}</b><br><span style="font-size:1.3rem;font-weight:900;color:{C_TEAL};">{fmt_r(row['Valor'], mil_mode)}</span><br>
                <span style="font-size:.85rem;color:#6c757d;">{int(row['Operacoes'])} operações · Ticket médio {fmt_r(row['TicketMedio'], mil_mode)}</span></div></div></div>""", unsafe_allow_html=True)
        if show_tables:
            st.dataframe(df_orig.style.format({"Valor": lambda v: fmt_r(v, mil_mode), "TicketMedio": lambda v: fmt_r(v, mil_mode)}),
                         hide_index=True, use_container_width=True)
elif show_charts:
    st.info("Ranking de originadores indisponível: aba de transações (ex.: 'ASSERTIF DIRETO') não encontrada.")

# ---------------------------------------------------------------------------
# Ranking Clientes
# ---------------------------------------------------------------------------
if show_charts and combined_tx is not None:
    df_cli = combined_tx.groupby("Cliente")["Valor"].sum().reset_index()
    df_cli = df_cli[df_cli["Valor"] > 0].sort_values("Valor", ascending=False)
    if not df_cli.empty:
        st.markdown(f'<div class="section-header" style="background:linear-gradient(135deg,{C_SUCCESS},{C_TEAL});"><h2>🏅 Ranking — Maiores Clientes</h2></div>', unsafe_allow_html=True)
        top15 = df_cli.head(15)
        fig_cli = go.Figure(go.Bar(y=top15["Cliente"], x=top15["Valor"], orientation="h",
                                    marker=dict(color=top15["Valor"], colorscale="Tealgrn"),
                                    text=[fmt_r(v, mil_mode) for v in top15["Valor"]], textposition="outside"))
        fig_cli.update_layout(height=650, paper_bgcolor="white", plot_bgcolor="rgba(0,0,0,0)", yaxis=dict(categoryorder="total ascending"), margin=dict(l=260))
        st.plotly_chart(fig_cli, use_container_width=True, config=PLOTLY_CONFIG)

# ---------------------------------------------------------------------------
# Ranking Despesas
# ---------------------------------------------------------------------------
if show_charts and combined_desp is not None:
    df_cat = combined_desp.groupby("Categoria")["Valor"].sum().reset_index().sort_values("Valor", ascending=False)
    if not df_cat.empty:
        st.markdown(f'<div class="section-header" style="background:linear-gradient(135deg,{C_DANGER},#ee5a52);"><h2>💸 Ranking — Maiores Despesas</h2></div>', unsafe_allow_html=True)
        top10 = df_cat.head(10)
        fig_d = go.Figure(go.Bar(x=top10["Categoria"], y=top10["Valor"], marker=dict(color=top10["Valor"], colorscale="Reds"),
                                  text=[fmt_r(v, mil_mode) for v in top10["Valor"]], textposition="outside"))
        fig_d.update_layout(height=520, paper_bgcolor="white", plot_bgcolor="rgba(0,0,0,0)", xaxis_tickangle=-30)
        st.plotly_chart(fig_d, use_container_width=True, config=PLOTLY_CONFIG)
elif show_charts:
    st.info("Ranking de despesas indisponível: aba 'DESPESAS' não encontrada.")

# ---------------------------------------------------------------------------
# Resumo Executivo
# ---------------------------------------------------------------------------
resumo_linhas = [
    ("Faturamento Bruto", receita_total, True),
    ("  Produção Direta", combined["ReceitaDireta"].sum(), False),
    ("  Portal MAAS", combined["ReceitaPortal"].sum(), False),
    ("Impostos Diretos", combined["Impostos"].sum(), False),
    ("Custo Operacional (D.A)", combined["Custo"].sum(), False),
    ("Co-Corretagem", combined["CoCorretagem"].sum(), False),
    ("Rebate AAI", combined["RebateAAI"].sum(), False),
    ("Custos Totais", custos_totais, True),
    ("(=) Margem de Contribuição", margem_contrib, True),
    ("Despesas Operacionais", combined["Despesas"].sum(), False),
    ("Folha + Terceiros", combined["Folha"].sum(), False),
    ("Despesas Totais", despesas_total, True),
    ("Resultado Operacional", resultado_operacional, True),
]
if show_tables:
    st.markdown(f'<div class="section-header" style="background:linear-gradient(135deg,{C_NAVY},{C_NAVY2});"><h2>📋 Resumo Executivo — DRE</h2></div>', unsafe_allow_html=True)
    html = "<table style='width:100%;border-collapse:collapse;'>"
    for label, val, bold in resumo_linhas:
        w = "700" if bold else "400"
        bg = "#e8f0fe" if bold else "white"
        html += f"""<tr style="background:{bg};"><td style="padding:9px 14px;font-weight:{w};border-bottom:1px solid #eee;">{label}</td>
        <td style="padding:9px 14px;text-align:right;font-weight:{w};border-bottom:1px solid #eee;">{fmt_r(val, mil_mode)}</td></tr>"""
    html += "</table>"
    st.markdown(html, unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# PDF — ReportLab (paginação automática, fontes padrão já suportam acentos PT-BR)
# ---------------------------------------------------------------------------
def build_pdf():
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="CoverTitle", fontSize=30, textColor=colors.white, alignment=TA_CENTER, fontName="Helvetica-Bold", leading=36))
    styles.add(ParagraphStyle(name="CoverSub", fontSize=14, textColor=colors.white, alignment=TA_CENTER, fontName="Helvetica"))
    styles.add(ParagraphStyle(name="SecHeader", fontSize=14, textColor=colors.white, alignment=TA_LEFT, fontName="Helvetica-Bold"))
    styles.add(ParagraphStyle(name="NoteTxt", fontSize=9, textColor=HexColor(C_NAVY), alignment=TA_JUSTIFY, fontName="Helvetica", leading=13))
    styles.add(ParagraphStyle(name="TOCEntry", fontSize=11, textColor=HexColor(C_NAVY), fontName="Helvetica"))

    buffer = io.BytesIO()
    doc = SimpleDocTemplate(buffer, pagesize=A4, rightMargin=1.2*cm, leftMargin=1.2*cm, topMargin=2.2*cm, bottomMargin=1.6*cm)
    elements = []

    def section_header(title, color=C_NAVY):
        t = Table([[Paragraph(title, styles["SecHeader"])]], colWidths=[18*cm])
        t.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), HexColor(color)), ("TOPPADDING", (0, 0), (-1, -1), 10),
                                ("BOTTOMPADDING", (0, 0), (-1, -1), 10), ("LEFTPADDING", (0, 0), (-1, -1), 16)]))
        return t

    def data_table(headers, rows, col_widths=None, highlight=None):
        tdata = [headers] + rows
        col_widths = col_widths or [18*cm/len(headers)]*len(headers)
        t = Table(tdata, colWidths=col_widths)
        cmds = [("BACKGROUND", (0, 0), (-1, 0), HexColor(C_NAVY)), ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, -1), 9),
                ("GRID", (0, 0), (-1, -1), 0.4, HexColor("#e0e0e0")), ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("TOPPADDING", (0, 0), (-1, -1), 6), ("BOTTOMPADDING", (0, 0), (-1, -1), 6)]
        for i in range(1, len(tdata)):
            cmds.append(("BACKGROUND", (0, i), (-1, i), HexColor("#f8f9fa") if i % 2 == 0 else colors.white))
        if highlight:
            for idx in highlight:
                cmds.append(("FONTNAME", (0, idx), (-1, idx), "Helvetica-Bold"))
                cmds.append(("BACKGROUND", (0, idx), (-1, idx), HexColor("#e8f0fe")))
        t.setStyle(TableStyle(cmds))
        return t

    # Capa
    cover_rows = [
        [Spacer(1, 2*cm)],
        [Paragraph("DASHBOARD FINANCEIRO", styles["CoverTitle"])],
        [Spacer(1, 0.3*cm)],
        [Paragraph("Relatório Executivo Premium", styles["CoverSub"])],
        [Spacer(1, 1*cm)],
        [Paragraph(f"Planilhas: {' + '.join(selected_files)}", styles["CoverSub"])],
        [Paragraph(f"Período: {periodo_label}", styles["CoverSub"])],
        [Spacer(1, 0.6*cm)],
        [Paragraph(f"<b>Status: {status}  ·  Margem: {margem_lucro:.0%}</b>", styles["CoverSub"])],
        [Spacer(1, 2*cm)],
        [Paragraph(f"Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}", styles["CoverSub"])],
    ]
    cover = Table(cover_rows, colWidths=[18*cm])
    cover.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), HexColor(C_NAVY)), ("ALIGN", (0, 0), (-1, -1), "CENTER")]))
    elements.append(cover)
    elements.append(Spacer(1, 0.6*cm))
    info = Table([["FATURAMENTO", fmt_r(receita_total, mil_mode), "RESULTADO OPERACIONAL", fmt_r(resultado_operacional, mil_mode)]], colWidths=[5*cm, 4*cm, 6*cm, 3*cm])
    info.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), HexColor("#f8f9fa")), ("BOX", (0, 0), (-1, -1), 1.5, HexColor(C_NAVY)),
                               ("GRID", (0, 0), (-1, -1), 0.5, HexColor("#e0e0e0")), ("ALIGN", (0, 0), (-1, -1), "CENTER"),
                               ("FONTNAME", (0, 0), (-1, -1), "Helvetica-Bold"), ("TOPPADDING", (0, 0), (-1, -1), 10), ("BOTTOMPADDING", (0, 0), (-1, -1), 10)]))
    elements.append(info)
    elements.append(PageBreak())

    # KPIs
    elements.append(section_header("INDICADORES PRINCIPAIS (KPIs)"))
    elements.append(Spacer(1, 12))
    kpi_row = [[k[0] for k in kpi_defs], [k[1] for k in kpi_defs]]
    kt = Table(kpi_row, colWidths=[3.6*cm]*5)
    kt.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, 0), HexColor(C_NAVY2)), ("TEXTCOLOR", (0, 0), (-1, -1), colors.white),
                             ("BACKGROUND", (0, 1), (-1, 1), HexColor(C_NAVY)), ("FONTNAME", (0, 0), (-1, 0), "Helvetica"),
                             ("FONTNAME", (0, 1), (-1, 1), "Helvetica-Bold"), ("FONTSIZE", (0, 0), (-1, 0), 8), ("FONTSIZE", (0, 1), (-1, 1), 11),
                             ("ALIGN", (0, 0), (-1, -1), "CENTER"), ("TOPPADDING", (0, 0), (-1, -1), 8), ("BOTTOMPADDING", (0, 0), (-1, -1), 8)]))
    elements.append(kt)
    elements.append(Spacer(1, 14))
    nota = Table([[Paragraph("<b>Faturamento:</b> Receita Direta+Portal · <b>Custos Totais:</b> Impostos+Custo D.A+Rebate AAI−Co-Corretagem · "
                              "<b>Margem de Contribuição:</b> Faturamento−Custos · <b>Despesas Totais:</b> Operacionais+Folha/Terceiros · "
                              "<b>Resultado Operacional:</b> Margem−Despesas", styles["NoteTxt"])]], colWidths=[18*cm])
    nota.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), HexColor("#f8f9fa")), ("BOX", (0, 0), (-1, -1), 1, HexColor(C_TEAL)),
                               ("TOPPADDING", (0, 0), (-1, -1), 10), ("BOTTOMPADDING", (0, 0), (-1, -1), 10), ("LEFTPADDING", (0, 0), (-1, -1), 12)]))
    elements.append(nota)
    elements.append(Spacer(1, 16))

    # Evolução mensal
    elements.append(section_header("EVOLUÇÃO MENSAL — RECEITA vs RESULTADO", C_NAVY2))
    elements.append(Spacer(1, 12))
    drawing = Drawing(480, 190)
    lc = HorizontalLineChart()
    lc.x, lc.y, lc.height, lc.width = 45, 35, 120, 400
    lc.data = [list(receita_mensal.values)]
    lc.categoryAxis.categoryNames = meses_full
    lc.valueAxis.valueMin = 0
    lc.valueAxis.valueMax = max(receita_mensal.max() * 1.2, 1)
    lc.lines[0].strokeColor = HexColor(C_NAVY)
    lc.lines[0].strokeWidth = 3
    lc.lines[0].symbol = makeMarker("Circle")
    drawing.add(lc)
    drawing.add(String(240, 175, "Receita Bruta Mensal (R$)", fontName="Helvetica-Bold", fontSize=10, textAnchor="middle"))
    elements.append(drawing)
    elements.append(Spacer(1, 10))
    ev_rows = [[m, fmt_r(r, mil_mode), ("-" if pd.isna(c) else f"{c:+.1f}%"), fmt_r(ro, mil_mode)]
               for m, r, c, ro in zip(meses_full, receita_mensal.values, crescimento.values, combined["ResultadoOperacional"].values)]
    elements.append(data_table(["Mês", "Receita Bruta", "Crescimento", "Resultado Operacional"], ev_rows, [4*cm, 5*cm, 4*cm, 5*cm]))
    elements.append(PageBreak())

    # Sócios
    elements.append(section_header("DISTRIBUIÇÃO DE RESULTADOS — SÓCIOS", C_TEAL))
    elements.append(Spacer(1, 12))
    soc_rows = [[m, fmt_r(p, mil_mode), fmt_r(mv, mil_mode)] for m, p, mv in zip(meses_full, combined["SocioPartner"], combined["SocioMaldivas"])]
    elements.append(data_table(["Mês", f"Partner ({partner_pct:.0%})", f"Maldivas ({maldivas_pct:.0%})"], soc_rows, [6*cm, 6*cm, 6*cm]))
    elements.append(Spacer(1, 16))

    # Rankings
    if combined_tx is not None:
        df_seg_pdf = combined_tx.groupby("Seguradora")["Valor"].sum().reset_index()
        df_seg_pdf = df_seg_pdf[df_seg_pdf["Valor"] > 0].sort_values("Valor", ascending=False).head(10)
        if not df_seg_pdf.empty:
            elements.append(section_header("RANKING — SEGURADORAS (TOP 10)", C_NAVY))
            elements.append(Spacer(1, 10))
            elements.append(data_table(["Seguradora", "Comissão"], [[r.Seguradora, fmt_r(r.Valor, mil_mode)] for r in df_seg_pdf.itertuples()], [12*cm, 6*cm]))
            elements.append(Spacer(1, 16))

        if ranking_df is not None:
            elements.append(section_header("RANKING — TOP ORIGINADORES", C_TEAL))
            elements.append(Spacer(1, 10))
            elements.append(data_table(["Originador", "Valor", "Operações", "Ticket Médio"],
                                        [[r.Originador, fmt_r(r.Valor, mil_mode), str(int(r.Operacoes)), fmt_r(r.TicketMedio, mil_mode)] for r in ranking_df.itertuples()],
                                        [6*cm, 4*cm, 4*cm, 4*cm]))
            elements.append(Spacer(1, 16))

        df_cli_pdf = combined_tx.groupby("Cliente")["Valor"].sum().reset_index()
        df_cli_pdf = df_cli_pdf[df_cli_pdf["Valor"] > 0].sort_values("Valor", ascending=False).head(10)
        if not df_cli_pdf.empty:
            elements.append(section_header("RANKING — CLIENTES (TOP 10)", C_SUCCESS))
            elements.append(Spacer(1, 10))
            elements.append(data_table(["Cliente", "Receita"], [[r.Cliente, fmt_r(r.Valor, mil_mode)] for r in df_cli_pdf.itertuples()], [12*cm, 6*cm]))
            elements.append(Spacer(1, 16))

    if combined_desp is not None:
        df_cat_pdf = combined_desp.groupby("Categoria")["Valor"].sum().reset_index().sort_values("Valor", ascending=False).head(10)
        if not df_cat_pdf.empty:
            elements.append(section_header("RANKING — DESPESAS (TOP 10)", C_DANGER))
            elements.append(Spacer(1, 10))
            elements.append(data_table(["Categoria", "Valor"], [[r.Categoria, fmt_r(r.Valor, mil_mode)] for r in df_cat_pdf.itertuples()], [12*cm, 6*cm]))
            elements.append(Spacer(1, 16))

    elements.append(PageBreak())

    # Resumo executivo
    elements.append(section_header("RESUMO EXECUTIVO — DRE", C_NAVY))
    elements.append(Spacer(1, 10))
    bold_idx = [i + 1 for i, (_, _, b) in enumerate(resumo_linhas) if b]
    elements.append(data_table(["Indicador", "Valor"], [[l, fmt_r(v, mil_mode)] for l, v, _ in resumo_linhas], [12*cm, 6*cm], highlight=bold_idx))

    def footer(canvas, doc_):
        canvas.saveState()
        canvas.setFillColor(HexColor(C_NAVY))
        canvas.rect(0, 0, A4[0], 1.1*cm, fill=True, stroke=False)
        canvas.setFillColor(colors.white)
        canvas.setFont("Helvetica", 8)
        canvas.drawString(1.2*cm, 0.4*cm, f"Gerado em {datetime.now().strftime('%d/%m/%Y')}")
        canvas.drawCentredString(A4[0]/2, 0.4*cm, "Dashboard Financeiro Premium")
        canvas.drawRightString(A4[0]-1.2*cm, 0.4*cm, f"Página {doc_.page}")
        canvas.restoreState()

    doc.build(elements, onFirstPage=footer, onLaterPages=footer)
    buffer.seek(0)
    return buffer.getvalue()

st.markdown("---")
st.markdown(f'<div class="section-header" style="background:linear-gradient(135deg,{C_SUCCESS},#00b894);"><h2>📥 Exportar Dashboard para PDF</h2></div>', unsafe_allow_html=True)
c1, c2, c3 = st.columns([1, 2, 1])
with c2:
    if st.button("📄 GERAR PDF PROFISSIONAL", type="primary", use_container_width=True):
        with st.spinner("Gerando PDF..."):
            pdf_bytes = build_pdf()
        st.success("PDF gerado com sucesso!")
        st.download_button("⬇️ Baixar PDF", data=pdf_bytes, file_name=f"dashboard_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                            mime="application/pdf", use_container_width=True)
