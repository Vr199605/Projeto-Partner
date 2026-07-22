# app.py
import re
import io
from datetime import datetime

import streamlit as st
import pandas as pd
import numpy as np
import openpyxl
import plotly.graph_objects as go
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from fpdf import FPDF

st.set_page_config(page_title="Dashboard Financeiro Premium", layout="wide", page_icon="📊")

# ---------------------------------------------------------------------------
# Paleta (extraída do modelo de referência)
# ---------------------------------------------------------------------------
C_PRIMARY = "#667eea"
C_PRIMARY_DARK = "#1d3a5e"
C_GOLD = "#FFD700"
C_SILVER = "#C0C0C0"
C_BRONZE = "#CD7F32"
C_GREEN = "#28a745"
C_GREEN_BG = "#e8f5e9"
C_RED = "#db3545"
C_RED_BG = "#fdeaea"
C_BLUE_BG = "#e8f0fe"
C_GRAY_BG = "#f8f9fa"

MESES_ORDEM = ["Jan", "Fev", "Mar", "Abr", "Mai", "Jun", "Jul", "Ago", "Set", "Out", "Nov", "Dez"]
MES_PREFIX = {"JAN": "Jan", "FEV": "Fev", "MAR": "Mar", "ABR": "Abr", "MAI": "Mai", "JUN": "Jun",
              "JUL": "Jul", "AGO": "Ago", "SET": "Set", "OUT": "Out", "NOV": "Nov", "DEZ": "Dez"}
QUARTER_OF = {**{m: "1º Tri" for m in ["Jan", "Fev", "Mar"]},
              **{m: "2º Tri" for m in ["Abr", "Mai", "Jun"]},
              **{m: "3º Tri" for m in ["Jul", "Ago", "Set"]},
              **{m: "4º Tri" for m in ["Out", "Nov", "Dez"]}}

# ---------------------------------------------------------------------------
# CSS
# ---------------------------------------------------------------------------
st.markdown(f"""
<style>
.hero {{
    background: linear-gradient(135deg, {C_PRIMARY} 0%, #764ba2 100%);
    padding: 28px 32px; border-radius: 14px; color: white; margin-bottom: 18px;
}}
.hero h1 {{ margin:0; font-size: 28px; }}
.hero p {{ margin:4px 0 0 0; opacity:.9; }}
.kpi-card {{
    border-radius: 14px; padding: 18px 16px; text-align:center;
    box-shadow: 0 2px 10px rgba(0,0,0,.08); background:white; border:1px solid #eee;
}}
.kpi-label {{ font-size:12px; font-weight:700; color:#6b747d; letter-spacing:.5px; text-transform:uppercase;}}
.kpi-value {{ font-size:26px; font-weight:800; color:{C_PRIMARY_DARK}; margin:6px 0 2px 0;}}
.kpi-sub {{ font-size:11px; color:#8c8c8c; }}
.section-title {{ font-size:19px; font-weight:800; color:{C_PRIMARY_DARK}; margin:22px 0 10px 0;
    border-bottom:2px solid {C_PRIMARY}; padding-bottom:6px; }}
.socio-card {{ border-radius:14px; padding:16px; color:white; }}
.rank-card {{ border-radius:14px; padding:14px 16px; display:flex; align-items:center; gap:14px;
    box-shadow:0 2px 8px rgba(0,0,0,.06); margin-bottom:10px; background:white; }}
.rank-medal {{ font-size:30px; }}
</style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Parsing helpers
# ---------------------------------------------------------------------------
def _norm(v):
    return str(v).strip().upper() if v is not None else ""

def find_row(ws, label, start=1, end=None):
    end = end or ws.max_row
    target = _norm(label)
    for r in range(start, end + 1):
        if target in _norm(ws.cell(row=r, column=2).value):
            return r
    return None

def find_header_row(ws):
    for r in range(1, ws.max_row + 1):
        vals = [_norm(ws.cell(row=r, column=c).value)[:3] for c in range(3, ws.max_column + 1)]
        if sum(1 for v in vals if v in [m.upper() for m in MES_PREFIX.values()]) >= 6:
            return r
    return None

def detect_month_name(sheet_name):
    clean = re.sub(r"[^A-Za-zÀ-ÿ]", "", sheet_name).upper()[:3]
    clean = clean.replace("Ç", "C")
    return MES_PREFIX.get(clean)

def row_series(ws, row, cols):
    return [ws.cell(row=row, column=c).value or 0 for c in cols]

# ---------------------------------------------------------------------------
# Parse "DRE 2026"
# ---------------------------------------------------------------------------
def parse_dre(ws):
    header_row = find_header_row(ws)
    if header_row is None:
        return None
    month_cols, ytd_col = [], None
    for c in range(3, ws.max_column + 1):
        v = _norm(ws.cell(row=header_row, column=c).value)
        if v == "YTD":
            ytd_col = c
        elif v[:3] in [m.upper() for m in MES_PREFIX.values()]:
            month_cols.append(c)
    if not month_cols:
        return None

    total_row = find_row(ws, "RECEITA BRUTA TOTAL")
    direta_sec = find_row(ws, "PRODUÇÃO DIRETA")
    portal_sec = find_row(ws, "PORTAL MAAS")
    resultado_row = find_row(ws, "RESULTADO OPERACIONAL", start=portal_sec)
    while resultado_row and "DISTRIBUI" in _norm(ws.cell(row=resultado_row, column=2).value):
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
        "d_ebitda": find_row(ws, "EBITDA SOCIETÁRIO", direta_sec, d_end),
        "p_receita": find_row(ws, "RECEITA BRUTA", portal_sec, p_end),
        "p_impostos": find_row(ws, "IMPOSTOS DIRETOS", portal_sec, p_end),
        "p_custo": find_row(ws, "CUSTO OPERACIONAL", portal_sec, p_end),
        "p_margem": find_row(ws, "MARGEM DE CONTRIBUIÇÃO", portal_sec, p_end),
        "p_despesas": find_row(ws, "DESPESAS", portal_sec, p_end),
        "p_folha": find_row(ws, "FOLHA+TERCEIROS", portal_sec, p_end),
        "p_ebitda": find_row(ws, "EBITDA SOCIETÁRIO", portal_sec, p_end),
    }
    socio_partner_row = find_row(ws, "Sócio Partner", start=resultado_row)
    socio_maldivas_row = find_row(ws, "Sócio Maldivas", start=resultado_row)
    valor_pagar_maldivas_row = find_row(ws, "Valor a pagar", start=resultado_row)

    def s(key):
        r = rows.get(key)
        return row_series(ws, r, month_cols) if r else [0] * len(month_cols)

    def rs(row):
        return row_series(ws, row, month_cols) if row else [0] * len(month_cols)

    meses = [MES_PREFIX[_norm(ws.cell(row=header_row, column=c).value)[:3]] for c in month_cols]

    df = pd.DataFrame({
        "ReceitaDireta": s("d_receita"), "ReceitaPortal": s("p_receita"),
        "Impostos": [a + b for a, b in zip(s("d_impostos"), s("p_impostos"))],
        "Custo": [a + b for a, b in zip(s("d_custo"), s("p_custo"))],
        "CoCorretagem": s("d_cocorretagem"), "RebateAAI": s("d_rebate"),
        "MargemContribuicao": [a + b for a, b in zip(s("d_margem"), s("p_margem"))],
        "Despesas": [a + b for a, b in zip(s("d_despesas"), s("p_despesas"))],
        "Folha": [a + b for a, b in zip(s("d_folha"), s("p_folha"))],
        "EBITDA_Direta": s("d_ebitda"), "EBITDA_Portal": s("p_ebitda"),
        "ResultadoOperacional": rs(resultado_row),
        "SocioPartner": rs(socio_partner_row), "SocioMaldivas": rs(socio_maldivas_row),
        "ValorPagarMaldivas": rs(valor_pagar_maldivas_row),
    }, index=meses)
    df = df.groupby(df.index).sum()
    df = df.reindex([m for m in MESES_ORDEM if m in df.index])
    return df

def parse_shares(wb):
    if "INPUTS" not in wb.sheetnames:
        return 0.7, 0.3
    ws = wb["INPUTS"]
    rp = find_row(ws, "Sócio Partner")
    rm = find_row(ws, "Sócio Maldivas")
    partner = ws.cell(row=rp, column=2).value if rp else 0.7
    maldivas = ws.cell(row=rm, column=2).value if rm else 0.3
    return float(partner or 0.7), float(maldivas or 0.3)

def parse_originadores(wb):
    if "BASE" not in wb.sheetnames:
        return None
    base_ws = wb["BASE"]
    code_col = name_col = None
    for c in range(1, base_ws.max_column + 1):
        h = _norm(base_ws.cell(row=1, column=c).value)
        if "COD INTERNO" in h:
            code_col = c
        if "NOME ASSESSOR" in h:
            name_col = c
    if not code_col or not name_col:
        return None
    code_to_name = {}
    for r in range(2, base_ws.max_row + 1):
        code = base_ws.cell(row=r, column=code_col).value
        name = base_ws.cell(row=r, column=name_col).value
        if code and name and code not in code_to_name:
            code_to_name[str(code).strip()] = str(name).strip()

    records = []
    for sheet_name in wb.sheetnames:
        mes = detect_month_name(sheet_name)
        if not mes:
            continue
        ws = wb[sheet_name]
        assessor_col = valor_col = None
        for c in range(1, ws.max_column + 1):
            h = _norm(ws.cell(row=2, column=c).value)
            if "ASSESSOR" in h and "COD" in h:
                assessor_col = c
            if "COMISS" in h and "D.A" in h and "PARTNER" not in h and "RECEBIDA" not in h:
                valor_col = c
        if not assessor_col or not valor_col:
            continue
        for r in range(3, ws.max_row + 1):
            code = ws.cell(row=r, column=assessor_col).value
            val = ws.cell(row=r, column=valor_col).value
            if not code or not val:
                continue
            name = code_to_name.get(str(code).strip(), str(code).strip())
            records.append((mes, name, float(val)))
    if not records:
        return None
    return pd.DataFrame(records, columns=["Mes", "Assessor", "Valor"])

@st.cache_data(show_spinner=False)
def parse_workbook(file_bytes, file_name):
    wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)
    if "DRE 2026" not in wb.sheetnames:
        return None
    df = parse_dre(wb["DRE 2026"])
    if df is None:
        return None
    shares = parse_shares(wb)
    tx = parse_originadores(wb)
    return {"file_name": file_name, "df": df, "shares": shares, "tx": tx}

# ---------------------------------------------------------------------------
# Agregação conforme filtros
# ---------------------------------------------------------------------------
def aggregate(parsed_list, selected_files, selected_months):
    sel = [p for p in parsed_list if p["file_name"] in selected_files]
    if not sel:
        return None
    parts = [p["df"].loc[p["df"].index.intersection(selected_months)] for p in sel]
    combined = pd.concat(parts).groupby(level=0).sum()
    combined = combined.reindex([m for m in MESES_ORDEM if m in combined.index])

    tx_parts = [p["tx"] for p in sel if p["tx"] is not None]
    combined_tx = pd.concat(tx_parts) if tx_parts else None
    if combined_tx is not None:
        combined_tx = combined_tx[combined_tx["Mes"].isin(selected_months)]

    partner_share = float(np.mean([p["shares"][0] for p in sel]))
    maldivas_share = float(np.mean([p["shares"][1] for p in sel]))
    return combined, combined_tx, (partner_share, maldivas_share)

def fmt_r(v):
    v = 0 if v is None or (isinstance(v, float) and np.isnan(v)) else v
    sign = "-" if v < 0 else ""
    return f"{sign}R$ {abs(v):,.0f}".replace(",", ".")

def fmt_pct(v):
    if v is None or (isinstance(v, float) and np.isnan(v)):
        return "-"
    return f"{v:.1%}"

# ---------------------------------------------------------------------------
# UI — Upload e filtros
# ---------------------------------------------------------------------------
st.sidebar.header("📁 Upload")
uploads = st.sidebar.file_uploader("Envie uma ou mais planilhas (.xlsx)", type=["xlsx"], accept_multiple_files=True)

if not uploads:
    st.info("Envie uma ou mais planilhas contendo a aba 'DRE 2026' para gerar o dashboard.")
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
all_months = sorted({m for p in parsed_list for m in p["df"].index}, key=lambda m: MESES_ORDEM.index(m))

st.sidebar.header("🔎 Filtros")
if "sel_files" not in st.session_state:
    st.session_state.sel_files = all_files
if "sel_months" not in st.session_state:
    st.session_state.sel_months = all_months

c1, c2 = st.sidebar.columns(2)
if c1.button("Todas as planilhas"):
    st.session_state.sel_files = all_files
if c2.button("Todos os meses"):
    st.session_state.sel_months = all_months

selected_files = st.sidebar.multiselect("Planilhas", all_files, key="sel_files")
selected_months = st.sidebar.multiselect("Meses", all_months, key="sel_months")

if not selected_files or not selected_months:
    st.warning("Selecione ao menos uma planilha e um mês.")
    st.stop()

combined, combined_tx, (partner_share, maldivas_share) = aggregate(parsed_list, selected_files, selected_months)

# ---------------------------------------------------------------------------
# KPIs
# ---------------------------------------------------------------------------
receita_total = combined["ReceitaDireta"].sum() + combined["ReceitaPortal"].sum()
resultado_operacional = combined["ResultadoOperacional"].sum()
margem_lucro = (resultado_operacional / receita_total) if receita_total else 0
status = "LUCRO" if resultado_operacional >= 0 else "DÉFICIT"

periodo_label = f"{selected_months[0]} a {selected_months[-1]} {2026}" if len(selected_months) > 1 else f"{selected_months[0]} 2026"

# ---------------------------------------------------------------------------
# Header
# ---------------------------------------------------------------------------
st.markdown(f"""
<div class="hero">
  <h1>📊 Dashboard Financeiro Premium</h1>
  <p>{" + ".join(selected_files)} · Período: {periodo_label}</p>
  <p style="font-size:12px;">Gerado em {datetime.now().strftime('%d/%m/%Y às %H:%M')}</p>
</div>
""", unsafe_allow_html=True)

st.markdown('<div class="section-title">💰 Indicadores Principais</div>', unsafe_allow_html=True)
k1, k2, k3, k4 = st.columns(4)
kpi_defs = [
    (k1, "Faturamento", fmt_r(receita_total), periodo_label),
    (k2, "Lucro Líquido", fmt_r(resultado_operacional), f"{margem_lucro:.0%} margem"),
    (k3, "Margem de Lucro", fmt_pct(margem_lucro), f"Status: {status}"),
    (k4, "EBITDA", fmt_r(resultado_operacional), "Resultado Operacional"),
]
for col, label, value, sub in kpi_defs:
    col.markdown(f"""<div class="kpi-card"><div class="kpi-label">{label}</div>
    <div class="kpi-value">{value}</div><div class="kpi-sub">{sub}</div></div>""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Evolução Mensal
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">📈 Evolução Mensal — Receita vs Resultado</div>', unsafe_allow_html=True)
receita_mensal = combined["ReceitaDireta"] + combined["ReceitaPortal"]
crescimento = receita_mensal.pct_change() * 100

colA, colB = st.columns(2)
with colA:
    tab = pd.DataFrame({
        "Mês": combined.index, "Receita Bruta": receita_mensal.values,
        "Crescimento": ["-"] + [f"{v:+.1f}%" for v in crescimento.values[1:]]
    })
    st.dataframe(tab.style.format({"Receita Bruta": fmt_r}), hide_index=True, use_container_width=True)
with colB:
    tab2 = pd.DataFrame({"Mês": combined.index, "Resultado Operacional": combined["ResultadoOperacional"].values})
    st.dataframe(tab2.style.format({"Resultado Operacional": fmt_r}), hide_index=True, use_container_width=True)

fig_evol = go.Figure()
fig_evol.add_bar(x=combined.index, y=combined["ReceitaDireta"], name="Receita Direta", marker_color=C_PRIMARY)
fig_evol.add_bar(x=combined.index, y=combined["ReceitaPortal"], name="Receita Portal MAAS", marker_color="#0F9ED5")
fig_evol.add_trace(go.Scatter(x=combined.index, y=combined["ResultadoOperacional"], name="Resultado Operacional",
                               mode="lines+markers", line=dict(color=C_GREEN, width=3), yaxis="y2"))
fig_evol.update_layout(barmode="group", yaxis_title="Receita (R$)",
                        yaxis2=dict(title="Resultado (R$)", overlaying="y", side="right"),
                        legend=dict(orientation="h", y=-0.2))
st.plotly_chart(fig_evol, use_container_width=True)

# ---------------------------------------------------------------------------
# Distribuição de Resultados — Sócios
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">🤝 Distribuição de Resultados — Sócios</div>', unsafe_allow_html=True)
socio_partner_total = combined["SocioPartner"].sum()
socio_maldivas_total = combined["SocioMaldivas"].sum()
denom = socio_partner_total + socio_maldivas_total
partner_pct = (socio_partner_total / denom) if denom else partner_share
maldivas_pct = 1 - partner_pct

s1, s2 = st.columns(2)
s1.markdown(f"""<div class="socio-card" style="background:linear-gradient(135deg,{C_PRIMARY},#764ba2);">
<b>👔 Sócio Partner ({partner_pct:.0%})</b><h2 style="margin:6px 0;">{fmt_r(socio_partner_total)}</h2>
<span style="font-size:12px;opacity:.9;">Participação majoritária no resultado</span></div>""", unsafe_allow_html=True)
s2.markdown(f"""<div class="socio-card" style="background:linear-gradient(135deg,#0F9ED5,#0B4A63);">
<b>🏝️ Sócio Maldivas ({maldivas_pct:.0%})</b><h2 style="margin:6px 0;">{fmt_r(socio_maldivas_total)}</h2>
<span style="font-size:12px;opacity:.9;">Participação operacional</span></div>""", unsafe_allow_html=True)

st.write("")
d1, d2 = st.columns(2)
with d1:
    st.dataframe(pd.DataFrame({"Mês": combined.index, "Valor Partner": combined["SocioPartner"].values})
                 .style.format({"Valor Partner": fmt_r}), hide_index=True, use_container_width=True)
with d2:
    st.dataframe(pd.DataFrame({"Mês": combined.index, "Valor Maldivas": combined["SocioMaldivas"].values})
                 .style.format({"Valor Maldivas": fmt_r}), hide_index=True, use_container_width=True)

# ---------------------------------------------------------------------------
# Ranking Top Originadores (se a planilha tiver aba BASE + abas mensais)
# ---------------------------------------------------------------------------
ranking_df = None
if combined_tx is not None and not combined_tx.empty:
    st.markdown('<div class="section-title">👥 Ranking — Top Originadores</div>', unsafe_allow_html=True)
    ranking_df = (combined_tx.groupby("Assessor")
                  .agg(Valor=("Valor", "sum"), Operacoes=("Valor", "count"))
                  .assign(TicketMedio=lambda d: d["Valor"] / d["Operacoes"])
                  .sort_values("Valor", ascending=False).head(3).reset_index())
    medals = ["🥇", "🥈", "🥉"]
    colors = [C_GOLD, C_SILVER, C_BRONZE]
    for i, row in ranking_df.iterrows():
        st.markdown(f"""<div class="rank-card" style="border-left:6px solid {colors[i]};">
        <div class="rank-medal">{medals[i]}</div>
        <div><b>{row['Assessor']}</b><br>
        <span style="font-size:20px;font-weight:800;color:{C_PRIMARY_DARK};">{fmt_r(row['Valor'])}</span><br>
        <span style="font-size:12px;color:#8c8c8c;">{int(row['Operacoes'])} operações | Ticket médio: {fmt_r(row['TicketMedio'])}</span>
        </div></div>""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Resumo Executivo
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">📋 Resumo Executivo</div>', unsafe_allow_html=True)
resumo_linhas = [
    ("Receita Bruta Total", receita_total, True),
    ("Receita Bruta – Produção Direta", combined["ReceitaDireta"].sum(), False),
    ("Receita Bruta – Portal MAAS", combined["ReceitaPortal"].sum(), False),
    ("Impostos Diretos", combined["Impostos"].sum(), False),
    ("Custo Operacional (D.A)", combined["Custo"].sum(), False),
    ("Co-Corretagem", combined["CoCorretagem"].sum(), False),
    ("Rebate AAI", combined["RebateAAI"].sum(), False),
    ("(=) Margem de Contribuição", combined["MargemContribuicao"].sum(), True),
    ("Despesas", combined["Despesas"].sum(), False),
    ("Folha + Terceiros", combined["Folha"].sum(), False),
    ("EBITDA Societário", combined["EBITDA_Direta"].sum() + combined["EBITDA_Portal"].sum(), True),
    ("Resultado Operacional Total", resultado_operacional, True),
]
resumo_html = "<table style='width:100%;border-collapse:collapse;'>"
for label, val, bold in resumo_linhas:
    w = "700" if bold else "400"
    bg = C_BLUE_BG if bold else "white"
    resumo_html += f"""<tr style="background:{bg};">
    <td style="padding:8px 12px;font-weight:{w};border-bottom:1px solid #eee;">{label}</td>
    <td style="padding:8px 12px;text-align:right;font-weight:{w};border-bottom:1px solid #eee;">{fmt_r(val)}</td></tr>"""
resumo_html += "</table>"
st.markdown(resumo_html, unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Resultado Trimestral (Maldivas)
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">📅 Resultado Trimestral — Valor a Receber (Maldivas)</div>', unsafe_allow_html=True)
quarter_series = combined.groupby([QUARTER_OF[m] for m in combined.index])["ValorPagarMaldivas"].sum()
quarter_series = quarter_series.reindex(["1º Tri", "2º Tri", "3º Tri", "4º Tri"]).fillna(0)
qcols = st.columns(4)
for col, q in zip(qcols, quarter_series.index):
    col.markdown(f"""<div class="kpi-card"><div class="kpi-label">{q}</div>
    <div class="kpi-value" style="font-size:20px;">{fmt_r(quarter_series[q])}</div></div>""", unsafe_allow_html=True)

# ---------------------------------------------------------------------------
# Composição (pizzas)
# ---------------------------------------------------------------------------
st.markdown('<div class="section-title">🧩 Composição (Período Selecionado)</div>', unsafe_allow_html=True)
p1, p2 = st.columns(2)
with p1:
    fig_p1 = go.Figure(go.Pie(labels=["Direta", "Portal MAAS"],
                               values=[combined["ReceitaDireta"].sum(), combined["ReceitaPortal"].sum()],
                               marker_colors=[C_PRIMARY, "#0F9ED5"], hole=.45))
    fig_p1.update_layout(title="Composição da Receita Bruta")
    st.plotly_chart(fig_p1, use_container_width=True)
with p2:
    fig_p2 = go.Figure(go.Pie(labels=["Sócio Partner", "Sócio Maldivas"],
                               values=[socio_partner_total, socio_maldivas_total],
                               marker_colors=[C_PRIMARY, "#0F9ED5"], hole=.45))
    fig_p2.update_layout(title="Distribuição do Resultado")
    st.plotly_chart(fig_p2, use_container_width=True)

# ---------------------------------------------------------------------------
# Exportação em PDF
# ---------------------------------------------------------------------------
def clean(txt):
    return str(txt).encode("latin-1", "replace").decode("latin-1")

def mpl_bar_line(combined):
    fig, ax1 = plt.subplots(figsize=(7, 3.2))
    x = np.arange(len(combined.index))
    ax1.bar(x - 0.2, combined["ReceitaDireta"], width=0.4, label="Receita Direta", color="#667eea")
    ax1.bar(x + 0.2, combined["ReceitaPortal"], width=0.4, label="Receita Portal", color="#0F9ED5")
    ax1.set_xticks(x); ax1.set_xticklabels(combined.index)
    ax2 = ax1.twinx()
    ax2.plot(x, combined["ResultadoOperacional"], color="#28a745", marker="o", label="Resultado Operacional")
    fig.legend(loc="upper center", ncol=3, bbox_to_anchor=(0.5, 1.05))
    fig.tight_layout()
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=150); plt.close(fig); buf.seek(0)
    return buf

def mpl_pie(labels, values, colors, title):
    fig, ax = plt.subplots(figsize=(3.4, 3.4))
    ax.pie(values, labels=labels, colors=colors, autopct="%1.0f%%", startangle=90)
    ax.set_title(title)
    fig.tight_layout()
    buf = io.BytesIO(); fig.savefig(buf, format="png", dpi=150); plt.close(fig); buf.seek(0)
    return buf

def build_pdf():
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()

    pdf.set_fill_color(102, 126, 234)
    pdf.rect(0, 0, 210, 28, "F")
    pdf.set_text_color(255, 255, 255)
    pdf.set_xy(10, 6)
    pdf.set_font("Helvetica", "B", 18)
    pdf.cell(0, 10, clean("Dashboard Financeiro Premium"), ln=1)
    pdf.set_font("Helvetica", "", 10)
    pdf.set_x(10)
    pdf.cell(0, 6, clean(f"{' + '.join(selected_files)}  |  Periodo: {periodo_label}"), ln=1)
    pdf.set_x(10)
    pdf.cell(0, 6, clean(f"Gerado em {datetime.now().strftime('%d/%m/%Y as %H:%M')}"), ln=1)

    pdf.set_text_color(0, 0, 0)
    pdf.ln(14)
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 8, clean("Indicadores Principais"), ln=1)

    kpis_pdf = [("Faturamento", fmt_r(receita_total)), ("Lucro Liquido", fmt_r(resultado_operacional)),
                ("Margem de Lucro", fmt_pct(margem_lucro)), ("EBITDA", fmt_r(resultado_operacional))]
    box_w = 47
    x0 = 10
    for i, (label, value) in enumerate(kpis_pdf):
        x = x0 + i * (box_w + 2)
        pdf.set_fill_color(232, 240, 254)
        pdf.rect(x, pdf.get_y(), box_w, 22, "F")
        pdf.set_xy(x, pdf.get_y() + 2)
        pdf.set_font("Helvetica", "", 8)
        pdf.multi_cell(box_w, 4, clean(label), align="C")
        pdf.set_xy(x, pdf.get_y())
        pdf.set_font("Helvetica", "B", 12)
        pdf.multi_cell(box_w, 6, clean(value), align="C")
    pdf.ln(18)

    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 8, clean("Evolucao Mensal"), ln=1)
    img_buf = mpl_bar_line(combined)
    pdf.image(img_buf, x=10, w=190)
    pdf.ln(4)

    if pdf.get_y() > 200:
        pdf.add_page()
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 8, clean("Distribuicao de Resultados - Socios"), ln=1)
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(95, 7, clean(f"Socio Partner ({partner_pct:.0%}): {fmt_r(socio_partner_total)}"))
    pdf.cell(95, 7, clean(f"Socio Maldivas ({maldivas_pct:.0%}): {fmt_r(socio_maldivas_total)}"), ln=1)
    pdf.ln(2)

    if ranking_df is not None:
        pdf.set_font("Helvetica", "B", 13)
        pdf.cell(0, 8, clean("Ranking - Top Originadores"), ln=1)
        pdf.set_font("Helvetica", "", 10)
        medals_txt = ["1o", "2o", "3o"]
        for i, row in ranking_df.iterrows():
            pdf.cell(0, 6, clean(f"{medals_txt[i]} {row['Assessor']} - {fmt_r(row['Valor'])} "
                                  f"({int(row['Operacoes'])} operacoes, ticket medio {fmt_r(row['TicketMedio'])})"), ln=1)
        pdf.ln(2)

    if pdf.get_y() > 200:
        pdf.add_page()
    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 8, clean("Resumo Executivo"), ln=1)
    pdf.set_font("Helvetica", "", 10)
    for label, val, bold in resumo_linhas:
        pdf.set_font("Helvetica", "B" if bold else "", 10)
        pdf.cell(130, 6, clean(label), border="B")
        pdf.cell(60, 6, clean(fmt_r(val)), border="B", align="R", ln=1)
    pdf.ln(4)

    pdf.set_font("Helvetica", "B", 13)
    pdf.cell(0, 8, clean("Resultado Trimestral - Valor a Receber (Maldivas)"), ln=1)
    pdf.set_font("Helvetica", "", 10)
    for q in quarter_series.index:
        pdf.cell(47, 7, clean(f"{q}: {fmt_r(quarter_series[q])}"))
    pdf.ln(10)

    pie_buf1 = mpl_pie(["Direta", "Portal MAAS"],
                        [combined["ReceitaDireta"].sum(), combined["ReceitaPortal"].sum()],
                        ["#667eea", "#0F9ED5"], "Composicao Receita")
    pie_buf2 = mpl_pie(["Partner", "Maldivas"], [socio_partner_total, socio_maldivas_total],
                        ["#667eea", "#0F9ED5"], "Distribuicao Resultado")
    if pdf.get_y() > 200:
        pdf.add_page()
    pdf.image(pie_buf1, x=10, w=90)
    pdf.image(pie_buf2, x=105, y=pdf.get_y() - 65, w=90)

    pdf.set_y(-15)
    pdf.set_font("Helvetica", "I", 8)
    pdf.set_text_color(120, 120, 120)
    pdf.cell(0, 10, clean("Documento gerado automaticamente pelo Dashboard Financeiro Premium"), align="C")

    return bytes(pdf.output(dest="S"))

st.markdown('<div class="section-title">📤 Exportar</div>', unsafe_allow_html=True)
pdf_bytes = build_pdf()
st.download_button("⬇️ Baixar Dashboard em PDF", data=pdf_bytes,
                    file_name=f"dashboard_financeiro_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                    mime="application/pdf")
