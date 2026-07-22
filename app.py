# app.py
import streamlit as st
import pandas as pd
import openpyxl
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="Dashboard Financeiro DRE 2026", layout="wide")

MESES_ORDEM = ["Jan","Fev","Mar","Abr","Mai","Jun","Jul","Ago","Set","Out","Nov","Dez"]

# ---------------------------------------------------------------------------
# Leitura e parsing da aba "DRE 2026"
# ---------------------------------------------------------------------------
def _norm(v):
    return str(v).strip().upper() if v is not None else ""

def find_row(ws, label, start=1, end=None, contains=True):
    """Acha a primeira linha cuja coluna B contenha (ou seja igual a) o label, dentro do intervalo."""
    end = end or ws.max_row
    target = _norm(label)
    for r in range(start, end + 1):
        cell = _norm(ws.cell(row=r, column=2).value)
        if (target in cell) if contains else (cell == target):
            return r
    return None

def find_header_row(ws):
    """Acha a linha com os rótulos dos meses (Jan, Fev, ...) + YTD."""
    for r in range(1, ws.max_row + 1):
        vals = [_norm(ws.cell(row=r, column=c).value) for c in range(3, ws.max_column + 1)]
        hits = sum(1 for v in vals if v[:3] in [m.upper()[:3] for m in MESES_ORDEM])
        if hits >= 6:
            return r
    return None

def row_values(ws, row, col_start, col_end):
    return [ws.cell(row=row, column=c).value or 0 for c in range(col_start, col_end + 1)]

def parse_dre_2026(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    if "DRE 2026" not in wb.sheetnames:
        return None
    ws = wb["DRE 2026"]

    header_row = find_header_row(ws)
    if header_row is None:
        return None

    # descobre colunas de meses (C..N) e a coluna YTD (O)
    month_cols = []
    ytd_col = None
    for c in range(3, ws.max_column + 1):
        v = _norm(ws.cell(row=header_row, column=c).value)
        if v == "YTD":
            ytd_col = c
        elif v[:3] in [m.upper()[:3] for m in MESES_ORDEM]:
            month_cols.append(c)
    col_start, col_end = min(month_cols), max(month_cols)

    total_row = find_row(ws, "RECEITA BRUTA TOTAL")
    direta_sec = find_row(ws, "PRODUÇÃO DIRETA")
    portal_sec = find_row(ws, "PORTAL MAAS")
    resultado_row = find_row(ws, "RESULTADO OPERACIONAL", start=portal_sec, contains=True)
    # garante que não pega "RESULTADO OPERACIONAL -DISTRIBUIÇÃO"
    while resultado_row and "DISTRIBUI" in _norm(ws.cell(row=resultado_row, column=2).value):
        resultado_row = find_row(ws, "RESULTADO OPERACIONAL", start=resultado_row + 1)

    def sub(label, start, end):
        return find_row(ws, label, start=start, end=end)

    d_receita = sub("RECEITA BRUTA", direta_sec, portal_sec)
    d_impostos = sub("IMPOSTOS DIRETOS", direta_sec, portal_sec)
    d_custo = sub("CUSTO OPERACIONAL", direta_sec, portal_sec)
    d_despesas = sub("DESPESAS", direta_sec, portal_sec)
    d_ebitda = sub("EBITDA SOCIETÁRIO", direta_sec, portal_sec)

    p_receita = sub("RECEITA BRUTA", portal_sec, resultado_row)
    p_impostos = sub("IMPOSTOS DIRETOS", portal_sec, resultado_row)
    p_custo = sub("CUSTO OPERACIONAL", portal_sec, resultado_row)
    p_despesas = sub("DESPESAS", portal_sec, resultado_row)
    p_ebitda = sub("EBITDA SOCIETÁRIO", portal_sec, resultado_row)

    dist_start = resultado_row
    socio_partner = find_row(ws, "Sócio Partner", start=dist_start)
    socio_maldivas = find_row(ws, "Sócio Maldivas", start=dist_start)

    def series(row):
        return row_values(ws, row, col_start, col_end) if row else [0] * len(month_cols)

    def ytd(row):
        return (ws.cell(row=row, column=ytd_col).value or 0) if row else 0

    meses_labels = [_norm(ws.cell(row=header_row, column=c).value).title() for c in month_cols]
    # corrige acentuação (Jan, Fev, Mar, Abr, Mai, Jun, Jul, Ago, Set, Out, Nov, Dez já batem)

    df = pd.DataFrame({
        "Mês": meses_labels,
        "Receita Direta": series(d_receita),
        "Receita Portal MAAS": series(p_receita),
        "Impostos Diretos": [a + b for a, b in zip(series(d_impostos), series(p_impostos))],
        "Custo Operacional": [a + b for a, b in zip(series(d_custo), series(p_custo))],
        "Despesas Totais": [a + b for a, b in zip(series(d_despesas), series(p_despesas))],
        "EBITDA Direta": series(d_ebitda),
        "EBITDA Portal MAAS": series(p_ebitda),
        "Resultado Operacional": series(resultado_row),
    })
    df["Receita Bruta Total"] = df["Receita Direta"] + df["Receita Portal MAAS"]
    df["Margem EBITDA %"] = df["Resultado Operacional"] / df["Receita Bruta Total"].replace(0, pd.NA)

    ytd_data = {
        "receita_total": ytd(total_row),
        "receita_direta": ytd(d_receita),
        "receita_portal": ytd(p_receita),
        "impostos": ytd(d_impostos) + ytd(p_impostos),
        "margem_contribuicao": ytd(sub("MARGEM DE CONTRIBUIÇÃO", direta_sec, portal_sec)) +
                                 ytd(sub("MARGEM DE CONTRIBUIÇÃO", portal_sec, resultado_row)),
        "despesas": ytd(d_despesas) + ytd(p_despesas),
        "ebitda": ytd(resultado_row),
        "socio_partner": ytd(socio_partner),
        "socio_maldivas": ytd(socio_maldivas),
    }
    ytd_data["margem_ebitda_pct"] = (ytd_data["ebitda"] / ytd_data["receita_total"]) if ytd_data["receita_total"] else 0

    return {"df": df, "ytd": ytd_data}

# ---------------------------------------------------------------------------
# UI
# ---------------------------------------------------------------------------
st.title("📊 Dashboard Financeiro - Partners")
st.caption("Demonstrativo de Resultado do Exercício 2026 · Produção Direta + Portal MAAS · Valores em R$")

uploaded_files = st.file_uploader(
    "Envie uma ou mais planilhas (.xlsx) contendo a aba 'DRE 2026'",
    type=["xlsx"], accept_multiple_files=True
)

if not uploaded_files:
    st.info("Aguardando upload de planilha(s)...")
    st.stop()

file_names = [f.name for f in uploaded_files]
if len(file_names) > 1:
    selected_name = st.selectbox("Selecione o arquivo para visualizar:", file_names)
else:
    selected_name = file_names[0]

selected_file = next(f for f in uploaded_files if f.name == selected_name)
data = parse_dre_2026(selected_file)

if data is None:
    st.error(f"A aba 'DRE 2026' não foi encontrada (ou não pôde ser lida) em '{selected_name}'.")
    st.stop()

df = data["df"]
ytd = data["ytd"]

meses_disponiveis = df["Mês"].tolist() + ["YTD"]
mes_ref = st.selectbox("Mês de Referência:", meses_disponiveis, index=len(meses_disponiveis) - 1)

if mes_ref == "YTD":
    receita_total = ytd["receita_total"]
    impostos = ytd["impostos"]
    margem_contrib = ytd["margem_contribuicao"]
    despesas = ytd["despesas"]
    ebitda = ytd["ebitda"]
else:
    row = df[df["Mês"] == mes_ref].iloc[0]
    receita_total = row["Receita Bruta Total"]
    impostos = row["Impostos Diretos"]
    margem_contrib = row["Receita Bruta Total"] - row["Impostos Diretos"] - row["Custo Operacional"]
    despesas = row["Despesas Totais"]
    ebitda = row["Resultado Operacional"]

margem_ebitda_pct = (ebitda / receita_total) if receita_total else 0

def fmt_r(v):
    return f"R$ {v:,.0f}".replace(",", ".")

col1, col2, col3, col4, col5, col6 = st.columns(6)
col1.metric("Receita Bruta Total", fmt_r(receita_total))
col2.metric("Impostos Diretos", fmt_r(impostos))
col3.metric("Margem de Contribuição", fmt_r(margem_contrib))
col4.metric("Despesas Totais", fmt_r(despesas))
col5.metric("EBITDA Societário", fmt_r(ebitda))
col6.metric("Margem EBITDA %", f"{margem_ebitda_pct:.1%}")

st.divider()

# --- Evolução Mensal: Receita Direta vs Portal MAAS ---
st.subheader("Evolução Mensal — Receita Bruta (Direta vs Portal MAAS)")
fig1 = go.Figure()
fig1.add_bar(x=df["Mês"], y=df["Receita Direta"], name="Receita Direta")
fig1.add_bar(x=df["Mês"], y=df["Receita Portal MAAS"], name="Receita Portal MAAS")
fig1.update_layout(barmode="group", yaxis_title="R$")
st.plotly_chart(fig1, use_container_width=True)

# --- EBITDA Mensal: Direta vs Portal ---
st.subheader("Evolução Mensal — EBITDA Societário (Direta vs Portal MAAS)")
fig2 = go.Figure()
fig2.add_bar(x=df["Mês"], y=df["EBITDA Direta"], name="EBITDA Direta")
fig2.add_bar(x=df["Mês"], y=df["EBITDA Portal MAAS"], name="EBITDA Portal MAAS")
fig2.update_layout(barmode="group", yaxis_title="R$")
st.plotly_chart(fig2, use_container_width=True)

# --- Tendência do Resultado Operacional ---
st.subheader("Tendência do Resultado Operacional")
fig3 = px.line(df, x="Mês", y="Resultado Operacional", markers=True)
fig3.update_layout(yaxis_title="R$")
st.plotly_chart(fig3, use_container_width=True)

st.divider()

# --- Composição YTD ---
st.subheader("Composição (Acumulado YTD)")
cpie1, cpie2 = st.columns(2)
with cpie1:
    fig4 = px.pie(
        names=["Direta", "Portal MAAS"],
        values=[ytd["receita_direta"], ytd["receita_portal"]],
        title="Composição da Receita Bruta (YTD)"
    )
    st.plotly_chart(fig4, use_container_width=True)
with cpie2:
    fig5 = px.pie(
        names=["Sócio Partner", "Sócio Maldivas"],
        values=[ytd["socio_partner"], ytd["socio_maldivas"]],
        title="Distribuição do Resultado Operacional (YTD)"
    )
    st.plotly_chart(fig5, use_container_width=True)

st.divider()

# --- Resumo Mensal Detalhado ---
st.subheader("Resumo Mensal Detalhado")
tabela = df.set_index("Mês")[[
    "Receita Bruta Total", "Impostos Diretos", "Custo Operacional",
    "Despesas Totais", "Resultado Operacional", "Margem EBITDA %"
]].T
tabela["YTD"] = [
    ytd["receita_total"], ytd["impostos"], "",
    ytd["despesas"], ytd["ebitda"], ytd["margem_ebitda_pct"]
]
st.dataframe(
    tabela.style.format(lambda v: f"{v:,.0f}".replace(",", ".") if isinstance(v, (int, float)) and abs(v) > 1
                         else (f"{v:.1%}" if isinstance(v, float) else v)),
    use_container_width=True
)
