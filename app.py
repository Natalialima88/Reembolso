# -*- coding: utf-8 -*-
"""
Painel Interativo - Acompanhamento de Reembolso (ASR) 2026
============================================================

Aplicativo Streamlit que lê o arquivo
"Acompanhamento_Reembolso_Pag_2026_CORRIGIDO.xlsx" (o mesmo arquivo da
aba "Indicadores") e oferece filtros interativos por:
  Canal, Empresa, Código da Empresa, Cód. do Beneficiário, Motivo,
  CPF/CNPJ Dentista, Nome do Prestador de Serviço, Cidade do Prestador,
  UF do Prestador, Auditor, Prazo Contratual Atendido (Sim/Não),
  Segmento (Santander / Itaú / Premium Free / Outros) e Mês.

E recalcula, em tempo real, TODOS os indicadores da aba "Indicadores":
  - Quantidade de ASR paga
  - Valor de reembolso solicitado / pago (R$)
  - Quantidade de beneficiários com reembolsos pagos
  - Santander / Itaú / Premium Free (R$ e quantidade)
  - ASRs dentro / fora do prazo contratual (qtd e %)
  - Solicitação de cadastro Astrein
  - Reembolso por motivo: Insuficiência de Rede / Cobrança Indevida /
    Deslocamento (R$ e quantidade)

Como rodar
----------
    pip install -r requirements.txt
    streamlit run app.py

Por padrão o app procura o arquivo
"Acompanhamento_Reembolso_Pag_2026_CORRIGIDO.xlsx" na mesma pasta deste
script. Se não encontrar (ou se você quiser usar outra base/mês mais
recente), ele mostra um campo para você enviar o arquivo .xlsx pela
própria tela do navegador - não precisa mexer no código.
"""

import io
import os
import unicodedata
from typing import Optional

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

# ----------------------------------------------------------------------
# Configuração geral
# ----------------------------------------------------------------------
st.set_page_config(
    page_title="Painel ASR - Acompanhamento de Reembolso",
    page_icon="📊",
    layout="wide",
)

DEFAULT_FILE = os.path.join(os.path.dirname(__file__), "Acompanhamento_Reembolso_Pag_2026_CORRIGIDO.xlsx")

MONTH_SHEETS = [
    "Janeiro 2026", "Fevereiro 2026", "Março 2026", "Abril 2026 ", "Maio 2026",
    "Junho 2026", "Julho 2026", "Agosto 2026", "Setembro 2026",
]
MONTH_LABEL = {
    "Janeiro 2026": "Jan", "Fevereiro 2026": "Fev", "Março 2026": "Mar",
    "Abril 2026 ": "Abr", "Maio 2026": "Mai", "Junho 2026": "Jun",
    "Julho 2026": "Jul", "Agosto 2026": "Ago", "Setembro 2026": "Set",
}
MONTH_ORDER = list(MONTH_LABEL.values())

# Nomes de coluna padronizados (posição -> nome final), porque algumas
# abas mensais têm o cabeçalho da coluna Y digitado errado ("PRAZO
# CONTRATUAL" em vez de "PRAZO ATENDIDO" - ver aba "Legenda e
# Metodologia" do arquivo original). Usamos a POSIÇÃO da coluna, não o
# texto do cabeçalho, para não depender dessa inconsistência.
COLS_BY_POSITION = {
    0: "CANAL",
    1: "N_SOLICITACAO",
    2: "COD_EMPRESA",
    3: "EMPRESA",
    4: "CPF_TITULAR",
    5: "COD_BENEFICIARIO",
    6: "VALOR_SOLICITADO",
    7: "VALOR_PAGO",
    8: "MOTIVO",
    9: "CPF_CNPJ_DENTISTA",
    10: "NOME_PRESTADOR",
    11: "CIDADE_PRESTADOR",
    12: "UF_PRESTADOR",
    13: "AUDITOR",
    14: "ASTREIN",
    22: "DT_PRAZO_CONTRATUAL",
    23: "DATA_PAGAMENTO_SAP",
    24: "PRAZO_ATENDIDO",  # cache da fórmula do Excel - não confiável, recalculado abaixo
}


def _norm(s) -> str:
    """Normaliza texto (maiúsculas, sem espaço nas pontas, sem acento) para comparações."""
    if s is None:
        return ""
    s = str(s).strip().upper()
    s = "".join(c for c in unicodedata.normalize("NFD", s) if unicodedata.category(c) != "Mn")
    return s


@st.cache_data(show_spinner="Lendo a planilha e calculando os indicadores...")
def load_data(file_bytes: bytes):
    xls = pd.ExcelFile(io.BytesIO(file_bytes))

    # ---- listas de códigos Santander / Itaú (mesma fonte da aba Indicadores) ----
    santander_codes, itau_codes = set(), set()
    if "Códigos Santander (ref)" in xls.sheet_names:
        df_s = pd.read_excel(xls, "Códigos Santander (ref)", header=0)
        santander_codes = {_norm(v) for v in df_s.iloc[:, 0].dropna()}
    if "Códigos Itaú (ref)" in xls.sheet_names:
        df_i = pd.read_excel(xls, "Códigos Itaú (ref)", header=0)
        itau_codes = {_norm(v) for v in df_i.iloc[:, 0].dropna()}

    frames = []
    for sheet in MONTH_SHEETS:
        if sheet not in xls.sheet_names:
            continue
        raw = pd.read_excel(xls, sheet, header=0)
        if raw.empty:
            continue
        n = raw.shape[1]
        rename = {raw.columns[pos]: name for pos, name in COLS_BY_POSITION.items() if pos < n}
        df = raw.rename(columns=rename)
        keep = [c for c in COLS_BY_POSITION.values() if c in df.columns]
        df = df[keep].copy()

        # só linhas de fato preenchidas (ignora eventual linha de
        # subtotal solta, sem número de solicitação)
        df = df[df["N_SOLICITACAO"].notna()].copy()

        df["MES"] = MONTH_LABEL[sheet]
        frames.append(df)

    if not frames:
        return pd.DataFrame()

    data = pd.concat(frames, ignore_index=True)
    data["MES"] = pd.Categorical(data["MES"], categories=MONTH_ORDER, ordered=True)

    for col in ("VALOR_SOLICITADO", "VALOR_PAGO"):
        data[col] = pd.to_numeric(data[col], errors="coerce").fillna(0.0)

    # "Prazo atendido" é recalculado aqui a partir das datas (mesma regra da
    # coluna Y de cada aba mensal: =SE(E(X<>"";W<>"");SE(X<=W;"SIM";"NÃO");""))
    # em vez de ler o valor em cache da fórmula do Excel, que pode estar
    # vazio se o arquivo ainda não foi recalculado/salvo no Excel.
    prazo_dt = pd.to_datetime(data.get("DT_PRAZO_CONTRATUAL"), errors="coerce")
    pagto_dt = pd.to_datetime(data.get("DATA_PAGAMENTO_SAP"), errors="coerce")
    ambas_preenchidas = prazo_dt.notna() & pagto_dt.notna()
    data["PRAZO_ATENDIDO"] = pd.NA
    dentro = (pagto_dt[ambas_preenchidas] <= prazo_dt[ambas_preenchidas]).map({True: "SIM", False: "NÃO"})
    data.loc[ambas_preenchidas, "PRAZO_ATENDIDO"] = dentro
    data["PRAZO_ATENDIDO_NORM"] = data["PRAZO_ATENDIDO"].map(_norm)
    data["MOTIVO_NORM"] = data.get("MOTIVO", pd.Series(index=data.index)).map(_norm)

    # ---- segmento: mesma regra usada na coluna auxiliar do Excel ----
    def segmento(row):
        cod = _norm(row.get("COD_EMPRESA"))
        if cod in santander_codes:
            return "Santander"
        if cod in itau_codes:
            return "Itaú"
        empresa = _norm(row.get("EMPRESA")).replace(" ", "")
        motivo = _norm(row.get("MOTIVO"))
        if motivo == "CONTRATUAL" and "INDIVIDUAL" in empresa:
            return "Premium Free"
        return "Outros"

    data["SEGMENTO"] = data.apply(segmento, axis=1)

    return data


def multiselect_filter(df: pd.DataFrame, col: str, label: str) -> pd.DataFrame:
    if col not in df.columns:
        return df
    options = sorted([v for v in df[col].dropna().unique().tolist()], key=lambda x: str(x))
    selected = st.sidebar.multiselect(label, options, default=[])
    if selected:
        return df[df[col].isin(selected)]
    return df


def brl(v: float) -> str:
    return "R$ " + f"{v:,.2f}".replace(",", "§").replace(".", ",").replace("§", ".")


# ----------------------------------------------------------------------
# Carregar arquivo
# ----------------------------------------------------------------------
st.title("📊 Painel Interativo — Acompanhamento de Reembolso (ASR)")

file_bytes: Optional[bytes] = None
if os.path.exists(DEFAULT_FILE):
    with open(DEFAULT_FILE, "rb") as f:
        file_bytes = f.read()
    st.caption(f"Base carregada automaticamente: `{os.path.basename(DEFAULT_FILE)}`. "
               f"Você pode enviar uma versão mais nova abaixo, se quiser.")

uploaded = st.file_uploader("Enviar/atualizar a planilha (.xlsx)", type=["xlsx"])
if uploaded is not None:
    file_bytes = uploaded.read()

if file_bytes is None:
    st.warning("Envie o arquivo `Acompanhamento_Reembolso_Pag_2026_CORRIGIDO.xlsx` para começar.")
    st.stop()

data_full = load_data(file_bytes)
if data_full.empty:
    st.error("Não encontrei nenhuma aba mensal reconhecível nesse arquivo.")
    st.stop()

# ----------------------------------------------------------------------
# Filtros (barra lateral)
# ----------------------------------------------------------------------
st.sidebar.header("Filtros")

meses_sel = st.sidebar.multiselect("Mês", MONTH_ORDER, default=[])
df = data_full if not meses_sel else data_full[data_full["MES"].isin(meses_sel)]

df = multiselect_filter(df, "CANAL", "Canal")
df = multiselect_filter(df, "EMPRESA", "Empresa")
df = multiselect_filter(df, "COD_EMPRESA", "Código da empresa")
df = multiselect_filter(df, "COD_BENEFICIARIO", "Cód. do beneficiário")
df = multiselect_filter(df, "MOTIVO", "Motivo")
df = multiselect_filter(df, "CPF_CNPJ_DENTISTA", "CPF/CNPJ dentista")
df = multiselect_filter(df, "NOME_PRESTADOR", "Nome do prestador de serviço")
df = multiselect_filter(df, "CIDADE_PRESTADOR", "Cidade do prestador")
df = multiselect_filter(df, "UF_PRESTADOR", "UF do prestador")
df = multiselect_filter(df, "AUDITOR", "Auditor")
df = multiselect_filter(df, "SEGMENTO", "Segmento (Santander/Itaú/Premium Free)")

prazo_sel = st.sidebar.radio("Prazo contratual atendido", ["Todos", "Sim", "Não"], index=0)
if prazo_sel != "Todos":
    df = df[df["PRAZO_ATENDIDO_NORM"] == _norm(prazo_sel)]

st.sidebar.markdown("---")
st.sidebar.caption(f"{len(df):,} ASR(s) selecionada(s) de {len(data_full):,} no total".replace(",", "."))

if df.empty:
    st.info("Nenhuma ASR encontrada para os filtros selecionados.")
    st.stop()

# ----------------------------------------------------------------------
# Indicadores (mesma lógica da aba "Indicadores" do Excel)
# ----------------------------------------------------------------------
qtd_asr = len(df)
valor_solicitado = df["VALOR_SOLICITADO"].sum()
valor_pago = df["VALOR_PAGO"].sum()
beneficiarios = df["COD_BENEFICIARIO"].nunique()
dentro_prazo = (df["PRAZO_ATENDIDO_NORM"] == "SIM").sum()
fora_prazo = (df["PRAZO_ATENDIDO_NORM"] == "NAO").sum()
pct_dentro = dentro_prazo / qtd_asr if qtd_asr else 0
ticket_medio = valor_pago / qtd_asr if qtd_asr else 0

if "ASTREIN" in df.columns:
    astrein_novos = df["ASTREIN"].apply(
        lambda v: pd.notna(v) and _norm(v) not in ("", "JA POSSUI")
    ).sum()
else:
    astrein_novos = 0

seg_valor = df.groupby("SEGMENTO")["VALOR_PAGO"].sum().reindex(
    ["Santander", "Itaú", "Premium Free", "Outros"], fill_value=0.0)
seg_qtd = df.groupby("SEGMENTO").size().reindex(
    ["Santander", "Itaú", "Premium Free", "Outros"], fill_value=0)

motivos_alvo = {
    "Insuficiência de Rede": "INSUFICIENCIA DE REDE",
    "Cobrança Indevida": "COBRANCA INDEVIDA",
    "Deslocamento": "DESLOCAMENTO",
}
motivo_valor, motivo_qtd = {}, {}
for label, norm_key in motivos_alvo.items():
    sub = df[df["MOTIVO_NORM"] == norm_key]
    motivo_valor[label] = sub["VALOR_PAGO"].sum()
    motivo_qtd[label] = len(sub)

# ----------------------------------------------------------------------
# KPIs
# ----------------------------------------------------------------------
st.subheader("Indicadores gerais")
c1, c2, c3, c4, c5, c6 = st.columns(6)
c1.metric("ASR paga (qtd)", f"{qtd_asr:,}".replace(",", "."))
c2.metric("Valor solicitado", brl(valor_solicitado))
c3.metric("Valor pago", brl(valor_pago))
c4.metric("Beneficiários únicos", f"{beneficiarios:,}".replace(",", "."))
c5.metric("% dentro do prazo", f"{pct_dentro:.1%}")
c6.metric("Ticket médio pago", brl(ticket_medio))

st.subheader("Prazo contratual e Astrein")
c1, c2, c3 = st.columns(3)
c1.metric("ASRs dentro do prazo", f"{dentro_prazo:,}".replace(",", "."))
c2.metric("ASRs fora do prazo", f"{fora_prazo:,}".replace(",", "."))
c3.metric("Solicitações de cadastro Astrein", f"{astrein_novos:,}".replace(",", "."))

st.subheader("Quantidade de ASR paga por segmento (Santander / Itaú / Premium Free)")
c1, c2, c3, c4 = st.columns(4)
for col, seg in zip((c1, c2, c3, c4), ["Santander", "Itaú", "Premium Free", "Outros"]):
    col.metric(f"{seg} — valor pago", brl(seg_valor.get(seg, 0.0)))
    col.metric(f"{seg} — qtd ASR", f"{int(seg_qtd.get(seg, 0)):,}".replace(",", "."))

st.subheader("Reembolso por motivo (Insuficiência de Rede / Cobrança Indevida / Deslocamento)")
c1, c2, c3 = st.columns(3)
for col, label in zip((c1, c2, c3), motivos_alvo.keys()):
    col.metric(f"{label} — R$", brl(motivo_valor[label]))
    col.metric(f"{label} — qtd", f"{motivo_qtd[label]:,}".replace(",", "."))

st.markdown("---")

# ----------------------------------------------------------------------
# Gráficos
# ----------------------------------------------------------------------
g1, g2 = st.columns(2)

with g1:
    fig = px.bar(
        x=list(motivo_valor.values()), y=list(motivo_valor.keys()), orientation="h",
        labels={"x": "Valor pago (R$)", "y": ""},
        title="Reembolso por motivo (R$)",
        text=[brl(v) for v in motivo_valor.values()],
    )
    fig.update_traces(marker_color="#2E75B6", textposition="outside")
    st.plotly_chart(fig, use_container_width=True)

with g2:
    seg_plot = seg_valor[seg_valor > 0]
    fig = px.pie(
        values=seg_plot.values, names=seg_plot.index, hole=0.55,
        title="Valor pago por segmento",
        color_discrete_sequence=["#2E75B6", "#ED7D31", "#A5A5A5", "#FFC000"],
    )
    fig.update_traces(textinfo="percent+label")
    st.plotly_chart(fig, use_container_width=True)

g3, g4 = st.columns(2)

with g3:
    trend = df.groupby("MES", observed=True).size().reindex(MONTH_ORDER, fill_value=0)
    fig = go.Figure(go.Scatter(x=trend.index, y=trend.values, mode="lines+markers+text",
                                text=trend.values, textposition="top center",
                                line=dict(color="#2E75B6", width=3)))
    fig.update_layout(title="ASR paga por mês", yaxis_title="Qtd ASR")
    st.plotly_chart(fig, use_container_width=True)

with g4:
    prazo_mes = df.assign(
        dentro=(df["PRAZO_ATENDIDO_NORM"] == "SIM").astype(int),
        fora=(df["PRAZO_ATENDIDO_NORM"] == "NAO").astype(int),
    ).groupby("MES", observed=True)[["dentro", "fora"]].sum().reindex(MONTH_ORDER, fill_value=0)
    fig = go.Figure()
    fig.add_bar(x=prazo_mes.index, y=prazo_mes["dentro"], name="Dentro do prazo", marker_color="#2E75B6")
    fig.add_bar(x=prazo_mes.index, y=prazo_mes["fora"], name="Fora do prazo", marker_color="#ED7D31")
    fig.update_layout(barmode="stack", title="ASR dentro x fora do prazo, por mês")
    st.plotly_chart(fig, use_container_width=True)

# ----------------------------------------------------------------------
# Tabela detalhada + download
# ----------------------------------------------------------------------
st.markdown("---")
st.subheader(f"ASRs selecionadas ({qtd_asr:,})".replace(",", "."))

cols_show = [c for c in [
    "MES", "CANAL", "N_SOLICITACAO", "COD_EMPRESA", "EMPRESA", "COD_BENEFICIARIO",
    "VALOR_SOLICITADO", "VALOR_PAGO", "MOTIVO", "CPF_CNPJ_DENTISTA", "NOME_PRESTADOR",
    "CIDADE_PRESTADOR", "UF_PRESTADOR", "AUDITOR", "PRAZO_ATENDIDO", "SEGMENTO",
] if c in df.columns]

st.dataframe(df[cols_show], use_container_width=True, hide_index=True)

csv = df[cols_show].to_csv(index=False, sep=";", decimal=",").encode("utf-8-sig")
st.download_button("⬇️ Baixar seleção (CSV)", data=csv, file_name="asr_filtrado.csv", mime="text/csv")
