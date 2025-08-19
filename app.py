# -*- coding: utf-8 -*-
# app.py — Comparativo de Gastos por Secretaria e Administração
# Requisitos: streamlit, pandas, openpyxl, plotly

import io
import os
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Comparativo de Gastos • Secretarias x Administração", layout="wide")

# =========================
# Utilidades
# =========================
MONTH_ORDER = ["JANEIRO","FEVEREIRO","MARÇO","ABRIL","MAIO","JUNHO","JULHO","AGOSTO","SETEMBRO","OUTUBRO","NOVEMBRO","DEZEMBRO"]
MONTH_TO_NUM = {m: i+1 for i, m in enumerate(MONTH_ORDER)}

def normalize_text(s: str) -> str:
    if not isinstance(s, str):
        s = "" if pd.isna(s) else str(s)
    return s.strip().upper()

def format_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

def order_months(df, col="MES"):
    df[col] = df[col].apply(normalize_text)
    df["MES_NUM"] = df[col].map(MONTH_TO_NUM)
    return df.dropna(subset=["MES_NUM"]).sort_values(["MES_NUM"])

def find_local_file(candidates):
    for name in candidates:
        if os.path.exists(name):
            return name
        alt = os.path.join("data", name)
        if os.path.exists(alt):
            return alt
    return None

# =========================
# Leitura de dados
# =========================
@st.cache_data(show_spinner=False)
def load_all_secretarias_from_planilha1(xlsx_file) -> pd.DataFrame:
    df_raw = pd.read_excel(xlsx_file, sheet_name="Planilha1", header=None)
    header_idx_candidates = df_raw.index[
        df_raw.iloc[:, 0].astype(str).str.upper().str.contains("MÊS", na=False)
    ].tolist()
    header_idx = header_idx_candidates[0] if header_idx_candidates else 1

    df = df_raw.iloc[header_idx+1:, 0:3].copy()
    df.columns = ["MES", "SECRETARIA", "VALOR"]

    df["MES"] = df["MES"].apply(normalize_text)
    df["SECRETARIA"] = df["SECRETARIA"].apply(normalize_text)
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce").fillna(0.0)

    df = df.dropna(subset=["MES", "SECRETARIA"], how="any")
    df = order_months(df, "MES")
    return df

@st.cache_data(show_spinner=False)
def load_admin_departments(xlsx_file) -> pd.DataFrame:
    df_raw = pd.read_excel(xlsx_file, sheet_name="GERAL", header=None)

    head_idx_candidates = df_raw.index[
        df_raw.iloc[:, 1].astype(str).str.upper().str.contains("BENEFICIARIO", na=False)
    ].tolist()
    if not head_idx_candidates:
        flat = df_raw.astype(str).apply(lambda col: col.str.upper().str.contains("BENEFICIARIO", na=False))
        if flat.any().any():
            head_idx_candidates = [flat.any(axis=1).idxmax()]
        else:
            return pd.DataFrame(columns=["BENEFICIARIO", "MES", "VALOR"])

    head_idx = head_idx_candidates[0]
    header_row = df_raw.iloc[head_idx].tolist()
    df = df_raw.iloc[head_idx+1:, :len(header_row)].copy()
    df.columns = [normalize_text(c) for c in header_row]

    cols = [c for c in df.columns if c and c != "NAN"]
    df = df[cols].copy()

    ben_col_candidates = [c for c in df.columns if "BENEFICIARIO" in c]
    if not ben_col_candidates:
        ben_col_candidates = [df.columns[0]]
    ben_col = ben_col_candidates[0]

    month_cols = [c for c in df.columns if c != ben_col]
    mapping_cols = {}
    month_cols_norm = []
    for c in month_cols:
        c_norm = normalize_text(c)
        if c_norm in MONTH_ORDER:
            month_cols_norm.append(c_norm)
            mapping_cols[c] = c_norm

    if not month_cols_norm:
        return pd.DataFrame(columns=["BENEFICIARIO", "MES", "VALOR"])

    df = df.rename(columns=mapping_cols)
    keep_cols = [ben_col] + month_cols_norm
    df = df[keep_cols].copy().rename(columns={ben_col: "BENEFICIARIO"})
    df["BENEFICIARIO"] = df["BENEFICIARIO"].apply(normalize_text)

    df_long = df.melt(id_vars=["BENEFICIARIO"], value_vars=month_cols_norm,
                      var_name="MES", value_name="VALOR")
    df_long["MES"] = df_long["MES"].apply(normalize_text)
    df_long["VALOR"] = pd.to_numeric(df_long["VALOR"], errors="coerce").fillna(0.0)
    df_long = df_long[df_long["BENEFICIARIO"].notna() & (df_long["BENEFICIARIO"] != "")]
    return order_months(df_long, "MES")

def df_download_excel(df: pd.DataFrame, filename: str) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="dados")
    return buffer.getvalue()

# =========================
# Sidebar: Upload / Fontes
# =========================
st.sidebar.title("⚙️ Fontes de Dados")
st.sidebar.caption("Envie as planilhas ou deixe os arquivos no repositório (raiz ou /data).")

default_all = find_local_file(["COMBUSTIVEL 2025.xlsx"])
default_admin = find_local_file(["Combustivel POR SECRETARIA.xlsx"])

uploaded_all = st.sidebar.file_uploader(
    "Planilha de TODAS as Secretarias (aba 'Planilha1')",
    type=["xlsx"], key="all_xlsx"
)
uploaded_admin = st.sidebar.file_uploader(
    "Planilha da Administração + Departamentos (aba 'GERAL')",
    type=["xlsx"], key="admin_xlsx"
)

all_source = uploaded_all if uploaded_all is not None else default_all
adm_source = uploaded_admin if uploaded_admin is not None else default_admin

# =========================
# Título
# =========================
st.title("📊 Comparativo de Gastos")
st.caption("Totais por Secretaria + Detalhamento da Secretaria de Administração e beneficiários/departamentos.")

# =========================
# Carrega dados
# =========================
c1, c2 = st.columns(2)
with c1:
    if all_source is None:
        st.error("Faltou a planilha de TODAS as Secretarias. Faça upload ou adicione **COMBUSTIVEL 2025.xlsx** (raiz ou /data).")
    else:
        st.success("Planilha de TODAS as Secretarias carregada.")
with c2:
    if adm_source is None:
        st.warning("Sem a planilha da Administração + Departamentos. O app funciona com a visão geral mesmo assim.")
    else:
        st.success("Planilha da Administração + Departamentos carregada.")

df_all = load_all_secretarias_from_planilha1(all_source) if all_source else pd.DataFrame(columns=["MES","SECRETARIA","VALOR","MES_NUM"])
df_adm = load_admin_departments(adm_source) if adm_source else pd.DataFrame(columns=["BENEFICIARIO","MES","VALOR","MES_NUM"])

# =========================
# Filtros
# =========================
st.subheader("🔎 Filtros")
colf1, colf2, colf3 = st.columns([1.2, 1.7, 1.1])

meses_disponiveis = sorted(
    list(set(df_all["MES"].dropna().tolist() + df_adm["MES"].dropna().tolist())),
    key=lambda m: MONTH_TO_NUM.get(m, 99)
)

with colf1:
    meses_sel = st.multiselect("Meses", options=meses_disponiveis, default=meses_disponiveis)

with colf2:
    secretarias_disponiveis = sorted(df_all["SECRETARIA"].dropna().unique().tolist())
    default_secs = [s for s in secretarias_disponiveis if s in ["ADMINISTRAÇÃO","OBRAS","SAÚDE"]][:3]
    secs_sel = st.multiselect("Secretarias (visão geral)", options=secretarias_disponiveis, default=default_secs)

with colf3:
    topn = st.number_input("Top N (ranking)", min_value=3, max_value=30, value=10, step=1)

df_all_f = df_all[df_all["MES"].isin(meses_sel)].copy() if meses_sel else df_all.copy()
df_adm_f = df_adm[df_adm["MES"].isin(meses_sel)].copy() if meses_sel else df_adm.copy()

# =========================
# Abas
# =========================
tabs = st.tabs(["📈 Visão Geral (Secretarias)", "⚖️ Administração x Demais", "🏛️ Administração por Beneficiário", "📄 Tabelas / Exportar"])

# Aba 1 — Visão Geral (Secretarias)
with tabs[0]:
    st.markdown("### 📈 Gastos por Secretaria (mensal)")
    if df_all_f.empty:
        st.info("Sem dados para exibir.")
    else:
        plot_df = (df_all_f[df_all_f["SECRETARIA"].isin(secs_sel)] if secs_sel else df_all_f)\
                    .groupby(["MES","SECRETARIA"], as_index=False)["VALOR"].sum()
        plot_df = order_months(plot_df, "MES")

        fig = px.line(plot_df, x="MES", y="VALOR", color="SECRETARIA", markers=True, title="Evolução Mensal por Secretaria")
        fig.update_layout(yaxis_title="Valor (R$)")
        st.plotly_chart(fig, use_container_width=True)

        st.markdown("#### Barras Empilhadas (participação por mês)")
        fig2 = px.bar(plot_df, x="MES", y="VALOR", color="SECRETARIA", title="Participação por Secretaria", barmode="stack")
        fig2.update_layout(yaxis_title="Valor (R$)")
        st.plotly_chart(fig2, use_container_width=True)

        # Ranking acumulado
        rank_df = plot_df.groupby("SECRETARIA", as_index=False)["VALOR"].sum().sort_values("VALOR", ascending=False).head(topn)
        rank_df["VALOR_fmt"] = rank_df["VALOR"].map(format_brl)
        st.markdown("#### 🔝 Ranking (soma nos meses filtrados)")
        st.dataframe(rank_df[["SECRETARIA","VALOR_fmt"]], use_container_width=True, hide_index=True)

        # >>> NOVO: Tabela completa por secretaria (colunas = meses) + TOTAL e TOTAL GERAL
        st.markdown("#### 🧮 Tabela — Valores por Secretaria (meses filtrados)")
        tbl = df_all_f.groupby(["SECRETARIA","MES"], as_index=False)["VALOR"].sum()
        tbl = order_months(tbl, "MES")
        pivot = tbl.pivot(index="SECRETARIA", columns="MES", values="VALOR").fillna(0.0)

        # ordena colunas por mês
        ordered_cols = sorted(pivot.columns, key=lambda m: MONTH_TO_NUM.get(m, 99))
        pivot = pivot[ordered_cols]

        # total por secretaria e total geral
        pivot["TOTAL"] = pivot.sum(axis=1)
        total_row = pd.DataFrame(pivot.sum(axis=0)).T
        total_row.index = ["TOTAL GERAL"]

        pivot_total = pd.concat([pivot, total_row], axis=0)

        display = pivot_total.applymap(lambda v: format_brl(v))
        st.dataframe(display, use_container_width=True)

        st.markdown(f"**Soma de todas as Secretarias (meses filtrados): {format_brl(pivot['TOTAL'].sum())}**")

# Aba 2 — Administração x Demais
with tabs[1]:
    st.markdown("### ⚖️ Administração vs Demais Secretarias")
    if df_all_f.empty:
        st.info("Sem dados para exibir.")
    else:
        total_mes = df_all_f.groupby("MES", as_index=False)["VALOR"].sum().rename(columns={"VALOR":"TOTAL"})
        admin_mes = df_all_f[df_all_f["SECRETARIA"] == "ADMINISTRAÇÃO"].groupby("MES", as_index=False)["VALOR"].sum().rename(columns={"VALOR":"ADMINISTRAÇÃO"})
        base = pd.merge(total_mes, admin_mes, on="MES", how="left")
        base["ADMINISTRAÇÃO"] = base["ADMINISTRAÇÃO"].fillna(0.0)
        base["DEMAIS"] = (base["TOTAL"] - base["ADMINISTRAÇÃO"]).clip(lower=0)
        base = order_months(base, "MES")

        base_long = base.melt(id_vars=["MES","MES_NUM"], value_vars=["ADMINISTRAÇÃO","DEMAIS"], var_name="GRUPO", value_name="VALOR")

        fig = px.bar(base_long, x="MES", y="VALOR", color="GRUPO", barmode="group", title="Administração x Demais (por mês)")
        fig.update_layout(yaxis_title="Valor (R$)")
        st.plotly_chart(fig, use_container_width=True)

        fig2 = px.line(base_long, x="MES", y="VALOR", color="GRUPO", markers=True, title="Evolução Mensal: Administração x Demais")
        fig2.update_layout(yaxis_title="Valor (R$)")
        st.plotly_chart(fig2, use_container_width=True)

        base_show = base[["MES","ADMINISTRAÇÃO","DEMAIS","TOTAL"]].copy()
        for c in ["ADMINISTRAÇÃO","DEMAIS","TOTAL"]:
            base_show[c] = base_show[c].map(format_brl)
        st.dataframe(base_show, use_container_width=True, hide_index=True)

# Aba 3 — Administração por Beneficiário
with tabs[2]:
    st.markdown("### 🏛️ Administração — Detalhe por Beneficiário")
    if df_adm_f.empty:
        st.info("Sem dados da Administração + Departamentos.")
    else:
        beneficiarios = sorted(df_adm_f["BENEFICIARIO"].unique().tolist())
        colb1, colb2 = st.columns([1.5, 2])
        with colb1:
            bens_sel = st.multiselect("Beneficiários/Departamentos", options=beneficiarios, default=beneficiarios[:10])
        with colb2:
            mes_unico = st.selectbox("Mês para ranking", options=meses_sel if meses_sel else sorted(df_adm_f["MES"].unique(), key=lambda m: MONTH_TO_NUM.get(m, 99)))

        df_plot = df_adm_f.copy()
        if bens_sel:
            df_plot = df_plot[df_plot["BENEFICIARIO"].isin(bens_sel)]

        df_plot = df_plot.groupby(["MES","BENEFICIARIO"], as_index=False)["VALOR"].sum()
        df_plot = order_months(df_plot, "MES")

        fig = px.line(df_plot, x="MES", y="VALOR", color="BENEFICIARIO", markers=True, title="Evolução Mensal por Beneficiário (Administração)")
        fig.update_layout(yaxis_title="Valor (R$)")
        st.plotly_chart(fig, use_container_width=True)

        # Top N no mês selecionado — >>> rótulos em REAIS <<<
        rank_mes = df_adm_f[df_adm_f["MES"] == mes_unico].groupby("BENEFICIARIO", as_index=False)["VALOR"].sum()
        rank_mes = rank_mes.sort_values("VALOR", ascending=False).head(topn)
        rank_mes["VALOR_fmt"] = rank_mes["VALOR"].map(format_brl)

        fig_bar = px.bar(
            rank_mes, x="BENEFICIARIO", y="VALOR",
            title=f"Top {topn} no mês {mes_unico}",
            text="VALOR_fmt"
        )
        fig_bar.update_traces(texttemplate="%{text}", textposition="outside", cliponaxis=False)
        fig_bar.update_layout(xaxis_title="", yaxis_title="Valor (R$)")
        st.plotly_chart(fig_bar, use_container_width=True)

        total_ben = df_adm_f.groupby("BENEFICIARIO", as_index=False)["VALOR"].sum().sort_values("VALOR", ascending=False).head(topn)
        total_ben["VALOR_fmt"] = total_ben["VALOR"].map(format_brl)
        st.markdown("#### 🔝 Ranking acumulado (meses filtrados)")
        st.dataframe(total_ben[["BENEFICIARIO","VALOR_fmt"]], use_container_width=True, hide_index=True)

# Aba 4 — Tabelas e Exportação
with tabs[3]:
    st.markdown("### 📄 Dados e Exportação")
    colx1, colx2 = st.columns(2)
    with colx1:
        st.markdown("**Secretarias (filtrado):**")
        show_all = df_all_f.copy()
        show_all["VALOR"] = show_all["VALOR"].round(2)
        st.dataframe(show_all, use_container_width=True, hide_index=True)
        if not show_all.empty:
            x1 = df_download_excel(show_all, "secretarias_filtrado.xlsx")
            st.download_button("⬇️ Baixar Excel (Secretarias)", data=x1, file_name="secretarias_filtrado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    with colx2:
        st.markdown("**Administração (filtrado):**")
        show_adm = df_adm_f.copy()
        show_adm["VALOR"] = show_adm["VALOR"].round(2)
        st.dataframe(show_adm, use_container_width=True, hide_index=True)
        if not show_adm.empty:
            x2 = df_download_excel(show_adm, "administracao_filtrado.xlsx")
            st.download_button("⬇️ Baixar Excel (Administração)", data=x2, file_name="administracao_filtrado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.caption("Coloque as planilhas na raiz ou em /data, ou faça upload na barra lateral.")
