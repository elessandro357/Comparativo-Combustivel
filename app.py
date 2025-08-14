# -*- coding: utf-8 -*-
# app.py — Comparativo de Gastos por Secretaria e Administração (Departamentos)
# Requisitos: streamlit, pandas, openpyxl, plotly

import io
import os
import unicodedata
import pandas as pd
import streamlit as st
import plotly.express as px

st.set_page_config(page_title="Comparativo de Gastos • Secretarias x Administração", layout="wide")

# =========================
# Utilidades
# =========================
MONTH_ORDER = [
    "JANEIRO", "FEVEREIRO", "MARÇO", "ABRIL", "MAIO", "JUNHO",
    "JULHO", "AGOSTO", "SETEMBRO", "OUTUBRO", "NOVEMBRO", "DEZEMBRO"
]
MONTH_TO_NUM = {m: i+1 for i, m in enumerate(MONTH_ORDER)}

def normalize_text(s: str) -> str:
    if not isinstance(s, str):
        s = "" if pd.isna(s) else str(s)
    s = s.strip()
    # Mantém acentos para meses/PT-BR, mas remove espaços extras e padroniza maiúsculas
    return s.upper()

def format_brl(v):
    try:
        return f"R$ {float(v):,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
    except Exception:
        return str(v)

def order_months(df, col="MES"):
    df[col] = df[col].apply(normalize_text)
    df["MES_NUM"] = df[col].map(MONTH_TO_NUM)
    df = df.dropna(subset=["MES_NUM"]).sort_values(["MES_NUM"])
    return df

# =========================
# Carregamento / Parsing
# =========================
@st.cache_data(show_spinner=False)
def load_all_secretarias_from_planilha1(xlsx_file) -> pd.DataFrame:
    """
    Lê a aba 'Planilha1' da planilha COMBUSTIVEL 2025.xlsx,
    que traz linhas no formato: [MÊS, SECRETARIA, VALOR].
    A planilha tem uma linha cabeçalho e pode ter linhas em branco no topo.
    """
    df_raw = pd.read_excel(xlsx_file, sheet_name="Planilha1", header=None)
    # Detecta a linha do cabeçalho (onde a primeira coluna é 'MÊS')
    header_idx_candidates = df_raw.index[
        df_raw.iloc[:, 0].astype(str).str.upper().str.contains("MÊS", na=False)
    ].tolist()
    header_idx = header_idx_candidates[0] if header_idx_candidates else 1

    # Pega as 3 primeiras colunas abaixo do cabeçalho
    df = df_raw.iloc[header_idx+1:, 0:3].copy()
    df.columns = ["MES", "SECRETARIA", "VALOR"]

    # Limpeza
    df["MES"] = df["MES"].apply(normalize_text)
    df["SECRETARIA"] = df["SECRETARIA"].apply(normalize_text)
    # VALOR para float
    df["VALOR"] = pd.to_numeric(df["VALOR"], errors="coerce").fillna(0.0)

    # Remove linhas vazias
    df = df.dropna(subset=["MES", "SECRETARIA"], how="any")
    # Ordena por mês
    df = order_months(df, "MES")
    return df

@st.cache_data(show_spinner=False)
def load_admin_departments(xlsx_file) -> pd.DataFrame:
    """
    Lê a planilha 'Combustivel POR SECRETARIA.xlsx' (aba 'GERAL').
    Estrutura observada:
      - Linha de cabeçalho onde a 2ª coluna = 'BENEFICIARIO'
      - Demais colunas são meses (ex.: FEVEREIRO, MARÇO, ABRIL, MAIO)
    Converte para formato longo: [BENEFICIARIO, MES, VALOR]
    """
    df_raw = pd.read_excel(xlsx_file, sheet_name="GERAL", header=None)

    # Localiza a linha do cabeçalho (onde a coluna 1 é 'BENEFICIARIO')
    head_idx_candidates = df_raw.index[
        df_raw.iloc[:, 1].astype(str).str.upper().str.contains("BENEFICIARIO", na=False)
    ].tolist()
    if not head_idx_candidates:
        # fallback: tenta encontrar qualquer linha com 'BENEFICIARIO'
        flat = df_raw.astype(str).apply(lambda col: col.str.upper().str.contains("BENEFICIARIO", na=False))
        if flat.any().any():
            head_idx_candidates = [flat.any(axis=1).idxmax()]
        else:
            # Se não achar cabeçalho, retorna vazio amigavelmente
            return pd.DataFrame(columns=["BENEFICIARIO", "MES", "VALOR"])

    head_idx = head_idx_candidates[0]
    header_row = df_raw.iloc[head_idx].tolist()

    # Cria DataFrame a partir da linha seguinte ao cabeçalho
    df = df_raw.iloc[head_idx+1:, :len(header_row)].copy()
    df.columns = [normalize_text(c) for c in header_row]

    # Esperado: primeira coluna vazia, segunda = BENEFICIARIO, demais = MESES
    # Seleciona somente colunas válidas (beneficiário + meses)
    cols = [c for c in df.columns if c and c != "NAN"]
    df = df[cols].copy()

    # Garante coluna 'BENEFICIARIO'
    # Pode estar em segunda coluna originalmente; protege contra variações
    ben_col_candidates = [c for c in df.columns if "BENEFICIARIO" in c]
    if not ben_col_candidates:
        # tenta um nome próximo
        ben_col_candidates = [df.columns[0]]
    ben_col = ben_col_candidates[0]

    # Colunas de meses = todas exceto beneficiário
    month_cols = [c for c in df.columns if c != ben_col]
    # Normaliza meses
    month_cols_norm = []
    mapping_cols = {}
    for c in month_cols:
        c_norm = normalize_text(c)
        # Mantém apenas colunas que pareçam mês (presentes no MONTH_ORDER)
        if c_norm in MONTH_ORDER:
            month_cols_norm.append(c_norm)
            mapping_cols[c] = c_norm
        else:
            # ignora colunas que não são mês
            pass

    if not month_cols_norm:
        # Sem meses válidos, retorna vazio amigavelmente
        return pd.DataFrame(columns=["BENEFICIARIO", "MES", "VALOR"])

    # Renomeia para meses padronizados
    df = df.rename(columns=mapping_cols)

    # Mantém somente beneficiário + meses válidos
    keep_cols = [ben_col] + month_cols_norm
    df = df[keep_cols].copy()
    df = df.rename(columns={ben_col: "BENEFICIARIO"})
    df["BENEFICIARIO"] = df["BENEFICIARIO"].apply(normalize_text)

    # Derrete para formato longo
    df_long = df.melt(id_vars=["BENEFICIARIO"], value_vars=month_cols_norm,
                      var_name="MES", value_name="VALOR")
    # Tipos
    df_long["MES"] = df_long["MES"].apply(normalize_text)
    df_long["VALOR"] = pd.to_numeric(df_long["VALOR"], errors="coerce").fillna(0.0)

    # Remove linhas sem beneficiário
    df_long = df_long[df_long["BENEFICIARIO"].notna() & (df_long["BENEFICIARIO"] != "")]
    df_long = order_months(df_long, "MES")
    return df_long

def df_download_excel(df: pd.DataFrame, filename: str) -> bytes:
    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="dados")
    return buffer.getvalue()

# =========================
# Sidebar: Upload / Fontes
# =========================
st.sidebar.title("⚙️ Fontes de Dados")
st.sidebar.caption("Se preferir, envie os arquivos. Caso não envie, tento usar os nomes padrão.")

default_all = "COMBUSTIVEL 2025.xlsx"
default_admin = "Combustivel POR SECRETARIA.xlsx"

uploaded_all = st.sidebar.file_uploader(
    "Planilha de TODAS as Secretarias (aba 'Planilha1')",
    type=["xlsx"], key="all_xlsx"
)
uploaded_admin = st.sidebar.file_uploader(
    "Planilha da Administração + Departamentos (aba 'GERAL')",
    type=["xlsx"], key="admin_xlsx"
)

# Decide fonte
all_source = uploaded_all if uploaded_all is not None else (default_all if os.path.exists(default_all) else None)
adm_source = uploaded_admin if uploaded_admin is not None else (default_admin if os.path.exists(default_admin) else None)

# =========================
# Título
# =========================
st.title("📊 Comparativo de Gastos")
st.caption("Totais por Secretaria + Detalhamento da Secretaria de Administração e beneficiários/departamentos.")

# =========================
# Carrega dados
# =========================
col_status1, col_status2 = st.columns(2)
with col_status1:
    if all_source is None:
        st.error("Não encontrei a planilha de TODAS as Secretarias. Envie o arquivo ou coloque **COMBUSTIVEL 2025.xlsx** na mesma pasta.")
    else:
        st.success("Planilha de TODAS as Secretarias carregada.")

with col_status2:
    if adm_source is None:
        st.warning("Planilha da Administração + Departamentos não encontrada. Você pode usar o app só com a visão geral das Secretarias.")
    else:
        st.success("Planilha da Administração + Departamentos carregada.")

df_all = load_all_secretarias_from_planilha1(all_source) if all_source else pd.DataFrame(columns=["MES","SECRETARIA","VALOR","MES_NUM"])
df_adm = load_admin_departments(adm_source) if adm_source else pd.DataFrame(columns=["BENEFICIARIO","MES","VALOR","MES_NUM"])

# =========================
# Filtros globais
# =========================
st.subheader("🔎 Filtros")
colf1, colf2, colf3 = st.columns([1.2, 1.7, 1.1])

# Meses disponíveis (na ordem correta)
meses_disponiveis = sorted(df_all["MES"].dropna().unique().tolist(), key=lambda m: MONTH_TO_NUM.get(m, 99))
if not meses_disponiveis:  # fallback para quando há só df_adm
    meses_disponiveis = sorted(df_adm["MES"].dropna().unique().tolist(), key=lambda m: MONTH_TO_NUM.get(m, 99))

with colf1:
    meses_sel = st.multiselect(
        "Meses", options=meses_disponiveis, default=meses_disponiveis
    )

with colf2:
    secretarias_disponiveis = sorted(df_all["SECRETARIA"].dropna().unique().tolist())
    secs_sel = st.multiselect(
        "Secretarias (para a visão geral)",
        options=secretarias_disponiveis,
        default=[s for s in secretarias_disponiveis if s in ["ADMINISTRAÇÃO","OBRAS","SAÚDE"]][:3]
    )

with colf3:
    topn = st.number_input("Top N (ranking)", min_value=3, max_value=30, value=10, step=1)

# Aplica filtros
df_all_f = df_all[df_all["MES"].isin(meses_sel)].copy()
df_adm_f = df_adm[df_adm["MES"].isin(meses_sel)].copy()

# =========================
# Abas
# =========================
tabs = st.tabs(["📈 Visão Geral (Secretarias)", "⚖️ Administração x Demais", "🏛️ Administração por Beneficiário", "📄 Tabelas / Exportar"])

# -------------------------
# Aba 1: Visão Geral
# -------------------------
with tabs[0]:
    st.markdown("### 📈 Gastos por Secretaria (mensal)")
    if df_all_f.empty:
        st.info("Sem dados para exibir.")
    else:
        if secs_sel:
            plot_df = df_all_f[df_all_f["SECRETARIA"].isin(secs_sel)].groupby(["MES","SECRETARIA"], as_index=False)["VALOR"].sum()
        else:
            plot_df = df_all_f.groupby(["MES","SECRETARIA"], as_index=False)["VALOR"].sum()

        plot_df = order_months(plot_df, "MES")
        fig = px.line(
            plot_df, x="MES", y="VALOR", color="SECRETARIA",
            markers=True, title="Evolução Mensal por Secretaria"
        )
        fig.update_layout(yaxis_title="Valor (R$)")
        st.plotly_chart(fig, use_container_width=True)

        # Barras empilhadas
        st.markdown("#### Barras Empilhadas (participação por mês)")
        fig2 = px.bar(
            plot_df, x="MES", y="VALOR", color="SECRETARIA",
            title="Participação por Secretaria em cada mês", barmode="stack"
        )
        fig2.update_layout(yaxis_title="Valor (R$)")
        st.plotly_chart(fig2, use_container_width=True)

        # Ranking total por secretaria (nos meses filtrados)
        rank_df = plot_df.groupby("SECRETARIA", as_index=False)["VALOR"].sum().sort_values("VALOR", ascending=False).head(topn)
        rank_df["VALOR_fmt"] = rank_df["VALOR"].map(format_brl)
        st.markdown("#### 🔝 Ranking (soma nos meses filtrados)")
        st.dataframe(rank_df[["SECRETARIA","VALOR_fmt"]], use_container_width=True, hide_index=True)

# -------------------------
# Aba 2: Administração x Demais
# -------------------------
with tabs[1]:
    st.markdown("### ⚖️ Administração vs Demais Secretarias")
    if df_all_f.empty:
        st.info("Sem dados para exibir.")
    else:
        # Agrega por mês
        total_mes = df_all_f.groupby("MES", as_index=False)["VALOR"].sum().rename(columns={"VALOR":"TOTAL"})
        admin_mes = df_all_f[df_all_f["SECRETARIA"] == "ADMINISTRAÇÃO"].groupby("MES", as_index=False)["VALOR"].sum().rename(columns={"VALOR":"ADMINISTRAÇÃO"})
        base = pd.merge(total_mes, admin_mes, on="MES", how="left")
        base["ADMINISTRAÇÃO"] = base["ADMINISTRAÇÃO"].fillna(0.0)
        base["DEMAIS"] = (base["TOTAL"] - base["ADMINISTRAÇÃO"]).clip(lower=0)

        base = order_months(base, "MES")
        base_long = base.melt(id_vars=["MES","MES_NUM"], value_vars=["ADMINISTRAÇÃO","DEMAIS"], var_name="GRUPO", value_name="VALOR")

        fig = px.bar(base_long, x="MES", y="VALOR", color="GRUPO", barmode="group",
                     title="Administração x Demais (por mês)")
        fig.update_layout(yaxis_title="Valor (R$)")
        st.plotly_chart(fig, use_container_width=True)

        # Linha comparativa
        fig2 = px.line(base_long, x="MES", y="VALOR", color="GRUPO", markers=True,
                       title="Evolução Mensal: Administração x Demais")
        fig2.update_layout(yaxis_title="Valor (R$)")
        st.plotly_chart(fig2, use_container_width=True)

        # Tabela resumo
        base_show = base[["MES","ADMINISTRAÇÃO","DEMAIS","TOTAL"]].copy()
        for c in ["ADMINISTRAÇÃO","DEMAIS","TOTAL"]:
            base_show[c] = base_show[c].map(format_brl)
        st.dataframe(base_show, use_container_width=True, hide_index=True)

# -------------------------
# Aba 3: Administração por Beneficiário (Departamentos/Pessoas/Órgãos)
# -------------------------
with tabs[2]:
    st.markdown("### 🏛️ Administração — Detalhe por Beneficiário")
    if df_adm_f.empty:
        st.info("Sem dados da planilha de Administração + Departamentos.")
    else:
        # Filtros locais
        beneficiarios = sorted(df_adm_f["BENEFICIARIO"].unique().tolist())
        colb1, colb2 = st.columns([1.5, 2])
        with colb1:
            bens_sel = st.multiselect("Beneficiários/Departamentos", options=beneficiarios, default=beneficiarios[:10])
        with colb2:
            mes_unico = st.selectbox("Mês para ranking de barras", options=meses_sel, index=0 if meses_sel else 0)

        # Linha/Áreas por beneficiário (mensal)
        df_plot = df_adm_f.copy()
        if bens_sel:
            df_plot = df_plot[df_plot["BENEFICIARIO"].isin(bens_sel)]

        df_plot = df_plot.groupby(["MES","BENEFICIARIO"], as_index=False)["VALOR"].sum()
        df_plot = order_months(df_plot, "MES")

        fig = px.line(df_plot, x="MES", y="VALOR", color="BENEFICIARIO", markers=True,
                      title="Evolução Mensal por Beneficiário (Administração)")
        fig.update_layout(yaxis_title="Valor (R$)")
        st.plotly_chart(fig, use_container_width=True)

        # Ranking no mês selecionado
        rank_mes = df_adm_f[df_adm_f["MES"] == mes_unico].groupby("BENEFICIARIO", as_index=False)["VALOR"].sum()
        rank_mes = rank_mes.sort_values("VALOR", ascending=False).head(topn)
        fig_bar = px.bar(rank_mes, x="BENEFICIARIO", y="VALOR", title=f"Top {topn} no mês {mes_unico}", text_auto=".2s")
        fig_bar.update_layout(xaxis_title="", yaxis_title="Valor (R$)")
        st.plotly_chart(fig_bar, use_container_width=True)

        # Acumulado (meses filtrados)
        total_ben = df_adm_f.groupby("BENEFICIARIO", as_index=False)["VALOR"].sum().sort_values("VALOR", ascending=False).head(topn)
        total_ben["VALOR_fmt"] = total_ben["VALOR"].map(format_brl)
        st.markdown("#### 🔝 Ranking acumulado (meses filtrados)")
        st.dataframe(total_ben[["BENEFICIARIO","VALOR_fmt"]], use_container_width=True, hide_index=True)

# -------------------------
# Aba 4: Tabelas e Exportação
# -------------------------
with tabs[3]:
    st.markdown("### 📄 Dados e Exportação")
    colx1, colx2 = st.columns(2)
    with colx1:
        st.markdown("**Tabela — Secretarias (filtrada):**")
        show_all = df_all_f.copy()
        show_all["VALOR"] = show_all["VALOR"].round(2)
        st.dataframe(show_all, use_container_width=True, hide_index=True)

        if not show_all.empty:
            x1 = df_download_excel(show_all, "secretarias_filtrado.xlsx")
            st.download_button("⬇️ Baixar Excel (Secretarias)", data=x1, file_name="secretarias_filtrado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    with colx2:
        st.markdown("**Tabela — Administração (Beneficiários, filtrada):**")
        show_adm = df_adm_f.copy()
        show_adm["VALOR"] = show_adm["VALOR"].round(2)
        st.dataframe(show_adm, use_container_width=True, hide_index=True)

        if not show_adm.empty:
            x2 = df_download_excel(show_adm, "administracao_filtrado.xlsx")
            st.download_button("⬇️ Baixar Excel (Administração)", data=x2, file_name="administracao_filtrado.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

# Rodapé
st.caption("Feito para você analisar rapidamente o gasto mensal por Secretaria e abrir a caixa-preta da Administração.")
