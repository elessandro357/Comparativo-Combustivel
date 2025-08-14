# Comparativo de Gastos — Secretarias x Administração (Streamlit)

Dash para comparar:
1) Gastos mensais por Secretaria (arquivo: COMBUSTIVEL 2025.xlsx, aba "Planilha1")
2) Gastos mensais da Secretaria de Administração por beneficiário/departamento (arquivo: Combustivel POR SECRETARIA.xlsx, aba "GERAL")

## Como rodar
python -m venv .venv
source .venv/bin/activate  # (Windows: .venv\Scripts\activate)
pip install -r requirements.txt
streamlit run app.py

## Uso
- Faça upload dos arquivos no sidebar **ou** coloque os dois arquivos na mesma pasta do app com estes nomes:
  - `COMBUSTIVEL 2025.xlsx`
  - `Combustivel POR SECRETARIA.xlsx`

O app identifica automaticamente os cabeçalhos, normaliza os meses (JANEIRO..DEZEMBRO) e gera:
- Visão geral por Secretaria (linha + barras empilhadas + ranking)
- Comparativo Administração x Demais
- Detalhamento Administração por beneficiário (linha, ranking no mês, acumulado)
- Exportação para Excel (dados filtrados)
