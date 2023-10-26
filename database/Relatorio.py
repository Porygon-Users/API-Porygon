import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide", page_title="Relatórios gerais", page_icon=":bar_chart:")

# Ler todas as planilhas do arquivo Excel
excel_file = pd.ExcelFile("infodados.xlsx")
sheet_names = excel_file.sheet_names

# Filtrar apenas as planilhas que contêm "turma" no nome
turma_sheets = [sheet_name for sheet_name in sheet_names if "turma" in sheet_name.lower()]

# Criar um filtro de seleção para as turmas
selected_sheet = st.sidebar.selectbox("Selecione a turma:", turma_sheets)

# Ler os dados da planilha selecionada
df_turma = pd.read_excel("infodados.xlsx", sheet_name=selected_sheet)

alunos_selecionados = []

st.sidebar.header("Utilize o filtro de alunos:")

# Botão para selecionar todos os alunos
if st.sidebar.button("Selecionar Todos"):
    alunos_selecionados = df_turma["Alunos"].unique()

# Botão para remover todas as seleções
if st.sidebar.button("Remover Todos"):
    alunos_selecionados = []

# Caixa de seleção para escolher alunos individualmente
alunos_unicos = df_turma["Alunos"].unique()
alunos_selecionados = st.sidebar.multiselect("Selecione os alunos:", alunos_unicos, default=alunos_selecionados)

# Filtrar o DataFrame com base nos alunos selecionados
df_selecao = df_turma[df_turma["Alunos"].isin(alunos_selecionados)]

st.dataframe(df_selecao)

col1, col2 = st.columns(2)
col3, col4, col5 = st.columns(3)

nota_por_aluno = px.bar(df_selecao, x="Alunos", y="Nota", title="Nota por aluno")
col1.plotly_chart(nota_por_aluno)
print("a")