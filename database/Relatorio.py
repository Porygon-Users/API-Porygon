import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide", page_title="Relatórios gerais", page_icon=":bar_chart:") 

st.sidebar.image("pbltex.jpg",caption = "Análise de dados") #inserção da imagem

st.title("Relatório das Turmas")

# Ler todas as planilhas do arquivo Excel
excel_file = pd.ExcelFile("infodados.xlsx")
sheet_names = excel_file.sheet_names

# Defina palavras-chave relacionadas às turmas
keywords = ["turma", "classe", "grupo"]  # Adicione outras palavras-chave, se necessário

# Filtrar as planilhas que contêm pelo menos uma palavra-chave e não contêm "fechada"
turma_sheets = [sheet_name for sheet_name in sheet_names if any(keyword in sheet_name.lower() for keyword in keywords) and "fechada" not in sheet_name.lower()]

# Criar um filtro de seleção para as turmas
selected_sheet = st.sidebar.selectbox("Selecione a turma:", turma_sheets)

# Ler os dados da planilha selecionada
df_turma = pd.read_excel("infodados.xlsx", sheet_name=selected_sheet)


alunos_selecionados = []    

# Adicione um seletor de ciclo
colunas_ciclo = [coluna for coluna in df_turma.columns if coluna.lower().startswith("ciclo")]
selected_ciclo = st.sidebar.selectbox("Selecione o ciclo:", colunas_ciclo)

st.sidebar.header("Utilize o filtro de alunos:")
# Inicialmente, definimos a variável de seleção de alunos como todos os alunos
alunos_selecionados = df_turma["Alunos"].unique()

# Verifique se o botão "Remover Todos" foi clicado
if st.sidebar.button("Remover Todos"):
    alunos_selecionados = []

# Verifique se o botão "Selecionar Todos" foi clicado
if st.sidebar.button("Selecionar Todos"):
    alunos_selecionados = df_turma["Alunos"].unique()

# Caixa de seleção para escolher alunos individualmente
alunos_unicos = df_turma["Alunos"].unique()
alunos_selecionados = st.sidebar.multiselect("Selecione os alunos:", alunos_unicos, default=alunos_selecionados)


# Filtrar o DataFrame com base nos alunos selecionados e no ciclo selecionado
df_selecao = df_turma[df_turma["Alunos"].isin(alunos_selecionados)][["Alunos", selected_ciclo, "Média"]]

# Preencher valores nulos na coluna "Média" com 0
df_selecao["Média"].fillna(0, inplace=True)

st.dataframe(df_selecao)

col1, col2 = st.columns(2)
col3, col4, col5 = st.columns(3)

# Criar um gráfico de barras para as notas do ciclo selecionado
nota_por_aluno = px.bar(df_selecao, x="Alunos", y=selected_ciclo, title=f"Nota do {selected_ciclo} por aluno")
col1.plotly_chart(nota_por_aluno)

# Criar um gráfico de colunas para a média de cada aluno
media_por_aluno = px.bar(df_selecao, x="Alunos", y="Média", title="Média por aluno")
col2.plotly_chart(media_por_aluno)

