import streamlit as st
import pandas as pd
import plotly.express as px


st.set_page_config(layout="wide", page_title= "Relatórios gerais",page_icon = ":bar_chart:") #configuração do layout da página

df_turma1 = pd.read_excel("infodados.xlsx", sheet_name = 'Turma 1 (fechada)')
df_turma2 = pd.read_excel("infodados.xlsx", sheet_name = "Turma 2")

#all_spreadsheets = pd.concat(df_turma1,df_turma2)
# df_turma1["Início do Ciclo"] = pd.to_datetime(df_turma1["Início do Ciclo"])

# df_turma1 = df_turma1.sort_values("Início do Ciclo")

# df_turma1["Month"] = df_turma1["Início do Ciclo"].apply(lambda x: str(x.year) + "-" + str(x.month))

# monthy = st.sidebar.selectbox("Ciclos", df_turma1["Month"].unique())

#df_filtered = df_turma1[df_turma1["Month"] == monthy] #Filtrar o gráfico com base no mês



alunos_selecionados = []

st.sidebar.header("Utilize o filtro: ")

# Botão para selecionar todos os alunos
if st.sidebar.button("Selecionar Todos"):
    alunos_selecionados = df_turma1["Alunos"].unique()

# Botão para remover todas as seleções
if st.sidebar.button("Remover Todos"):
    alunos_selecionados = []

# Caixa de seleção para escolher alunos individualmente
alunos_unicos = df_turma1["Alunos"].unique()
alunos_selecionados = st.sidebar.multiselect("Selecione os alunos: ", alunos_unicos, default=alunos_selecionados)

# Filtrar o DataFrame com base nos alunos selecionados
df_selecao = df_turma1[df_turma1["Alunos"].isin(alunos_selecionados)]

st.dataframe(df_selecao)
col1, col2 = st.columns(2)
col3, col4, col5 = st.columns(3)
    
nota_por_aluno = px.bar(df_selecao, x = "Alunos", y = "Nota", title = "Nota por aluno")
col1.plotly_chart(nota_por_aluno)










