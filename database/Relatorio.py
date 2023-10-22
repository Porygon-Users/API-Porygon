import streamlit as st
import pandas as pd
import plotly.express as px

excel_file_path = os.path.join(dir_path, "infodados.xlsx")

st.set_page_config(layout="wide") #configuração do site para ficar no layout correto

df_turma1 = pd.read_excel("C:\Users\Isaque\Documents\API-Porygon\database\infodados.xlsx", sheet_name='Turma 1 (fechada)')
#df_turma2 = pd.read_excel("infodados.xlsx", sheet_name= "Turma 2")

#all_spreadsheets = pd.concat(df_turma1,df_turma2)
# df_turma1["Início do Ciclo"] = pd.to_datetime(df_turma1["Início do Ciclo"])

# df_turma1 = df_turma1.sort_values("Início do Ciclo")

# df_turma1["Month"] = df_turma1["Início do Ciclo"].apply(lambda x: str(x.year) + "-" + str(x.month))

# monthy = st.sidebar.selectbox("Ciclos", df_turma1["Month"].unique())

#df_filtered = df_turma1[df_turma1["Month"] == monthy] #Filtrar o gráfico com base no mês

aluno = st.sidebar.selectbox("Alunos", df_turma1["Alunos"].unique())


filtro_do_aluno = df_turma1[df_turma1["Alunos"] == aluno]

alunos_notas_1 = df_turma1

st.write(alunos_notas_1)

col1, col2 = st.columns(2)
col3, col4, col5 = st.columns(3)

nota_por_aluno = px.bar(alunos_notas_1, x = "Alunos", y = "Nota", title = "Nota por aluno")
col1.plotly_chart(nota_por_aluno)









