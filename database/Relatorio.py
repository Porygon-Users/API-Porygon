import streamlit as st
import pandas as pd
import plotly.express as px

st.set_page_config(layout="wide") #configuração do site para ficar no layout correto

df_turma1 = pd.read_excel("infodados.xlsx", sheet_name='Turma 1 (fechada)')
df_turma2 = pd.read_excel("infodados.xlsx", sheet_name= "Turma 2")

#all_spreadsheets = pd.concat(df_turma1,df_turma2)
df_turma1["Início do Ciclo"] = pd.to_datetime(df_turma1["Início do Ciclo"])

df_turma1 = df_turma1.sort_values("Início do Ciclo")

df_turma1["Month"] = df_turma1["Início do Ciclo"].apply(lambda x: str(x.year) + "-" + str(x.month))

monthy = st.sidebar.selectbox("Ciclos", df_turma1["Month"].unique())

df_filtered = df_turma1[df_turma1["Month"] == monthy] 

col1, col2 = st.columns(2)
col3, col4, col5 = st.columns(3)

df_filtered



#df_turma1 = df_turma1.sort_values(df_turma1["Início do Ciclo"])

#df_turma1["Monthy"] = df_turma1["Início do Ciclo"].apply(lambda x: str(x.year) + )




