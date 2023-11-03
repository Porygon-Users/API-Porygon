import streamlit as st
import pandas as pd
import plotly.express as px


# Função para carregar os dados do arquivo Excel
def carregar_dados_excel():
    excel_file = pd.ExcelFile("infodados.xlsx")
    sheet_names = excel_file.sheet_names
    return excel_file, sheet_names

# Função para filtrar as planilhas de turma
def filtrar_planilhas_turma(excel_file, sheet_names):
    keywords = ["turma", "classe", "grupo"]
    turma_sheets = [sheet_name for sheet_name in sheet_names if any(keyword in sheet_name.lower() for keyword in keywords) and "fechada" not in sheet_name.lower()]
    return turma_sheets

# Função para selecionar a planilha de turma
def selecionar_planilha_turma(turma_sheets):
    selected_sheet = st.sidebar.selectbox("Selecione a turma:", turma_sheets)
    return selected_sheet

# Função para aplicar o filtro de alunos
def aplicar_filtro_alunos(df_turma):
    alunos_unicos = df_turma["Alunos"].unique()
    
    # Add a "Select All" button
    select_all_button = st.sidebar.checkbox("Todos os alunos")
    
    if select_all_button:
        selected_alunos = alunos_unicos
    else:
        # Use multiselect if the "Selecionar Todos" checkbox is not selected
        selected_alunos = st.sidebar.multiselect("Selecione os alunos:", alunos_unicos)
    
    return selected_alunos



# Função para filtrar o DataFrame com base nos alunos e ciclo selecionados
def filtrar_dataframe(df_turma, alunos_selecionados, selected_ciclo):
    df_selecao = df_turma[df_turma["Alunos"].isin(alunos_selecionados)][["Alunos", selected_ciclo, "Média", "Professores", "Início do Curso", "Fim do Curso"]]
    return df_selecao

# Função para criar gráficos
def criar_graficos(df_selecao, selected_ciclo):
    col1, col2 = st.columns(2)
    
    nota_por_aluno = px.bar(df_selecao, x="Alunos", y=selected_ciclo, title=f"Nota do {selected_ciclo} por aluno")
    col1.plotly_chart(nota_por_aluno)

    media_por_aluno = px.bar(df_selecao, x="Alunos", y="Média", title="Média por aluno")
    col2.plotly_chart(media_por_aluno)

# Configurações da página
st.set_page_config(layout="wide", page_title="Relatórios gerais", page_icon=":bar_chart:")
st.sidebar.image("pbltex.jpg", caption="Análise de dados")
st.title("Relatório das Turmas")

# Carregar os dados do arquivo Excel
excel_file, sheet_names = carregar_dados_excel()

# Filtrar as planilhas de turma
turma_sheets = filtrar_planilhas_turma(excel_file, sheet_names)

# Selecionar a planilha de turma
selected_sheet = selecionar_planilha_turma(turma_sheets)

# Ler os dados da planilha selecionada
df_turma = pd.read_excel("infodados.xlsx", sheet_name=selected_sheet)

# Adicionar um seletor de ciclo
colunas_ciclo = [coluna for coluna in df_turma.columns if coluna.lower().startswith("ciclo")]
selected_ciclo = st.sidebar.selectbox("Selecione o ciclo:", colunas_ciclo)

# Aplicar o filtro de alunos
alunos_selecionados = aplicar_filtro_alunos(df_turma)

# Filtrar o DataFrame com base nos alunos e ciclo selecionados
df_selecao = filtrar_dataframe(df_turma, alunos_selecionados, selected_ciclo)

# Criar gráficos
criar_graficos(df_selecao, selected_ciclo)

st.dataframe(df_selecao)