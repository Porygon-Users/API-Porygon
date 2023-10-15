import openpyxl
from openpyxl.styles import Font, Alignment
import os

# Especifique o caminho completo para o arquivo Excel
caminho_arquivo = "database/infodados.xlsx"

# Verifique se o arquivo Excel já existe
if os.path.exists(caminho_arquivo):
    # Se o arquivo existe, carregue-o em vez de criar um novo
    arquivo_excel = openpyxl.load_workbook(caminho_arquivo)
    # Acesse a planilha existente
    planilha = arquivo_excel.active
else:
    # Se o arquivo não existe, crie um novo
    arquivo_excel = openpyxl.Workbook()
    planilha = arquivo_excel.active
    planilha.title = "Cadastro"

# Defina estilos para o cabeçalho
cabecalho_fonte = Font(bold=True)
cabecalho_alinhamento = Alignment(horizontal='center')

# Adicione cabeçalhos se a planilha for nova
if not os.path.exists("infodados.xlsx"):
    planilha['A1'] = "NOME"
    planilha['B1'] = "CPF"
    planilha['C1'] = "E-MAIL"
    planilha['D1'] = "FUNÇÃO"
    planilha['E1'] = "LOGIN"
    planilha['F1'] = "SENHA"

    # Aplicar estilos ao cabeçalho
    for cell in planilha['1:1']:
        cell.font = cabecalho_fonte
        cell.alignment = cabecalho_alinhamento

    # Aumentar a largura das colunas
    planilha.column_dimensions['A'].width = 20
    planilha.column_dimensions['B'].width = 20
    planilha.column_dimensions['C'].width = 25
    planilha.column_dimensions['D'].width = 20
    planilha.column_dimensions['E'].width = 20
    planilha.column_dimensions['F'].width = 20

# Salve o arquivo Excel com o caminho completo
arquivo_excel.save(caminho_arquivo)

# Feche o arquivo
arquivo_excel.close()