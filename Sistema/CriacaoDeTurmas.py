#Bibliote que fará que o python tenha uma conexão com o Excel
import openpyxl

#Criar uma planilha
book = openpyxl.load_workbook('dados.xlsx')
#Criando uma página
book.create_sheet('Turmas')
#Selecionando a página
turmas_page = book['Turmas']
turmas_page.append(['Nome da matéria', 'aluno', 'matéria', 'Nome do professor'])

#salvar a planilha
book.save('dados.xlsx')






