import openpyxl
import random
import string

# Função para gerar RP aleatório para professores
def gerar_rp():
    return ''.join(random.choice(string.digits) for _ in range(10))

# Função para cadastrar um novo professor
def cadastrar_professor():
    try:
        wb = openpyxl.load_workbook('dados.xlsx')
    except FileNotFoundError:
        print("A planilha 'dados.xlsx' não foi encontrada. Certifique-se de criar a planilha com os alunos primeiro.")
        return

    professor_sheet = wb["Prof"]
    
    if professor_sheet is None:
        print("Planilha 'Prof' não encontrada.")
        return
    
    nome = input("Digite o nome completo do professor: ")
    email = input("Digite o email do professor: ")
    rp = gerar_rp()
    
    professor_sheet.append([rp, nome, email])
    
    wb.save('dados.xlsx')
    print(f"RP do Professor: {rp}")
    print("Professor cadastrado com sucesso!")

# Chamando a função para cadastrar um novo professor
cadastrar_professor()
