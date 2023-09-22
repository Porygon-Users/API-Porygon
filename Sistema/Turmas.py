import openpyxl

# Nome do arquivo Excel
arquivo_excel = "Dados Cadastrais.xlsx"

# Função para criar uma nova aba de turma
def criar_nova_turma(book, numero_turma):
    nova_aba = book.create_sheet(f"Turma {numero_turma}")
    nova_aba.append(["Aluno", "Nome do Professor"])
    book.save(arquivo_excel)

# Função para exibir o número de turmas
def mostrar_numero_de_turmas(book):
    turmas_existentes = sum(1 for sheet in book.sheetnames if sheet.startswith("Turma "))
    print(f"\nTotal de turmas: {turmas_existentes}")

# Carregar o arquivo Excel (se existir) ou criar um novo
try:
    book = openpyxl.load_workbook(arquivo_excel)
except FileNotFoundError:
    book = openpyxl.Workbook()
    book.save(arquivo_excel)

# Loop principal
while True:
    print("\nEscolha uma opção:")
    print("\n1 - Criar nova turma")
    print("2 - Visualizar número de turmas")
    print("3 - Sair", "\n")
    
    opcao = input('Digite o número da opção: ')
    
    if opcao == "1":
        quantidade_turmas = int(input("Digite a quantidade de turmas que deseja criar: ", "\n"))
        for _ in range(quantidade_turmas):
            proxima_turma = sum(1 for sheet in book.sheetnames if sheet.startswith("Turma ")) + 1
            criar_nova_turma(book, proxima_turma)
            print(f"Turma {proxima_turma} criada com sucesso.")
    elif opcao == "2":
        mostrar_numero_de_turmas(book)
    elif opcao == "3":
        break
    else:
        print("\nOpção inválida.", "\n")

print("Encerrando o programa.", "\n")
