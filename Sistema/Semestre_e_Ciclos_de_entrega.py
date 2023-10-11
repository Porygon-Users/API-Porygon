import openpyxl
from datetime import datetime, timedelta

# Função para adicionar a data de início do curso e calcular o fim do curso
def adicionar_data_e_ciclos(planilha, turma_destino):
    aba_turma = planilha[turma_destino]

    data_inicio = datetime.strptime(input("Digite a data de início do curso (DD/MM/AAAA): "), "%d/%m/%Y")
    data_fim = datetime.strptime(input("Digite a data de término do curso (DD/MM/AAAA): "), "%d/%m/%Y")

    # Adicionar data de início e término do curso na planilha
    aba_turma.cell(row=1, column=7).value = "Início do Curso"
    aba_turma.cell(row=1, column=8).value = "Fim do Curso"
    aba_turma.cell(row=2, column=7).value = data_inicio.strftime('%d/%m/%Y')
    aba_turma.cell(row=2, column=8).value = data_fim.strftime('%d/%m/%Y')

    qtd_ciclos = int(input("Quantos ciclos você deseja: "))

    # Ajuste para garantir que todos os ciclos, exceto o último, tenham a mesma quantidade de dias
    duracao_ciclo = (data_fim - data_inicio) / qtd_ciclos

    ciclos = []

    for i in range(qtd_ciclos - 1):  # Até o penúltimo ciclo
        ciclo_nome = f"Ciclo {i + 1}"
        ciclo_inicio = data_inicio + i * duracao_ciclo
        ciclo_fim = ciclo_inicio + duracao_ciclo - timedelta(days=1)
        ciclos.append((ciclo_nome, ciclo_inicio, ciclo_fim))

    # Último ciclo inclui o último dia do curso
    ultimo_ciclo_nome = f"Ciclo {qtd_ciclos}"
    ultimo_ciclo_inicio = data_inicio + (qtd_ciclos - 1) * duracao_ciclo
    ultimo_ciclo_fim = data_fim
    ciclos.append((ultimo_ciclo_nome, ultimo_ciclo_inicio, ultimo_ciclo_fim))

    # Adicionar cabeçalhos para os dias do ciclo, início e término do ciclo
    aba_turma.cell(row=1, column=9).value = "Início do Ciclo"
    aba_turma.cell(row=1, column=10).value = "Término do Ciclo"
    aba_turma.cell(row=1, column=11).value = "Dias do Ciclo"

    # Calcular e adicionar a quantidade de dias do ciclo
    for i, (ciclo_nome, ciclo_inicio, ciclo_fim) in enumerate(ciclos):
        aba_turma.cell(row=i + 2, column=9).value = ciclo_inicio.strftime('%d/%m/%Y')
        aba_turma.cell(row=i + 2, column=10).value = ciclo_fim.strftime('%d/%m/%Y')

        # Calcular e adicionar a quantidade de dias do ciclo
        dias_ciclo = (ciclo_fim - ciclo_inicio).days + 1
        aba_turma.cell(row=i + 2, column=11).value = dias_ciclo

    planilha.save('Dados Cadastrais.xlsx')
    print("Data de início, fim do curso e ciclos adicionados com sucesso.")

# Abrir a planilha
planilha = openpyxl.load_workbook('Dados Cadastrais.xlsx')

while True:
    print("\nOpções:")
    print("\n1. Adicionar data de início do curso e criar ciclos")
    print("2. Sair do programa", "\n")

    escolha = input("Escolha uma das opções: ")

    if escolha == '1':
        abas_turmas = [sheet for sheet in planilha.sheetnames if sheet.startswith('Turma ')]

        if not abas_turmas:
            print("\nNão foram encontradas abas de turma na planilha.")
        else:
            print("\nTurmas existentes:", "\n")
            for i, turma in enumerate(abas_turmas, start=1):
                print(f"{i}. {turma}")

            num_turma = int(input("Escolha o número da turma: "))
            if 1 <= num_turma <= len(abas_turmas):
                turma_desejada = abas_turmas[num_turma - 1]
                adicionar_data_e_ciclos(planilha, turma_desejada)
            else:
                print("Número de turma inválido.")

    elif escolha == '2':
        break
    else:
        print("Opção inválida. Escolha 1 para adicionar a data de início do curso e criar ciclos ou 2 para sair.")

print("\nPrograma encerrado.", "\n")
