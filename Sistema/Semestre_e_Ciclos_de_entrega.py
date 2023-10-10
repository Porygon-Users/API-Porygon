import openpyxl
from datetime import datetime, timedelta

# Definir data_inicio e data_fim como None inicialmente
data_inicio = None
data_fim = None

# Função para adicionar a data de início do curso e calcular o fim do curso
def adicionar_data_inicio_e_calcula_fim(planilha, turma_destino, data_inicio, data_fim):
    aba_turma = planilha[turma_destino]

    coluna_inicio = None
    coluna_fim = None

    for row in aba_turma.iter_rows(min_row=1, max_row=1):
        for cell in row:
            if cell.value == "Início do Curso":
                coluna_inicio = cell.column
            elif cell.value == "Fim do Curso":
                coluna_fim = cell.column

    if not coluna_inicio or not coluna_fim:
        print("Não foram encontradas colunas com os cabeçalhos 'Início do Curso' e 'Fim do Curso'.")
        return

    if not (data_inicio >= data_inicio and data_inicio <= data_fim):
        print("A data de início do curso está fora do período do curso.")
        return

    if not (data_fim >= data_inicio and data_fim <= data_fim):
        print("A data de término do curso está fora do período do curso.")
        return

    for row in aba_turma.iter_rows(min_row=2, max_row=aba_turma.max_row, min_col=coluna_inicio, max_col=coluna_inicio):
        row[0].value = data_inicio.strftime("%d/%m/%Y")

    for row in aba_turma.iter_rows(min_row=2, max_row=aba_turma.max_row, min_col=coluna_fim, max_col=coluna_fim):
        row[0].value = data_fim.strftime("%d/%m/%Y")

    planilha.save('Dados Cadastrais.xlsx')

# Função para criar ciclos de entregas dentro do curso
def adicionar_ciclos(planilha, turma_destino, data_inicio, data_fim):
    aba_turma = planilha[turma_destino]

    qtd_ciclos = int(input("Quantos ciclos você deseja: "))

    # Ajuste para garantir que o último ciclo inclua o último dia do curso
    duracao_ciclo = (data_fim - data_inicio) / (qtd_ciclos - 1)

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

    for i, (ciclo_nome, ciclo_inicio, ciclo_fim) in enumerate(ciclos):
        aba_turma.cell(row=i + 2, column=9).value = f"{ciclo_inicio.strftime('%d/%m/%Y')} - {ciclo_nome}"
        aba_turma.cell(row=i + 2, column=10).value = f"{ciclo_fim.strftime('%d/%m/%Y')} - {ciclo_nome}"

    aba_turma.cell(row=1, column=9).value = "Ciclo de Início"
    aba_turma.cell(row=1, column=10).value = "Ciclo de Término"

    planilha.save('Dados Cadastrais.xlsx')
    print("Ciclos de entrega adicionados com sucesso.")
    return qtd_ciclos

# Abrir a planilha
planilha = openpyxl.load_workbook('Dados Cadastrais.xlsx')

while True:
    print("\nOpções:")
    print("\n1. Adicionar data de início do curso")
    print("2. Criar ciclos de entrega")
    print("3. Sair do programa", "\n")

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
                data_inicio = datetime.strptime(input("Digite a data de início do curso(DD/MM/AAAA): "), "%d/%m/%Y")
                data_fim = datetime.strptime(input("Digite a data de término do curso (DD/MM/AAAA): "), "%d/%m/%Y")
                adicionar_data_inicio_e_calcula_fim(planilha, turma_desejada, data_inicio, data_fim)
                print("\nData de início e fim do curso adicionadas com sucesso.")
            else:
                print("Número de turma inválido.")
    elif escolha == '2':
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
                adicionar_ciclos(planilha, turma_desejada, data_inicio, data_fim)
            else:
                print("Número de turma inválido.")

    elif escolha == '3':
        break
    else:
        print("Opção inválida. Escolha 1 para adicionar a data de início do curso, 2 para adicionar ciclos de entrega ou 3 para sair.")

print("\nPrograma encerrado.", "\n")
