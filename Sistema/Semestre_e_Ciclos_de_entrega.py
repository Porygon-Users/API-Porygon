import openpyxl
from datetime import datetime, timedelta

# Função para adicionar a data de início do semestre e calcular o fim do semestre
def adicionar_data_inicio_e_calcula_fim(planilha, turma_destino, data_inicio_semestre, data_fim_semestre):
    aba_turma = planilha[turma_destino]

    # Converter a data de início e fim para objetos datetime
    data_inicio = datetime.strptime(data_inicio_semestre, "%d/%m/%Y")
    data_fim = datetime.strptime(data_fim_semestre, "%d/%m/%Y")

    # Encontrar as colunas "Inicio de Semestre" e "Fim de Semestre"
    coluna_inicio = None
    coluna_fim = None

    for row in aba_turma.iter_rows(min_row=1, max_row=1):
        for cell in row:
            if cell.value == "Inicio de Semestre":
                coluna_inicio = cell.column
            elif cell.value == "Fim de Semestre":
                coluna_fim = cell.column

    if not coluna_inicio or not coluna_fim:
        print("Não foram encontradas colunas com os cabeçalhos 'Inicio de Semestre' e 'Fim de Semestre'.")
        return

    # Validar se a data de início está dentro do semestre
    if not (data_inicio >= data_inicio and data_inicio <= data_fim):
        print("A data de início do semestre está fora do período do semestre.")
        return

    # Validar se a data de fim está dentro do semestre
    if not (data_fim >= data_inicio and data_fim <= data_fim):
        print("A data de término do semestre está fora do período do semestre.")
        return

    # Adicionar a data de início na coluna "Início de Semestre"
    for row in aba_turma.iter_rows(min_row=2, max_row=aba_turma.max_row, min_col=coluna_inicio, max_col=coluna_inicio):
        row[0].value = data_inicio.strftime("%d/%m/%Y")
    
    # Adicionar a data de fim na coluna "Fim de Semestre"
    for row in aba_turma.iter_rows(min_row=2, max_row=aba_turma.max_row, min_col=coluna_fim, max_col=coluna_fim):
        row[0].value = data_fim.strftime("%d/%m/%Y")

    planilha.save('Dados Cadastrais.xlsx')

# Função para criar ciclos de entregas dentro do semestre
def adicionar_ciclos(planilha, turma_destino, data_inicio_semestre, data_fim_semestre):
    aba_turma = planilha[turma_destino]

    # Converter a data de início e fim para objetos datetime
    data_inicio_semestre = datetime.strptime(data_inicio_semestre, "%d/%m/%Y")
    data_fim_semestre = datetime.strptime(data_fim_semestre, "%d/%m/%Y")

    # Solicitar ao usuário que forneça as datas de início e término para cada ciclo
    ciclos_datas = []
    
    qtd_ciclos = int(input("Quantos ciclos você deseja: "))
    
    for i in range(qtd_ciclos):
        ciclo_nome = f"Ciclo {i+1}"

        # Solicitar e validar a data de início do ciclo
        while True:
            ciclo_inicio = input(f"Digite a data de início para o {ciclo_nome} (DD/MM/AAAA): ")
            ciclo_inicio = datetime.strptime(ciclo_inicio, "%d/%m/%Y")
            if ciclo_inicio >= data_inicio_semestre and ciclo_inicio <= data_fim_semestre:
                break
            else:
                print("A data de início do ciclo está fora do período do semestre.")

        # Solicitar e validar a data de término do ciclo
        while True:
            ciclo_fim = input(f"Digite a data de término para o {ciclo_nome} (DD/MM/AAAA): ")
            ciclo_fim = datetime.strptime(ciclo_fim, "%d/%m/%Y")
            if ciclo_fim >= data_inicio_semestre and ciclo_fim <= data_fim_semestre:
                break
            else:
                print("A data de término do ciclo está fora do período do semestre.")

        ciclos_datas.append((ciclo_nome, ciclo_inicio, ciclo_fim))

    # Inicializa listas para armazenar os ciclos de início e término
    ciclos_inicio = []
    ciclos_termino = []

    # Iterar sobre os ciclos e adicionar as datas nas listas
    for i, (ciclo_nome, ciclo_inicio, ciclo_fim) in enumerate(ciclos_datas):
        ciclos_inicio.append(f"{ciclo_inicio.strftime('%d/%m/%Y')} - Ciclo {i+1}")
        ciclos_termino.append(f"{ciclo_fim.strftime('%d/%m/%Y')} - Ciclo {i+1}")

    # Escrever os ciclos de início e término na planilha
    for i in range(len(ciclos_inicio)):
        aba_turma.cell(row=i+2, column=9).value = ciclos_inicio[i]
        aba_turma.cell(row=i+2, column=10).value = ciclos_termino[i]

    # Adicionar cabeçalhos
    aba_turma.cell(row=1, column=9).value = "Ciclo de Início"
    aba_turma.cell(row=1, column=10).value = "Ciclo de Término"

    planilha.save('Dados Cadastrais.xlsx')
    print("Ciclos de entrega adicionados com sucesso.")
    return qtd_ciclos

# Abrir a planilha
planilha = openpyxl.load_workbook('Dados Cadastrais.xlsx')

while True:
    print("\nOpções:")
    print("\n1. Adicionar data de início do semestre")
    print("2. Criar ciclos de entrega")
    print("3. Sair do programa", "\n")

    escolha = input("Escolha uma das opções: ")

    if escolha == '1':
        # Listar as abas de turma disponíveis
        abas_turmas = [sheet for sheet in planilha.sheetnames if sheet.startswith('Turma ')]

        if not abas_turmas:
            print("\nNão foram encontradas abas de turma na planilha.")
        else:
            print("\nTurmas existentes:", "\n")
            for i, turma in enumerate(abas_turmas, start=1):
                print(f"{i}. {turma}", "\n")

            # Solicitar ao usuário qual turma deseja adicionar a data
            num_turma = int(input("Escolha o número da turma: "))
            if 1 <= num_turma <= len(abas_turmas):
                turma_desejada = abas_turmas[num_turma - 1]
                data_inicio_semestre = input("Digite a data de início do semestre (DD/MM/AAAA): ")
                data_fim_semestre = input("Digite a data de término do semestre (DD/MM/AAAA): ")
                adicionar_data_inicio_e_calcula_fim(planilha, turma_desejada, data_inicio_semestre, data_fim_semestre)
                print("\nData de início e fim do semestre adicionadas com sucesso.")
            else:
                print("Número de turma inválido.")
    elif escolha == '2':
        abas_turmas = [sheet for sheet in planilha.sheetnames if sheet.startswith('Turma ')]

        if not abas_turmas:
            print("\nNão foram encontradas abas de turma na planilha.")
        else:
            print("\nTurmas existentes:", "\n")
            for i, turma in enumerate(abas_turmas, start=1):
                print(f"{i}. {turma}", "\n")

            num_turma = int(input("Escolha o número da turma: "))
            if 1 <= num_turma <= len(abas_turmas):
                turma_desejada = abas_turmas[num_turma - 1]
                data_inicio_semestre = input("Digite a data de início do semestre (DD/MM/AAAA): ")
                data_fim_semestre = input("Digite a data de término do semestre (DD/MM/AAAA): ")
                adicionar_ciclos(planilha, turma_desejada, data_inicio_semestre, data_fim_semestre)
            else:
                print("Número de turma inválido.")
    
    elif escolha == '3':
        break
    else:
        print("Opção inválida. Escolha 1 para adicionar a data de início do semestre, 2 para adicionar ciclos de entrega ou 3 para sair.")

print("\nPrograma encerrado.", "\n")