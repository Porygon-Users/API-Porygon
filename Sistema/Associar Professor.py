import openpyxl

# Função para verificar se um professor já está alocado em alguma turma
def professor_em_turma(planilha, professor_chave):
    abas_turmas = [sheet for sheet in planilha.sheetnames if sheet.startswith('Turma ')]
    
    for turma in abas_turmas:
        aba_turma = planilha[turma]
        for row in aba_turma.iter_rows(min_row=2, max_row=aba_turma.max_row, min_col=1, max_col=3):
            nome = row[0].value
            cpf = row[1].value
            professor_turma_chave = (nome, cpf)
            
            if professor_turma_chave == professor_chave:
                return True
    
    return False

# Função para copiar os professores da aba "Professores" para uma turma específica
def adicionar_professores_a_turma(planilha, turma_destino, professores_adicionados, nomes_professores):
    aba_professores = planilha['Professores']
    aba_turma = planilha[turma_destino]

    for row in aba_professores.iter_rows(min_row=2, max_row=aba_professores.max_row, min_col=1, max_col=3):
        nome = row[0].value
        cpf = row[1].value
        email = row[2].value
        professor_chave = (nome, cpf)

        # Verificar se o professor já está alocado em alguma turma
        if professor_em_turma(planilha, professor_chave):
            continue

        # Verificar se o professor já foi adicionado a alguma turma
        if nome in nomes_professores:
            # Adicionar o nome do professor à coluna D (quarta coluna) e o CPF à coluna E (quinta coluna) da turma
            aba_turma.append([None, None, None, nome, cpf])
            professores_adicionados.add(nome)

    planilha.save('Dados Cadastrais.xlsx')

# Abrir a planilha
planilha = openpyxl.load_workbook('Dados Cadastrais.xlsx')

# Conjunto para manter o controle dos nomes de professores já adicionados
professores_adicionados = set()

while True:
    print("\nOpções:")
    print("1. Adicionar professores às turmas")
    print("2. Ver professores disponíveis e não alocados")
    print("3. Sair do programa")
    
    escolha = input("Escolha a opção (1, 2 ou 3): ")
    
    if escolha == '1':
        # Listar as abas de turma disponíveis
        abas_turmas = [sheet for sheet in planilha.sheetnames if sheet.startswith('Turma ')]
        
        if not abas_turmas:
            print("Não foram encontradas abas de turma na planilha.")
        else:
            print("Turmas existentes:")
            for i, turma in enumerate(abas_turmas, start=1):
                print(f"{i}. {turma}")

            # Solicitar ao usuário qual turma deseja adicionar os professores
            turma_desejada = input("Em qual turma deseja adicionar os professores: ")
            if turma_desejada in abas_turmas:
                nomes_professores = input("Digite os nomes dos professores (separados por vírgula): ")
                nomes_professores = [nome.strip() for nome in nomes_professores.split(',')]
                adicionar_professores_a_turma(planilha, turma_desejada, professores_adicionados, nomes_professores)
                print(f"Professores adicionados com sucesso à {turma_desejada}")
            else:
                print("Turma não encontrada.")
    elif escolha == '2':
        professores_nao_alocados = set()
        
        # Verificar professores na aba "Professores" que não foram alocados a nenhuma turma
        aba_professores = planilha['Professores']
        for row in aba_professores.iter_rows(min_row=2, max_row=aba_professores.max_row, min_col=1, max_col=3):
            nome = row[0].value
            cpf = row[1].value
            email = row[2].value
            professor_chave = (nome, cpf)
            
            if not professor_em_turma(planilha, professor_chave):
                professores_nao_alocados.add(nome)
                
        if professores_nao_alocados:
            print("Professores disponíveis para alocação:")
            for nome in professores_nao_alocados:
                print(nome)
        else:
            print("Não há professores disponíveis para alocação.")
    elif escolha == '3':
        break
    else:
        print("Opção inválida. Escolha 1 para adicionar professores, 2 para verificar professores disponíveis ou 3 para sair.")

print("Programa encerrado.")
