import openpyxl

# Função para criar grupos em uma turma
def criar_grupos(planilha, turma_nome, num_alunos_por_grupo):
    # Carregar o arquivo da planilha
    wb = openpyxl.load_workbook(planilha)
    
    # Selecionar a aba da turma
    try:
        sheet = wb[turma_nome]
    except KeyError:
        print(f'A aba "{turma_nome}" não foi encontrada na planilha.')
        return
    
    # Obter a lista de alunos na turma
    alunos = [cell.value for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=1) for cell in row if cell.value]
    
    # Exibir a quantidade de alunos na turma
    num_alunos = len(alunos)
    print(f'A turma "{turma_nome}" possui {num_alunos} alunos.')
    
    
    # Verificar se é possível criar grupos com o número especificado de alunos
    if num_alunos % num_alunos_por_grupo != 0:
        print(f'O número de alunos ({num_alunos}) não é divisível pelo número de alunos por grupo ({num_alunos_por_grupo}).')
        return
    
    # Calcular o número total de grupos
    num_grupos = num_alunos // num_alunos_por_grupo
    
    # Criar grupos
    grupos = []
    for i in range(num_grupos):
        grupo_alunos = alunos[i * num_alunos_por_grupo: (i + 1) * num_alunos_por_grupo]
        grupos.append(f'Grupo {i + 1} - {", ".join(grupo_alunos)}')
    
    # Adicionar os grupos à coluna C da planilha
    for i, grupo in enumerate(grupos):
        sheet.cell(row=i + 2, column=5, value=grupo)
    
    # Salvar a planilha
    wb.save(planilha)
    print(f'Foram criados {num_grupos} grupos com {num_alunos_por_grupo} alunos cada na "{turma_nome}".')

# Função para contar grupos em uma turma
def contar_grupos(planilha, turma_nome):
    # Carregar o arquivo da planilha
    wb = openpyxl.load_workbook(planilha)
    
    # Selecionar a aba da turma
    try:
        sheet = wb[turma_nome]
    except KeyError:
        print(f'A aba "{turma_nome}" não foi encontrada na planilha.')
        return
    
    # Contar grupos na coluna C da turma
    num_grupos = sum(1 for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=3, max_col=3) if row[0].value)
    print(f'A turma "{turma_nome}" possui {num_grupos} grupos.')
    alunos_por_grupo = len(sheet.cell(row=2, column=3).value.split(','))  # Assumindo que todos os grupos têm o mesmo número de alunos
    print(f'Cada grupo possui {alunos_por_grupo} alunos.')

# Função para contar alunos em uma turma
def contar_alunos(planilha, turma_nome):
    # Carregar o arquivo da planilha
    wb = openpyxl.load_workbook(planilha)
    
    # Selecionar a aba da turma
    try:
        sheet = wb[turma_nome]
    except KeyError:
        print(f'A aba "{turma_nome}" não foi encontrada na planilha.')
        return
    
    # Contar alunos na turma
    num_alunos = sum(1 for row in sheet.iter_rows(min_row=2, max_row=sheet.max_row, min_col=1, max_col=1) if row[0].value)
    print(f'A turma "{turma_nome}" possui {num_alunos} alunos.')
    return num_alunos  # Retornar a quantidade de alunos na turma

# Função principal do programa
def main():
    planilha = 'Dados Cadastrais.xlsx'
    
    while True:
        print("\nOpções:")
        print("\n1. Criar grupos")
        print("2. Ver quantos grupos têm em uma turma")
        print("3. Sair", "\n")
        
        escolha = input("Escolha uma opção: ")
        
        if escolha == '1':
            turma = input('\nDigite o nome da turma (ex: Turma 1): ')
            alunos_na_turma = contar_alunos(planilha, turma)  # Mostrar a quantidade de alunos na turma
            if alunos_na_turma is not None:
                if alunos_na_turma > 0:
                    alunos_por_grupo = int(input('Digite o número de alunos por grupo: '))
                    criar_grupos(planilha, turma, alunos_por_grupo)
        elif escolha == '2':
            turma = input('Digite o nome da turma que deseja ver: ')
            contar_grupos(planilha, turma)
        elif escolha == '3':
            print("\nSaindo do programa.", "\n")
            break
        else:
            print("\nOpção inválida. Tente novamente.", "\n")

if __name__ == "__main__":
    main()