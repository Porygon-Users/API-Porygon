import openpyxl
from openpyxl.styles import Font, Alignment  # Importe as classes Font e Alignment
from datetime import datetime, timedelta
import os

# Construa o caminho completo para o arquivo Excel no diretório 'database'
diretorio_atual = os.path.dirname(os.path.abspath(__file__))
caminho_arquivo_excel = os.path.join(diretorio_atual, '..', 'database', 'infodados.xlsx')

# Verificar se o arquivo Excel existe e criar um novo se não existir
if os.path.exists(caminho_arquivo_excel):
    planilha = openpyxl.load_workbook(caminho_arquivo_excel)
else:
    planilha = openpyxl.Workbook()
    planilha.save(caminho_arquivo_excel)

# Função para adicionar a data de início do curso e calcular o fim do curso
def adicionar_data_e_ciclos(planilha, turma_destino):
    aba_turma = planilha[turma_destino]

    data_inicio = datetime.strptime(input("\nDigite a data de início do curso (DD/MM/AAAA): "), "%d/%m/%Y")
    data_fim = datetime.strptime(input("Digite a data de término do curso (DD/MM/AAAA): "), "%d/%m/%Y")

    qtd_ciclos = int(input("\nQuantos ciclos você deseja: "))

    # Calcular a duração de cada ciclo
    duracao_ciclo = (data_fim - data_inicio) / qtd_ciclos

    ciclos = []

    for i in range(qtd_ciclos):
        ciclo_nome = f"Ciclo {i + 1}"
        ciclo_inicio = data_inicio + i * duracao_ciclo
        ciclo_fim = ciclo_inicio + duracao_ciclo - timedelta(days=1)
        ciclos.append((ciclo_nome, ciclo_inicio, ciclo_fim))

    # Adicionar cabeçalhos para as datas de início e término do curso
    aba_turma.cell(row=1, column=7).value = "Início do Curso"
    aba_turma.cell(row=1, column=8).value = "Fim do Curso"

    # Preencher as datas de início e término do curso
    aba_turma.cell(row=2, column=7).value = data_inicio.strftime('%d/%m/%Y')
    aba_turma.cell(row=2, column=8).value = data_fim.strftime('%d/%m/%Y')

    # Adicionar cabeçalhos para os dias do ciclo, início e término do ciclo
    aba_turma.cell(row=1, column=9).value = "Início do Ciclo"
    aba_turma.cell(row=1, column=10).value = "Término do Ciclo"
    aba_turma.cell(row=1, column=11).value = "Dias do Ciclo"

    # Estilizar os cabeçalhos
    cabecalho_fonte = Font(bold=True, size=11)  # Negrito e tamanho de fonte razoável
    cabecalho_alinhamento = Alignment(horizontal='center')  # Centralizar o texto

    for coluna in range(7, 12):  # Aplicar estilos para as colunas 7 a 11
        aba_turma.cell(row=1, column=coluna).font = cabecalho_fonte
        aba_turma.cell(row=1, column=coluna).alignment = cabecalho_alinhamento

    # Ajustar a largura das colunas (por exemplo, para as colunas 7 a 11)
    largura_coluna = 20  # Ajuste o valor conforme necessário
    for coluna in range(7, 12):
        aba_turma.column_dimensions[openpyxl.utils.get_column_letter(coluna)].width = largura_coluna

    # Calcular e adicionar a quantidade de dias do ciclo
    for i, (ciclo_nome, ciclo_inicio, ciclo_fim) in enumerate(ciclos):
        aba_turma.cell(row=i + 2, column=9).value = ciclo_inicio.strftime('%d/%m/%Y')
        aba_turma.cell(row=i + 2, column=10).value = ciclo_fim.strftime('%d/%m/%Y')

        # Calcular e adicionar a quantidade de dias do ciclo
        dias_ciclo = (ciclo_fim - ciclo_inicio).days + 1
        aba_turma.cell(row=i + 2, column=11).value = dias_ciclo

        # Exibir a quantidade de dias do ciclo
        print(f"\n{ciclo_nome} terá {dias_ciclo} dias.")

    # Salvar a planilha no arquivo infodados.xlsx
    caminho_arquivo_excel = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'database', 'infodados.xlsx')
    planilha.save(caminho_arquivo_excel)
    print("\n----Data de início, fim do curso e ciclos adicionados com sucesso!!----")

# Abrir a planilha
caminho_arquivo_excel = os.path.join(os.path.dirname(os.path.abspath(__file__)), '..', 'database', 'infodados.xlsx')

if os.path.exists(caminho_arquivo_excel):
    planilha = openpyxl.load_workbook(caminho_arquivo_excel)
else:
    planilha = openpyxl.Workbook()
    planilha.save(caminho_arquivo_excel)

while True:
    print("\nOpções:")
    print("\n1. Adicionar data de início do curso e criar ciclos")
    print("2. Sair do programa", "\n")

    escolha = input("Escolha uma das opções: ")

    if escolha == '1':
        abas_turmas = [sheet for sheet in planilha.sheetnames if sheet.startswith('Turma ')]

        if not abas_turmas:
            print("\nNão foram encontradas abas de turma na planilha")
        else:
            print("\nTurmas existentes:", "\n")
            for i, turma in enumerate(abas_turmas, start=1):
                print(f"{i}. {turma}")

            num_turma = int(input("\nEscolha o número da turma: "))
            if 1 <= num_turma <= len(abas_turmas):
                turma_desejada = abas_turmas[num_turma - 1]
                adicionar_data_e_ciclos(planilha, turma_desejada)
            else:
                print("\nNúmero de turma inválido.")

    elif escolha == '2':
        break
    else:
        print("Opção inválida. Escolha 1 para adicionar a data de início do curso e criar ciclos ou 2 para sair.")

print("\nPrograma encerrado.", "\n")
