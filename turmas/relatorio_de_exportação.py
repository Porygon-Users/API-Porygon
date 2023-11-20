import os
import pandas as pd
import openpyxl

def obter_turmas(planilha_path):
    # Carrega a planilha
    xls = pd.ExcelFile(planilha_path)

    # Filtra apenas as abas que representam turmas
    turmas = [sheet_name for sheet_name in xls.sheet_names if sheet_name.lower().startswith('turma')]

    return turmas

def obter_peso_ciclo(ciclo, df_pesos):
    # Busca o peso correspondente ao ciclo no DataFrame de pesos
    try:
        peso_ciclo = df_pesos.loc[df_pesos['Ciclo'] == ciclo, 'Peso'].values[0]
        return peso_ciclo
    except IndexError:
        # Se não encontrar um peso correspondente, emite um aviso e retorna 0
        print(f"Aviso: Não foram encontrados pesos para o {ciclo} na aba 'Pesos'.")
        return 0

def gerar_relatorio_turmas(planilha_path, df_pesos):
    # Obtém as turmas disponíveis
    turmas = obter_turmas(planilha_path)

    if not turmas:
        print("Nenhuma turma encontrada na planilha.")
        return

    # Imprime as opções de turmas
    print("Turmas disponíveis:")
    for i, turma in enumerate(turmas, 1):
        print(f"{i}. {turma}")

    # Pede ao usuário para selecionar uma turma
    while True:
        try:
            escolha_turma = int(input("Escolha uma turma (digite o número): "))
            turma_selecionada = turmas[escolha_turma - 1]
            break
        except (ValueError, IndexError):
            print("Escolha inválida. Tente novamente.")

    # Cria um DataFrame com os dados da turma selecionada
    df_turma = pd.read_excel(planilha_path, sheet_name=turma_selecionada)

    # Cria um DataFrame vazio para o relatório consolidado
    df_relatorio = pd.DataFrame(columns=['Aluno', 'Início do curso', 'Término do curso', 'Pesos', 'Média'])

    # Adiciona as informações do relatório para cada aluno
    for index, row in df_turma.iterrows():
        aluno = row['Alunos']
        inicio_curso = row['Início do Curso']
        fim_curso = row['Fim do Curso']

        # Adiciona informações básicas ao DataFrame do relatório
        df_relatorio = pd.concat([df_relatorio, pd.DataFrame({
            'Aluno': [aluno],
            'Início do curso': [inicio_curso],
            'Término do curso': [fim_curso],
            'Pesos': [0],  # Inicializa com 0, será atualizado posteriormente
            'Média': [row['Média']],  # Adiciona a média
        })], ignore_index=True)

        # Itera pelos ciclos presentes na turma
        for col in df_turma.columns:
            if col.startswith('Ciclo'):
                ciclo = f"Notas Ciclo {col.split(' ')[-1]}"
                nota = row[col]

                # Verifica se o ciclo já exibiu o aviso
                if ciclo not in df_relatorio.columns:
                    # Adiciona uma nova coluna para o ciclo
                    df_relatorio[ciclo] = None

                # Adiciona a nota ao DataFrame do relatório
                df_relatorio.loc[df_relatorio['Aluno'] == aluno, ciclo] = nota

                # Atualiza o valor de 'Pesos' para o ciclo
                peso_ciclo = obter_peso_ciclo(ciclo, df_pesos)
                df_relatorio.loc[df_relatorio['Aluno'] == aluno, 'Pesos'] = peso_ciclo

    # Atualiza o valor de 'Pesos' para todos os alunos no DataFrame do relatório
    df_relatorio['Pesos'] = df_pesos['Peso'].values

    # Reorganiza as colunas para ter 'Aluno', 'Início do curso', 'Término do curso', 'Pesos', 'Média' no início
    colunas_organizadas = ['Aluno', 'Início do curso', 'Término do curso', 'Pesos', 'Média'] + sorted(df_relatorio.columns[5:])
    df_relatorio = df_relatorio[colunas_organizadas]

    # Caminho para o novo arquivo Excel (na pasta 'database')
    novo_arquivo_path = os.path.join('..', 'database', f'relatorio_de_turmas.xlsx')

    # Adiciona o DataFrame do relatório ao novo arquivo
    with pd.ExcelWriter(novo_arquivo_path, engine='openpyxl') as writer:
        df_relatorio.to_excel(writer, sheet_name='Relatorio_Turma', index=False)

    print(f"\nRelatório da turma '{turma_selecionada}' gerado em '{novo_arquivo_path}'.")

    # Pergunta se deseja selecionar outra turma
    outra_turma = input("Deseja selecionar outra turma? (s/n): ")
    if outra_turma.lower() == 's':
        gerar_relatorio_turmas(planilha_path, df_pesos)
    else:
        print("Saindo.")

if __name__ == "__main__":
    # Caminho para a planilha do Excel
    caminho_planilha = os.path.join('..', 'database', 'infodados.xlsx')

    # Caminho para a aba 'Pesos'
    caminho_pesos = os.path.join('..', 'database', 'infodados.xlsx')

    # Cria um DataFrame com os pesos dos ciclos
    df_pesos = pd.read_excel(caminho_pesos, sheet_name='Pesos')

    # Chama a função para gerar relatórios de turmas, passando o DataFrame de pesos
    gerar_relatorio_turmas(caminho_planilha, df_pesos)
