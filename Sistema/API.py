import random

total_alunos = int(input("Digite o número total de alunos: "))
max_alunos_por_sala = int (input ("Digite o máximo de alunos por sala: "))


def dividir_alunos(total_alunos):
    if total_alunos == 0:
        print("O número de alunos deve ser maior que zero.")
        return

    min_professores_por_sala = 1
    max_professores_por_sala = 3

    salas = []
    alunos_restantes = total_alunos

    while alunos_restantes > 0:
        # Calcule o número de alunos para esta sala (mínimo 1, máximo 40)
        alunos_na_sala = min(max_alunos_por_sala , alunos_restantes)

        # Calcule o número de professores para esta sala (entre min e max)
        professores_na_sala = random.randint(min_professores_por_sala, max_professores_por_sala)

        # Adicione a sala à lista de salas
        sala = {
            "Alunos": alunos_na_sala,
            "Professores": professores_na_sala,
            "Grupos": []
        }
        salas.append(sala)

        # Atualize o número de alunos restantes
        alunos_restantes -= alunos_na_sala

    # Distribua os alunos em grupos
    tamanho_grupo = 10
    for sala in salas:
        alunos_na_sala = sala["Alunos"]
        grupos = [tamanho_grupo] * (alunos_na_sala // tamanho_grupo)
        alunos_restantes_na_sala = alunos_na_sala % tamanho_grupo

        for i in range(alunos_restantes_na_sala):
            grupos[i % len(grupos)] += 1

        sala["Grupos"] = grupos

    # Exiba a distribuição
    for i, sala in enumerate(salas, start=1):
        print(f"Sala {i} - Alunos: {sala['Alunos']}, Professores: {sala['Professores']}")
        for j, grupo in enumerate(sala["Grupos"], start=1):
            print(f"Grupo {j}: {grupo} alunos")

dividir_alunos(total_alunos)
