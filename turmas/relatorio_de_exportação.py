import pandas as pd
import openpyxl as openpy

#Depois de ter definido todas as coisas necessárias
print("Você deseja:\n 1- Emitir um relatório\n2- Sair")
op = int(input("Insira a opção desejada: "))

while True:
    if op == 2:
        break
    elif op == 1:
        #Inserir lista de turmas disponíveis para a seleção
        #Solicitar o número da turma desejada
        turma = input("Informe a turma desejada para o relatório: ")
    else:
        print("Insira uma opção válida.")