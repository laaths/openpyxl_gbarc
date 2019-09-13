from openpyxl import load_workbook, Workbook
from pyxl_class import loadArq, verifArq
from pyxl_msgs import msgs
import os
proc = "PAE 0836982"

plan = loadArq()

print(plan.formatCircuitos())
#print(plan.origPlan)
#print(plan.origAba)
#print(plan.destPlan)
#print(plan.destAba)
#print(plan.proc)
#plan.confProg()


# Menu de opções
def menuOpt():
    while True:
        msgs().getMenu()
        opt = input("OPÇÃO DESEJADA: ").upper()
        if opt == 'Q':
            break
        elif opt == 'A':  # COPIAR E COLAR LINHA DE UMA PLANILHA PRA OUTRA
            pass
        elif opt == 'B':  # PROCURAR
            pass
        elif opt == 'C':  # INSERIR LOOPBACK
            pass
        elif opt == 'D':  # INDEFINIDO
            pass
        elif opt == 'E':  # MOSTRA PLANILHAS UTILIZADAS
            pass
        elif opt == 'F':  # DEVELOPER
            pass
        elif opt == 'T':  # OPÇÃO PARA TESTES DE FUNÇÕES
            pass
        else:
            msgs().getMsg1()