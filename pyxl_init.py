from openpyxl import load_workbook, Workbook
from pyxl_class import excell
from pyxl_msgs import msgs

def callClass(call):
    while True:
        try:
            loadclass = call
            return loadclass
        except PermissionError:
            input("\n FECHAR ARQUIVO ABERTO \n")

# Menu de opções
def menuOpt():
    while True:
        input( "\nPRESSIONE ENTER PARA INICIAR" )
        msgs().getMenu()
        opt = input("OPÇÃO DESEJADA: ").upper()
        if opt == 'Q':
            break
        elif opt == 'A':  # COPIAR E COLAR LINHA DE UMA PLANILHA PRA OUTRA
            callClass(excell().insertPlanDados())
        elif opt == 'B':  # PROCURA SIMLPES
            callClass(excell().procItemSimple(input("Digite UL/CCT/IP: ")))
        elif opt == 'C':  # INDEFINIDO
            pass
        elif opt == 'D':  # INDEFINIDO
            pass
        elif opt == 'E':  # MOSTRA PLANILHAS UTILIZADAS
            callClass(excell().confProg())
        elif opt == 'F':  # DEVELOPER
            pass
        elif opt == 'T':  # OPÇÃO PARA TESTES DE FUNÇÕES
            pass
        else:
            msgs().getMsg1()

menuOpt()