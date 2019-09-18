from openpyxl import load_workbook, Workbook
from pyxl_class import excell, verifArq, sshConnect
from pyxl_msgs import msgs
import os


#ssh = sshConnect()
#ssh.ssh.exec_command("ssh nho-rs-ser-a01")
#ssh.exec_cmd("ssh nho-rs-ser-a01")
#ssh.exec_cmd("telnet router 5293 10.252.102.194")
#ssh.exec_cmd("tr521506")
#ssh.exec_cmd("rspoa040")
#ssh.exec_cmd("dis ip int br")



def callClass(call):
    while True:
        try:
            loadclass = call
            return loadclass
        except PermissionError:
            input("\n FECHAR ARQUIVO ABERTO \n")

# Menu de opções
def menuOpt():
    input(" PRESSIONE ENTER PARA INICIAR ")
    while True:
        msgs().getMenu()
        opt = input("OPÇÃO DESEJADA: ").upper()
        if opt == 'Q':
            break
        elif opt == 'A':  # COPIAR E COLAR LINHA DE UMA PLANILHA PRA OUTRA
            callClass(excell().insertPlanDados())
        elif opt == 'B':  # PROCURA SIMLPES
            callClass(excell()).procItemSimple(input("Digite UL/CCT/IP: "))
        elif opt == 'C':  # INSERIR LOOPBACK
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