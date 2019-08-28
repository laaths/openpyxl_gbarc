from teste_class import loadArq, saveArq
from teste_msg import msgs

orig = ['planilhas/copia.xlsx', 'COMPLETA']
dest = ['planilhas/lotericas_controle.xlsx', 'COMPLETA']
proc = '18-017811-3'


def procPlan(planilha, aba, proc):
    plan = loadArq(planilha, aba)
    ws = plan.loadSheet()
    tt = plan.listCols(proc, ws)
    return tt

def insertPlan(orig_plan, orig_aba, dest_plan, dest_aba, proc):
    orig_dados = procPlan(orig_plan, orig_aba, proc)
    dest_dados = procPlan(dest_plan, dest_aba, proc)
    if orig_dados is False:
        print("Não Encontrado")
        return
    elif dest_dados != False:
        print("Existente na planilha")
        return
    else:
        plan = loadArq(dest_plan, dest_aba)
        wb = plan.loadPlan()
        aba = wb[dest_aba]
        ws = wb.active
        plan.insertDados(aba, ws, orig_dados)
        wb.save(dest_plan)
        return



insertPlan(orig[0], orig[1], dest[0], dest[1], proc)

def menuOpt():
    while True:
        msgs().getMenu()
        opt = input(str("OPÇÃO DESEJADA: ").upper())
        if opt == 'Q':
            pass
        elif opt == 'A': # COPIAR E COLAR LINHA DE UMA PLANILHA PRA OUTRA
            pass
        elif opt == 'B': # PROCURAR
            pass
        elif opt == 'C': # INSERIR LOOPBACK
            pass
        elif opt == 'D':
            pass
        elif opt == 'E':
            pass
        elif opt == 'F':
            pass
        else:
            msgs().getMsg1()
