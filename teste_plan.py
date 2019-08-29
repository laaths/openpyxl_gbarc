from teste_class import loadArq, saveArq
from teste_msg import msgs

orig = ['planilhas/copia.xlsx', 'Planilha1']
dest = ['planilhas/lotericas_controle.xlsx', 'COMPLETA']
proc = 'PAE 0845598'

# Procura e retorna lista com os dados
def procPlan(planilha, aba, proc):
    plan = loadArq( planilha, aba )
    ws = plan.loadSheet()
    tt = plan.listCols( proc, ws )
    return tt


def insertPlan(orig_plan, orig_aba, dest_plan, dest_aba, proc):
    orig_dados = procPlan( orig_plan, orig_aba, proc ) # Retorna dados procurado
    dest_dados = procPlan( dest_plan, dest_aba, proc ) # Retorna dados procurado
    if orig_dados is False:
        print( "Não Encontrado" )
        return
    elif dest_dados != False:
        print( "Existente na planilha" )
        return
    else:
        plandest = loadArq( dest_plan, dest_aba )
        wbdest = plandest.loadPlan()
        planorig = loadArq( orig_plan, orig_aba )
        orig_panel = plandest.listColsPanel( planorig.loadSheet() ) # Retorna nomes da coluna do painel fixo
        dest_panel = plandest.listColsPanel( plandest.loadSheet() ) # Retorna nomes da coluna do painel fixo
        ws = wbdest.active
        dad_ext = plandest.insertDadosRow( dest_panel, orig_panel, ws, orig_dados ) # Retorna lista do conteudo ordenado pela painel de destino
        ws.append( dad_ext ) # Adiciona dados de uma lista na proxima linha vazia
        wbdest.save( dest_plan ) # Salva o conteudo na planilha
        return


def localizarDados():
    pass

# print(procPlan(dest[0], dest[1], proc))
# print(painel(orig[0], orig[1]))

# Menu de opções
def menuOpt():
    while True:
        msgs().getMenu()
        opt = input( str( "OPÇÃO DESEJADA: " ).upper() )
        if opt == 'Q':
            pass
        elif opt == 'A':  # COPIAR E COLAR LINHA DE UMA PLANILHA PRA OUTRA
            insertPlan( orig[0], orig[1], dest[0], dest[1], proc )
        elif opt == 'B':  # PROCURAR
            pass
        elif opt == 'C':  # INSERIR LOOPBACK
            pass
        elif opt == 'D':
            pass
        elif opt == 'E':
            localizarDados()
        elif opt == 'F':
            pass
        else:
            msgs().getMsg1()

menuOpt()