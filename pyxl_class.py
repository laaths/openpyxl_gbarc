from openpyxl import load_workbook, Workbook
import os
# Class para carregamento do arquivo de leitura

class verifArq():
    def __init__(self):
        self.origPlan = self.loadConfPlanOrig()[0]
        self.origAba = self.loadConfPlanOrig()[1]
        self.destPlan = self.loadConfPlanDest()[0]
        self.destAba = self.loadConfPlanDest()[1]
        self.proc = self.loadConfPlanProc()

    # Verificação do arquivo de texto
    def verifDirArq(self):
        dados = [
            "config/config.xlsx",
            "config1",
            "PLANILHA DE ORIGEM DE DADOS",
            "CAMINHO DA PASTA",
            "NOME DA PLANILHA XLSX",
            "ABA DA PLANILHA",
            "PLANILHA DE DESTINO DE DADOS",
            "LISTA DE CIRCUITOS PARA COPIA A PARTIR DA LINHA 10"
        ]
        try:
            wb = load_workbook( dados[0] )  # NOME DO ARQUIVO
            return wb
        except FileNotFoundError:
            try:
                os.mkdir( "config" )
            except FileExistsError:
                pass
            wb = Workbook()
            ws = wb.active
            ws.title = dados[1]
            ws['A1'] = dados[2]
            ws['A2'] = dados[3]
            ws['B2'] = dados[4]
            ws['C2'] = dados[5]
            ws['A5'] = dados[6]
            ws['A6'] = dados[3]
            ws['B6'] = dados[4]
            ws['C6'] = dados[5]
            ws['A9'] = dados[7]
            wb.save( dados[0] )
            return wb

    # CARREGA PLANILHA DE ORIGEM DOS DADOS
    def loadConfPlanOrig(self):
        wb = self.verifDirArq()
        ws = wb.active
        orig = [ws.cell( row=3, column=1 ).value + ws.cell( row=3, column=2 ).value, ws.cell( row=3, column=3 ).value]
        return orig

    # CARREGA PLANILHA DE DESTINO DOS DADOS
    def loadConfPlanDest(self):
        wb = self.verifDirArq()
        ws = wb.active
        orig = [ws.cell( row=7, column=1 ).value + ws.cell( row=7, column=2 ).value, ws.cell( row=7, column=3 ).value]
        return orig

    # LISTA DE CIRCUITOS PARA TRANSFERIR DE PLANILHAS
    def loadConfPlanProc(self):
        wb = self.verifDirArq()
        ws = wb.active
        cctlst = []
        for x in range( ws.max_row ):
            if ws.cell( row=x + 10, column=1 ).value is None:
                pass
            else:
                cctlst.append( ws.cell( row=x + 10, column=1 ).value )
        return cctlst
    # FORMATA ENTRADA DE CIRCUITOS
    def formatCircuitos(self):
        proccompl = []
        for x in range(len(self.proc)):
            if self.proc[x][0:3] == "PAE" and (self.proc[x][3] == " " or self.proc[x][3] == "_") and self.proc[x][4:11].isdigit() :
                proccompl.append(self.proc[x])
            elif len(self.proc[x]) >= 12:
                for y in range(len(self.proc[x])):
                    try:
                        if (self.proc[x][y]+self.proc[x][y+1]+self.proc[x][y+2]) == 'PAE' or (self.proc[x][y]+self.proc[x][y+1]+self.proc[x][y+2]) == 'PLT':
                            proccompl.append(self.proc[x][y:y+11])
                    except IndexError:
                        pass
            elif len(self.proc[x]) <= 11:
                proccompl.append(self.proc[x])
            else:
                pass
        self.proc = proccompl
        return self.proc

    # LEITURA DO ARQUIVO DE CONFIGURAÇÃO
    def confProg(self):
        wb = self.verifDirArq()
        ws = wb.active
        if ws.cell( row=3, column=2 ).value is None:
            print( 'CADASTRE AS PLANILHAS A SEREM UTILIZADA!' )
        else:
            print( '\n', ws.cell( row=1, column=1 ).value, '\n',
                   ws.cell( row=3, column=2 ).value,
                   '\n', ws.cell( row=2, column=1 ).value, '\n',
                   ws.cell( row=3, column=1 ).value )

            print( '\n', ws.cell( row=5, column=1 ).value, '\n',
                   ws.cell( row=7, column=2 ).value,
                   '\n', ws.cell( row=6, column=1 ).value, '\n',
                   ws.cell( row=7, column=1 ).value )
            print("\nCIRCUITOS A PROCURAR:")
            for x in range(len(self.proc)):
                print(self.proc[x])

class loadArq( verifArq ):
    def __init__(self):
        verifArq.__init__(self)
        self.plan1 = 'plan'

    # Criar lista com itens da linha
    def listCols(self, proc, load):
        cols = []
        for cl in range(load.max_column ):
            for ln in range( load.max_row ):
                position = load.cell(row=ln+1, column=cl+1).value
                if position == proc:
                    for cl2 in range(load.max_column ):
                        cols.append(load.cell(row=ln+1, column=cl2+1).value)
                    return cols
                else:
                    pass
        return False

    # Criar lista com Colunas do painel
    def listColsPanel(self, load):
        cols = []
        for cl in range(load.max_column):
            cols.append(load.cell(row=1, column=cl+1).value)
        return cols


    # Criar lista com itens da coluna
    def listRows(self, proc, load):
        rows = []
        for ln in range(load.max_row):
            pass

    # Inserir dados na planilha
    def insertDadosRow(self, dest_panel, orig_panel, ws, orig_dados):
        dados = []
        #for plan in range(ws.max_row):
        for cl in range(len(dest_panel)):
            if dest_panel[cl] in orig_panel:
                #print('Existe na planilha')
                dados.append(orig_dados[orig_panel.index(dest_panel[cl])])
            elif dest_panel[cl] not in orig_panel:
                dados.append(None)
            else:
                pass
        return dados
