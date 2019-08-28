from openpyxl import load_workbook, Workbook

class loadArq:
    def __init__(self, planilha, aba_sheet):
        self.plan = planilha
        self.sheet = aba_sheet
        self.alfab = '0ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    def loadPlan(self):
        return load_workbook(self.plan)

    def loadSheet(self):
        return self.loadPlan()[self.sheet]

    def activePlan(self):
        return self.loadPlan().active

    def savePlan(self):
        return self.loadPlan().save(self.plan)

    def listCols(self, proc, load):
        cols = []
        for cl in range(load.max_column ):
            for ln in range( load.max_row ):
                position = load[self.alfab[cl + 1] + str( ln + 1 )].value
                if position == proc:
                    for cl2 in range(load.max_column ):
                        cols.append(load[self.alfab[cl2+1]+str(ln+1)].value)
                    return cols
                else:
                    pass
        return False

    def listRows(self, proc, load):
        rows = []
        for ln in range(load.max_row):
            pass

    def insertDados(self, aba, ws, dados):
        x = 2
        while True:
            cel = str(x)
            if aba['G' + cel].value == None:
                for y in range(len(dados)):
                    ws[self.alfab[y+1]+cel] = dados[y]
                break
            else:
                x+=1
        return ws

class saveArq:
    def __init__(self, planilha, aba_sheet):
        self.plan = planilha
        self.sheet = aba_sheet

    def cadastramento(self, aba, ws):
        x = 2
        while True:
            cel = str( x )
            if aba['A' + cel].value == None:

                break
            else:
                x += 1
        return