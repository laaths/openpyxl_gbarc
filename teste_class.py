from openpyxl import load_workbook, Workbook

class loadArq:
    def __init__(self, planilha, aba_sheet):
        self.plan = planilha
        self.sheet = aba_sheet
        self.alfab = '0ABCDEFGHIJKLMNOPQRSTUVWXYZ'

    def loadPlan(self):
        wb = load_workbook(self.plan)
        return wb

    def loadSheet(self):
        ws = self.loadPlan()[self.sheet]
        return ws

    def listItem(self, procura, load):
        lista = []
        for cl in range(load.max_column ):
            for ln in range( load.max_row ):
                position = load[self.alfab[cl + 1] + str( ln + 1 )].value
                if position == procura:
                    for cl2 in range(load.max_column ):
                        lista.append(load[self.alfab[cl2+1]+str(ln+1)].value)
                    return lista
                else:
                    pass


class saveArq:
    def __init__(self, planilha, aba_sheet):
        self.plan = planilha
        self.sheet = aba_sheet