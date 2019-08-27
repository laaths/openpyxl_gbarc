from teste_class import loadArq, saveArq
from teste_msg import msgs


ent = '18-009892-6'

exc = loadArq('planilhas/copia.xlsx', 'COMPLETA')

wb = exc.loadPlan()
ws = exc.loadSheet()

tt = exc.listItem(ent, ws)

print(tt)
